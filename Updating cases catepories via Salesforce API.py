# This code will automatically update cases' categories field in your Salesforce system.
# The code is using the model that was built previously and it's inserting the model's classification to the Salesforce relevant field via API.

# To review the model development, please visit: https://github.com/aloncohen1/My-Projects/blob/master/NLP%20Project.ipynb

# Made by: Alon Cohen
# Bigalon1990@gmail.com

from simple_salesforce import Salesforce
import requests
import base64
import json
import salesforce_reporting
import pandas as pd
import pickle
import re

# Connect to Salesforce API
sf_production = Salesforce(username="xxx@xxxxxx.com"
                           , password="xxxxxxxx"
                           , security_token="xxxxxxxxx", )

# Connect to Salesforce API - REPORTS
reports_sf = salesforce_reporting.Connection(username="xxx@xxxxxx.com"
                                             , password="xxxxxxxx"
                                             , security_token="xxxxxxxxx", )
report = reports_sf.get_report(
    'Ener the report ID', )  ### Create a report of all the uncataloged cases and enter his ID here

parser = salesforce_reporting.ReportParser(report)

# Extract the report of the uncataloged cases and transform it into a Pandas data frame
cases_to_catalog = pd.DataFrame(
    columns=['Case Number', 'Case ID', 'Subject', 'Description', 'Case Comments', 'Topic - for PC'],
    index=range(len(parser.records())))
for i in range(len(parser.records())):
    cases_to_catalog['Case Number'][i] = parser.records()[i][0]
    cases_to_catalog['Case ID'][i] = parser.records()[i][1]
    cases_to_catalog['Subject'][i] = parser.records()[i][2]
    cases_to_catalog['Description'][i] = parser.records()[i][3]
    cases_to_catalog['Case Comments'][i] = parser.records()[i][4]

cases_to_catalog = cases_to_catalog.rename(index=str, columns={"Subject": "Title"})


# Load the trained models
def load_obj(name):
    with open(name + '.pkl', 'rb') as f:
        return pickle.load(f)


pipeline_lr_linking = load_obj('pipeline_lr_linking')
pipeline_lr_others = load_obj('pipeline_lr_others')

# Create dictionary of cases
cases_to_catalog_dict = dict()
count = 0
for i in cases_to_catalog['Case ID']:
    if i not in cases_to_catalog_dict.keys():
        cases_to_catalog_dict[(cases_to_catalog['Case ID'][count])] = cases_to_catalog['Topic - for PC'][count]
    count += 1
len(cases_to_catalog_dict)


# Create functions that clean the text
def my_function(raw):
    raw = raw.lower()
    raw = raw.replace(']', '')
    raw = raw.replace('[', '')
    raw = raw.replace(')', '')
    raw = raw.replace('(', '')
    raw = raw.replace(':', '')
    raw = raw.replace('.', '')
    raw = raw.replace(',', '')
    raw = raw.replace('  ', ' ')
    raw = raw.replace('"', '')
    raw = raw.replace('\n', ' ')
    raw = raw.replace('\t', ' ')
    raw = raw.replace('?', '')
    raw = re.sub(r"http\S+", "", raw)
    raw = re.sub('https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+', "", raw)
    raw = re.sub(" \d{13} ", " isbn ", raw)
    raw = re.sub(" \d{10} ", " isbn ", raw)
    raw = re.sub(" \d{7}\d{1}[\dx] ", " issn ", raw)
    raw = re.sub(" \d{4}[-]\d{3}[\dx] ", " issn ", raw)
    raw = re.sub(" 10\.\S+ ", " doi ", raw)
    raw = re.sub('<.*>', "", raw)
    raw = re.sub('\S+@\S+', "email", raw)
    raw = re.sub('[0-9]+', "", raw)
    raw = re.sub(r'(\d+/\d+/\d+)', "date", raw)
    for i in raw.split():
        if len(i) > 22:
            raw = raw.replace(i, '')
    raw = raw.replace('$', '')
    raw = raw.replace('!', '')
    raw = raw.replace("'", '')
    raw = raw.replace("->", '')
    raw = raw.replace('&', '')
    raw = raw.replace('/', '')
    raw = raw.replace('%', ' ')
    raw = raw.replace(' - ', ' ')
    raw = raw.replace('+', '')
    raw = raw.replace('_', '')
    raw = raw.replace('@', '')
    raw = raw.replace('--', '')
    raw = raw.replace('#', '')
    raw = raw.replace('=', '')
    raw = raw.replace('â', '')
    raw = raw.replace('*', '')
    raw = raw.replace('-', '')
    raw = raw.replace(';', '')
    raw = raw.replace('<', '')
    raw = raw.replace('>', '')
    raw = raw.replace('ß', '')
    ' '.join(raw.split())
    raw = raw.replace('  ', ' ')
    raw = raw.replace('  ', ' ')

    return raw


# Create data frame that flattens the information (from several comments to one long string)
merged_cases_to_catalog = pd.DataFrame(columns=['Case ID', 'Mixed_Comments', 'Categorie'], )
cases_by_id = cases_to_catalog.groupby('Case ID')

counter = 0
for i in cases_to_catalog_dict.keys():
    case = cases_by_id.get_group(i)

    mix_comment = ''

    for x in case['Title']:
        title = x
    for y in case['Description']:
        description = y

    mix_title = my_function(str(title)) + ' ' + my_function(str(description))
    mix_comment += mix_title
    for comment in case['Case Comments']:
        if len(str(comment).split()) > 4:
            mix_comment += (' ' + my_function(str(comment)))

    for t in case['Topic - for PC']:
        topic = t
    merged_cases_to_catalog.loc[counter] = [i, mix_comment, topic]
    counter += 1

# Condition - the update will occur only if there are cases to update
if len(merged_cases_to_catalog) > 0:

    # Predict the cases topics using the trained models
    first_prediction_for_catalog = pipeline_lr_linking.predict(merged_cases_to_catalog['Mixed_Comments'])
    second_prediction_for_catalog = pipeline_lr_others.predict(merged_cases_to_catalog['Mixed_Comments'])

    # Create a file that aggregate the data and the prediction
    final_prediction = pd.DataFrame(columns=['Case ID', 'Mixed_Comments', 'Linking/Not Linking', 'General Prediction',
                                             'Final Predicted Categorie'], )
    final_prediction['Case ID'] = merged_cases_to_catalog['Case ID']
    final_prediction['Mixed_Comments'] = merged_cases_to_catalog['Mixed_Comments']
    final_prediction['Linking/Not Linking'] = first_prediction_for_catalog
    final_prediction['General Prediction'] = second_prediction_for_catalog
    final_prediction['Final Predicted Categorie'] = second_prediction_for_catalog

    # Cataloging to "Data linking" or "Other" - if "Other", then the second prediction will be counted
    for i in range(len(final_prediction)):
        if final_prediction['Linking/Not Linking'][i] == 'Data linking':
            final_prediction['Final Predicted Categorie'][i] = 'Data linking'

    # Create a finale data frame for the update
    update_file = pd.DataFrame(columns=['Topic - for PC'], index=final_prediction['Case ID'])
    update_file['Topic - for PC'] = list(final_prediction['Final Predicted Categorie'])

    # Update the cases topics using Salesforce API
    counter = 0
    for i in update_file.index:
        if sf_production.Case.get(i)['bl_New_Category__c'] == None:
            sf_production.Case.update(i, {'bl_New_Category__c': 'General'})
        sf_production.Case.update(i, {'Topic_for_PC__c': update_file['Topic - for PC'][i]})
        counter += 1
print(str(counter) + ' ' + 'Cases has been updated')
