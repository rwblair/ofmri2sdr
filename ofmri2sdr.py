import csv
import json
from pprint import pprint

from openpyxl import Workbook

api_file = './ofmri.json'
out_file = './ofmri.xlsx'

api_fp = open(api_file)
api_json = json.load(api_fp)

column_count = 0
max_investigators_len = 0 
max_publications_len = 0
max_tasks_len = 0
max_links_len = 0

for dataset in api_json:
    if len(dataset['investigator_set']) > max_investigators_len:
        max_investigators_len = len(dataset['investigator_set'])

    if len(dataset['publicationpubmedlink_set']) > max_publications_len:
        max_publications_len = len(dataset['publicationpubmedlink_set'])
    
    if len(dataset['task_set']) > max_tasks_len:
        max_tasks_len = len(dataset['task_set'])
   
    if len(dataset['link_set']) > max_links_len:
        max_links_len = len(dataset['link_set']) 

columns = ['Title']
for i in range(max_investigators_len):
    columns.append('Author')
    columns.append('Role')
columns.append('Date')
columns.append('Abstract')
columns.append('Preferred Citation')
for i in range(max_publications_len):
    columns.append('Link Title to Related Content')
    columns.append('Link to Related Content')
columns.append('Contact Email')
for i in range(max_tasks_len):
    columns.append('Keyword')
for i in range(max_links_len):
    columns.append('Filename')
    columns.append('File Description')

wb = Workbook(write_only=True)
ws = wb.create_sheet()

ws.append(columns)

for dataset in api_json:
    output = []
    output.append(dataset['project_name'])

    investigator_count = 0
    for investigator in dataset['investigator_set']:
        output.append(investigator['investigator'])
        output.append("Principal investigator")
        investigator_count += 1

    for i in range(max_investigators_len - investigator_count):
        output.append('')
        output.append('')

    # date
    earliest = None
    for revision in dataset['revision_set']:
        if not earliest and revision['date_set']:
            earliest = revision['date_set']
        elif revision['date_set'] < earliest:
            earliest = revision['date_set']
        else:
            pass
    output.append('earliest')

    output.append(dataset['summary'])

    # citation
    output.append('')
    
    publication_count = 0
    for pub in dataset['publicationpubmedlink_set']:
        output.append(pub['title'])
        output.append(pub['url'])
        publication_count += 1
    
    for i in range(max_publications_len - publication_count):
        output.append('')
        output.append('')

    #email
    output.append('submissions@openfmri.org')
    
    task_count = 0
    for task in dataset['task_set']:
        output.append(task['name'])
        task_count += 1

    for i in range(max_tasks_len - task_count):
        output.append('')

    link_count = 0
    for link in dataset['link_set']:
        filename = link['title']
        rev = link['revision']
        if rev:
            filename += ': ' + rev
        output.append(filename)
        output.append(link['url'])
        link_count += 1
    
    for i in range(max_links_len - link_count):
        output.append('')
        output.append('')

    ws.append(output)
    #print(len(output))

wb.save(out_file)
