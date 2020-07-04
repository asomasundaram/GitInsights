import requests
import json
import time
from time import mktime
from jsonpath_ng import parse
from datetime import datetime
from openpyxl import load_workbook


wb = load_workbook('Analytics.xlsx')
#ws = wb['Data-PR-Issues-Commits']
ws = wb.active
repos = ["tensorflow/tensorflow","pytorch/pytorch","opencv/opencv","explosion/SpaCy", "ant-design/ant-design"]
default_hdr = {'Authorization': 'Token 3d8f0c2b3d3e0d069cba6767fa9a13285fa2599b'}
commit_hdr = {'Authorization': 'Token 3d8f0c2b3d3e0d069cba6767fa9a13285fa2599b','Accept':'application/vnd.github.cloak-preview'}
def issues():
    global repos
    last_row_index=2
    for repo in repos:
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " created:", last_row_index, 4, default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:closed created:", last_row_index, 5, default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:open created:", last_row_index, 6, default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:closed closed:", last_row_index, 7,default_hdr)
        last_row_index=last_row_index+6

        
def pr():
    global repos
    last_row_index=2
    for repo in repos:
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr is:closed created:", last_row_index, 9,default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr linked:issue created:", last_row_index,  10,default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr is:merged created:", last_row_index, 11,default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr interactions:0 created:", last_row_index, 12,default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr interactions:1..10 created:", last_row_index, 13,default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr interactions:11..20 created:",last_row_index,  14,default_hdr)
        write_to_excel("https://api.github.com/search/issues?q=" + repo, " is:pr interactions:21..* created:", last_row_index, 15,default_hdr)
        last_row_index=last_row_index+6

def commits():
    global repos
    last_row_index=2
    for repo in repos:
        global row_index
        write_to_excel("https://api.github.com/search/commits?q=" + repo, " merge:false author-date:", last_row_index, 16, commit_hdr)
        write_to_excel("https://api.github.com/search/commits?q=" + repo, " merge:true author-date:", last_row_index, 17,commit_hdr)
        write_to_excel("https://api.github.com/search/commits?q=" + repo, " author-date:", last_row_index, 18,commit_hdr)
        last_row_index=last_row_index+6

def write_to_excel(repo, item_type, row_index, column_index, hdr):
        MAX_YEAR=2021
        year = 2015
        while year < MAX_YEAR:
                url =  repo + item_type ;
                url = url + str(year) + "-01-01.." + str(year) + "-12-31"
                print(url)
                rate_limit()
                response = requests.get(url,headers=hdr)

                if response.status_code == 200:
                    issues = json.loads(response.content.decode('utf-8'))
                    
                    global ws
                    #c1 = ws.cell(row=row_index, column=1)
                    #c1.value = repo.split('/')[0]
                    #c2 = ws.cell(row=row_index, column=2)
                    #c2.value = repo.split('/')[1]
                    #c3 = ws.cell(row=row_index, column=3)
                    #c3.value=year
                    cc = ws.cell(row_index, column_index)
                    cc.value = issues['total_count']
                    row_index = row_index+1
                else:
                    print("Error getting data " + str(response.status_code))
                    return 
                
                year=year+1

def write_headers():
    global repos
    global ws
    row_index=2
    for repo in repos:
        MAX_YEAR=2021
        year = 2015
        while year < MAX_YEAR:
                c1 = ws.cell(row=row_index, column=1)
                c1.value = repo.split('/')[0]
                c2 = ws.cell(row=row_index, column=2)
                c2.value = repo.split('/')[1]
                c3 = ws.cell(row=row_index, column=3)
                c3.value=year
                row_index = row_index+1
                
                year=year+1

def rate_limit(dbg=False):
    url = "https://api.github.com/rate_limit"
    response = requests.get(url,headers={'Authorization': 'Token 3d8f0c2b3d3e0d069cba6767fa9a13285fa2599b'})
    rate_json = json.loads(response.text)
    json_expression=parse('$.resources.search.reset')
    match=json_expression.find(rate_json)   
    reset_time = time.localtime(match[0].value)
    dt_string = time.strftime("%m/%d/%Y %H:%M:%S", reset_time)
    now = datetime.now()
        
    json_expression_remaining=parse('$.resources.search.remaining')
    match=json_expression_remaining.find(rate_json) 
    if dbg:
        print(rate_json)
        print("reset date and time =", dt_string)
        print("now =", now)
        print('Remaining' + str(match[0].value))

    if (match[0].value == 1):
        dt = datetime.fromtimestamp(mktime(reset_time))
        print ('Sleeping ' + str((dt-now).total_seconds()))
        time.sleep((dt-now).total_seconds()+1)
        
rate_limit(True)
#write_headers()
#issues()
pr()
#commits()
wb.save("Analytics.xlsx")
