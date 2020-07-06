import requests
import json
import time
from time import mktime
from jsonpath_ng import parse
from datetime import datetime
from openpyxl import load_workbook


wb = load_workbook('Analytics.xlsx')
#ws = wb.active
repos = ["tensorflow/tensorflow","pytorch/pytorch","opencv/opencv","explosion/SpaCy", "ant-design/ant-design"]
default_hdr = {'Authorization': 'Token f28c09ab095ac30084cb1796c0d49f39ac74c49a'}
commit_hdr = {'Authorization': 'Token f28c09ab095ac30084cb1796c0d49f39ac74c49a','Accept':'application/vnd.github.cloak-preview'}
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
        global ws
        ws = wb['Data-PR-Issues-Commits']
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
                    print("Error getting data for " + url + " " + str(response.status_code))
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
    response = requests.get(url,headers=default_hdr)
    rate_json = json.loads(response.text)
    if response.status_code == 200:
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
    else:
        print(response.text)


def write_contributors():
    global repos
    global wb
    ws = wb['Contributors-data']
    #global wb
    row_index=2
 
    for repo in repos :
        url = "https://api.github.com/repos/"+repo+"/stats/contributors"
        response = requests.get(url,headers=default_hdr)
        
        if (response.status_code == 200):
            conjson = json.loads(response.text)
            #with open('cont.json') as json_file:
            #    conjson = json.load(json_file)
            json_exp = parse('$[*]')
            additions_tot=0
            deletions_tot=0
            commits_tot=0
            lists= [match.value for match in json_exp.find(conjson)]
            for l in lists:
                if (type(l) == type(dict())):
                    for (k, v) in l.items():
                        if (k=="author"):
                            for a,a2 in v.items():
                                if (a=="login"):
                                    print(a2)
                                    print(additions_tot)
                                    print(deletions_tot)
                                    print(commits_tot)
                                    c1 = ws.cell(row=row_index, column=1)
                                    c1.value = repo.split('/')[0]
                                    c2 = ws.cell(row=row_index, column=2)
                                    c2.value = repo.split('/')[1]
                                    cc = ws.cell(row_index, column=3)
                                    cc.value = a2
                                    c4 = ws.cell(row_index, column=4)
                                    c4.value = additions_tot
                                    c5 = ws.cell(row_index, column=5)
                                    c5.value = deletions_tot
                                    c6 = ws.cell(row_index, column=6)
                                    c6.value = commits_tot
                                    
                                    row_index = row_index+1
                                    additions_tot=0
                                    deletions_tot=0
                                    commits_tot=0
                        elif (k == "weeks"):
                            for v1 in v:
                                additions_tot=additions_tot+v1["a"]
                                deletions_tot=deletions_tot+v1["d"]
                                commits_tot=commits_tot+v1["c"]
        else:
            print("Error getting data for contributors" + url + " "+response.text)    
    
    
    
    
    

def test_contributors():
    global repos
    #global wb
    #row_index=2
 
    #for repo in repos :
    url = "https://api.github.com/repos/"+"tensorflow/tensorflow"+"/stats/contributors"
    response = requests.get(url,headers=default_hdr)
    
    if (response.status_code == 200):
        #conjson = json.loads(response.text)
        with open('cont.json') as json_file:
            conjson = json.load(json_file)
        json_exp = parse('$[*]')
        additions_tot=0
        deletions_tot=0
        commits_tot=0
        lists= [match.value for match in json_exp.find(conjson)]
        for l in lists:
            if (type(l) == type(dict())):
                for (k, v) in l.items():
                    if (k=="author"):
                        for a,a2 in v.items():
                            if (a=="login"):
                                print(a2)
                                print(additions_tot)
                                print(deletions_tot)
                                print(commits_tot)
                                additions_tot=0
                                deletions_tot=0
                                commits_tot=0
                    elif (k == "weeks"):
                        for v1 in v:
                            print(v1["w"])
                            additions_tot=additions_tot+v1["a"]
                            deletions_tot=deletions_tot+v1["d"]
                            commits_tot=commits_tot+v1["c"]
                            #print("a "+str(v1["a"]))
                            #print(v1["d"])
                            #print(v1["c"])
    else:
        print("Error getting data for contributors" + url + " "+response.text)

def write_commit_activity():
    global repos
    global wb
    start_row_index=2
    ws = wb['Commit-Activity']
 
    for repo in repos :
        url = "https://api.github.com/repos/"+repo+"/stats/commit_activity"
        response = requests.get(url,headers=default_hdr)
        
        if (response.status_code == 200):
            conjson = json.loads(response.text)
            row_index=start_row_index
            week_exp = parse('$.[*].week')
            for match in week_exp.find(conjson):
                c1 = ws.cell(row=row_index, column=1)
                c1.value = repo.split('/')[0]
                c2 = ws.cell(row=row_index, column=2)
                c2.value = repo.split('/')[1]
                week_date = time.localtime(match.value)
                dt_string = time.strftime("%m/%d/%Y", week_date)
                print(f'{match.value}')
                c3 = ws.cell(row_index, column=3)
                c3.value=dt_string
                c4 = ws.cell(row_index, column=4)
                c4.value = time.strftime("%Y", week_date)
                row_index = row_index+1
            
            row_index=start_row_index    
            total_exp = parse('$.[*].total')
            for match in total_exp.find(conjson):
                print(f'{match.value}')
                c5 = ws.cell(row_index, column=5)
                c5.value=f'{match.value}'
                row_index=row_index+1
                
            start_row_index=row_index

        else:
            print ("Commit_activity Call Failed " + response.text)
            
def write_code_frequency():
    global repos
    global wb
    row_index=2
 
    for repo in repos :
        url = "https://api.github.com/repos/"+repo+"/stats/code_frequency"
        response = requests.get(url,headers=default_hdr)
        
        if (response.status_code == 200):
            conjson = json.loads(response.text)
            ws = wb['Code-Frequency']
            json_exp = parse('$.[*]')
            #for match in json_exp.find(conjson):
            week_lists= [match.value for match in json_exp.find(conjson)]
            for week_list in week_lists:
                c1 = ws.cell(row=row_index, column=1)
                c1.value = repo.split('/')[0]
                c2 = ws.cell(row=row_index, column=2)
                c2.value = repo.split('/')[1]
                c3 = ws.cell(row_index, column=3)
                week_date = time.localtime(week_list[0])
                dt_string = time.strftime("%m/%d/%Y", week_date)
                print (repo + " " + dt_string)
                c3.value = dt_string
                c4 = ws.cell(row_index, column=4)
                c4.value = week_list[1]
                c5 = ws.cell(row_index, column=5)
                c5.value = week_list[2]
                c6 = ws.cell(row_index, column=6)
                c6.value = time.strftime("%Y", week_date)
                row_index = row_index+1
        else:
            print ("Code Frequency Call Failled " + response.text)

                 
#rate_limit(True)
#write_headers()
#issues()
#pr()
#commits()
#write_code_frequency()
#write_commit_activity()
write_contributors()
#test_contributors()
wb.save("Analytics.xlsx")

