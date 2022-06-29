## Important!!!
## Everytime you run the script clear everything from the google drive (even trash) and Certificate_Final_URL google sheet
## Clear the google sheet output page as well
## Change the api token as well

# Import all the libraries

import pandas as pd
import numpy as np
import os.path
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.cm import ScalarMappable
from matplotlib.lines import Line2D
from mpl_toolkits.axes_grid1.inset_locator import inset_axes
from textwrap import wrap
from PIL import Image, ImageFont, ImageDraw 
import gspread 
from google.oauth2 import service_account
import json
import requests
from Google import Create_Service
import openpyxl

# Load our data set which is the google sheet

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'gsheet.json'

credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


gc = gspread.authorize(credentials)

worksheet = gc.open_by_key('1NNGRq0cYgk_1YS98Tc4pdkTDD-4dFQ8bz0T2fDciUo0').sheet1
rows = worksheet.get_all_values()

# Convert google sheet data to data frame

df = pd.DataFrame(rows[1:], columns=rows[0])
df[df==""] = np.NaN
df.fillna(method="ffill",inplace=True)

cols=list(df.columns)
not_used=[]
for i in range(0,3):
    not_used.append(cols[i])

for i in range(5,11):
    not_used.append(cols[i])

print("These columns are of no use: ",not_used)

start=0
end=0
for j in range(0,len(cols)):
    if (cols[j]=='Ques1'):
        start=j
        
    elif (cols[j]=='Ques10'):
        end=j
        

action=[]

counts = df.pivot_table(index=['Name'], aggfunc='size')
counts = pd.DataFrame(counts) 
counts.index.name = 'Name'
counts.reset_index(inplace=True) 
counts.columns = ['Name','Counts']
df = df.merge(counts)


def remove_duplicate(test_list):
    res = []
    for i in test_list:
        if i not in res:
            res.append(i)
    return res

def frequency(l):
    max = 0
    res = l[0]
    for i in l:
        freq = l.count(i)
        if freq > max:
            max = freq
            res = i
    return res

# Get a list of all names
name=[]
name_list=[]

for i in range(0,len(df)):
    name.append(df.loc[i,"Name"])
    name_list.append(df.loc[i,"Name"])
    
name=remove_duplicate(name)

program_list=df['orgid'].tolist()
action_list=df['Type of action'].tolist()

def program_check(name):
    program_final=[]
    for i in range(0,len(name_list)):
        if (name_list[i]==name):
             program_final.append(program_list[i])

    if (len(program_final)>1):
        final_program=frequency(program_final)
        return final_program
    else:
        return program_final[0]    

program_final=[]

for i in range(0,len(name)):
    program_final.append(program_check(name[i]))



def action_check(name):
    action=[]
    for i in range(0,len(df)):
        if (df.loc[i,'Name']==name):
             action.append(action_list[i])

    if (len(action)>1):
        final_action=frequency(action)
        return final_action
    else:
        return action[0]

action_final=[]

for i in range(0,len(name)):
    action_final.append(action_check(name[i]))

def isNaN(string):
    return string != string    

#Certificate Generation function
def certificate_generate(name_list,name,program,final_action,start,end):
    
    index=name_list.index(name)
    chart_list=[]

    # Generate the skill map based on the question answered in the whatsapp chatbot

    ##Entrepreneurship -    Q8
    ##Data Orientation -    Q1
    ##Hands-on skills -     Q2
    ##Citizenship -         Q9
    ##Critical Thinking -   Q6
    ##Problem Solving -     Q4
    ##Communication Colab - Q3
    ##Grit -                Q7
    ##Applied Empathy -     Q10
    ##Communication -       Q5
    
    
    for j in range(start,end+1):
        l=[]
        if (df.columns[j]=="Ques1"):
            l.append("Data Orientation")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)
        
        elif (df.columns[j]=="Ques2"):
            l.append("HandsOn")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)

        elif (df.columns[j]=="Ques3"):
            l.append("Communication Colab")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)

             
        elif (df.columns[j]=="Ques4"):
            l.append("Problem Solving")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)
    
        elif (df.columns[j]=="Ques5"):
            l.append("Communication")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)

        elif (df.columns[j]=="Ques6"):
            l.append("Critical Thinking")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)

        elif (df.columns[j]=="Ques7"):
            l.append("Grit")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)

        elif (df.columns[j]=="Ques8"):
            l.append("Entrepreneurship")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)
             
        elif (df.columns[j]=="Ques9"):
            l.append("Citizenship")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)
            
        elif (df.columns[j]=="Ques10"):
            l.append("Applied Empathy")
            if (isNaN(df.loc[index,df.columns[j]])==True):
                l.append(0)
            elif ((int(df.loc[index,df.columns[j]])==6)):
                l.append(0)
            else:
                l.append(int(df.loc[index,df.columns[j]])*500)
            chart_list.append(l)
    
    
    # Generate the skill map using a circular bar plot
    
    chart_data = pd.DataFrame(chart_list, columns = ['Skill', 'Score'])

    chart_data = chart_data.sort_values("Score", ascending=False)

    ANGLES = np.linspace(0.05, 2 * np.pi - 0.05, len(chart_data), endpoint=False)

    LENGTHS = chart_data["Score"].values

    REGION = chart_data["Skill"].values

    
    plt.rcParams.update({"font.family": "Arial"})

    plt.rcParams["text.color"] = 'black'

    plt.rc("axes", unicode_minus=False)

    COLORS = ["#54A8A9","#FF5733","#C81B1B"]

    cmap = mpl.colors.LinearSegmentedColormap.from_list("my color", COLORS, N=256)

    fig, ax = plt.subplots(figsize=(9, 12.6), subplot_kw={"projection": "polar"})

    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    ax.set_theta_offset(1.2 * np.pi / 2)
    ax.set_ylim(0,3500)

    ax.bar(ANGLES, LENGTHS, color=COLORS, alpha=0.9, width=0.52, zorder=10)

    stepsize=250

    start, end = ax.get_ylim()
    ax.set_yticks(np.arange(start, end, stepsize))
    ax.set_xticks(ANGLES)
    ax.set_xticklabels(REGION, size=12);
    ax.set_yticklabels([]);

    XTICKS = ax.xaxis.get_major_ticks()
    for tick in XTICKS:
        tick.set_pad(40)

    ax.xaxis.grid(True,color='#154360')
    ax.yaxis.grid(True,color='#154360')


    plt.savefig('certificate_graph.png')

    ##Congratulations! This is to certify that ______
    ##was part of the ____ program. The skills activated is as below 
    ##You are a ___. You are at engaged level.
   
    char=""
    # Generate the character to be displayed in the certificate
    if (final_action=="Report"):
        im2 = Image.open("ReportingRhino.png")
        char="Reporting Rhino"
        
    if (final_action=="Hands On"):
        im2 = Image.open("HandsonHippo.png")
        char="Handson Hippo"
        
    if (final_action=="Tech Build"):
        im2 = Image.open("TechnoTiger.png")
        char="Techno Tiger"
        
    if (final_action=="Action"):
        im2 = Image.open("ActionAnt.png")
        char="Action Ant"
        
    if (final_action=="Campaign"):
        im2 = Image.open("CampaignChameleon.png")
        char="Campaign Chameleon"
        
    if (final_action=="Build"):
        im2 = Image.open("BuilderBear.png")
        char="Builder Bear"
        
    else:
        im2 = Image.open("CuriousCat.png")
        char="Curious Cat"

    empty_img = Image.open("blank_certificate.png")
     
    # Text on the certificate

    text1 = "Congratulations! This is to certify that "

    text2 = name+" has achieved the following skills"

    if (program!=''):
        text3 = "for the "+program+" Program."  

    text4 = "You are a "+char+". You are at Level 1"
    

    font = ImageFont.truetype('/fonts/Arial/arialbd.ttf', size=55)

    font1 = ImageFont.truetype('/fonts/Arial/arialbd.ttf', size=20)

    image_editable = ImageDraw.Draw(empty_img)
    image_editable.text((375,500), text1, (255, 255, 255), font=font)
    image_editable.text((375,575), text2, (255, 255, 255), font=font)

    if (program!=''):
        image_editable.text((375,650), text3, (255, 255, 255), font=font)
        
    image_editable.text((1050,2100), text4, (0, 0, 0), font=font1)
    
    empty_img.save("blank_certificate_result.png")


    im1 = Image.open("blank_certificate_result.png")

    im3 = Image.open("certificate_graph.png")

    im2 = im2.resize((400,400))
    im3 = im3.resize((500,700))

    empty_img = im1.copy()
    empty_img.paste(im2, (1100, 1600))
    empty_img.paste(im3, (400, 1475))
    empty_img.save(name+"_Certificate.png")

    return name+"_Certificate.png"    
    
# Generate the certificate for everyone 

cert_list=[]

for items in range(0,len(name)):
    cert_list.append(certificate_generate(name_list,name[items],program_final[items],action_final[items],start,end))
    


# upload certificate on google drive

headers = {"Authorization": "Bearer ya29.a0ARrdaM_kLP6laP6P6WLNAaNwD9dNFd5NdkOPqkwcSBf_JyM-JwUoRiKsl92XHmSKAio5Owwj1JItElBPbLnm5h57sgsnJ0zyPBPe6Mh7TzckJ1R6BVgDWzjLH5VQt0NBF9FcDEm7NPy60wAGhwHQyUIRV3YB"}

for i in range(0,len(cert_list)):
    para = {
        "name": cert_list[i],
        "parents":['1le9lrYxth7N_b1reAgjHZYrdHFypq-_p'],
    }
    files = {
        'data': ('metadata', json.dumps(para), 'application/json; charset=UTF-8'),
        'file': open("./"+cert_list[i], "rb")
    }
    r = requests.post(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        headers=headers,
        files=files
    )

print("Uploaded on google drive !!")



CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# Update Sharing Setting
file_id = '1le9lrYxth7N_b1reAgjHZYrdHFypq-_p'
query = f"parents = '{file_id}'"

response = service.files().list(q=query).execute()
files = response.get('files')
nextPageToken = response.get('nextPageToken')


while nextPageToken:
    response = service.files().list(q=query).execute()
    files = response.get('files')
    nextPageToken = response.get('nextPageToken')

ans=[]

for i in range(1,len(name)+1):
    ans.append(name[-i])
 
drive_data = pd.DataFrame(files)

drive_data.insert(loc = 0,
          column = 'Names',
          value = ans)


# upload all g drive image url to a new google sheet
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

spreadsheet_id = '1NNGRq0cYgk_1YS98Tc4pdkTDD-4dFQ8bz0T2fDciUo0'

response_date = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        valueInputOption='RAW',
        range='Output!A1:E1',
        body=dict(
            majorDimension='ROWS',
            values=drive_data.T.reset_index().T.values.tolist())
    ).execute()

print("Updated Google Sheet with Drive URLS!!")
