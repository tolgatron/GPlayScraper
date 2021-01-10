from bs4 import BeautifulSoup
import requests
import os
import random
import datetime
from Google import Create_Service

## Spreadsheets API Set-Up
CLIENT_SECRET_FILE = 'credentials.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
sheet_id = '1BsGXaS2tbC-6jXu0QyW2qV0hbaeYtB0KSplRedCvzeQ'
my_sheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
print('Spreadsheet API set!')

## Create New Tab with the Date
tab_name = f'Top 100 Chart {datetime.datetime.now().strftime("%d/%m/%y")}'
body = {'requests': [{'addSheet': {'properties': {'title': tab_name,'tabColor': {'red': random.uniform(0,1),'green': random.uniform(0,1),'blue': random.uniform(0,1)}}}}]}
service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,body=body).execute()
print(f'Spreadsheet {tab_name} created!')

worksheet_name = f'{tab_name}!'
cell_range_insert = 'A1'
titles=(
    ('Game Rank', 'Game Name', 'Publisher', 'Promo Text', 'Store Link', 'Star Rating', 'Download Count', 'Ratings Count', 'Genre', 'Package Size', 'Version', 'Last Update Date'),
)
value_range_body={
    'majorDimension' : 'ROWS',
    'values' : titles
}
service.spreadsheets().values().update(
    spreadsheetId = sheet_id,
    valueInputOption = 'USER_ENTERED',
    range=worksheet_name + cell_range_insert,
    body=value_range_body
).execute()    
print('Headers created in spreadsheet!')

# request_body = {
#     'requests': [
#         {
#             'repeatCell':{
#                 'range':{
#                     'sheetId' :sheet_id,
#                     'startRowIndex' : 0,
#                     'endRowIndex' : 1
#                 },
#                 'cell':{
#                     'userEnteredFormat':{
#                         'textFormat':{
#                             'fontSize':14,
#                             'bold': True
#                         }
#                     }
#                 },
#             'fields': 'userEnteredFormat(horizontalAlignment, textFormat)'
#             }
#         }
#     ]
# }

# service.spreadsheets().batchUpdate(
#     spreadsheetId=sheet_id,
#     body=request_body
# ).execute()

## Code for Scraper

## Set Lists
nameList =[]
rankintList =[]
publisherList=[]
fTextList=[]
linkList=[]
rankList =[]
downloadsList=[]
rCountList=[]
genreList=[]
psizeList=[]
vrsnList=[]
lastupList=[]
print('Lists created!')

## Set Up Scraper
googlePlay_text = requests.get('https://play.google.com/store/apps/collection/cluster?clp=0g4cChoKFHRvcHNlbGxpbmdfZnJlZV9HQU1FEAcYAw%3D%3D:S:ANO1ljJ_Y5U&gsr=Ch_SDhwKGgoUdG9wc2VsbGluZ19mcmVlX0dBTUUQBxgD:S:ANO1ljL4b8c').text
gPlay_soup = BeautifulSoup(googlePlay_text, 'lxml')
gPlay_games = gPlay_soup.find_all('div', class_ = 'ImZGtf mpg5gc')
rankint = 0
print('Scraper Set!')

## Scrape
for index, game in enumerate(gPlay_games):
    # Variables
    rankint+=1
    rankintList.append(rankint)
    lastup =''
    psize = ''
    downloads = ''
    vrsn= ''
    # Find Name
    name = game.find('div', class_ ='WsMG1c nnK0zc').text
    nameList.append(name)
    # Find Publisher
    publisher = game.find('div', class_ = 'b8cIId ReQCgd KoLSrc').text
    publisherList.append(publisher)
    # Promo Text
    fText = game.find('div', class_ = 'b8cIId f5NCO').text
    fTextList.append(fText)
    # Store Link
    link_container = game.find('div', class_ = 'b8cIId f5NCO')
    link=link_container.find('a', href = True).attrs['href']
    linkList.append(link)
    # Star Rating
    rank_container = game.find('div', class_ = 'pf5lIe')
    rank = rank_container.find('div').attrs['aria-label']
    rankList.append(rank)
    # Get SubLink
    sub_link = requests.get(f"https://play.google.com{link}").text
    sub_soup = BeautifulSoup(sub_link, 'lxml')
    # Genre
    genre = sub_soup.find_all('span', class_ = 'T32cc UAO9ie')[1].text
    genreList.append(genre)
    # Ratings Count
    rCount = sub_soup.find('span', class_ = 'AYi5wd TBRnV').text
    rCountList.append(rCount)
    # Find Sub infos
    add_info = sub_soup.find_all('div', class_ = 'JHTxhe IQ1z0d')
    for infos in add_info:
        valuelar = sub_soup.find_all('span', class_ = 'htlgb')
        ozellik=[]
        for valeu in valuelar:   
            ozellik.append(valeu.text)
        lastup = ozellik[0]
        lastupList.append(lastup)
        psize = ozellik[2]
        psizeList.append(psize)
        downloads = ozellik[4]
        downloadsList.append(downloads)
        vrsn = ozellik[6]
        vrsnList.append(vrsn)
    with open(f'posts/{index}.txt','a', encoding='utf-8') as f:
        f.write(f'''
            ======== {name} ============
            Rank :              {rankint}
            Name :              {name}
            Publisher :         {publisher}
            Promo Text :        {fText}
            Store Link :        https://play.google.com{link}
            Star Rating :       {rank}
            Downloads :         {downloads}
            Ratings Count :     {rCount}
            Genre :             {genre}
            Size :              {psize}
            Current Version :   {vrsn}
            Last Update :       {lastup}
        ''')
    with open('posts/master.txt','a', encoding='utf-8') as f:
        f.write(f'''
            ======== {name} ============
            Rank :              {rankint}
            Name :              {name}
            Publisher :         {publisher}
            Promo Text :        {fText}
            Store Link :        https://play.google.com{link}
            Star Rating :       {rank}
            Downloads :         {downloads}
            Ratings Count :     {rCount}
            Genre :             {genre}
            Size :              {psize}
            Current Version :   {vrsn}
            Last Update :       {lastup}
            ''')
    print(f'File saved: {index}.txt!') 
    print('Also saved to master.txt!')
print('Done with scraping!')

# Rank Integer
rank_resource  ={
    'majorDimension':'COLUMNS',
    'values': tuple(rankintList)
}
rank_range = 'A2'
service.spreadsheets().values().update(
    spreadsheetId = sheet_id,
    valueInputOption ='USER_ENTERED',
    range = worksheet_name + rank_range,
    body = rank_resource
).execute()

# Game Name
name_resource={
    'majorDimension':'COLUMNS',
    'values': tuple(nameList)
}
name_range = 'B2'
service.spreadsheets().values().update(
    spreadsheetId = sheet_id,
    valueInputOption = 'USER_ENTERED',
    range = worksheet_name + name_range,
    body = name_resource  
).execute()

