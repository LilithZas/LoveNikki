import requests
from bs4 import BeautifulSoup
import sys
from openpyxl import Workbook
import pandas as pd
import xlrd

def lnData1(sType, sID):
    ItemType = {
        'Hair' : 'H',
        'Dress': 'D',
        'Coat': 'C',
        'Top': 'T',
        'Bottom': 'B',
        'Hosiery': 'P',
        'Shoes': 'S',
        'Makeup': 'M',
        'Accessory': 'A',
        'Soul': 'X',
    }
    
    itemTop1 = []
    
    
    # Obtaining item information
    url = 'https://ln.nikkis.info/wardrobe/' + sType + '/' + str(sID)
    response = requests.get(url)
    soup = BeautifulSoup(response.text,'lxml')
    
    # Obtaining item name
    itemName = soup.find('h4', class_='header pink-text text-lighten-2').text
    itemName = itemName[:-12]
    
    # Obtain item ID
    itemInfo = soup.find('strong').text
    itemID = [int(itemInfo) for itemInfo in str.split(itemInfo) if itemInfo.isdigit()]
    itemTag = str.split(itemInfo)
    itemTag = ItemType[itemTag[0]]
    
    # Obtain top scoring info
    itemTop = soup.find_all('div', class_='collapsible-header')
    for i in itemTop:
        itemTop1.append(i.text[15:])
        
    # Obtain rarity
    rarity = soup.find('span', class_='grey-text')
    rarity = rarity.text
    
    
    # Obtaining source
    itemSource = soup.find_all('h5', class_='item-section-head')
    itemSourceList = []
    source = 0
    for i in itemSource:
        itemSourceList.append(i.text)
    if any('Customization' in s for s in itemSourceList):
        source = 'Customization'
    if any('Evolution' in s for s in itemSourceList):
        source = 'Evolution'
    if any('Crafted from' in s for s in itemSourceList):
        source = 'Crafted'
    if any('Obtained from' in s for s in itemSourceList):
        source = soup.find('li', class_='collection-item')
        source = source.text
    if (source == 0):
        source = 'Special Event'
    
    # Obtaining percentage own 
    url = 'https://my.nikkis.info/stats/own/ln/clothes/' + itemTag + str(itemID[0])
    response = requests.get(url)
    soup = BeautifulSoup(response.text,'lxml')
    percentageOwn = soup.text + '%'
    

        
    return itemName,int(rarity),sType,itemID,percentageOwn,itemTop1,source

def dataToDf(data):
    topScoring = data[5]
    topScoring = [w.replace('\xa0',' ') for w in topScoring]
    topInfo = [0,0,0]
    count = 0
    for i in topScoring:
        if ('Chapter' in i):
            tempTop = topScoring[count]
            tempTop = [int(s) for s in str.split(tempTop) if s.isdigit()]
            topInfo[0] = tempTop[0]
        if ('Commission' in i):
            tempTop = topScoring[count]
            tempTop = [int(s) for s in str.split(tempTop) if s.isdigit()]
            topInfo[1] = tempTop[0]
        if ('Stylist' in i):
            tempTop = topScoring[count]
            tempTop = [int(s) for s in str.split(tempTop) if s.isdigit()]
            topInfo[2] = tempTop[0]
        count += 1
    data1 = []
    for i in data[0:5]:
        data1.append(i)
    for i in topInfo:
        data1.append(i)
    for i in data[6:]:
        data1.append(i)
    return data1

def obtainTopAddress():
    url = 'https://ln.nikkis.info/top/'
    response = requests.get(url)
    soup = BeautifulSoup(response.text,'lxml')
    links = []
    link = soup.find_all('a', class_='witem collection-item avatar icon-room col s12 m6 l6')
    for i in link:
        links.append(i.get('href'))
    links2 = []
    for i in links:
        links2.append(i.split('/')[2:])
    return links2
  
output = []
address = obtainTopAddress()
for i in address:
    try:
        data = lnData1(i[0], i[1])
        data = dataToDf(data)
        output.append(data)
    # reader = pd.read_excel(r'output.xlsx')
    # df = pd.DataFrame(output)
    # df.to_excel('output.xlsx',mode='a',index=False,header=False, startrow=len(reader)+1)
        print('Finished processing item ID: ' + str(i))
    except:
        print('Item ID ' + str(i) + ' not found')
columnName = ['Name','Rarity','Type','ID','Ownership','Top Chapter','Top Commission','Top Arena','Obtained by']
df = pd.DataFrame(output,columns=columnName)
df.to_excel("output.xlsx")
# with open('output.csv', 'a') as f:
#     df.to_csv(f, header=False)





