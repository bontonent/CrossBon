# Library
import pandas as pd
from tqdm import tqdm
import numpy as np
import copy
#copy in clipboard
import pyperclip
#pyperclip.copy(mes[1])

# convert from excel to data
df_jd = pd.read_excel("data_JD.xlsx")
articles_excel = pd.read_excel("need_articles.xlsx")
articles = articles_excel["Номер 2"] # main key

# create convenient format data
mesive = []
data = np.array(df_jd)
for work in data:
    mesive.append(list(work))
work_mesive = copy.deepcopy(mesive[:])

# Create key
keyboard = {}
keyboard_db = {}
resould = {}
for article in articles:
    keyboard[article] = [article]
    keyboard_db[article] = []
    resould[article] = []

# create check if don't have copy we add element
def add_in_keyboard_right(key, value):
    error = 0
    for element in keyboard[key]:
        if element == value:
            error = 1
    if error == 0:
        keyboard[key].append(value)
def add_in_keyboard_left(key, value):
    error = 0
    for element in keyboard[key]:
        if element == value:
            error = 1
    if error == 0:
        keyboard[key].insert(0,value)

temps = []
temp = []
def search_left(key, value):
    temps = []
    temps.append(value)
    while 0 == 0:
        temp = []
        for tem  in temps:
            #print(tem)
            work_value = tem
            for i, mes in enumerate(mesive):
                if mes[0] == work_value:
                    add_in_keyboard_left(key,mes[1])
                    temp.append(mes[1])
        if temp == []:
            break;
        else:
            temps = temp
def search_right(key, value):
    temps = []
    temps.append(value)
    while 0 == 0:
        temp = []
        for tem  in temps:
            #print(tem)
            work_value = tem
            for i, mes in enumerate(mesive):
                if mes[1] == work_value:
                    add_in_keyboard_left(key,mes[0])
        if temp == []:
            break;
        else:
            temps = temp

print("Choose necessary data")
for i, mes in enumerate(tqdm(mesive)):
    for art in articles:
        if mes[0] == art:
            add_in_keyboard_left(mes[0],mes[1])
            search_left(mes[0],mes[1])
            #print("Need add up ",mes)
        if mes[1] == art:
            add_in_keyboard_right(mes[1],mes[0])
            search_right(mes[1],mes[0])
            #print("Need add low ",mes)

# for DB
def add_in_keyboard_db(key, value):
    error = 0
    for element in keyboard_db[key]:
        if element == value:
            error = 1
    if error == 0:
        keyboard_db[key].append(value)

# keyboard_db element use for add data from our DB
df_have = pd.read_excel("We_have.xlsx")
for i, data_need in enumerate(df_have.iloc[:,0]):
    element_dd = data_need.replace("DD ","")
    #print(element_dd)
    error = 0
    for article in articles:
        if article == element_dd:
            error = 1
            #print(element_dd)
            break
    if error == 1:
        #print(element_dd," ", (df_have.iloc[i,2]))
        add_in_keyboard_db(element_dd,df_have.iloc[i,2])


for article in tqdm(articles):
    #print(i," ",article)
    #print(keyboard[article]," ",keyboard_db[article])
    for keybo in keyboard[article]:
        error = 0
        for keybo_db in keyboard_db[article]:
            if keybo == keybo_db:
                error = 1
        if error == 0:
            resould[article].append(keybo)

place_A = []
place_FH = []
for i, article in enumerate(articles):
    if len(resould[article]) != 0:
        # print(article," ",resould[article])
        for resou_art in resould[article]:
            place_A.append(article)
            place_FH.append(resou_art)

df = pd.DataFrame(np.nan,index = range(len(place_A)),columns = range(8),dtype=object)
#df.iloc[:,0] = np.array(place_A)

for i,place in enumerate(place_A):
    df.loc[i,0] = "".join(("DD ",str(place)))
for i,place in enumerate(place_FH):
    df.loc[i,5] = str(place)
    df.loc[i,7] = str(place)
df.loc[:,3] = "OE"
df.loc[:,6] = "JOHN DEERE"


#print(df)
df.to_excel("DATA.xlsx",index = False,header =False)




