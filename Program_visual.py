import PyQt6
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import *
from PyQt6.QtGui import QEnterEvent, QIcon
import sys
from tqdm import tqdm
import pandas as pd
import numpy as np

def main_call(text_code, text, get_sts):
    text_code = text_code
    if text_code == None:
        get_art = False
    else:
        get_art = True
    # read files, if don't have file. Will say error
    print("Try read excel files")
    try:
        df = pd.read_excel("Article_need.xlsx")
    except:
        print("Need Article_need.xlsx")
        return
    try:
        df = pd.read_excel("Cross_have.xlsx")
    except:
        print("Need Cross_have.xlsx")
        return
    try:
        df = pd.read_excel("DATA_JD.xlsx")
    except:
        print("Need DATA_JD.xlsx")
        return

    # -------------------------- not delete under -----------------------------
    # You want get data from Article code (DD code) get code

    # --------------------- will be response button -----------------|
    # convert from excel to data
    df_jd = pd.read_excel("DATA_JD.xlsx", dtype=str, header=None)
    articles_excel = pd.read_excel("Article_need.xlsx")
    articles = articles_excel["Номер"]  # main key

    ### create convenient format data
    mesive = []
    data = np.array(df_jd)
    for work in data:
        mesive.append(list(work))

    # Create key
    keyboard = {}
    keyboard_db = {}
    resould = {}
    ## next version add button where add element
    ## keyboard[article] = [article] without [DD ]
    if get_art == True:
        for i, article in enumerate(articles):
            if get_sts == 1:
                if text_code == 0:
                    print("work 0")
                    keyboard[article] = [article[article.find(" ") + 1:]]
                    # print(article, article[article.find(" ")+1:])
                elif text_code == 1:  # replace left one
                    # print(article[article.find(" ")+1:])
                    el_art = (article[article.find(" ") + 1:].find(text) + len(text))
                    arg1 = str(article[article.find(" ") + 1:][:el_art].replace(text, ""))
                    arg2 = str(article[article.find(" ") + 1:][el_art:])
                    # print(arg1,arg2)
                    pull = "".join((arg1, arg2))

                    keyboard[article] = [pull]
                    # print(article, [article[article.find(" ")+1:].replace(str(text),"")])
                elif text_code == 2:  # replace all
                    keyboard[article] = [article[article.find(" ") + 1:].replace(str(text), "")]
                    # print(article, [article[article.find(" ")+1:].replace(str(text),"")])
                elif text_code == 3:  # replace right one
                    cut_article = article[article.find(" ") + 1:]
                    cut_art_re = cut_article[::-1]
                    text_re = text[::-1]

                    el_art = (cut_art_re.find(text_re) + len(text_re))
                    arg1 = str(cut_art_re[:el_art].replace(text_re, ""))
                    arg2 = str(cut_art_re[el_art:])
                    pull = "".join((arg1, arg2))[::-1]
                    keyboard[article] = [pull]
                    # print(article, [pull])
                else:
                    keyboard[article] = []
            else:
                keyboard[article] = []
            keyboard_db[article] = []
            resould[article] = []
    else:
        for article in articles:
            print("work none")
            keyboard[article] = []
            keyboard_db[article] = []
            resould[article] = []

    ### create convenient format data
    mesive = []
    data = np.array(df_jd)
    for work in data:
        mesive.append(list(work))

    # for DB
    def add_in_keyboard_db(key, value):
        error = 0
        for element in keyboard_db[key]:
            if element == value:
                error = 1
        if error == 0:
            keyboard_db[key].append(value)

    # keyboard_db element use for add data from our DB
    df_have = pd.read_excel("Cross_have.xlsx")
    ion = 0
    for i, data_need in enumerate(df_have.loc[:, "Код товара"]):
        element_dd = data_need
        # print(element_dd)
        error = 0
        for article in articles:
            if article == element_dd:
                error = 1
                # print(element_dd)
                break
        if error == 1:
            for ion, column in enumerate(df_have.columns):
                if column == "Номер перекрестной ссылки":
                    code = ion
                    # print(code)
            add_in_keyboard_db(element_dd, df_have.iloc[i, code])

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
            keyboard[key].insert(0, value)
    # Cycle for search all data and add all data in under key elements
    # {key: [here]}
    # Cycle for search all data and add all data in under key elements
    # {key: [here]}temps = []
    # Cycle for search all data and add all data in under key elements
    # {key: [here]}temps = []
    temp = []

    # Cycle for search all data and add all data in under key elements
    # {key: [here]}temps = []
    temp = []

    def search_left(key, value):
        temps = []
        temps.append(value)
        add_in_keyboard_left(key, value)
        while 0 == 0:
            temp = []
            for tem in temps:
                # print(tem)
                work_value = tem
                for i, mes in enumerate(mesive):
                    if mes[0] == work_value:
                        add_in_keyboard_left(key, mes[1])
                        temp.append(mes[1])
            if temp == []:
                break;
            else:
                temps = temp

    def search_right(key, value):
        temps = []
        temps.append(value)
        add_in_keyboard_right(key, value)
        while 0 == 0:
            temp = []
            for tem in temps:
                # print(tem)
                work_value = tem
                for i, mes in enumerate(mesive):
                    if mes[1] == work_value:
                        add_in_keyboard_left(key, mes[0])
                        temp.append(mes[0])
            if temp == []:
                break;
            else:
                temps = temp

    # call all def for two side
    print("Choose necessary data")
    for i, mes in enumerate(tqdm(mesive)):
        for art in articles:
            if get_art == True:
                if text_code == 0:
                    art_element_work = art[art.find(" ") + 1:]
                    if mes[0] == art_element_work:
                        search_left(art, mes[1])
                    elif mes[1] == art_element_work:
                        search_right(art, mes[0])
                elif text_code == 1:
                    el_art = (art[art.find(" ") + 1:].find(text) + len(text))
                    arg1 = str(art[art.find(" ") + 1:][:el_art].replace(text, ""))
                    arg2 = str(art[art.find(" ") + 1:][el_art:])
                    # print(arg1,arg2)
                    pull = "".join((arg1, arg2))

                    art_element_work = pull
                    if mes[0] == art_element_work:
                        search_left(art, mes[1])
                    elif mes[1] == art_element_work:
                        search_right(art, mes[0])

                elif text_code == 2:
                    art_element_work = art[art.find(" ") + 1:].replace(str(text), "")
                    if mes[0] == art_element_work:
                        search_left(art, mes[1])
                    elif mes[1] == art_element_work:
                        search_right(art, mes[0])
                elif text_code == 3:
                    cut_art = art[art.find(" ") + 1:]
                    cut_art_re = cut_art[::-1]
                    text_re = text[::-1]

                    ele_art = (cut_art_re.find(text_re) + len(text_re))
                    arg1 = str(cut_art_re[:ele_art].replace(text_re, ""))
                    arg2 = str(cut_art_re[ele_art:])
                    pull = "".join((arg1, arg2))[::-1]
                    atr_element_work = pull
                    if mes[0] == atr_element_work:
                        search_left(art, mes[1])
                    elif mes[1] == atr_element_work:
                        search_right(art, mes[0])
            for value_db in keyboard_db[art]:
                value_db = value_db
                if mes[0] == value_db:
                    search_left(art, mes[1])
                    continue
                elif mes[1] == value_db:
                    search_right(art, mes[0])
                    continue

    # work with our Data
    for article in tqdm(articles):
        for keybo in keyboard[article]:
            error = 0
            for keybo_db in keyboard_db[article]:
                if keybo == keybo_db:
                    error = 1
            if error == 0:
                resould[article].append(keybo)

    # Start work with view file xlsx
    place_A = []
    place_FH = []
    # place_G = []
    for i, article in enumerate(articles):
        if len(resould[article]) != 0:
            # print(article," ",resould[article])
            for resou_art in resould[article]:
                place_A.append(article)
                place_FH.append(resou_art)
                # place_G.append(key_brand[article])

    df = pd.DataFrame(np.nan, index=range(len(place_A)), columns=range(8), dtype=object)
    # df.iloc[:,0] = np.array(place_A)

    for i, place in enumerate(place_A):
        df.loc[i, 0] = str(place)
    for i, place in enumerate(place_FH):
        df.loc[i, 5] = str(place)
        df.loc[i, 7] = str(place)
    df.loc[:, 3] = "OE"
    df.loc[:, 6] = "JOHN DEERE"

    # print(df)
    df.to_excel("ANSWER.xlsx", index=False, header=False)

    from openpyxl import load_workbook

    # covert file in data
    workbook = load_workbook("ANSWER.xlsx")
    worksheet = workbook.active  # active sheet(in jupyter notebook break xlsx file(don't let change it is file)

    # make size width for my data
    worksheet.column_dimensions['A'].width = 20
    worksheet.column_dimensions['D'].width = 5
    worksheet.column_dimensions['F'].width = 15
    worksheet.column_dimensions['G'].width = 12
    worksheet.column_dimensions['H'].width = 15
    # Save
    workbook.save("ANSWER.xlsx")
    print(df)

class MyWindow(QMainWindow):
    def __init__(self):

        self.activate_get = 0
        #main
        super().__init__()
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        self.setGeometry(200, 200, 100, 200)
        self.setWindowTitle("CrossBon")

        layout = QGridLayout()
        # describe first editor
        self.label1 = QLabel()
        self.label1.setText("Delete text from Article code")
        self.label1.setFixedHeight(12)

        #entry for first number photo
        self.edit_place = QLineEdit()
        self.edit_place.setPlaceholderText("Enter...")

        self.labelb_1 = QLabel()
        self.labelb_1.setText("Use code Articles")
        self.labelb_1.setFixedHeight(12)
        self.switch_box_1 = QCheckBox()
        self.switch_box_1.stateChanged.connect(self.Clicke_box_1)

        self.labelb_2 = QLabel()
        self.labelb_2.setText("Add article same to same")
        self.labelb_2.setFixedHeight(12)
        self.switch_box_2 = QCheckBox()
        self.switch_box_2.stateChanged.connect(self.Clicke_box_2)

        #self.Clicked_on_1(switch_box_1.checkState())

        self.labelb1 = QLabel()
        self.labelb1.setText("First element")
        self.labelb1.setFixedHeight(12)
        self.switch_box1 = QCheckBox()
        self.switch_box1.stateChanged.connect(self.Clicke_box1)
        self.labelb2 = QLabel()
        self.labelb2.setText("All same elements")
        self.labelb2.setFixedHeight(12)
        self.switch_box2 = QCheckBox()
        self.switch_box2.stateChanged.connect(self.Clicke_box2)
        self.labelb3 = QLabel()
        self.labelb3.setText("Finally element")
        self.labelb3.setFixedHeight(12)
        self.switch_box3 = QCheckBox()
        self.switch_box3.stateChanged.connect(self.Clicke_box3)

        #for button 1
        button1 = QPushButton("Start")
        button1.clicked.connect(self.start_app)
        # for button 2
        # All element in place
        layout.addWidget(self.label1,0,2)
        layout.addWidget(self.edit_place, 1, 1)
        layout.addWidget(self.labelb1, 2, 1)
        layout.addWidget(self.switch_box1, 2, 2)
        layout.addWidget(self.labelb2, 3, 1)
        layout.addWidget(self.switch_box2, 3, 2)
        layout.addWidget(self.labelb3, 4, 1)
        layout.addWidget(self.switch_box3, 4, 2)
        layout.addWidget(self.labelb_1, 5, 1)
        layout.addWidget(self.switch_box_1,5,0)
        layout.addWidget(self.labelb_2, 5, 3)
        layout.addWidget(self.switch_box_2, 5, 4)
        layout.addWidget(button1,3,3)
        central_widget.setLayout(layout)
    def Clicke_box_1(self,status):
        if status == 2:
            self.activate_get = 0
        elif status == 0:
            self.switch_box1.setChecked(False)
            self.switch_box2.setChecked(False)
            self.switch_box3.setChecked(False)
            self.switch_box_2.setChecked(False)
            self.activate_get = -1

    def Clicke_box_2(self, status):
        self.switch_box_1.setChecked(True)

    def Clicke_box1(self,status):
        #if str(self.switch_box2.checkState()) == "CheckState.Checked":
        self.switch_box2.setChecked(False)
        #if str(self.switch_box3.checkState()) == "CheckState.Checked":
        self.switch_box3.setChecked(False)
        #if str(self.switch_box_1.checkState()) == "CheckState.Unchecked":
        self.switch_box_1.setChecked(True)
        #print(self.switch_box1.checkState())
        self.activate_get = 1

    def Clicke_box2(self,status):
        #if str(self.switch_box1.checkState()) == "CheckState.Checked":
        self.switch_box1.setChecked(False)
        #if str(self.switch_box3.checkState()) == "CheckState.Checked":
        self.switch_box3.setChecked(False)
        #if str(self.switch_box_1.checkState()) == "CheckState.Unchecked":
        self.switch_box_1.setChecked(True)

    def Clicke_box3(self, status):
        #if str(self.switch_box1.checkState()) == "CheckState.Checked":
        self.switch_box1.setChecked(False)
        #if str(self.switch_box2.checkState()) == "CheckState.Checked":
        self.switch_box2.setChecked(False)
        #if str(self.switch_box_1.checkState()) == "CheckState.Unchecked":
        self.switch_box_1.setChecked(True)

    def start_app(self):
        check_el_2 = 0
        print("------------------------------")
        print(self.switch_box_1.checkState())
        print(self.switch_box_2.checkState())
        print(self.switch_box1.checkState())
        print(self.switch_box2.checkState())
        print(self.switch_box3.checkState())
        print("------------------------------")

        if str(self.switch_box_2.checkState()) == "CheckState.Checked":
            check_el_2 = 1
        elif str(self.switch_box_2.checkState()) == "UncheckState.Checked":
            check_el_2 = 0

        if str(self.switch_box1.checkState()) == "CheckState.Checked":
            main_call(1, str(self.edit_place.text()),check_el_2)
        elif str(self.switch_box2.checkState()) == "CheckState.Checked":
            main_call(2, str(self.edit_place.text()),check_el_2)
        elif str(self.switch_box3.checkState()) == "CheckState.Checked":
            main_call(3, str(self.edit_place.text()),check_el_2)
        elif str(self.switch_box_1.checkState()) == "CheckState.Checked":
            main_call(int(0), None, check_el_2)
        else:
            main_call(None, None, None)

# Create and run the application
app = QApplication(sys.argv)
win = MyWindow()
win.show()
sys.exit(app.exec())
