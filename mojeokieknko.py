import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from asyncio import events
import testExcel

# from cv2 import exp
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
import openpyxl
import os
import json
class Hit:
    def __init__(self, timeStamp, deviceType,deviceOs,deviceApp,deviceplayer,_type,status):
        self._timeStamp = timeStamp 
        self._deviceType = deviceType
        self._deviceOs = deviceOs
        self._deviceApp = deviceApp
        self._deviceplayer = deviceplayer
        self._type = _type
        self._status = status

    def printSelf(self):
        print(self._timeStamp,self._deviceType,self._deviceOs,self._deviceApp,self._deviceplayer,self._type,self._status)


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.geometry("300x200")
        self.title('Kibana Webscrapper')
        self.resizable(0, 0)

        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=7)
        self.data={}
        self.settings={}
        self.currentPath = os.getcwd()
        self.homedir = os.path.expanduser("~")
        self.appdataPath = self.homedir+'\\Appdata\\Local'
        if not (os.path.exists(self.homedir+'\\Appdata\\Local\\AutomatyzacjaHitow')):
            print("powinno sie zrobic")
            os.chdir(self.appdataPath)
            os.mkdir('AutomatyzacjaHitow')
        os.chdir(self.appdataPath+'\\AutomatyzacjaHitow')

        try:
            with open('settings.json','r+') as f:
                data = f.read()
                self.settings=json.loads(data)
                os.chdir(self.currentPath)
        except:
            os.chdir(self.currentPath)
            self.settings={"chromePath":self.currentPath,
            "outputPath":self.currentPath,
            "idNumber":123,
            "name":"Maciej Dabkowski",
            "platform":"android"}
        


        #C:\Users\gumko\AppData\Roaming
        self.create_widgets()

    def selectChromePath(self):
        pass

    def selectOutPutPath(self):
        pass

    def openConfigurationPath(self):
        pass

    def generateExcel(self):
        pass
    def saveSettings(self):
        print("saving")
        chromePath  =    self.ChromePathText.cget("text")
        outPutPath  =    self.OutputPathText.cget("text")
        idNumber    =    self.IdText.get()
        name        =    self.ImieEntry.get()
        platform    =    self.DeviceTypeEntry.get()

        self.settings={"chromePath":chromePath,
                    "outputPath":outPutPath,
                    "idNumber":idNumber,
                    "name":name,
                    "platform":platform}

        self.currentPath = os.getcwd()
        self.homedir = os.path.expanduser("~")
        self.appdataPath = self.homedir+'\\Appdata\\Local'
        if not (os.path.exists(self.homedir+'\\Appdata\\Local\\AutomatyzacjaHitow')):
            print("powinno sie zrobic")
            os.chdir(self.appdataPath)
            os.mkdir('AutomatyzacjaHitow')
        os.chdir(self.appdataPath+'\\AutomatyzacjaHitow')
       
        try:
            with open('settings.json','w+') as f:
                jsonObj = json.dumps(self.settings)
                f.write(jsonObj)
        except:
            os.chdir(self.currentPath)

    def runKibanaTest(self):
        print("kibana")
        while(True):
            self.saveSettings()
            driver = None
            print(os.getcwd())
            try:
                DRIVER_PATH = self.settings['chromePath']+ '\\' +'chromedriver' #dopisać wczytywanie 
                print(DRIVER_PATH)
                driver = webdriver.Chrome(DRIVER_PATH)
            except:
                messagebox.showerror("Błąd", "Nie odnaleziono webdrivera")
                break
            _id = self.settings['idNumber']
            print("ID",_id)
            try:
                driver.get("http://kibana.kantei-bd.redefine.pl/app/kibana#/discover?_g=(refreshInterval:(pause:!f,value:10000),time:(from:now-1h,mode:quick,to:now))&_a=(columns:!(data.userAgentData.deviceType,data.userAgentData.os,data.userAgentData.application,data.userAgentData.player,type,data.status),filters:!(('$state':(store:appState),meta:(alias:!n,disabled:!f,index:af768b90-c5ac-11ea-bebc-19fef6ef9198,key:object.id,negate:!f,params:(query:'[{}]',type:phrase),type:phrase,value:'{}'),query:(match:(object.id:(query:'{}',type:phrase))))),index:be38bb90-c5ab-11ea-bebc-19fef6ef9198,interval:auto,query:(language:lucene,query:''),sort:!(emissionDate,desc))".format(_id,_id,_id))
                flagIsFound = False
            except:
                messagebox.showerror("Błąd", "Nie udało się odnaleźć strony aplikacji")
                break
            for i in range(35):
                print("time",i)
                time.sleep(1)
                try:
                    lenTemp = len(driver.find_elements_by_xpath(".//td[@class='kbnDocTableCell__dataField eui-textBreakAll eui-textBreakWord']"))
                    if lenTemp>0:
                        flagIsFound = True
                        break
                except:
                    pass
            if(flagIsFound==False):
                messagebox.showerror("Błąd", "Nie udało się wczytać informacji ze strony")
                break
            times = driver.find_elements(by=By.XPATH,value = "//td[@class='eui-textNoWrap']")
            values = driver.find_elements_by_xpath(".//td[@class='kbnDocTableCell__dataField eui-textBreakAll eui-textBreakWord']")
            collumns = driver.find_elements_by_xpath(".//span[@i18n-id='common.ui.docTable.tableHeader.timeHeaderCellTitle']")
            listOfHits = []
            for item in range(0,len(values),6):
                tempItem = Hit(times[int(item/6)].text,values[item+0].text,values[item+1].text,values[item+2].text,values[item+3].text,values[item+4].text,values[item+5].text)
                listOfHits.append(tempItem)
            for item in listOfHits:
                item.printSelf()
            #handaling excel

            testExcel.makeExcel(self.settings['outputPath']+"\\test1",listOfHits)
            break
        print("koniec")


    def create_widgets(self):
        # username
        self.ChromePathBtn = ttk.Button(self, text="ChromePath:", command=self.selectChromePath)
        self.ChromePathBtn.grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)

        self.ChromePathText = ttk.Label(self,text=self.settings["chromePath"])
        self.ChromePathText.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)

        # output Path
        self.OutputPathBtn = ttk.Button(self, text="OutputPathBtn:", command=self.selectOutPutPath)
        self.OutputPathBtn.grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)

        self.OutputPathText = ttk.Label(self,text=self.settings["outputPath"])
        self.OutputPathText.grid(column=1, row=1, sticky=tk.E, padx=5, pady=5)

        # Id
        IdLabel = ttk.Label(self, text="ID:")
        IdLabel.grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)

        self.IdText = ttk.Entry(self)
        self.IdText.grid(column=1, row=2, sticky=tk.E, padx=5, pady=5)
        self.IdText.insert(tk.END,self.settings["idNumber"])
        
        self.ImieLabel = ttk.Label(self, text="Imie i Nazwisko:")
        self.ImieLabel.grid(column=0, row=3, sticky=tk.W, padx=5, pady=5)

        self.ImieEntry = ttk.Entry(self)
        self.ImieEntry.grid(column=1, row=3, sticky=tk.E, padx=5, pady=5)
        self.ImieEntry.insert(tk.END,self.settings["name"])
        
        
        self.DeviceTypeLabel = ttk.Label(self, text="Platforma")
        self.DeviceTypeLabel.grid(column=0, row=4, sticky=tk.W, padx=5, pady=5)

        self.DeviceTypeEntry = ttk.Entry(self)
        self.DeviceTypeEntry.grid(column=1, row=4, sticky=tk.E, padx=5, pady=5)
        self.DeviceTypeEntry.insert(tk.END,self.settings["platform"])

        # Generate
        self.generateBtn = ttk.Button(self, text="Generate", command=self.runKibanaTest)
        self.generateBtn.grid(column=0, row=5, columnspan=2,sticky="WSEN", padx=5, pady=5)

        #Output
        self.Output = ttk.Label(self,text="Wynik")
        self.Output.grid(column=0, row=6, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.OutputText = tk.Label(self,text="asd")
        self.OutputText.grid(column=0, row=7, columnspan=2, sticky=tk.W, padx=5, pady=5)


if __name__ == "__main__":
    app = App()
    app.mainloop()