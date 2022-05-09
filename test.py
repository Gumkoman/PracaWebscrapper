from asyncio import events
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time


def generateRaport(listOfHits):
    pass

class Hit:
    def __init__(self, timeStamp, _id,event,device,eventEt,ip,city):
        self._timeStamp = timeStamp
        self._id = _id  
        self._event = event
        self._device = device
        self._eventEt = eventEt
        self._ip = ip
        self._city = city
    def printSelf(self):
        print(self._timeStamp,self._id,self._event,self._device,self._eventEt,self._ip,self._city)


 
DRIVER_PATH = 'chromedriver'
driver = webdriver.Chrome(executable_path=DRIVER_PATH)
# driver.get("http://kibana.kantei-bd.redefine.pl/app/kibana#/discover?_g=()&_a=(columns:!(id,event.os,event.event_name,event.dv,event.et),index:appevent,interval:auto,query:(language:lucene,query:''),sort:!(headers.receiveTimestamp,desc))")
# driver.get("http://kibana.kantei-bd.redefine.pl/app/kibana#/discover?_g=()&_a=(columns:!(id,event.event_name,event.dv,event.et,headers.additionalProperties.client-ip),index:appevent,interval:ms,query:(language:lucene,query:''),sort:!(headers.receiveTimestamp,desc))")
driver.get("http://kibana.kantei-bd.redefine.pl/app/kibana#/discover?_g=()&_a=(columns:!(id,event.event_name,event.dv,event.et,headers.additionalProperties.client-ip,headers.additionalProperties.client-city),index:appevent,interval:ms,query:(language:lucene,query:''),sort:!(headers.receiveTimestamp,desc))")

print("#0")
for i in range(35):
    print("time",i)
    time.sleep(1)


times = driver.find_elements(by=By.XPATH,value = "//td[@class='eui-textNoWrap']")
values = driver.find_elements_by_xpath(".//td[@class='kbnDocTableCell__dataField eui-textBreakAll eui-textBreakWord']")
collumns = driver.find_elements_by_xpath(".//span[@i18n-id='common.ui.docTable.tableHeader.timeHeaderCellTitle']")

print("Columns",len(collumns))
for col in collumns:
    print(col.text)
print("########################")
print("times",len(times))
for time1 in times:
    # realTime = time1.find_element_by_xpath("//span[.='ng-non-bindable']")
    print(time1.text)
print("########################")
print("values",len(values))
# for val in values:
#     # value = val.find_element_by_xpath("//span[.='ng-non-bindable']")
#     print(val.text)
listOfHits = []
for item in range(0,len(values),6):
    tempItem = Hit(times[int(item/6)].text,values[item+0].text,values[item+1].text,values[item+2].text,values[item+3].text,values[item+4].text,values[item+5].text)
    listOfHits.append(tempItem)

uniqueIPs=[]
for hit in listOfHits:
    flag = True
    for ip in uniqueIPs:
        if ip == hit._ip:
            flag = False
    if flag:
        uniqueIPs.append(hit._ip)

for i,item in enumerate(uniqueIPs):
    print(i,item,end=" ")
    for hit in listOfHits:
        if(item == hit._ip):
            print(hit._city)
            break
number = input("Wybierz numer IP z którego korzystasz (proszę podać numer z powyższej listy):")
number = int(number)
print(uniqueIPs[number])
driver.get("http://kibana.kantei-bd.redefine.pl/app/kibana#/discover?_g=()&_a=(columns:!(id,event.event_name,event.dv,event.et,headers.additionalProperties.client-ip),filters:!(('$state':(store:appState),meta:(alias:!n,disabled:!f,index:appevent,key:headers.additionalProperties.client-ip,negate:!f,params:(query:'[{}]',type:phrase),type:phrase,value:'{}'),query:(match:(headers.additionalProperties.client-ip:(query:'{}',type:phrase))))),index:appevent,interval:ms,query:(language:lucene,query:''),sort:!(headers.receiveTimestamp,desc))".format(uniqueIPs[number],uniqueIPs[number],uniqueIPs[number]))

for i in range(5):
    print("time",i)
    time.sleep(1)

times = driver.find_elements(by=By.XPATH,value = "//td[@class='eui-textNoWrap']")
values = driver.find_elements_by_xpath(".//td[@class='kbnDocTableCell__dataField eui-textBreakAll eui-textBreakWord']")
collumns = driver.find_elements_by_xpath(".//span[@i18n-id='common.ui.docTable.tableHeader.timeHeaderCellTitle']")

listOfHits = []
for item in range(0,len(values),5):
    tempItem = Hit(times[int(item/5)].text,values[item+0].text,values[item+1].text,values[item+2].text,values[item+3].text,values[item+4].text,"-")
    listOfHits.append(tempItem)

for element in listOfHits:
    element.printSelf()

generateRaport(listOfHits)

# for i in range(len(values)):
#     if(i % 5 == 0):
#         print("")
#     print(values[i].text,end="\t")

print("Koniec")

