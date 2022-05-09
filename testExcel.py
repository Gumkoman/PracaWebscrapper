import xlsxwriter
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


event1 = [
"Rozpoczęcie odtwarzania",
"Sprawdzenie licencji DRM",
"Rozpoczęcie odtwarzania materiału (pierwsza klatka materiału)",
"Hit cykliczny podczas odtwarzania",
"Zmiana jakości",
"Zapauzowanie playera",
"Odpauzowanie playera",
"Zamknięcie playera",
"Zakończenie odtwarzania materiału (ostatnia klatka)",
"Seekowanie materiału",
"Przerwanie odtwarzania materiału",
"Poczas restartu aplikacji?",
"Rozpoczęcie buforowania",
"Zakończenie buforowania",
"Nieobsłużony błąd odtwarzania",
"Baner widoczny na playerze"
]
event2 = [
    "Rozpoczęcie odtwarzania",
    "Sprawdzenie licencji DRM",
    "Rozpoczęcie odtwarzania materiału (pierwsza klatka materiału)",
    "Hit cykliczny podczas odtwarzania",
    "Zmiana jakości",
    "Zapauzowanie playera",
    "Odpauzowanie playera",
    "Zamknięcie playera",
    "Zakończenie odtwarzania materiału (ostatnia klatka)",
    "Seekowanie materiału",
    "Przerwanie odtwarzania materiału",
    "Poczas restartu aplikacji?",
    "Rozpoczęcie buforowania",
    "Zakończenie buforowania",
    "Nieobsłużony błąd odtwarzania",
    "Baner widoczny na playerze"
]
event3 = [
    "Rozpoczęcie odtwarzania",
    "Sprawdzenie licencji DRM",
    "Rozpoczęcie odtwarzania materiału (pierwsza klatka materiału)",
    "Hit cykliczny podczas odtwarzania",
    "Zmiana jakości",
    "Zapauzowanie playera",
    "Odpauzowanie playera",
    "Zamknięcie playera",
    "Seekowanie materiału",
    "Przerwanie odtwarzania materiału",
    "Poczas restartu aplikacji?",
    "Rozpoczęcie buforowania",
    "Zakończenie buforowania",
    "Nieobsłużony błąd odtwarzania",
    "Baner widoczny na playerze"
]

przypadki1 = [
'try',
'(*begin) DRM',
'begin',
'cycle',
'change quality',
'pause',
'unpause',
'close',
'end - brak w przypadku odtworzenia następnego odcinka',
'seek',
'interrupted',
'restart_app',
'buffering start',
'buffering end',
'generic error',
'banner']

przypadki2 = [
    "try",
    "(*begin) DRM",
    "begin",
    "cycle",
    "change quality",
    "pause",
    "unpause",
    "close",
    "end",
    "seek",
    "interrupted",
    "restart_app",
    "buffering start",
    "buffering end",
    "generic error",
    "banner"
]
przypadki3=[
    "try",
    "(*begin) DRM",
    "begin",
    "cycle",
    "change quality",
    "pause",
    "unpause",
    "close",
    "seek",
    "interrupted",
    "restart_app",
    "buffering start",
    "buffering end",
    "generic error",
    "banner"
]

eventName1 =[
"PlayerInitialized",
"PlayerBeganDrm",
"PlayerBeganPlay",
"PlayerCycleHit",
"PlayerQualityChanged",
"PlayerPaused",
"PlayerUnPaused",
"PlayerClosed",
"PlayerEndedPlay",
"PlayerSeek",
"PlayerInterrupted",
"PlayerRestartedApp",
"PlayerStartedBuffering",
"PlayerStoppedBuffering",
"PlayerGenericError",
"PlayerShowedBanner"]
eventName2= [
    "PlayerInitialized",
    "PlayerBeganDrm",
    "PlayerBeganPlay",
    "PlayerCycleHit",
    "PlayerQualityChanged",
    "PlayerPaused",
    "PlayerUnPaused",
    "PlayerClosed",
    "PlayerEndedPlay",
    "PlayerSeek",
    "PlayerInterrupted",
    "PlayerRestartedApp",
    "PlayerStartedBuffering",
    "PlayerStoppedBuffering",
    "PlayerGenericError",
    "PlayerShowedBanner"
]
eventName3=[
    "PlayerInitialized",
    "PlayerBeganDrm",
    "PlayerBeganPlay",
    "PlayerCycleHit",
    "PlayerQualityChanged",
    "PlayerPaused",
    "PlayerUnPaused",
    "PlayerClosed",
    "PlayerSeek",
    "PlayerInterrupted",
    "PlayerRestartedApp",
    "PlayerStartedBuffering",
    "PlayerStoppedBuffering",
    "PlayerGenericError",
    "PlayerShowedBanner"
]

desc = [
"Uruchomienie aplikacji",
"Nawigacja do dowolnego widoku zakończona sukcesem",
"Nawigacja do dowolnego widoku zakończona błędem",
"Zalogowanie użytkownika na konto",
"Błąd podczas logowania",
"Wylogowanie z aplikacji",
"Zrzucenie aplikacji w tło",
"Przywrócenie aplikacji z tła"
]
eventType1 = [
    "hit AppStarted",
    "hit AppUserNavigated",
    "hit AppUserNavigated",
    "hit AppUserLogged",
    "hit AppUserLogged",
    "hit AppUserLoggedOut",
    "hit AppPaused",
    "hit AppResumed"

]


def makeExcel(name,data):
    workbook = xlsxwriter.Workbook(name+'.xlsx')
    worksheet = workbook.add_worksheet("Player Events")

    redBackgroundFormat = workbook.add_format({'bold':True,'bg_color':'red',"font_size":14})
    lightGreyBackgroundFormat = workbook.add_format({'bold':True,'bg_color':'#d3d3d3','font_color':'red'})
    darkGreyBackgroundFormat = workbook.add_format({'bold':True,'bg_color':'#808080','font_color':'black'})
    greenBackgroundFormat = workbook.add_format({'bold':True,'bg_color':'#81B622','font_color':'black'})
    boldFormat = workbook.add_format({'bold':True,'font_color':'black'})

    worksheet.set_column(0,4,35)

    worksheet.merge_range(0,0,0,4,"Fromularz Testów player Events",redBackgroundFormat)
    worksheet.merge_range(1,0,1,4,"Data:")
    worksheet.merge_range(2,0,2,4,"Platforma")
    worksheet.merge_range(3,0,3,4,"Imie i nazwisko")


    row = 6
    col = 1

    worksheet.write(row-2,col-1,"Sekcja",darkGreyBackgroundFormat)
    worksheet.write(row-2,col,"Zdarzenie",darkGreyBackgroundFormat)
    worksheet.write(row-2,col+1,"Przypadki",darkGreyBackgroundFormat)
    worksheet.write(row-2,col+2,"Nazwa Eventu",darkGreyBackgroundFormat)
    worksheet.write(row-2,col+3,"Wynik",darkGreyBackgroundFormat)

    worksheet.merge_range(row-1,col-1,row-1,col+3,"Materiał VOD sekwencyjny (serial) ",lightGreyBackgroundFormat)

    for i in range(16):
        worksheet.write(row,col,event1[i])
        worksheet.write(row,col+1,przypadki1[i])
        worksheet.write(row,col+2,eventName1[i])
        for j in data:
            if eventName1[i].lower() == j._type:
                if j._status == "success":
                    worksheet.write(row,col+3,"Sukces",greenBackgroundFormat)

        row = row+1
    row = 23
    col = 1
    
    worksheet.merge_range(row-1,col-1,row-1,col+3,"Materiał VOD (film)",lightGreyBackgroundFormat)
    
    for i in range(16):
        worksheet.write(row,col,event2[i])
        worksheet.write(row,col+1,przypadki2[i])
        worksheet.write(row,col+2,eventName2[i])
        for j in data:
            if eventName1[i].lower() == j._type:
                if j._status == "success":
                    worksheet.write(row,col+3,"Sukces",greenBackgroundFormat)
        row = row+1
    row = 40
    col = 1
    
    worksheet.merge_range(row-1,col-1,row-1,col+3,"Materiał LIVE/Kanał TV - DRM",lightGreyBackgroundFormat)

    for i in range(15):
        worksheet.write(row,col,event3[i])
        worksheet.write(row,col+1,przypadki3[i])
        worksheet.write(row,col+2,eventName3[i])
        for j in data:
            if eventName1[i].lower() == j._type:
                if j._status == "success":
                    worksheet.write(row,col+3,"Sukces",greenBackgroundFormat)
        row = row+1
    for i in range(5):
        worksheet.write(row,col-1+i," ",darkGreyBackgroundFormat)

    worksheet2  = workbook.add_worksheet("Activity Events")
    worksheet2.set_column(0,4,35)

    worksheet2.merge_range(0,0,0,2,"Fromularz Testów player Events",redBackgroundFormat)
    worksheet2.merge_range(1,0,1,2,"Data:")
    worksheet2.merge_range(2,0,2,2,"Platforma")
    worksheet2.merge_range(3,0,3,2,"Imie i nazwisko")
    worksheet2.write(4,0,"Opis",boldFormat)
    worksheet2.write(4,1,"Typ Eventu",boldFormat)
    worksheet2.write(4,2,"Wynik",boldFormat)
    row = 6
    col = 0
    for i in range(8):
        worksheet2.write(row,col,desc[i])
        worksheet2.write(row,col+1,eventType1[i])
        for j in data:
            if eventType1[i].lower().split(" ")[1].strip() == j._type:
                if j._status == "success":
                    worksheet.write(row,col+3,"Sukces",greenBackgroundFormat)
        row = row+1
    workbook.close()

if __name__ == "__main__":
    makeExcel("test1","")