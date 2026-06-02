#Hauptdatei mit dem eigentlichen Prüfvorgang


#Import der Module
#Import von Modul xml.etree.ElementTree um durch die Zeilen der records zu iterieren
import xml.etree.ElementTree as ET
#Import von Modul xlsxwriter, zum Erstellen eines Excel Files als Output
import xlsxwriter
#import von Functions.py für Zugriff auf die dort geschriebenen Funktionen und Variablen
import Functions

# Variabeln definieren
#mmsid = None
content = [] #Liste für mmsid und Fehlertext, die später ins Excel geschrieben werden

#Erstellen des Excel Files für den Output
workbook = xlsxwriter.Workbook("Output_Datei.xlsx")
worksheet = workbook.add_worksheet("Output_Datei")



#Parsing der Titeldatensätze
for event, elem in ET.iterparse("DissenTestfile.xml", events=("end",)):
  #mmsid abholen
    if elem.tag.endswith("Identifier"):
        mmsid = elem.text

  #vergleicht 502 $o und 990
    if elem.attrib.get("tag") == "502":
        for sub in elem:
            if sub.attrib.get("code") == "o":
                DissNrA = sub.text
    if elem.attrib.get("tag") == "990":
        for sub in elem:
            if sub.text[0:3] == "HSG":
                DissNrB = sub.text     
    if elem.tag.endswith("recordPosition"):        
        if DissNrA != DissNrB[14:18]:
            # Liste content wird mit Sub-Liste ergänzt - Index 0 mmsid und Index 1 Fehlermeldung
            content.append ([mmsid, f"Dissen Nr. in 502 '{DissNrA}' und 990 '{DissNrB}' stimmen nicht überein"])
        else:
            continue


#Schreibt Fehler ins Excel basierend auf der Liste content
for row, (mmsid, fehler) in enumerate(content):     
    worksheet.write(row, 0, mmsid) #Index 0 in content
    worksheet.write(row, 1, fehler) #Index 1 in content


workbook.close()