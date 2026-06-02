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
            print(f"Im Datensatz {mmsid} stimmen die Nummer in 502 $o {DissNrA} und die Nummer in 990 {DissNrB} nicht überein.")
        else:
            continue

