#Hauptdatei mit dem eigentlichen Prüfvorgang


#Import der Module
#Import von Modul xml.etree.ElementTree um durch die Zeilen der records zu iterieren
import xml.etree.ElementTree as ET
#Import von Modul xlsxwriter, zum Erstellen eines Excel Files als Output
import xlsxwriter
#import von Functions.py für Zugriff auf die dort geschriebenen Funktionen und Variablen
import Functions


# Variabeln definieren
ns = {"marc": "http://www.loc.gov/MARC21/slim"}
content = [("MMSID", "Fehlertext")] #Liste für mmsid und Fehlertext, die später ins Excel geschrieben werden inkl. subListe für "Titelzeile"

fehlerliste = [] #Liste für die Fehler ohne die MMSID pro Datensatz
mmsid = None #Identifikator der mmsid
DissNrA = None #Diss Nummer aus 502
DissNrB = None #Inhalt von Feld 990
unterfeld = None

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
            fehlerliste.append ( f"Dissen Nr. in 502 '{DissNrA}' und 990 '{DissNrB}' stimmen nicht überein")

    
  #prüft ob 245 $c mit korrekter Phrase startet
    if elem.attrib.get("tag") == "245":
        for sub in elem:
            if sub.attrib.get("code") == "c" and not sub.text.startswith(("submitted by", "vorgelegt von")):
                fehlerliste.append(f"Fehlerhafter Text in 245 $c: '{sub.text}'")

  #Prüft Vorhandensein einer Verknüpfung mit der GND in Feld 100
    if elem.attrib.get("tag") == "100":
        hat_0 = elem.find('./marc:subfield[@code="0"]', ns) is not None
        if hat_0 == False:
            fehlerliste.append("Kein Verknüfung in Feld 100")
        else:
            for sub in elem:    
                if sub.attrib.get("code") == "0" and not sub.text.startswith("(DE-588)"):
                    fehlerliste.append(f"keine GND-Verknüpfung in Feld 100, stattdessen: {sub.text}")






  #Record Ende
    if elem.tag.endswith("recordPosition"):
        # Zusammenführen aller Fehler aus fehlerliste mit mmsid und hinzufügen zu Liste content
        for fehler in fehlerliste:
            content.append([mmsid, fehler])
        # Zurücksetzen aller Variablen für den nächsten Record
        mmsid = None
        DissNrA = None
        DissNrB = None
        fehlerliste = [] 



#Schreibt Fehler ins Excel basierend auf der Liste content
for row, (mmsid, fehler) in enumerate(content):     
    worksheet.write(row, 0, mmsid) #Index 0 in content
    worksheet.write(row, 1, fehler) #Index 1 in content



workbook.close()