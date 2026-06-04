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
DissNr = None #Diss Nummer aus 502
DissJahr = None #Diss Abgabe Jahr 
Inh_990 = None #Inhalt von Feld 990
hat_082 = False #Anzeige für Vorhandensein von 082
Inhalt_655 = [] #Sprachen, in denen 655 im Record steht

#Erstellen des Excel Files für den Output
workbook = xlsxwriter.Workbook("Output_Datei.xlsx")
worksheet = workbook.add_worksheet("Output_Datei")



#Parsing der Titeldatensätze
for event, elem in ET.iterparse("DissenTestfile.xml", events=("end",)):
  #mmsid abholen
    if elem.tag.endswith("Identifier"):
        mmsid = elem.text

  #auslesen von 502 $o und 990 für Vergleich der Diss Nr und Jahr
    if elem.attrib.get("tag") == "502":
        for sub in elem:
            if sub.attrib.get("code") == "o":
                DissNr = sub.text
            if sub.attrib.get("code") == "d":
                DissJahr = sub.text
    if elem.attrib.get("tag") == "990":
        for sub in elem:
            if sub.text[0:3] == "HSG":
                Inh_990 = sub.text

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

  #Prüft Vorhandensein einer Verknüpfung mit der GND in Feld 700
    if elem.attrib.get("tag") == "700":
        hat_0 = elem.find('./marc:subfield[@code="0"]', ns) is not None
        if hat_0 == False:
            fehlerliste.append("Kein Verknüfung in Feld 700")
        else:
            for sub in elem:    
                if sub.attrib.get("code") == "0" and not sub.text.startswith("(DE-588)"):
                    fehlerliste.append(f"keine GND-Verknüpfung in Feld 700, stattdessen: {sub.text}")
 
  #Prüft ob die aktuelle Zeile ein Feld 082 ist
    test_082 = elem.find('./marc:datafield[@tag="082"]', ns) is not None
    #setzt Prüfvarible hat_082 auf True falls 082 gefunden
    if test_082 == True:
        hat_082 = True

  #Prüft ob 655 in Deutsch, Französisch und Italienisch vorhanden sind und ergänzt pro Sprach die Liste
    if elem.attrib.get("tag") == "655" and elem.attrib.get("ind2") == "7":
        for sub in elem:
            if sub.attrib.get("code") == "a":
                if sub.text == "Hochschulschrift":
                    Inhalt_655.append("de")
                if sub.text == "Tesi":
                    Inhalt_655.append("it")
                if sub.text == "Thèses et écrits académiques":
                    Inhalt_655.append("fr")



  #Record Ende
    if elem.tag.endswith("recordPosition"):
        #prüft Inhalt von Inhalt_655
        if sorted(Inhalt_655) != sorted(["de", "it", "fr"]):
            fehlerliste.append(f"ein oder mehrere 655 fehlerhaft, korrekt: {Inhalt_655}")
        if Inhalt_655 == []:
            fehlerliste.append("keine 655 gefunden")

        # Verleich der Diss Nr aus 502 $o und 990
        if DissNr != Inh_990[14:18]:
            fehlerliste.append (f"Dissen Nr. in 502 '{DissNr}' und 990 '{Inh_990}' stimmen nicht überein")
        
        # Vergleich Diss Publikationsjahr aus 502 und 990
        if DissJahr != Inh_990[9:13]:
            fehlerliste.append(f"Publikationsjahr in 502 '{DissJahr}' und 990 '{Inh_990[9:13]}' stimmen nicht überein")


        #fügt Fehlertext zur Fehlerliste hinzu, falls die Prüfvarible für 082 in diesem Datensatz nicht auf True gesetzt wurde
        if hat_082 == False:
            fehlerliste.append("kein Feld 082 gefunden")
        
    # Zusammenführen aller Fehler aus fehlerliste mit mmsid und hinzufügen zu Liste content
        for fehler in fehlerliste:
            content.append([mmsid, fehler])
    # Zurücksetzen aller Variablen für den nächsten Record
        mmsid = None
        DissNr = None
        DissJahr = None
        Inh_990 = None
        hat_082 = False
        Inhalt_655 = []
        fehlerliste = [] 



#Schreibt Fehler ins Excel basierend auf der Liste content
for row, (mmsid, fehler) in enumerate(content):     
    worksheet.write(row, 0, mmsid) #Index 0 in content
    worksheet.write(row, 1, fehler) #Index 1 in content



workbook.close()