#Funktionssammlung für alle Punkte, die getestet werden können
#Pro Punkt eine Funktion



#Testfunktion, kann später gelöscht werden
def testfunktion():
    print("Das ist ein Test.")

def get_mmsid(elem):
    if elem.tag.endswith("Identifier"):
        mmsid = elem.text

def DissNr_502_990(elem, mmsid):
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
            print("keine Fehler")