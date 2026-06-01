#Hauptdatei mit dem eigentlichen Prüfvorgang


#Import der Module
import xml.etree.ElementTree as ET

#Parsing der Titeldatensätze
for event, elem in ET.iterparse(" ", events=("start", "end",)):
    