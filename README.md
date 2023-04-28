# Kontoauszug-Parser
Dieses Python-Skript extrahiert aus comdirect Finanzreport-PDFs die Girokonto-Ums�tze und speichert sie als Tabelle im CSV-Format.
# Installation
Das Skript erfordert eine funktionierende Installation von Python 3 mit den Modulen pandas und pypdf. Der einfachste Weg, diese Umgebung herzustellen, ist der Umgebungsmanager [micromamba](https://mamba.readthedocs.io/en/latest/installation.html).
* Installiere micromamba nach den Anweisungen dort
* `micromamba create -f environment.yml -n comdiparse`
* `micromamba activate comdiparse`
# Ausf�hrung
* Falls die Finanzreport PDFs in demselben Ordner liegen wie das Skript, oder in einem Unterordner: `python parsecomdi.py`
* Falls die Finanzreport PDFs anderswo liegen:  `python parsecomdi.py -p "C:\Pfad\zu\PDFs"`
Optional kann man mit -o auch den Namen der Ausgabedatei steuern: `python parsecomdi.py -o "girokonto-2023.csv"`
# Weitere Auswertung der Tabelle
## Mit Excel
Excel sollte mit der CSV-Datei wie sie ist arbeiten k�nnen. Da das CSV-Format nur Tabellen beschreiben kann, bietet es sich f�r Analysen nat�rlich an, die weitere Arbeit stattdessen im XLSX-Format abzuspeichern.
## Mit Pandas
```
import pandas
df = pandas.read_csv("girokonto.csv", encoding="cp1252", header=0, sep=";")
print(df) # mach was du willst damit
```