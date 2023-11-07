# Kontoauszug-Parser
Dieses Python-Skript extrahiert aus comdirect Finanzreport-PDFs die Girokonto-Umsätze und speichert sie als Tabelle im CSV-Format.

Anhand von Stichwörtern werden die Umsätze simpel kategorisiert. Das smarter zu gestalten ist eine Hoffnung für die Zukunft.
# Installation
Das Skript erfordert eine funktionierende Installation von Python 3 mit den Modulen pandas und pypdf. Der einfachste Weg, diese Umgebung herzustellen, ist der Umgebungsmanager conda. Conda funktioniert unabhängig von evtl. anderen existierenden Python-Installationen.
* Dieses Repo runterladen und entpacken (grüner Knopf rechts oben: Code -> Download ZIP)
* Installiere [conda/miniforge](https://github.com/conda-forge/miniforge#miniforge3) nach den Anweisungen dort
* Öffne eine miniforge-Kommandozeile
* `cd pfad/zu/parse-kontoauszug`
* `conda create -f environment.yml`
# Ausführung
## Per mitgeliefertem Skript
* `run-default.bat` bearbeiten: Die drei Pfade zu python.exe, parsecomdi.py und dem Ordner mit den PDFs durch diejenigen auf dem eigenen System ersetzen.
  * Conda erstellt die Umgebung standardmäßig im Pfad `%HOME%\.conda\envs\parse-kontoauszug`, wie es im Skript schon steht. Falls das nicht der Fall ist, gibt `conda env list` Auskunft, wo die Umgebung und damit python.exe liegt.  
* `run-default.bat` ausführen.
## Direkt aus der Kommandozeile
* `conda activate parse-kontoauszug`
* `python parsecomdi.py` falls die Finanzreport PDFs in demselben Ordner liegen wie das Skript, oder in einem Unterordner.
* `python parsecomdi.py -p "C:\Pfad\zu\PDFs"` falls die Finanzreport PDFs anderswo liegen.
* `python parsecomdi.py -o "girokonto-2023.csv"` um den Namen der Ausgabedatei zu steuern.
# Ergebnis
Erwartete Ausgabe in der Konsole:
![Erwartete Ausgabe in der Konsole](readme/console-out.jpg)

CSV-Ausgabe geöffnet in Excel:
![Tabelle in Excel](readme/output-excel.jpg)
# Weitere Auswertung der Tabelle
## Mit Excel
Excel sollte mit der CSV-Datei wie sie ist arbeiten können. Da das CSV-Format nur simple Tabellen beschreiben kann, bietet es sich für Analysen natürlich an, die weitere Arbeit stattdessen im XLSX-Format abzuspeichern.
## Mit Pandas
Die Tabelle wird um zu Excel nett zu sein im Windows-Encoding gespeichert. Das muss man Pandas explizit angeben:
```
import pandas
df = pandas.read_csv("girokonto.csv", encoding="cp1252", header=0, sep=";")
print(df) # mach was du willst damit
```
