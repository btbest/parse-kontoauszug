@echo off
chcp 1252 >nul
SET pythonexe="C:%HOME%\.conda\envs\parse-kontoauszug\python.exe"
SET scriptpath="C:%HOME%\parse-kontoauszug\parsecomdi.py"
SET docspath="C:%HOME%\Documents\comdirect-finanzreports"
SET outname=girokonto.csv
echo.
echo Alle Finanzreport-pdfs im Zielordner werden gelesen.
echo Zielordner: %docspath%
echo Die Zieldatei %outname% wird im Verzeichnis dieses scripts abgespeichert.
echo Dateien jetzt lesen?
pause
%pythonexe% %scriptpath% -p %docspath% -o %outname%
echo Vorgang abgeschlossen. Beliebige Taste druecken um zu schliessen.
pause