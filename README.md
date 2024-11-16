Es handelt sich hier um einen automatisierten Workflow, der dafür sorgt, dass PDF und CSV Dateien automatisiert per E-Mail versandt werden. Die Arbeitsschritte erfolgen im Hintergrund.
Die PDF Dateien werden aus Excel-Dateien erstellt.

Hauptfunktion:
Erstellung von CSV und PDF Dateien und Versand per E-Mail (bestehend aus: Versand_PDF_CSV.bat, Versand_PDF_CSV.py, Versand_PDF_CSV_ausfuehren.vbs



Weitere Funktionen:
-Excel-Datei aktualisieren ( bestehend aus workbook_save.bat / workbook_save.py / workbook_save.vbs)
-Erinnerungsmail an bearbeiter, dass der vorgegebene Liefertermin bald erreicht ist in zwei Prioritätsstufen (bestehend aus Verfuegintern1.bat / Verfuegintern2.bat / Verfuegintern1.py / Verfuegintern2.py / Verfuegintern1_ausfuehren.vbs / Verfuegintern2_ausfuehren.vbs)


ACHTUNG - alle oben aufgelisteten Dateien müssen mit einer "einfachen Aufgabe" aus der Windows Aufgabenplanung abgerufen werden!
Damit die Aufgaben auch im Hintergrund ausgeführt werden und das CMD-Fenster ausgeblendet wird, muss unter dem Reiter "Aktionen"...
- innerhalb des Input Feldes von "Programm/Script" der Inhalt "wscript.exe" stehen.
- unter "Argumente hinzufügen (optional)" der Pfad zur jeweiligen VBS-Datei stehen.
- ein zeitlicher Trigger hinterlegt werden z.B. täglich um 14 Uhr
