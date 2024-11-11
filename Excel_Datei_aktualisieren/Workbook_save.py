import win32com.client
import time

# Name, den Sie löschen möchten
target_name = "Print_Area"



# Excel-Applikation starten
xlsApp = win32com.client.Dispatch("Excel.Application")



# Excel-Datei auswählen
file_path = r"C:\Users\user\Desktop\Workflows\Verfuegbarkeitsliste\Ausfuehrende_Datei_Verfueg.xlsx"
xlsWb = xlsApp.Workbooks.Open(file_path)

# Überprüfen, ob der Name existiert
for name in xlsWb.Names:
    if name.Name == target_name:
        # Wenn der Name existiert, löschen
        name.Delete()
        break  # Brechen Sie die Schleife ab, da der Name gefunden und gelöscht wurde



# Alle Abfragen aktualisieren
xlsWb.RefreshAll()

# Pause für 10 Sekunden
time.sleep(10)

# Excel-Datei speichern
xlsApp.DisplayAlerts = False
xlsWb.Save()

# Excel-Applikation beenden
xlsApp.Quit()