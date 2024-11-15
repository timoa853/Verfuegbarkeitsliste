import win32com.client

def verfuegbenachrichtigung1():
    # Variable dimensionieren
    xlsApp = win32com.client.Dispatch("Excel.Application")
    xlsApp.Visible = False

    # Excel-Datei auswählen
    file_path = r"C:\path\to\your\excel-file.xlsx"
    xlsWb = xlsApp.Workbooks.Open(file_path)

    # Referenz auf das Arbeitsblatt und die Liste erhalten
    arbeitsblatt = xlsWb.Sheets("Versand Verfüg")
    liste = arbeitsblatt.ListObjects("Prf.Verfüg")

    # Schleife über alle Zeilen der Liste - vergleicht aktuelles Datum mit dem mitgeteilten Verfügbarkeitstermin
    for zeile in range(1, liste.ListRows.Count + 1):
        # Prüfen, ob Zelle in Liste mit X gefüllt ist - wurde mit einer Wenn-Funktion gelöst (Wenn z.B. Verfügbarkeitstermin in einer Woche erreicht ist)
        if liste.ListColumns("Baldaktualisieren").DataBodyRange(zeile).Value == liste.ListColumns(
                "Zwecks Codierung").DataBodyRange(zeile).Value:
            # Mail versenden
            verfuegbenachrichtigung1_2(
                liste.ListColumns("Baldaktualisieren").DataBodyRange(zeile).Value,
                liste.ListColumns("Mailadresse").DataBodyRange(zeile).Value
            )

    # Excel-Datei speichern
    xlsApp.DisplayAlerts = False
    #xlsWb.Save() - Dieser Befehl wurde aufgrund eines Fehlers ausgeblendet

    # Excel-Applikation beenden
    xlsApp.Quit()

def verfuegbenachrichtigung1_2(baldaktualisieren, mailadresse):
    # Outlook-Applikation starten
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    oMail = outlook_app.CreateItem(0)  # 0 bedeutet MailItem

    # Mail formatieren - Mailadresse wird aus Excel-Datei übernommen. Kann auch direkt in den Code implementiert werden
    oMail.To = mailadresse
    oMail.Subject = "Verfügbarkeitsliste bald aktualisieren"
    oMail.BodyFormat = 2  # 2 bedeutet olFormatHTML
    oMail.HTMLBody = "<br><br><b><div style=background-color:yellow>Priorität: Mittel</div><br><br> Bitte Verfügbarkeitsliste aktualisieren</b><br><br>Bitte innerhalb einer Woche anpassen und Lieferanten kontaktieren sofern kein neuer Liefertermin bekannt ist<br><br><br><br>Mit freundlichen Grüßen<br><br>Automatische Nachricht"

    oMail.Send()

if __name__ == "__main__":
    verfuegbenachrichtigung1()
