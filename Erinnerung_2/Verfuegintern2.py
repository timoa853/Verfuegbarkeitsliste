import win32com.client

def verfuegbenachrichtigung1():
    # Variable dimensionieren
    xlsApp = win32com.client.Dispatch("Excel.Application")
    xlsApp.Visible = False

    # Excel-Datei auswählen
    file_path = r"C:\Users\user\Desktop\Workflows\Verfuegbarkeitsliste\Ausfuehrende_Datei_Verfueg.xlsx"
    xlsWb = xlsApp.Workbooks.Open(file_path)

    # Referenz auf das Arbeitsblatt und die Liste erhalten
    arbeitsblatt = xlsWb.Sheets("Versand Verfüg")
    liste = arbeitsblatt.ListObjects("Prf.Verfüg")

    # Schleife über alle Zeilen der Liste
    for zeile in range(1, liste.ListRows.Count + 1):
        # Prüfen, ob Zelle in Liste mit X gefüllt ist
        if liste.ListColumns("Aktualisieren").DataBodyRange(zeile).Value == liste.ListColumns(
                "Zwecks Codierung").DataBodyRange(zeile).Value:
            # Mail versenden
            verfuegbenachrichtigung2_2(
                liste.ListColumns("Aktualisieren").DataBodyRange(zeile).Value,
                liste.ListColumns("Mailadresse").DataBodyRange(zeile).Value
            )

    # Excel-Datei speichern
    xlsApp.DisplayAlerts = False
    #xlsWb.Save() - Dieser Befehl wurde aufgrund eines Fehlers ausgeblendet

    # Excel-Applikation beenden
    xlsApp.Quit()

def verfuegbenachrichtigung2_2(baldaktualisieren, mailadresse):
    # Outlook-Applikation starten
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    oMail = outlook_app.CreateItem(0)  # 0 bedeutet MailItem

    # Mail formatieren
    oMail.To = mailadresse
    oMail.Subject = "Bitte umgehend Verfügbarkeitsliste anpassen"
    oMail.BodyFormat = 2  # 2 bedeutet olFormatHTML
    oMail.HTMLBody = "<br><br><b><div style=background-color:red><font color= white>Priorität: Hoch</font></div><br><br> Bitte Verfügbarkeitsliste schleunigst aktualisieren</b><br><br>Bitte in den nächsten Tagen prüfen und aktualisieren, sowie den Lieferanten des fehlenden Teiles kontaktieren, sofern kein neuer Liefertermin bekannt ist<br><br><b>Pfad: Allgemein_Server\\PREISLISTEN\\Jahr\\Preisliste_Master</b><br><br>Mit freundlichen Grüßen<br><br>DINO - Automatische Nachricht"

    oMail.Send()

if __name__ == "__main__":
    verfuegbenachrichtigung1()