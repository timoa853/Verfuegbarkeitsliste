import win32com.client

def email_spiel_preis_pdf():
    # Excel-Applikation starten
    xlsApp = win32com.client.Dispatch("Excel.Application")

    # Sichtbarkeit setzen
    xlsApp.Visible = False

    # Excel-Datei auswählen
    file_path = r"C:\path\to\your\excelsheet.xlsx"
    xlsWb = xlsApp.Workbooks.Open(file_path)

    # PDF erstellen
    xlsWb.Sheets("Versand Verfüg").Activate()
    datei_name = xlsWb.Sheets("Versand Verfüg").Range("U3").Value + xlsWb.Sheets("Versand Verfüg").Range("U2").Value + ".pdf" #Dateiname besteht zum Teil aus einer Excelzelle - das kann hier frei geändert werden
    xlsWb.Sheets("Versand Verfüg").Range("A1:H374").ExportAsFixedFormat(
        Type=0, Filename=datei_name, Quality=0, IncludeDocProperties=True, IgnorePrintAreas=False, OpenAfterPublish=False
    ) #Range A1:H374 ist der Druckberech der in der PDF zu angezeigt wird. Der Rest wird ausgeblendet

    # Für CSV: Druckbereich kopieren
    temp_workbook = xlsApp.Workbooks.Add()
    xlsWb.Sheets("Versand Verfüg").Range("A1:H374").Copy() #Range A1:H374 ist der Druckberech der in die CSV übertragen wird
    temp_workbook.Sheets(1).Range("A1").PasteSpecial(3)
    temp_workbook.Sheets(1).Range("C:C").NumberFormat = "0"  # Format der Spalte C auf Zahl setzen

    # Für CSV: temporäre Arbeitsmappe aktivieren und speichern
    temp_workbook.Activate()
    csv_datei_name = xlsWb.Sheets("Versand Verfüg").Range("U3").Value + xlsWb.Sheets("Versand Verfüg").Range("U2").Value + ".csv" #Dateiname besteht zum Teil aus einer Excelzelle - das kann hier frei geändert werden

    # Unterdrücken der Meldung "Möchten Sie die Datei ersetzen"
    xlsApp.DisplayAlerts = False
    temp_workbook.SaveAs(Filename=csv_datei_name, FileFormat=6)  # 6 entspricht xlCSV
    xlsApp.DisplayAlerts = True

    # Schließen der temporären Arbeitsmappe
    temp_workbook.Close(False)
    temp_workbook = None

    # Email erstellen und versenden
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    outlook_mail_item = outlook_app.CreateItem(0)
    my_attachments = outlook_mail_item.Attachments

    # E-Mail-Adressen in kleinere Gruppen aufteilen - Bitte hier die korrekten E-Mail-Adressen eintragen
    alle_adressen = [
        "test1@testmail.de", "test2@testmail.de", "test3@testmail.de", "test4@testmail.de", "test5@testmail.de", "test6@testmail.de",
        "test7@testmail.de", "test8@testmail.de", "test9@testmail.de", "test10@testmail.de", "test11@testmail.de", "test12@testmail.de",
        "test13@testmail.de",
    ]

    gruppen_groesse = 10  # Anzahl der Empfänger pro E-Mail, um sicherzustellen, dass auch alle E-Mails versandt werden
    adress_gruppen = [alle_adressen[i:i + gruppen_groesse] for i in range(0, len(alle_adressen), gruppen_groesse)]

    for gruppe in adress_gruppen:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        outlook_mail_item = outlook_app.CreateItem(0)
        my_attachments = outlook_mail_item.Attachments

        # Separate Zuweisungen
        outlook_mail_item.To = "your-email-address@testmail.de" #Hauptempfänger ist die eigene E-Mail Adresse aus Datenschutzgründen
        outlook_mail_item.BCC = "; ".join(gruppe)  # BCC-Feld mit den E-Mail-Adressen der Gruppe
        outlook_mail_item.Subject = "Aktuelle Verfügbarkeitsliste"
        outlook_mail_item.BodyFormat = 2  # 2 entspricht olFormatHTML
        outlook_mail_item.HTMLBody = """
            <br><br>Sehr geehrte Damen und Herren<br><br>
            anbei erhalten Sie unsere aktuelle Verfügbarkeitsliste als PDF und CSV.<br>
            BITTE GEBEN SIE UNS GERNE EINE RÜCKMELDUNG, SOFERN DATEN NICHT VOLLSTÄNDIG ANGEZEIGT WERDEN ODER FEHLERCODES ERSICHTLICH SIND.<br><br>
            <hr>Dear Sir / Madam,<br><br>
            Enclosed you will find our current availability list as PDF and CSV.<br>
            PLEASE FEEL FREE TO GIVE US FEEDBACK IF DATA IS NOT COMPLETELY DISPLAYED OR ERROR CODES ARE VISIBLE.<br><br><br>
            <i><b>Achtung, diese Nachricht wurde automatisch generiert | This message was genereted automatically</b></i><br><br>
            Mit freundlichen Grüßen | With kind regards
        """

        # Anhänge hinzufügen
        my_attachments.Add(datei_name)
        my_attachments.Add(csv_datei_name)

        # E-Mail senden
        outlook_mail_item.Send()

    # Excel-Datei speichern
    xlsWb.Save()

    # Excel-Applikation beenden
    xlsApp.Quit()

if __name__ == "__main__":
    # Führen Sie das Python-Skript aus
    email_spiel_preis_pdf()
