Option Explicit

Dim ExcelApp, ExcelWorkbook, OutlookApp, OutlookMail

' Set the path to your Excel file
Dim excelFilePath
excelFilePath = "C:\Users\suraj lotankar\Downloads\File.xlsx"

' Set the sheet name and range to read data from
Dim sheetName, startCell
sheetName = "Sheet1"
startCell = "A1"

' Set Outlook email details
Dim recipient, subject, body
recipient = "sonalilotankar99@outlook.com"
subject = "Data from Excel Sheet"

' Initialize Excel application
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible = False ' Set to True if you want to see Excel

' Open the Excel workbook
Set ExcelWorkbook = ExcelApp.Workbooks.Open(excelFilePath)

' Get data and headers from the Excel sheet
Dim data, headers, col
headers = ExcelWorkbook.Worksheets(sheetName).Range(startCell).CurrentRegion.Rows(1).Value
data = ExcelWorkbook.Worksheets(sheetName).Range(startCell).CurrentRegion.Offset(1, 0).Resize(ExcelWorkbook.Worksheets(sheetName).Range(startCell).CurrentRegion.Rows.Count - 1).Value

' Close Excel workbook and release resources
ExcelWorkbook.Close
Set ExcelWorkbook = Nothing
ExcelApp.Quit
Set ExcelApp = Nothing

' Build the email body with HTML-formatted Excel data
body = "<html><body>"
body = body & "<p>Hello,</p>"
body = body & "<p>Please find the data from the Excel sheet below:</p>"
body = body & "<table border='1' cellpadding='5' cellspacing='0'><tr>"

' Add dynamic table headers
Dim header
For Each header In headers
    body = body & "<th>" & header & "</th>"
Next

body = body & "</tr>"

' Add dynamic table data
Dim row
For row = LBound(data, 1) To UBound(data, 1)
    body = body & "<tr>"
    For col = LBound(data, 2) To UBound(data, 2)
        body = body & "<td>" & data(row, col) & "</td>"
    Next
    body = body & "</tr>"
Next

body = body & "</table>"
body = body & "<p>Best regards,<br/>Sonali Lotankar</p>"
body = body & "</body></html>"

' Initialize Outlook application
Set OutlookApp = CreateObject("Outlook.Application")

' Create a new email
Set OutlookMail = OutlookApp.CreateItem(0)
With OutlookMail
    .To = recipient
    .Subject = subject
    .HTMLBody = body
    .Attachments.Add excelFilePath ' Attach the Excel file
    .Send ' Uncomment if you want to send automatically, or use .Display to show the email before sending
End With

' Release Outlook resources
Set OutlookMail = Nothing
Set OutlookApp = Nothing

WScript.Echo "Email sent successfully!"

' End of script

