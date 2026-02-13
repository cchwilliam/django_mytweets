Attribute VB_Name = "OutlookEmailReport"
Option Explicit

' Generates an Excel report with:
' 1) Email count per date
' 2) Top sender within a user-specified period
Public Sub GenerateEmailReport()
    Dim startDateInput As String
    Dim endDateInput As String
    Dim startDate As Date
    Dim endDate As Date

    startDateInput = InputBox("Enter start date (YYYY-MM-DD):", "Report Start Date")
    If Len(Trim$(startDateInput)) = 0 Then Exit Sub

    endDateInput = InputBox("Enter end date (YYYY-MM-DD):", "Report End Date")
    If Len(Trim$(endDateInput)) = 0 Then Exit Sub

    On Error GoTo DateError
    startDate = DateValue(startDateInput)
    endDate = DateValue(endDateInput) + TimeValue("23:59:59")
    On Error GoTo 0

    If endDate < startDate Then
        MsgBox "End date must be after start date.", vbExclamation
        Exit Sub
    End If

    Dim inbox As Outlook.Folder
    Dim items As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim mailItem As Outlook.MailItem

    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    Set items = inbox.Items
    items.Sort "[ReceivedTime]", True

    Dim filter As String
    filter = "[ReceivedTime] >= '" & Format(startDate, "ddddd h:nn AMPM") & "' AND [ReceivedTime] <= '" & Format(endDate, "ddddd h:nn AMPM") & "'"
    Set filteredItems = items.Restrict(filter)

    Dim dateCounts As Object
    Dim senderCounts As Object
    Set dateCounts = CreateObject("Scripting.Dictionary")
    Set senderCounts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To filteredItems.Count
        If TypeName(filteredItems.Item(i)) = "MailItem" Then
            Set mailItem = filteredItems.Item(i)
            Dim receivedDate As String
            receivedDate = Format(mailItem.ReceivedTime, "yyyy-mm-dd")

            If Not dateCounts.Exists(receivedDate) Then
                dateCounts.Add receivedDate, 0
            End If
            dateCounts(receivedDate) = dateCounts(receivedDate) + 1

            Dim sender As String
            sender = Trim$(mailItem.SenderEmailAddress)
            If Len(sender) = 0 Then sender = "(Unknown)"

            If Not senderCounts.Exists(sender) Then
                senderCounts.Add sender, 0
            End If
            senderCounts(sender) = senderCounts(sender) + 1
        End If
    Next i

    Dim excelApp As Object
    Dim workbook As Object
    Dim worksheet As Object

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    Set workbook = excelApp.Workbooks.Add
    Set worksheet = workbook.Worksheets(1)

    worksheet.Cells(1, 1).Value = "Date"
    worksheet.Cells(1, 2).Value = "Email Count"

    Dim rowIndex As Long
    rowIndex = 2

    Dim dateKey As Variant
    For Each dateKey In dateCounts.Keys
        worksheet.Cells(rowIndex, 1).Value = dateKey
        worksheet.Cells(rowIndex, 2).Value = dateCounts(dateKey)
        rowIndex = rowIndex + 1
    Next dateKey

    worksheet.Columns(1).AutoFit
    worksheet.Columns(2).AutoFit

    worksheet.Cells(1, 4).Value = "Top Sender"
    worksheet.Cells(1, 5).Value = "Email Count"

    Dim topSender As String
    Dim topCount As Long
    Dim senderKey As Variant

    topSender = "(None)"
    topCount = 0

    For Each senderKey In senderCounts.Keys
        If senderCounts(senderKey) > topCount Then
            topCount = senderCounts(senderKey)
            topSender = senderKey
        End If
    Next senderKey

    worksheet.Cells(2, 4).Value = topSender
    worksheet.Cells(2, 5).Value = topCount
    worksheet.Columns(4).AutoFit
    worksheet.Columns(5).AutoFit

    MsgBox "Report generated for " & Format(startDate, "yyyy-mm-dd") & " to " & Format(endDate, "yyyy-mm-dd") & ".", vbInformation
    Exit Sub

DateError:
    MsgBox "Invalid date format. Please enter dates as YYYY-MM-DD.", vbExclamation
End Sub
