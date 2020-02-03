VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#Const OSType = "Win"
' #Const OSType = "Mac"

Dim wordApp As Word.Application

'Private Sub Workbook_Open()
'Dim timesheet_InternalUsage As Worksheet
'Dim timesheet_Data As Worksheet
'Dim dictionaryData As String
'Dim noValue As String
'
'Set timesheet_InternalUsage = ThisWorkbook.Sheets("Internal_usage_1")
'dictionaryData = timesheet_InternalUsage.Cells(1083, 2).Text
'noValue = timesheet_InternalUsage.Cells(1302, "B").Text
'Set timesheet_Data = ThisWorkbook.Sheets(dictionaryData)
'
'
'timesheet_Data.Cells(205, "B").Value = noValue
'timesheet_Data.Cells(206, "B").Value = noValue
'timesheet_Data.Cells(207, "B").Value = noValue
'timesheet_Data.Cells(209, "B").Value = noValue
'timesheet_Data.Cells(210, "B").Value = noValue
'timesheet_Data.Cells(211, "B").Value = noValue
'timesheet_Data.Cells(212, "B").Value = noValue
'timesheet_Data.Cells(213, "B").Value = noValue
'timesheet_Data.Cells(214, "B").Value = noValue
'timesheet_Data.Cells(215, "B").Value = noValue
'timesheet_Data.Cells(216, "B").Value = noValue
'timesheet_Data.Cells(217, "B").Value = noValue
'
'End Sub

Sub GenerateReports()

Dim result As String
Dim dirToResult As String, saveDir As String
Dim selectionRange As Range
Dim excelDCSheet As Worksheet
Dim timesheet_Data As Worksheet
Dim excelTimeReportingSheet As Worksheet
Dim wdDoc As Word.Document
Dim RangeEndOfDoc As Word.Range
Dim i As Integer
Dim i1 As Integer
Dim i2 As Integer
Dim j As Integer
Dim j1 As Integer
Dim intX As Integer
Dim intY As Integer
Dim docTemplateFileName As String
Dim replaceVarValue As String
Dim stringReplaced As String
Dim myStoryRange As Variant
Dim deviceSite As String
Dim hostname As String
Dim tblRemoteConnection As Table
Dim blnFound As Boolean
Dim writtenSum As String
Dim surname As String
Dim fullname As String
Dim instance As Excel.Application
Dim invoiceExcel As Workbook
Dim servicesHTML As String
Dim sumHours As Double
Dim lastDayOfMonth As String
Dim year As String
Dim month As String
Dim monthWritten As String
Dim additionalExpences As String
Dim employeeName As String
Dim bt1_start As String
Dim bt1_end As String
Dim bt1_summ As String
Dim bt2_start As String
Dim bt2_end As String
Dim bt2_summ As String
Dim bt3_start As String
Dim bt3_end As String
Dim bt3_summ As String
Dim agreementNumber As String
Dim titleResponsibilities As String
Dim service_summ As String
Dim currentVersion As String
Dim dictionaryData As String
Dim dictionaryService As String
Dim detectedOS As String
Dim fsSeparator As String
Dim emailTo As String
Dim emailSubj As String
Dim emailBody As String
Dim emailAttach1 As String
Dim emailAttach2 As String
Dim emailAttach3 As String
Dim uuid_ts As String
Dim uuid_act As String
Dim uuid_invoice As String
Dim stub As String

Randomize

currentVersion = ""
dirToTemplates = ""
dirToResult = ""
uuid_ts = GetUUID()
uuid_act = GetUUID()
uuid_invoice = GetUUID()
' --- should not been used
dirToRunConfigs = ""

#If (OSType = "Win") Then
    fsSeparator = "\"
#ElseIf (OSType = "Mac") Then
    fsSeparator = ":"
#End If

If (SelfCheck() = True) Then
' --- All questions "NA" or "yes"
    
    saveDir = ActiveWorkbook.Path & fsSeparator & Format(Date, "dd.mm.yyyy") & "_" & Format(Time, "hh-mm-ss") & fsSeparator
    MkDir (saveDir)
    
    Set timesheet_InternalUsage = ThisWorkbook.Sheets("Internal_usage_1")
    
    dictionaryData = timesheet_InternalUsage.Cells(1083, 2).Text
    dictionaryService = timesheet_InternalUsage.Cells(1084, 2).Text
    
    
    Set timesheet_Data = ThisWorkbook.Sheets(dictionaryData)
    Set excelTimeReportingSheet = ThisWorkbook.Sheets(dictionaryService)
    
    currentVersion = timesheet_Data.Cells(170, "B").Value
    
    servicesHTML = ""
    For i2 = 5 To 48
        sumHours = 0
        For i1 = 3 To 50
            If IsNumeric(excelTimeReportingSheet.Cells(i2, i1).Value) Then
                sumHours = sumHours + excelTimeReportingSheet.Cells(i2, i1).Value
            End If
        Next i1
        If (sumHours > 0) Then
            servicesHTML = servicesHTML & "<br>* " & excelTimeReportingSheet.Cells(i2, 2).Text
        End If
    Next i2
    
breakCycle:
    sumHours = 0
    
    employeeName = timesheet_Data.Cells(3, 2).Text
    If (employeeName = "Name") Then
        MsgBox ("Choose your name")
        Exit Sub
    End If
    
    year = timesheet_Data.Cells(5, 2).Text
    month = timesheet_Data.Cells(56, 2).Text
    monthWritten = timesheet_Data.Cells(52, 2).Text
    lastDayOfMonth = timesheet_Data.Cells(51, 4).Text
    additionalExpences = timesheet_Data.Cells(6, 2).Text
    bt1_start = timesheet_Data.Cells(151, 2).Text
    bt1_end = timesheet_Data.Cells(152, 2).Text
    bt1_summ = timesheet_Data.Cells(11, 2).Text
    bt2_start = timesheet_Data.Cells(153, 2).Text
    bt2_end = timesheet_Data.Cells(154, 2).Text
    bt2_summ = timesheet_Data.Cells(12, 2).Text
    bt3_start = timesheet_Data.Cells(155, 2).Text
    bt3_end = timesheet_Data.Cells(156, 2).Text
    bt3_summ = timesheet_Data.Cells(13, 2).Text
    service_summ = timesheet_Data.Cells(10, 2).Text
    agreementNumber = Application.WorksheetFunction.VLookup(employeeName, timesheet_Data.Range("persons_table"), 5, False)
    titleResponsibilities = Application.WorksheetFunction.VLookup(employeeName, timesheet_Data.Range("persons_table"), 6, False)
    
    timesheet_Data.Unprotect ("flint")
    excelTimeReportingSheet.Unprotect ("flint")
    
    ' --- excelTimeReportingSheet.Rows(66).PageBreak = xlPageBreakManual
    
    ' --- generate qr-code and save it in TimeSheet
    timesheet_Data.Cells(172, "B").Value = uuid_ts
    timesheet_Data.Cells(173, "B").Value = uuid_act
    timesheet_Data.Cells(174, "B").Value = uuid_invoice
    stub = EncodeBarcode(2, "B69", "http://checkuuid.bestbounty.ru/check_uuid/" & uuid_ts, 51, 1)
    
    Set wordApp = CreateObject("Word.Application")
    ThisWorkbook.Sheets(dictionaryData).Activate
    
    writtenSumm = timesheet_Data.Cells(50, 2).Text
    fullname = timesheet_Data.Cells(3, 2).Text
    surname = timesheet_Data.Cells(3, 2).Text
    surname = Mid(surname, 1, InStr(1, surname, " ", vbTextCompare) - 1)
    
    
    On Error GoTo ErrorOpeningAct
    Set wdDoc = wordApp.Documents.Open(ActiveWorkbook.Path & fsSeparator & "Act_template_" & currentVersion & ".doc", True, False)
    On Error Resume Next
    
    wordApp.Visible = False
    wordApp.ScreenUpdating = False
    wdDoc.SaveAs (saveDir & "Act_" & surname & "_" & year & month & lastDayOfMonth & ".doc")
                
    result = ReplaceText("day", timesheet_Data.Cells(51, 4).Text)
    result = ReplaceText("month", monthWritten)
    result = ReplaceText("year", year)
    result = ReplaceText("first_day_of_month", "1 " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text)
    result = ReplaceText("last_day_of_month", lastDayOfMonth & " " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text)
    result = ReplaceText("total", timesheet_Data.Cells(165, 2).Text)
    result = ReplaceText("total_written", timesheet_Data.Cells(169, 2).Text)
    result = ReplaceText("total_rub", timesheet_Data.Cells(165, 3).Text)
    result = ReplaceText("total_coop", timesheet_Data.Cells(165, 5).Text)
    result = ReplaceText("act", timesheet_Data.Cells(9, 2).Text)
    result = ReplaceText("service", service_summ)
    result = ReplaceText("agreement", "N" + agreementNumber + titleResponsibilities)
    
    result = ReplaceText("bt1_start", bt1_start)
    result = ReplaceText("bt1_end", bt1_end)
    result = ReplaceText("bt1_summ", bt1_summ)
    result = ReplaceText("bt2_start", bt2_start)
    result = ReplaceText("bt2_end", bt2_end)
    result = ReplaceText("bt2_summ", bt2_summ)
    result = ReplaceText("bt3_start", bt3_start)
    result = ReplaceText("bt3_end", bt3_end)
    result = ReplaceText("bt3_summ", bt3_summ)
    
    If (bt3_summ = 0) Then
        wdDoc.Tables(1).Rows(5).Delete
    End If
    If (bt2_summ = 0) Then
        wdDoc.Tables(1).Rows(4).Delete
    End If
    If (bt1_summ = 0) Then
        wdDoc.Tables(1).Rows(3).Delete
    End If
    
    ' --- paste qr-code
    stub = EncodeBarcode(7, "$A$10", "http://checkuuid.bestbounty.ru/check_uuid/" & uuid_act, 51, 1)
    Set RangeEndOfDoc = wdDoc.Content
    RangeEndOfDoc.Collapse Direction:=wdCollapseEnd
    RangeEndOfDoc.InsertBreak Type:=wdPageBreak
    RangeEndOfDoc.Paste
    
    wdDoc.Close (True)
    wordApp.Quit (True)
    Set wordApp = Nothing
    
    ' --- EncodeBarcode placed here because after opening new "Excel.App" it gonna be active
    ' --- Inside EncodeBarcode there is call to ActiveWorkbook, which have to be Timesheet
    ' --- Summary: EncodeBarcode have to be called before opening Invoice
    stub = EncodeBarcode(7, "$A$30", "http://checkuuid.bestbounty.ru/check_uuid/" & uuid_invoice, 51, 1)
    
    Dim Proceed As Boolean
    Dim xlApp As Excel.Application
    ' Dim xlWorkBooks As Excel.Workbooks
    Dim xlWorkBook As Excel.Workbook
    Dim invoice_InputData As Excel.Worksheet
    Dim xlCells As Excel.Range
    
    ' Set xlApp = New Excel.Application
    Set xlApp = GetObject(, "Excel.Application")
    xlApp.DisplayAlerts = False
    Set xlWorkBooks = xlApp.Workbooks
    
    On Error GoTo ErrorOpeningInvoice
    Set xlWorkBook = xlWorkBooks.Open(ActiveWorkbook.Path & fsSeparator & "Invoice_template_" & currentVersion & ".xls", True, False)
    On Error Resume Next
    
    ' xlApp.Visible = False
    xlWorkBook.SaveAs (saveDir & "Invoice_" & surname & "_" & year & month & lastDayOfMonth & ".xls")
    Set invoice_InputData = xlWorkBook.Sheets("input_data")
    Set invoice_TDSheet = xlWorkBook.Sheets("TDSheet")
    
    ' Start period
    invoice_InputData.Cells(1, 2).Value = "1 " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text
    ' End period
    invoice_InputData.Cells(2, 2).Value = timesheet_Data.Cells(51, 4).Text & " " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text
    ' Rub
    invoice_InputData.Cells(4, 2).Value = timesheet_Data.Cells(165, 3).Value
    ' Coop
    invoice_InputData.Cells(5, 2).Value = Round(timesheet_Data.Cells(165, 5).Value, 0)
    ' Total written
    invoice_InputData.Cells(6, 2).Value = timesheet_Data.Cells(169, 2).Text
    ' Act number
    invoice_InputData.Cells(7, 2).Value = timesheet_Data.Cells(9, 2).Text
    ' Agreement number
    invoice_InputData.Cells(15, 2).Value = agreementNumber
    ' Title responsibilities
    invoice_InputData.Cells(16, 2).Value = titleResponsibilities
    ' Last day of month
    invoice_InputData.Cells(8, 2).Value = timesheet_Data.Cells(51, 4).Text & " " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text
    ' Service
    invoice_InputData.Cells(10, 4).Value = service_summ
    ' Business trip 1 start
    invoice_InputData.Cells(11, 2).Value = bt1_start
    ' Business trip 1 end
    invoice_InputData.Cells(11, 3).Value = bt1_end
    ' Business trip 1 summ
    invoice_InputData.Cells(11, 4).Value = bt1_summ
    ' Business trip 2 start
    invoice_InputData.Cells(12, 2).Value = bt2_start
    ' Business trip 2 end
    invoice_InputData.Cells(12, 3).Value = bt2_end
    ' Business trip 2 summ
    invoice_InputData.Cells(12, 4).Value = bt2_summ
    ' Business trip 3 start
    invoice_InputData.Cells(13, 2).Value = bt3_start
    ' Business trip 3 end
    invoice_InputData.Cells(13, 3).Value = bt3_end
    ' Business trip 3 summ
    invoice_InputData.Cells(13, 4).Value = bt3_summ
    
    If (bt3_summ = 0) Then
     invoice_TDSheet.Rows("26:26").EntireRow.Hidden = True
    End If
    If (bt2_summ = 0) Then
     invoice_TDSheet.Rows("25:25").EntireRow.Hidden = True
    End If
    If (bt1_summ = 0) Then
     invoice_TDSheet.Rows("24:24").EntireRow.Hidden = True
    End If
    
    ' --- Insert page break to get correct print layout
    ' --- QR-code have to be printed on the 2-nd sheet
    invoice_TDSheet.Rows(45).PageBreak = xlPageBreakManual
    
    ' --- Paste QR-code
    invoice_TDSheet.Range("A48").Select
    invoice_TDSheet.Paste
    ' stub = EncodeBarcode(1, "$A$48", "http://checkuuid.bestbounty.ru/check_uuid/" & uuid_invoice, 51, 1)
    
    ' --- scroll page all the way up
    invoice_TDSheet.Range("A1").Select
    
    ' --- print page properties
    With invoice_TDSheet.PageSetup
     .Orientation = xlPortrait
     .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = False
    End With
    
    ' --- Save and Close Invoice
    xlWorkBook.Close (True)
    
    ' --- print page properties
    With excelTimeReportingSheet.PageSetup
     .Orientation = xlLandscape
     .Zoom = False
     .FitToPagesWide = 1
     .FitToPagesTall = False
    End With
    
    ' --- Save and Close Timesheet
    excelTimeReportingSheet.Protect ("flint")
    timesheet_Data.Protect ("flint")
    ActiveWorkbook.SaveCopyAs (saveDir & "Timesheet_" & surname & "_" & year & month & lastDayOfMonth & ".xlsm")
    
    emailTo = "finance_rus@flintmail.com"
    emailSubj = "Invoice_Act_Timesheet: " & fullname & " îò " & timesheet_Data.Cells(51, 4).Text & " " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text
    emailAttach1 = saveDir & "Invoice_" & surname & "_" & year & month & lastDayOfMonth & ".xls"
    emailAttach2 = saveDir & "Act_" & surname & "_" & year & month & lastDayOfMonth & ".doc"
    emailAttach3 = saveDir & "Timesheet_" & surname & "_" & year & month & lastDayOfMonth & ".xlsm"
    emailBody = timesheet_InternalUsage.Cells(1085, 2).Text _
                & "1 " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text & timesheet_InternalUsage.Cells(1086, 2).Text _
                & timesheet_Data.Cells(51, 4).Text & " " & timesheet_Data.Cells(52, 2).Text & " " & timesheet_Data.Cells(5, 2).Text _
                & timesheet_InternalUsage.Cells(1087, 2).Text & timesheet_Data.Cells(23, 2).Text & timesheet_InternalUsage.Cells(1088, 2).Text & timesheet_Data.Cells(17, 2).Text & ")<br>"
    If (bt1_summ <> 0) Then
    emailBody = emailBody & timesheet_InternalUsage.Cells(1089, 2).Text & " 1 c " & bt1_start & timesheet_InternalUsage.Cells(1086, 2).Text & bt1_end & timesheet_InternalUsage.Cells(1087, 2).Text & bt1_summ & timesheet_InternalUsage.Cells(1090, 2).Text & " <br>"
    End If
    If (bt2_summ <> 0) Then
    emailBody = emailBody & timesheet_InternalUsage.Cells(1089, 2).Text & " 2 c " & bt2_start & timesheet_InternalUsage.Cells(1086, 2).Text & bt2_end & timesheet_InternalUsage.Cells(1087, 2).Text & bt2_summ & timesheet_InternalUsage.Cells(1090, 2).Text & " <br>"
    End If
    If (bt3_summ <> 0) Then
    emailBody = emailBody & timesheet_InternalUsage.Cells(1089, 2).Text & " 3 c " & bt3_start & timesheet_InternalUsage.Cells(1086, 2).Text & bt3_end & timesheet_InternalUsage.Cells(1087, 2).Text & bt3_summ & timesheet_InternalUsage.Cells(1090, 2).Text & " <br>"
    End If
    emailBody = emailBody & "<br>" & timesheet_InternalUsage.Cells(1091, 2).Text & servicesHTML _
                & "<br><br> " & fullname
    
    
    #If (OSType = "Win") Then
    
    Dim strFile As String
    Dim OutApp As Outlook.Application
    Dim objOutlookMsg As Outlook.MailItem
    Dim objOutlookRecip As Recipient
    Dim Recipients As Recipients
    
    Set OutApp = CreateObject("Outlook.Application")
    Set objOutlookMsg = OutApp.CreateItem(olMailItem)
    
    Set Recipients = objOutlookMsg.Recipients
    Set objOutlookRecip = Recipients.Add(emailTo)
    objOutlookRecip.Type = 1
    
    With objOutlookMsg
    '.SentOnBehalfOfName = "sales@domain.com"
    .Subject = emailSubj
    .HTMLBody = emailBody
    
    'Resolve each Recipient's name.
    For Each objOutlookRecip In objOutlookMsg.Recipients
    objOutlookRecip.Resolve
    Next
    
    .Attachments.Add emailAttach1
    .Attachments.Add emailAttach2
    .Attachments.Add emailAttach3
    .Display
    End With
    
    'objOutlookMsg.Send
    
    Set OutApp = Nothing
    #End If
    
    #If (OSType = "Mac") Then
         MailFromMacwithOutlook bodycontent:=emailBody, _
                        mailsubject:=emailSubj, _
                        toaddress:=emailTo, _
                        ccaddress:="", _
                        bccaddress:="", _
                        attachment:=emailAttach1, _
                        displaymail:=True
    
    #End If
    
    ThisWorkbook.Close
    xlApp.Quit

End If '---  If (SelfCheck() = True)

GoTo EndGeneratingDocSet

ErrorOpeningAct:
    MsgBox ("Error opening Act template")
    GoTo EndGeneratingDocSet
ErrorOpeningInvoice:
    MsgBox ("Error opening Invoice template")
    GoTo EndGeneratingDocSet

EndGeneratingDocSet:
End Sub



Function SelfCheck() As Boolean
    Dim timesheet_InternalUsage As Worksheet
    Dim timesheet_Data As Worksheet
    Dim noValue As String
    Dim naValue As String
    Dim dictionaryData As String
    
    SelfCheck = True
    
    Set timesheet_InternalUsage = ThisWorkbook.Sheets("Internal_usage_1")
    dictionaryData = timesheet_InternalUsage.Cells(1083, 2).Text
    noValue = timesheet_InternalUsage.Cells(1302, "B").Text
    naValue = timesheet_InternalUsage.Cells(1303, "B").Text
    Set timesheet_Data = ThisWorkbook.Sheets(dictionaryData)
    
    If ((timesheet_Data.Cells(11, "B").Value + timesheet_Data.Cells(12, "B").Value + timesheet_Data.Cells(13, "B").Value) = 0) Then
        timesheet_Data.Range("B209:B217") = naValue
    End If
    
    If (timesheet_Data.Cells(205, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1302, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(206, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1303, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(207, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1304, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(209, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1306, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(210, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1307, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(211, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1308, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(212, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1309, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(213, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1310, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(214, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1311, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(215, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1312, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(216, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1313, "E")
        SelfCheck = False
    ElseIf (timesheet_Data.Cells(217, "B").Value = noValue) Then
        MsgBox timesheet_InternalUsage.Cells(1314, "E")
        SelfCheck = False
    End If
    
End Function


Function ReplaceText(findStr As String, replaceStr As String) As String

SearchAgain:    With wordApp.ActiveDocument.Content.Find
                    .ClearFormatting
                    .Text = "<<" + findStr + ">>"
                    .Wrap = wdFindContinue
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Replacement.Highlight = False
                    .Replacement.Text = replaceStr
                    .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
'                    If .Execute Then
'                       wordapp.Selection.Range.Text = replaceStr
'                       GoTo SearchAgain
'                    End If
                End With
End Function



Function ImportTextFile(strFile As String) As String

    If Dir(strFile) <> "" Then
        Open strFile For Input As #1
        ImportTextFile = Input$(LOF(1), 1)
        Close #1
    End If

End Function


#If (OSType = "Mac") Then
Function MailFromMacwithOutlook(bodycontent As String, mailsubject As String, _
            toaddress As String, ccaddress As String, bccaddress As String, _
            attachment As String, displaymail As Boolean)



    Dim scriptToRun As String

    scriptToRun = scriptToRun & "tell application " & _
                  Chr(34) & "Microsoft Outlook" & Chr(34) & Chr(13)

    scriptToRun = scriptToRun & _
     "set NewMail to make new outgoing message with properties" & _
       "{content:""" & bodycontent & """, subject:""" & mailsubject & """}" & Chr(13)

    If toaddress <> "" Then scriptToRun = scriptToRun & _
       "make new to recipient at NewMail with properties" & _
       "{email address:{address:""" & toaddress & """}}" & Chr(13)

    If ccaddress <> "" Then scriptToRun = scriptToRun & _
       "make new cc recipient at NewMail with properties" & _
       "{email address:{address:""" & ccaddress & """}}" & Chr(13)

    If bccaddress <> "" Then scriptToRun = scriptToRun & _
       "make new bcc recipient at NewMail with properties" & _
       "{email address:{address:""" & bccaddress & """}}" & Chr(13)

    If attachment <> "" Then
     scriptToRun = scriptToRun & "make new attachment at NewMail with properties" & _
              "{file:""" & attachment & """ as alias}" & Chr(13)
    End If

    If displaymail = False Then
        scriptToRun = scriptToRun & "send NewMail" & Chr(13)
    Else
        scriptToRun = scriptToRun & "open NewMail" & Chr(13)
    End If
    scriptToRun = scriptToRun & "end tell" & Chr(13)

    If Len(toaddress) + Len(ccaddress) + Len(bccaddress) = 0 Or mailsubject = "" Then
        MsgBox "There is no To, CC or BCC address or Subject for this mail"
        Exit Function
    Else
        On Error Resume Next
        MacScript (scriptToRun)
        On Error GoTo 0
    End If
End Function

#End If


Sub DirectoryListing()

Dim varDirectory As Variant
Dim flag As Boolean
Dim i As Integer
Dim strDirectory As String

strDirectory = ActiveWorkbook.Path & "\*.xlsm"
i = 1
flag = True
varDirectory = Dir(strDirectory, vbNormal)

While flag = True
    If varDirectory = "" Then
        flag = False
    Else
        Debug.Print strDirectory + varDirectory
        'returns the next file or directory in the path
        varDirectory = Dir
        i = i + 1
    End If
Wend

End Sub

Function GetUUID() As String
    Dim result As String
    
    result = "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxxF"
    
    For i = 1 To Len(result)
        rand = Fix(Rnd() * 16)
        If (Mid(result, i, 1) = "x") Then
            Mid(result, i, 1) = Hex(rand)
        ElseIf (Mid(result, i, 1) = "y") Then
            Mid(result, i, 1) = Hex((rand And 3) Or 8)
        End If
    Next i
    
    GetUUID = result
End Function
