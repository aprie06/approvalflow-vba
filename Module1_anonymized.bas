Attribute VB_Name = "Module1"
' ==========================================
' APPROVALFLOW - OPTIMIZED VBA MODULE
' ==========================================
' Author: Alexis Prieto
' Last Updated: April 29 2026
' ==========================================
Option Explicit
Private overridePayPeriodDate As Date  ' Set to 0 to use current date


' ==========================================
' GLOBAL VARIABLES FOR IN-MEMORY LOOKUP
' ==========================================
Dim lookupData As Object

' ==========================================
' LOOKUP DATA FUNCTIONS
' ==========================================

Sub LoadLookupDataIntoMemory()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim sourceSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim acesID As String
    Dim studentEmail As String
    Dim studentName As String
    Dim supervisorEmail As String
    Dim supervisorName As String
    Dim employerName As String
    Dim studentCCAddr As String
    
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Initialize dictionary
    Set lookupData = CreateObject("Scripting.Dictionary")
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set sourceSheet = excelWorkbook.Sheets("Updated_Supervisor email")
    
    lastRow = sourceSheet.cells(sourceSheet.rows.count, 2).End(-4162).Row
    
    For i = 2 To lastRow
        ' Column structure: A=Banner ID, B=Employee Email @district.edu, C=Student Name,
        ' D=ACES ID, E=College, F=Employer, G=Supervisor Name, H=Supervisor Email,
        ' I=Student Email @student.district.edu (CC), J1=Pay Period Anchor
        studentEmail = LCase(Trim(sourceSheet.cells(i, 2).Value))
        studentName = Trim(sourceSheet.cells(i, 3).Value)
        acesID = LCase(Trim(sourceSheet.cells(i, 4).Value))
        employerName = Trim(sourceSheet.cells(i, 6).Value)
        supervisorName = Trim(sourceSheet.cells(i, 7).Value)
        supervisorEmail = LCase(Trim(sourceSheet.cells(i, 8).Value))
        studentCCAddr = LCase(Trim(sourceSheet.cells(i, 9).Value))
        
        ' Skip empty rows
        If studentEmail <> "" And studentName <> "" Then
            ' Store by ACES ID
            If acesID <> "" And Not lookupData.Exists(acesID) Then
                lookupData.Add acesID, Array(studentName, supervisorEmail, supervisorName, employerName, studentCCAddr)
            End If
            
            ' Store by full email
            If Not lookupData.Exists(studentEmail) Then
                lookupData.Add studentEmail, Array(studentName, supervisorEmail, supervisorName, employerName, studentCCAddr)
            End If
            
            ' Store by supervisor email for FastGetSupervisorName and FastGetEmployer
            If supervisorEmail <> "" And Not lookupData.Exists("sup_" & supervisorEmail) Then
                lookupData.Add "sup_" & supervisorEmail, Array(supervisorName, employerName)
            End If
        End If
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading lookup data: " & Err.Description, vbCritical
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub

Sub ClearLookupData()
    Set lookupData = Nothing
End Sub


' ==========================================
' EMAIL EXTRACTION FUNCTIONS
' ==========================================

'Function ExtractEmailAddress(ByVal inputString As String) As String
    'Dim regEx As Object
    'Dim matches As Object
    
    'On Error GoTo SimpleExtract
    
    'Set regEx = CreateObject("VBScript.RegExp")
    'regEx.Pattern = "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
    'regEx.Global = False
    'regEx.IgnoreCase = True
    
    'If regEx.Test(inputString) Then
        'Set matches = regEx.Execute(inputString)
        'ExtractEmailAddress = matches(0).Value
    'Else
        'ExtractEmailAddress = Trim(inputString)
    'End If
    'Exit Function
    
'SimpleExtract:
    ' Fallback: just return trimmed input
    'ExtractEmailAddress = Trim(inputString)
'End Function


Function DetermineResponseType(ByVal emailBody As String) As String
    Dim upperBody As String
    Dim cleanBody As String
    
    cleanBody = emailBody

Dim cutPos As Long

' Strip at common reply separators
cutPos = InStr(emailBody, "From:")
If cutPos > 20 Then cleanBody = Left(emailBody, cutPos - 1)

cutPos = InStr(cleanBody, "________________________________")
If cutPos > 1 Then cleanBody = Left(cleanBody, cutPos - 1)

cutPos = InStr(cleanBody, "-----Original Message-----")
If cutPos > 1 Then cleanBody = Left(cleanBody, cutPos - 1)

cutPos = InStr(cleanBody, "Sent:")
If cutPos > 20 Then cleanBody = Left(cleanBody, cutPos - 1)

cutPos = InStr(cleanBody, "On ")
If cutPos > 5 Then cleanBody = Left(cleanBody, cutPos - 1)
    upperBody = UCase(cleanBody)
    
    If InStr(upperBody, "NOT APPROVED") > 0 Then
        DetermineResponseType = "REJECTED"
    ElseIf InStr(upperBody, "REJECT") > 0 Then
        DetermineResponseType = "REJECTED"
    ElseIf InStr(upperBody, "DENIED") > 0 Then
        DetermineResponseType = "REJECTED"
    ElseIf InStr(upperBody, "APPROVE") > 0 Then
        DetermineResponseType = "APPROVED"
    ElseIf InStr(upperBody, "LOOKS GOOD") > 0 Then
        DetermineResponseType = "APPROVED"
    ElseIf InStr(upperBody, "CONFIRMED") > 0 Then
        DetermineResponseType = "APPROVED"
    ElseIf InStr(upperBody, "OK") > 0 And Len(Trim(cleanBody)) < 50 Then
        DetermineResponseType = "APPROVED"
    Else
        DetermineResponseType = "PENDING REVIEW"
    End If
End Function


' ==========================================
' STUDENT NAME EXTRACTION FUNCTIONS
' ==========================================
Function ExtractStudentNameFromReplySubject(ByVal subject As String) As String
    Dim cleanSubject As String
    Dim dashPos As Long
    Dim colonPos As Long
    Dim forPos As Long
    Dim studentName As String
    
    ' Remove RE: FW: [EXTERNAL] etc.
    cleanSubject = subject
    cleanSubject = Replace(cleanSubject, "RE:", "", , , vbTextCompare)
    cleanSubject = Replace(cleanSubject, "Re:", "", , , vbTextCompare)
    cleanSubject = Replace(cleanSubject, "FW:", "", , , vbTextCompare)
    cleanSubject = Replace(cleanSubject, "Fw:", "", , , vbTextCompare)
    cleanSubject = Replace(cleanSubject, "[EXTERNAL]", "", , , vbTextCompare)
    cleanSubject = Trim(cleanSubject)
    
    ' Pattern: "Timesheet Approval Request for Student Name"
    forPos = InStr(1, cleanSubject, " for ", vbTextCompare)
    If forPos > 0 Then
        studentName = Trim(Mid(cleanSubject, forPos + 5))
        Dim parenPos As Long
        parenPos = InStr(studentName, " (")
        If parenPos > 0 Then studentName = Trim(Left(studentName, parenPos - 1))
        ExtractStudentNameFromReplySubject = studentName
        Exit Function
    End If
    
    ' Fallback: Pattern "Timesheet Approval - Student Name"
    dashPos = InStr(cleanSubject, " - ")
    If dashPos > 0 Then
        studentName = Trim(Mid(cleanSubject, dashPos + 3))
        parenPos = InStr(studentName, " (")
        If parenPos > 0 Then studentName = Trim(Left(studentName, parenPos - 1))
        ExtractStudentNameFromReplySubject = studentName
        Exit Function
    End If
    
    ' Fallback: Pattern "INTERN_PROGRAM: Student Name"
    colonPos = InStr(cleanSubject, ": ")
    If colonPos > 0 Then
        studentName = Trim(Mid(cleanSubject, colonPos + 2))
        ExtractStudentNameFromReplySubject = studentName
        Exit Function
    End If
    
    ExtractStudentNameFromReplySubject = ""
End Function


' ==========================================
' EXTRACT STUDENT NAME FROM EMAIL BODY
' Looks for student name in the forwarded content
' ==========================================

Function ExtractStudentNameFromBody(emailBody As String) As String
    Dim lines() As String
    Dim i As Long
    Dim line As String
    Dim startPos As Long
    
    ExtractStudentNameFromBody = ""
    
    ' Look for common patterns in email body
    
    ' Pattern: "Student: [Name]" or "Student Name: [Name]"
    startPos = InStr(1, emailBody, "Student:", vbTextCompare)
    If startPos > 0 Then
        Dim afterStudent As String
        afterStudent = Mid(emailBody, startPos + 8, 100)
        
        ' Get text until newline or
        Dim endPos As Long
        endPos = InStr(afterStudent, vbCrLf)
        If endPos = 0 Then endPos = InStr(afterStudent, vbLf)
        If endPos = 0 Then endPos = InStr(afterStudent, "<")
        If endPos = 0 Then endPos = 50
        
        ExtractStudentNameFromBody = Trim(Left(afterStudent, endPos - 1))
        
        ' Clean up
        ExtractStudentNameFromBody = Replace(ExtractStudentNameFromBody, "Name:", "")
        ExtractStudentNameFromBody = Trim(ExtractStudentNameFromBody)
        
        If Len(ExtractStudentNameFromBody) > 2 And Len(ExtractStudentNameFromBody) < 50 Then
            Exit Function
        Else
            ExtractStudentNameFromBody = ""
        End If
    End If
    
    ' Pattern: Look for name in table cell after "Name" header
    startPos = InStr(1, emailBody, ">Name<", vbTextCompare)
    If startPos > 0 Then
        ' Find next TD content
        Dim tdStart As Long
        Dim tdEnd As Long
        tdStart = InStr(startPos, emailBody, "<td")
        If tdStart > 0 Then
            tdStart = InStr(tdStart, emailBody, ">") + 1
            tdEnd = InStr(tdStart, emailBody, "</td>")
            If tdEnd > tdStart Then
                ExtractStudentNameFromBody = Trim(Mid(emailBody, tdStart, tdEnd - tdStart))
                ExtractStudentNameFromBody = StripHTML(ExtractStudentNameFromBody)
                If Len(ExtractStudentNameFromBody) > 2 And Len(ExtractStudentNameFromBody) < 50 Then
                    Exit Function
                Else
                    ExtractStudentNameFromBody = ""
                End If
            End If
        End If
    End If
End Function


' ==========================================
' LOOKUP FUNCTIONS (for fallback when no Sent_Log match)
' ==========================================

Function GetStudentName(ByVal studentEmail As String) As String
    Dim cleanEmail As String
    Dim acesID As String
    Dim data As Variant
    
    cleanEmail = LCase(Trim(studentEmail))
    
    ' Extract ACES ID
    If InStr(cleanEmail, "@") > 0 Then
        acesID = Left(cleanEmail, InStr(cleanEmail, "@") - 1)
    Else
        acesID = cleanEmail
    End If
    
    ' Check lookup data
    If Not lookupData Is Nothing Then
        If lookupData.Exists(acesID) Then
            data = lookupData(acesID)
            GetStudentName = data(0)
            Exit Function
        ElseIf lookupData.Exists(cleanEmail) Then
            data = lookupData(cleanEmail)
            GetStudentName = data(0)
            Exit Function
        End If
    End If
    
    GetStudentName = "Name Not Found"
End Function

Function GetStudentEmailByName(ByVal studentName As String) As String
    Dim key As Variant
    Dim data As Variant
    Dim searchName As String
    
    searchName = LCase(Trim(studentName))
    
    If Not lookupData Is Nothing Then
        For Each key In lookupData.Keys
            data = lookupData(key)
            If InStr(1, LCase(data(0)), searchName, vbTextCompare) > 0 Then
                ' Found a match, return the key (which is the email or ACES ID)
                If InStr(key, "@") > 0 Then
                    GetStudentEmailByName = key
                    Exit Function
                End If
            End If
        Next key
    End If
    
    GetStudentEmailByName = ""
End Function

Function GetSupervisor(ByVal supervisorEmail As String) As String
    Dim key As Variant
    Dim data As Variant
    Dim cleanEmail As String
    
    cleanEmail = LCase(Trim(supervisorEmail))
    
    If Not lookupData Is Nothing Then
        For Each key In lookupData.Keys
            data = lookupData(key)
            If LCase(data(1)) = cleanEmail Then
                GetSupervisor = data(2)
                Exit Function
            End If
        Next key
    End If
    
    GetSupervisor = "Unknown"
End Function

Function GetEmployer(ByVal supervisorEmail As String) As String
    Dim key As Variant
    Dim data As Variant
    Dim cleanEmail As String
    
    cleanEmail = LCase(Trim(supervisorEmail))
    
    If Not lookupData Is Nothing Then
        For Each key In lookupData.Keys
            data = lookupData(key)
            If LCase(data(1)) = cleanEmail Then
                GetEmployer = data(3)
                Exit Function
            End If
        Next key
    End If
    
    GetEmployer = "Unknown"
End Function

' ==========================================
' ARCHIVE FUNCTION
' ==========================================

Sub ArchiveSubmissionData(excelWorkbook As Object, trackingSheet As Object)
    Dim archiveSheet As Object
    Dim lastRowTracking As Long
    Dim lastRowArchive As Long
    
    On Error Resume Next
    Set archiveSheet = excelWorkbook.Sheets("Submission_Archive")
    If archiveSheet Is Nothing Then
        Set archiveSheet = excelWorkbook.Sheets.Add(After:=excelWorkbook.Sheets(excelWorkbook.Sheets.count))
        archiveSheet.Name = "Submission_Archive"
        
        ' Copy headers
        trackingSheet.Range("A1:K1").Copy archiveSheet.Range("A1")
    End If
    On Error GoTo 0
    
    ' Get last rows
    lastRowTracking = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    lastRowArchive = archiveSheet.cells(archiveSheet.rows.count, 2).End(-4162).Row
    If lastRowArchive < 1 Then lastRowArchive = 1
    
    ' Copy data to archive
    If lastRowTracking > 1 Then
        trackingSheet.Range("A2:K" & lastRowTracking).Copy archiveSheet.cells(lastRowArchive + 1, 1)
        
        ' Clear tracking sheet data (keep headers)
        trackingSheet.Range("A2:K" & lastRowTracking).Clear
    End If
End Sub

' ==========================================
' PAY PERIOD DATE FUNCTIONS
' ==========================================
Function GetPayPeriodStartDate() As Date
    Dim targetDate As Date
    Dim currentDay As Long
    Dim currentMonth As Long
    Dim currentYear As Long
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim anchorDate As Date
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Function
    
    If overridePayPeriodDate > 0 Then
        targetDate = overridePayPeriodDate
    Else
        On Error Resume Next
        Set excelApp = CreateObject("Excel.Application")
        excelApp.Visible = False
        excelApp.DisplayAlerts = False
        Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
        anchorDate = excelWorkbook.Sheets("Updated_Supervisor email").Range("J1").Value
        excelWorkbook.Close False
        excelApp.Quit
        On Error GoTo 0
        
        If IsDate(anchorDate) And anchorDate > 0 Then
            targetDate = anchorDate
        Else
            targetDate = Date
        End If
    End If
    
    currentDay = Day(targetDate)
    currentMonth = Month(targetDate)
    currentYear = Year(targetDate)
    
    If currentDay <= 15 Then
        GetPayPeriodStartDate = DateSerial(currentYear, currentMonth, 1)
    Else
        GetPayPeriodStartDate = DateSerial(currentYear, currentMonth, 16)
    End If
End Function

Function GetPayPeriodEndDate() As Date
    Dim targetDate As Date
    Dim currentDay As Long
    Dim currentMonth As Long
    Dim currentYear As Long
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim anchorDate As Date
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Function
    
    If overridePayPeriodDate > 0 Then
        targetDate = overridePayPeriodDate
    Else
        On Error Resume Next
        Set excelApp = CreateObject("Excel.Application")
        excelApp.Visible = False
        excelApp.DisplayAlerts = False
        Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
        anchorDate = excelWorkbook.Sheets("Updated_Supervisor email").Range("J1").Value
        excelWorkbook.Close False
        excelApp.Quit
        On Error GoTo 0
        
        If IsDate(anchorDate) And anchorDate > 0 Then
            targetDate = anchorDate
        Else
            targetDate = Date
        End If
    End If
    
    currentDay = Day(targetDate)
    currentMonth = Month(targetDate)
    currentYear = Year(targetDate)
    
    If currentDay <= 15 Then
        GetPayPeriodEndDate = DateSerial(currentYear, currentMonth, 15)
    Else
        GetPayPeriodEndDate = DateSerial(currentYear, currentMonth + 1, 0)
    End If
End Function

Function GetPayPeriodLabel() As String
    Dim startDate As Date
    Dim endDate As Date
    
    startDate = GetPayPeriodStartDate()
    endDate = GetPayPeriodEndDate()
    
    GetPayPeriodLabel = Format(startDate, "mmm") & " " & Day(startDate) & "-" & Day(endDate) & ", " & Year(endDate)
End Function


' ==========================================
' MAIN: PROCESS TIMESHEET NOTIFICATIONS
' ==========================================
Sub MoveItems()
    Dim ol As Outlook.Application
    Dim myNameSpace As Outlook.NameSpace
    Dim myInbox As Outlook.folder
    Dim myDestFolder As Outlook.folder
    Dim sentFolder As Outlook.folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim htmlDoc As MSHTML.HTMLDocument
    Dim htmlBody As String
    Dim tagelements As MSHTML.IHTMLElementCollection
    Dim tagelement As MSHTML.IHTMLElement
    Dim super_email As String
    Dim studentEmail As String
    Dim forwardItem As Outlook.mailItem
    Dim tempbody As String
    Dim signaturePath As String
    Dim signatureFileName As String
    Dim signature As String
    Dim fso As Object
    Dim signatureFile As Object
    Dim emailArray As Variant
    Dim count As Long
    Dim i As Long
    Dim successCount As Long
    Dim failCount As Long
    
    
    tempbody = "<p>Greetings INTERN_PROGRAM Supervisor,</p>" & _
    "<p>Please review the intern timesheet submissions below and <b>reply</b> with ONE of the following responses.</p>" & _
    "<p><u>- <b>COPY AND PASTE ONE OF THE RESPONSES WHEN YOU REPLY</b></u></p>" & _
    "<div style='background-color:#E8F5E9; padding:15px; margin:15px 0; border-left:5px solid #4CAF50;'>" & _
    "<p style='font-size:16px; margin:5px 0;'><b>Type one of these in your reply:</b></p>" & _
    "<p style='font-size:18px; margin:10px 0;'><b>APPROVED</b> - if hours are correct</p>" & _
    "<p style='font-size:18px; margin:10px 0;'><b>REJECTED</b> - if hours are incorrect</p>" & _
    "<p style='font-size:18px; margin:10px 0;'><b>CORRECTIONS</b> - if you need to provide adjusted hours</p>" & _
    "</div>" & _
    "<p><b>If you type CORRECTIONS, please provide the corrected schedule:</b></p>" & _
    "<p style='background-color:#FFF9C4; padding:10px;'>Example: Monday 1/06/26 8am-4pm, Tuesday 1/07/26 9am-5pm</p>" & _
    "<p>-Important Reminders:</p>" & _
    "<p>- Interns can work up to 20 hours per work week</p>" & _
    "<p>- Work week: Saturday 12:00 am - Friday 11:59 pm</p>" & _
    "<p>- Semi-Monthly pay periods: 1st-15th and 16th-last day of month</p>" & _
    "<p>- Interns <u><b>DO NOT</b></u> receive Holiday pay - only paid for hours worked</p>" & _
    "<p><u><span style='background-color: #FFFF00'>-Charges apply: Overage charges apply for hours beyond the 20-hour weekly limit</span></u></p>" & _
    "<p><u><span style='background-color: #FFFF00'>**URGENT: Please respond within 24 hours to avoid payment delays.</span></u></p>" & _
    "<p><b><span style='background-color: #26F7FD'>Attention INTERN_PROGRAM Interns:</span> For transparency, I'm sharing this so you can advocate for timely supervisor approval.</b></p>"
    
    Set ol = Outlook.Application
    Set myNameSpace = ol.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(6)

On Error Resume Next
Set myDestFolder = myInbox.Folders("Student_Time_Sheet")
Set sentFolder = myInbox.Folders("Sent_super")
    On Error GoTo ErrorHandler
    
    If myDestFolder Is Nothing Then
        MsgBox "ERROR: Student_Time_Sheet folder not found!", vbCritical
        Exit Sub
    End If
    
    If sentFolder Is Nothing Then
        MsgBox "ERROR: Sent_super folder not found!", vbCritical
        Exit Sub
    End If
    
    Set myItems = myDestFolder.Items
    myItems.Sort "[ReceivedTime]", True
    Set htmlDoc = New MSHTML.HTMLDocument
    
    count = 0
    successCount = 0
    failCount = 0
    
    For Each myItem In myItems
        If count >= 20 Then Exit For
        
        If TypeOf myItem Is Outlook.mailItem Then
            Dim oMail As Outlook.mailItem
            Set oMail = myItem
            
            On Error Resume Next
            htmlBody = oMail.htmlBody
            If Err.Number <> 0 Then
                failCount = failCount + 1
                On Error GoTo ErrorHandler
                GoTo NextItem
            End If
            On Error GoTo ErrorHandler
            
            htmlDoc.Body.innerHTML = htmlBody
            Set tagelements = htmlDoc.getElementsByTagName("td")
            
            If tagelements.Length > 5 Then
                Set tagelement = tagelements.Item(5)
                studentEmail = Trim(tagelement.innerText)
                
                super_email = excel_vlooklook(studentEmail)
                
                If super_email <> "NA" And super_email <> "" Then
    emailArray = Split(super_email, ";")
    Set forwardItem = oMail.Forward
    
    For i = LBound(emailArray) To UBound(emailArray)
        Dim recipEmail As String
        recipEmail = Trim(emailArray(i))
        If recipEmail <> "" Then
            forwardItem.Recipients.Add recipEmail
        End If
    Next i
    
    ' CC both employee email and student email
    Dim studentCC As String
    studentCC = FastGetStudentCCEmail(studentEmail)
    If studentCC <> "" Then
        forwardItem.CC = studentEmail & "; " & studentCC
    Else
        forwardItem.CC = studentEmail
    End If
    
    forwardItem.htmlBody = tempbody & htmlBody & signature
    
    On Error Resume Next
    forwardItem.SendUsingAccount = ol.Session.Accounts("DST-INTERNPAYROLL@district.edu")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrorHandler
    
    '''Call excel_log(forwardItem.To, forwardItem.subject, studentEmail, Now'''
    Dim submissionDate As Date
submissionDate = ExtractTimesheetStatusDate(htmlBody)

If submissionDate > 0 Then
    If submissionDate < GetPayPeriodCutoffStart() Or submissionDate > GetPayPeriodCutoffEnd() Then
        failCount = failCount + 1
        GoTo NextItem
    End If
Else
    submissionDate = Now
End If

Call excel_log(forwardItem.To, forwardItem.subject, studentEmail, submissionDate)
    
    
    On Error Resume Next
    forwardItem.Recipients.ResolveAll ''ResolveAll is required after adding recipients due to the Office update applied in April 2026. Check this line first if this ever breaks again after another update''
    forwardItem.Send
    If Err.Number <> 0 Then
        failCount = failCount + 1
        Err.Clear
    Else
        successCount = successCount + 1
        oMail.Move sentFolder
    End If
    On Error GoTo ErrorHandler
Else
    Set forwardItem = oMail.Forward
    forwardItem.subject = forwardItem.subject & " Missing supervisor email"
    forwardItem.Recipients.Add "DST-INTERNPAYROLL@district.edu"
    Call excel_log(forwardItem.To, forwardItem.subject, studentEmail, Now)
    On Error Resume Next
    forwardItem.Send
    If Err.Number = 0 Then
        oMail.Move sentFolder
        failCount = failCount + 1
    End If
    On Error GoTo ErrorHandler
End If
            Else
                failCount = failCount + 1
            End If
            
NextItem:
            count = count + 1
        End If
    Next
    
    MsgBox "Processed " & count & " timesheets" & vbCrLf & _
           "Successful: " & successCount & vbCrLf & _
           "Failed: " & failCount, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


Sub ReadSupervisorRepliesFolder()
    Dim myNameSpace As Outlook.NameSpace
    Dim sharedMailbox As Outlook.folder
    Dim sharedInbox As Outlook.folder
    Dim replyFolder As Outlook.folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim oMail As Outlook.mailItem
    Dim processedCount As Long
    Dim skippedCount As Long
    Dim duplicateCount As Long
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim replySheet As Object
    Dim sentLogSheet As Object
    Dim currentRow As Long
    Dim supervisorEmail As String
    Dim responseType As String
    Dim cleanSupervisorEmail As String
    Dim sentLastRow As Long
    Dim sentRow As Long
    Dim sentDate As Date
    Dim timeDiff As Double
    Dim studentName As String
    Dim supervisorName As String
    Dim employerName As String
    Dim studentEmail As String
    Dim loggedReplies As Object
    Dim replyKey As String
    Dim bestMatchRow As Long
    Dim bestMatchScore As Double
    Dim timeScore As Double
    Dim nameScore As Double
    Dim totalScore As Double
    Dim parsedName As String
    Dim cutoffStart As Date
    Dim cutoffEnd As Date
    Dim excelFilePath As String

    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub

    On Error GoTo ErrorHandler

    cutoffStart = GetPayPeriodCutoffStart()
    cutoffEnd = GetPayPeriodCutoffEnd()

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")

    replySheet.Range("A2:G" & replySheet.rows.count).ClearContents
    currentRow = 2

    sentLastRow = sentLogSheet.cells(sentLogSheet.rows.count, 1).End(-4162).Row

    Set loggedReplies = CreateObject("Scripting.Dictionary")

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set sharedMailbox = myNameSpace.Folders("DST-INTERNPAYROLL@district.edu")
    Set sharedInbox = sharedMailbox.Folders("Inbox")

    On Error Resume Next
    Set replyFolder = sharedInbox.Folders("Supervisor_Replies")
    On Error GoTo ErrorHandler

    If replyFolder Is Nothing Then
        MsgBox "Supervisor_Replies folder not found!", vbCritical
        GoTo Cleanup
    End If

    Set myItems = replyFolder.Items
    myItems.Sort "[ReceivedTime]", True

    processedCount = 0
    skippedCount = 0
    duplicateCount = 0

    For Each myItem In myItems
        If TypeOf myItem Is Outlook.mailItem Then
            Set oMail = myItem

            If oMail.ReceivedTime < cutoffStart Then Exit For
            If oMail.ReceivedTime > cutoffEnd Then GoTo NextReplyItem

            supervisorEmail = GetSMTPAddress(oMail)
            cleanSupervisorEmail = LCase(Trim(supervisorEmail))

            responseType = DetermineResponseType(oMail.Body)

           bestMatchRow = 0
bestMatchScore = -1

' PRIMARY MATCH: Extract intern email from CC field
Dim ccEmail As String
ccEmail = ""
Dim ccAddresses As String
ccAddresses = oMail.CC

If ccAddresses <> "" Then
    Dim ccParts As Variant
    ccParts = Split(ccAddresses, ";")
    Dim ccPart As Variant
    For Each ccPart In ccParts
        Dim trimmedCC As String
        trimmedCC = LCase(Trim(ccPart))
        ' Look for district.edu address in CC - that is the intern
        If InStr(trimmedCC, "@district.edu") > 0 Or InStr(trimmedCC, "@student.district.edu") > 0 Then
            ' Extract just the email address
            If InStr(trimmedCC, "<") > 0 Then
                trimmedCC = Mid(trimmedCC, InStr(trimmedCC, "<") + 1)
                trimmedCC = Left(trimmedCC, InStr(trimmedCC, ">") - 1)
            End If
            ccEmail = Trim(trimmedCC)
            Exit For
        End If
    Next ccPart
End If

' If we found intern email in CC, match directly to Sent_Log
If ccEmail <> "" Then
    Dim ccParsedName As String
    ccParsedName = LCase(ExtractStudentNameFromReplySubject(oMail.subject))
    
    For sentRow = 2 To sentLastRow
        Dim sentStudentEmail As String
        sentStudentEmail = LCase(Trim(sentLogSheet.cells(sentRow, 6).Value))
        If InStr(sentStudentEmail, ccEmail) > 0 Or InStr(ccEmail, sentStudentEmail) > 0 Then
            Dim ccStudentName As String
            ccStudentName = LCase(sentLogSheet.cells(sentRow, 1).Value)
            If ccParsedName = "" Or InStr(ccStudentName, Left(ccParsedName, 4)) > 0 Then
                On Error Resume Next
                sentDate = sentLogSheet.cells(sentRow, 7).Value
                On Error GoTo ErrorHandler
                If IsDate(sentDate) Then
                    timeDiff = oMail.ReceivedTime - sentDate
                    If timeDiff > 0 And timeDiff <= 14 Then
                        bestMatchRow = sentRow
                        bestMatchScore = 99
                        Exit For
                    End If
                End If
            End If
        End If
    Next sentRow
End If

' FALLBACK: If CC match failed, use supervisor email + name scoring
If bestMatchRow = 0 Then
    parsedName = ExtractStudentNameFromReplySubject(oMail.subject)
    If InStr(parsedName, " - ") > 0 Then parsedName = Trim(Left(parsedName, InStr(parsedName, " - ") - 1))

    For sentRow = 2 To sentLastRow
        If InStr(1, LCase(Trim(sentLogSheet.cells(sentRow, 3).Value)), cleanSupervisorEmail, vbTextCompare) > 0 Then

            On Error Resume Next
            sentDate = sentLogSheet.cells(sentRow, 7).Value
            On Error GoTo ErrorHandler

            If IsDate(sentDate) Then
                timeDiff = oMail.ReceivedTime - sentDate

                If timeDiff > 0 And timeDiff <= 14 Then
                    studentName = sentLogSheet.cells(sentRow, 1).Value

                    If studentName <> "Name Not Found" And studentName <> "" Then
                        timeScore = 1 - (timeDiff / 14)
                        nameScore = 0
                        If parsedName <> "" Then
                            If InStr(1, LCase(studentName), LCase(parsedName), vbTextCompare) > 0 Then
                                nameScore = 2
                            Else
                                Dim nameParts As Variant
                                nameParts = Split(LCase(parsedName), " ")
                                If UBound(nameParts) >= 0 Then
                                    If InStr(1, LCase(studentName), nameParts(0), vbTextCompare) > 0 Then
                                        nameScore = 1
                                    End If
                                End If
                            End If
                        End If
                        totalScore = timeScore + nameScore
                        If totalScore > bestMatchScore Then
                            bestMatchScore = totalScore
                            bestMatchRow = sentRow
                        End If
                    End If
                End If
            End If
        End If
    Next sentRow
End If

            If bestMatchRow > 0 Then
                studentName = sentLogSheet.cells(bestMatchRow, 1).Value
                supervisorName = sentLogSheet.cells(bestMatchRow, 2).Value
                employerName = sentLogSheet.cells(bestMatchRow, 4).Value
                studentEmail = sentLogSheet.cells(bestMatchRow, 6).Value

                replyKey = LCase(studentEmail) & "|" & cleanSupervisorEmail & "|" & Format(oMail.ReceivedTime, "yyyymmddhhmmss")

                If Not loggedReplies.Exists(replyKey) Then
                    loggedReplies.Add replyKey, True

                    replySheet.cells(currentRow, 1).Value = studentName
                    replySheet.cells(currentRow, 2).Value = supervisorName
                    replySheet.cells(currentRow, 3).Value = sentLogSheet.cells(bestMatchRow, 3).Value
                    replySheet.cells(currentRow, 4).Value = employerName
                    replySheet.cells(currentRow, 5).Value = oMail.ReceivedTime
                    replySheet.cells(currentRow, 6).Value = responseType
                    replySheet.cells(currentRow, 7).Value = studentEmail & "|" & sentLogSheet.cells(bestMatchRow, 3).Value

                    currentRow = currentRow + 1
                    processedCount = processedCount + 1
                Else
                    duplicateCount = duplicateCount + 1
                End If
            Else
                skippedCount = skippedCount + 1
            End If
        End If

NextReplyItem:
    Next

Cleanup:
    If Not excelWorkbook Is Nothing Then
        excelWorkbook.Save
        excelWorkbook.Close False
    End If
    If Not excelApp Is Nothing Then excelApp.Quit

    MsgBox "Reply_Log rebuilt from Supervisor_Replies folder!" & vbCrLf & _
           "Entries logged: " & processedCount & vbCrLf & _
           "Skipped (Name Not Found): " & skippedCount & vbCrLf & _
           "Duplicates prevented: " & duplicateCount, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Function ExtractStudentNameFromSubject(subjectLine As String) As String
    ' Handles two formats:
    ' "Timesheet Approval Required - Lastname, Firstname"
    ' "FINAL NOTICE: Immediate Approval Required - Lastname, Firstname - Payment Delayed"
    
    Dim parts() As String
    Dim candidate As String
    
    ' Strip RE: or FW: prefix if present
    subjectLine = Trim(subjectLine)
    If Left(LCase(subjectLine), 3) = "re:" Then subjectLine = Trim(Mid(subjectLine, 4))
    If Left(LCase(subjectLine), 3) = "fw:" Then subjectLine = Trim(Mid(subjectLine, 4))
    
    ' Split on " - "
    parts = Split(subjectLine, " - ")
    
    If UBound(parts) >= 1 Then
        ' Standard format: name is after first " - "
        ' FINAL NOTICE format: name is after first " - ", "Payment Delayed" is after second
        candidate = Trim(parts(1))
        
        ' If candidate contains "Payment Delayed", it parsed wrong -- skip
        If InStr(1, LCase(candidate), "payment delayed", vbTextCompare) > 0 Then
            ExtractStudentNameFromSubject = ""
            Exit Function
        End If
        
        ' Validate it looks like a name (contains a comma or space)
        If InStr(candidate, ",") > 0 Or InStr(candidate, " ") > 0 Then
            ExtractStudentNameFromSubject = candidate
            Exit Function
        End If
    End If
    
    ExtractStudentNameFromSubject = ""
End Function

Function ExtractTimesheetStatusDate(htmlBody As String) As Date
    Dim plainText As String
    Dim searchPhrase As String
    Dim pos As Long
    Dim dateStr As String
    Dim extractedDate As Date
    
    ExtractTimesheetStatusDate = 0
    
    plainText = StripHTML(htmlBody)
    searchPhrase = "Time Sheet Status as of "
    
    pos = InStr(1, plainText, searchPhrase, vbTextCompare)
    If pos = 0 Then
        searchPhrase = "Timesheet Status as of "
        pos = InStr(1, plainText, searchPhrase, vbTextCompare)
    End If
    
    If pos = 0 Then Exit Function
    
    dateStr = Trim(Mid(plainText, pos + Len(searchPhrase), 12))
    
    ' Remove trailing period or other punctuation
    dateStr = Replace(dateStr, ".", "")
    dateStr = Trim(dateStr)
    
    On Error Resume Next
    If IsDate(dateStr) Then
        extractedDate = CDate(dateStr)
        ExtractTimesheetStatusDate = extractedDate
    End If
    On Error GoTo 0
End Function





Sub ProcessSupervisorReplies()
    Dim myNameSpace As Outlook.NameSpace
    Dim sharedMailbox As Outlook.folder
    Dim sharedInbox As Outlook.folder
    Dim replyFolder As Outlook.folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim oMail As Outlook.mailItem
    Dim processedCount As Long
    Dim skippedCount As Long
    Dim scannedCount As Long
    Dim emailsToMove As Collection
    Dim emailToMove As Object
    
    ' Load lookup data if not already loaded
    If lookupData Is Nothing Then
        Call LoadLookupDataIntoMemory
    End If
    
    On Error GoTo ErrorHandler
    
    processedCount = 0
    skippedCount = 0
    scannedCount = 0
    Dim cutoffStart As Date
    Dim cutoffEnd As Date
    cutoffStart = GetPayPeriodCutoffStart()
    cutoffEnd = GetPayPeriodCutoffEnd()
    Set emailsToMove = New Collection
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    
    On Error Resume Next
    Set sharedMailbox = myNameSpace.Folders("DST-INTERNPAYROLL@district.edu")
    On Error GoTo ErrorHandler
    
    If sharedMailbox Is Nothing Then
        MsgBox "ERROR: Cannot access shared mailbox", vbCritical
        Exit Sub
    End If
    
    Set sharedInbox = sharedMailbox.Folders("Inbox")
    
    On Error Resume Next
    Set replyFolder = sharedInbox.Folders("Supervisor_Replies")
    If replyFolder Is Nothing Then
        Set replyFolder = sharedInbox.Folders.Add("Supervisor_Replies")
    End If
    On Error GoTo ErrorHandler
    
    Set myItems = sharedInbox.Items
    myItems.Sort "[ReceivedTime]", True
    
    ' First pass: identify replies and log them
    For Each myItem In myItems
        If scannedCount >= 50 Then Exit For
        If TypeOf myItem Is Outlook.mailItem Then
            Set oMail = myItem
            If oMail.ReceivedTime < cutoffStart Then GoTo NextProcessItem
            If oMail.ReceivedTime > cutoffEnd Then GoTo NextProcessItem
            scannedCount = scannedCount + 1
            
            Dim isReply As Boolean
isReply = (InStr(1, oMail.Body, "APPROVE", vbTextCompare) > 0 Or _
           InStr(1, oMail.Body, "REJECT", vbTextCompare) > 0 Or _
           InStr(1, oMail.Body, "CORRECTIONS", vbTextCompare) > 0)
            
            If isReply Then
                On Error Resume Next
                Dim logSuccess As Boolean
                logSuccess = LogSupervisorReply(oMail)
                If logSuccess Then
                    emailsToMove.Add oMail
                    processedCount = processedCount + 1
                Else
                    skippedCount = skippedCount + 1
                End If
                On Error GoTo ErrorHandler
            End If
        End If
NextProcessItem:
    Next
    
    ' Second pass: move emails (cannot move while iterating)
    For Each emailToMove In emailsToMove
        On Error Resume Next
        emailToMove.Move replyFolder
        Err.Clear
        On Error GoTo ErrorHandler
    Next
    
    MsgBox "Processed " & processedCount & " supervisor replies" & vbCrLf & _
           "Skipped: " & skippedCount, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ==========================================
' LOGGING FUNCTIONS
' ==========================================
Function excel_log(sentto As String, subject As String, student As String, date_sent As Date) As Boolean
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelSheet As Object
    Dim lastRow As Long
    Dim lookupKey As String
    Dim studentName As String
    Dim supervisorName As String
    Dim employerName As String
    Dim fso As Object
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Function
    
    On Error GoTo ErrorHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(excelFilePath) Then
        excel_log = False
        Exit Function
    End If
    
studentName = FastGetStudentName(student)
supervisorName = FastGetSupervisorName(sentto)
employerName = FastGetEmployer(sentto)
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set excelSheet = excelWorkbook.Sheets("Sent_Log")
    
    lastRow = 1
    If excelSheet.cells(1, 1).Value = "" Or excelSheet.cells(1, 1).Value = "Student Name" Then
        On Error Resume Next
        lastRow = excelSheet.cells(excelSheet.rows.count, 1).End(-4162).Row
        On Error GoTo ErrorHandler
        If lastRow >= 1048576 Or excelSheet.cells(lastRow, 1).Value = "Student Name" Then
            lastRow = 1
        End If
    End If
    
    If excelSheet.cells(lastRow + 1, 1).Value = "Student Name" Then
        lastRow = lastRow + 1
    End If
    
    lookupKey = student & "|" & sentto
    
    Dim writeRow As Long
    writeRow = lastRow + 1
    If writeRow = 1 And excelSheet.cells(1, 1).Value <> "" Then
        writeRow = 2
    End If
    
    excelSheet.cells(writeRow, 1).Value = studentName
    excelSheet.cells(writeRow, 2).Value = supervisorName
    excelSheet.cells(writeRow, 3).Value = sentto
    excelSheet.cells(writeRow, 4).Value = employerName
    excelSheet.cells(writeRow, 5).Value = subject
    excelSheet.cells(writeRow, 6).Value = student
    excelSheet.cells(writeRow, 7).Value = date_sent
    excelSheet.cells(writeRow, 8).Value = lookupKey
    
    excelWorkbook.Save
    excelWorkbook.Close
    excelApp.Quit
    excel_log = True
    
    Set excelSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Function
    
ErrorHandler:
    excel_log = False
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    On Error GoTo 0
End Function

Function LogSupervisorReply(replyEmail As Outlook.mailItem) As Boolean
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim replySheet As Object
    Dim sentLogSheet As Object
    Dim lastRow As Long
    Dim sentLastRow As Long
    Dim sentRow As Long
    Dim supervisorEmail As String
    Dim supervisorName As String
    Dim studentEmail As String
    Dim studentName As String
    Dim employerName As String
    Dim responseType As String
    Dim lookupKey As String
    Dim fso As Object
    Dim i As Long
    Dim cleanSupervisorEmail As String
    Dim sentDate As Date
    Dim timeDiff As Double
    
    ' Variables for best match
    Dim bestMatchRow As Long
    Dim bestMatchDiff As Double
    
    ' Variables for duplicate check
    Dim replyKey As String
    Dim existingKey As String
    
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Function
    
    On Error GoTo ErrorHandler
    
    LogSupervisorReply = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(excelFilePath) Then Exit Function
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    supervisorEmail = GetSMTPAddress(replyEmail)
    cleanSupervisorEmail = LCase(Trim(supervisorEmail))
    responseType = DetermineResponseType(replyEmail.Body)
    sentLastRow = sentLogSheet.cells(sentLogSheet.rows.count, 1).End(-4162).Row
    
' Extract intern name from reply subject first
Dim parsedName As String
parsedName = ExtractStudentNameFromReplySubject(replyEmail.subject)
If InStr(parsedName, " - ") > 0 Then parsedName = Trim(Left(parsedName, InStr(parsedName, " - ") - 1))
parsedName = LCase(Trim(parsedName))

' FIND THE BEST MATCH in Sent_Log
bestMatchRow = 0
bestMatchDiff = 999
Dim bestMatchScore As Double
bestMatchScore = -1

For sentRow = 2 To sentLastRow
    If InStr(1, LCase(Trim(sentLogSheet.cells(sentRow, 3).Value)), cleanSupervisorEmail, vbTextCompare) > 0 Then
        
        On Error Resume Next
        sentDate = sentLogSheet.cells(sentRow, 7).Value
        On Error GoTo ErrorHandler
        
        If IsDate(sentDate) Then
            timeDiff = replyEmail.ReceivedTime - sentDate
            
            If timeDiff > 0 And timeDiff <= 14 Then
                studentName = sentLogSheet.cells(sentRow, 1).Value
                
                If studentName <> "Name Not Found" And studentName <> "" Then
                    Dim nameScore As Double
                    nameScore = 0
                    If parsedName <> "" Then
                        If InStr(1, LCase(studentName), parsedName, vbTextCompare) > 0 Then
                            nameScore = 1
                        ElseIf parsedName <> "" Then
                            ' Try matching on last name only
                            Dim nameParts As Variant
                            nameParts = Split(parsedName, " ")
                            If UBound(nameParts) >= 0 Then
                                If InStr(1, LCase(studentName), nameParts(0), vbTextCompare) > 0 Then
                                    nameScore = 0.5
                                End If
                            End If
                        End If
                    End If
                    
                    Dim timeScore As Double
                    timeScore = 1 - (timeDiff / 14)
                    
                    Dim totalScore As Double
                    totalScore = (nameScore * 2) + timeScore
                    
                    If totalScore > bestMatchScore Then
                        bestMatchScore = totalScore
                        bestMatchRow = sentRow
                        bestMatchDiff = timeDiff
                    End If
                End If
            End If
        End If
    End If
Next sentRow
    
    ' If no match found, try to get info from CC or subject
    If bestMatchRow = 0 Then
        ' Fallback to original logic
        studentEmail = ""
        
        If replyEmail.CC <> "" Then
            studentEmail = ExtractEmailAddress(replyEmail.CC)
        End If
        
        If studentEmail = "" Then
            Dim extractedName As String
            extractedName = ExtractStudentNameFromSubject(replyEmail.subject)
            If extractedName <> "" Then
                studentEmail = GetStudentEmailByName(extractedName)
            End If
        End If
        
        If studentEmail = "" Then
            studentEmail = "UNKNOWN"
            studentName = "Name Not Found"
            supervisorName = GetSupervisor(supervisorEmail)
            employerName = GetEmployer(supervisorEmail)
        Else
            studentName = GetStudentName(studentEmail)
            supervisorName = GetSupervisor(supervisorEmail)
            employerName = GetEmployer(supervisorEmail)
        End If
        
        lookupKey = studentEmail & "|" & supervisorEmail
    Else
        ' Use data from best match
        studentName = sentLogSheet.cells(bestMatchRow, 1).Value
        supervisorName = sentLogSheet.cells(bestMatchRow, 2).Value
        supervisorEmail = sentLogSheet.cells(bestMatchRow, 3).Value
        employerName = sentLogSheet.cells(bestMatchRow, 4).Value
        studentEmail = sentLogSheet.cells(bestMatchRow, 6).Value
        lookupKey = studentEmail & "|" & supervisorEmail
    End If
    
    ' CHECK FOR DUPLICATES before logging
    ' Create a unique key for this reply
    replyKey = LCase(studentEmail) & "|" & LCase(cleanSupervisorEmail) & "|" & Format(replyEmail.ReceivedTime, "yyyymmddhhmmss")
    
    ' Get Reply_Log last row
    On Error Resume Next
    lastRow = replySheet.cells(replySheet.rows.count, 1).End(-4162).Row
    If lastRow >= 1048576 Then lastRow = 1
    If replySheet.cells(lastRow, 1).Value = "" And lastRow > 1 Then
        For i = lastRow To 1 Step -1
            If replySheet.cells(i, 1).Value <> "" Then
                lastRow = i
                Exit For
            End If
        Next i
        If i = 0 Then lastRow = 1
    End If
    On Error GoTo ErrorHandler
    
    ' Check if this exact reply already exists
    For i = 2 To lastRow
        ' Check by lookup key and reply date
        existingKey = LCase(Trim(replySheet.cells(i, 7).Value))
        Dim existingDate As Date
        On Error Resume Next
        existingDate = replySheet.cells(i, 5).Value
        On Error GoTo ErrorHandler
        
        ' If same student-supervisor combo and same reply date (within 1 minute), skip
        If existingKey = LCase(lookupKey) Then
            If IsDate(existingDate) Then
                If Abs(DateDiff("n", existingDate, replyEmail.ReceivedTime)) < 2 Then
                    ' Duplicate found, skip logging
                    excelWorkbook.Close False
                    excelApp.Quit
                    LogSupervisorReply = True  ' Return True so email still gets moved
                    Exit Function
                End If
            End If
        End If
    Next i
    
    ' No duplicate found, log the reply
    replySheet.cells(lastRow + 1, 1).Value = studentName
    replySheet.cells(lastRow + 1, 2).Value = supervisorName
    replySheet.cells(lastRow + 1, 3).Value = supervisorEmail
    replySheet.cells(lastRow + 1, 4).Value = employerName
    replySheet.cells(lastRow + 1, 5).Value = replyEmail.ReceivedTime
    replySheet.cells(lastRow + 1, 6).Value = responseType
    replySheet.cells(lastRow + 1, 7).Value = lookupKey
    
    excelWorkbook.Save
    excelWorkbook.Close False
    excelApp.Quit
    
    LogSupervisorReply = True
    
    Set replySheet = Nothing
    Set sentLogSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Function
    
ErrorHandler:
    LogSupervisorReply = False
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    On Error GoTo 0
End Function




Public Function excel_vlooklook(email As String) As String
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelSheet As Object
    Dim result As Variant
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Function
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set excelSheet = excelWorkbook.Sheets("Updated_Supervisor email")
    
    result = excelApp.VLookup(email, excelSheet.Range("B:H"), 7, False)
    
    If IsError(result) Then
        excel_vlooklook = "NA"
    Else
        excel_vlooklook = result
    End If
    
    excelWorkbook.Close False
    excelApp.Quit
    Set excelSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Function
    
ErrorHandler:
    excel_vlooklook = "NA"
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Function


Function FastGetStudentName(email As String) As String
    Dim cleanEmail As String
    Dim acesID As String
    Dim data As Variant
    
    cleanEmail = LCase(Trim(ExtractEmailAddress(email)))
    
    If InStr(cleanEmail, "@") > 0 Then
        acesID = Left(cleanEmail, InStr(cleanEmail, "@") - 1)
    Else
        acesID = cleanEmail
    End If
    
    If lookupData.Exists(acesID) Then
        data = lookupData(acesID)
        FastGetStudentName = data(0)
        Exit Function
    End If
    
    If lookupData.Exists(cleanEmail) Then
        data = lookupData(cleanEmail)
        FastGetStudentName = data(0)
        Exit Function
    End If
    
    FastGetStudentName = "Name Not Found"
End Function

Function FastGetSupervisorName(email As String) As String
    Dim cleanEmail As String
    Dim data As Variant
    
    cleanEmail = LCase(Trim(ExtractEmailAddress(email)))
    
    If lookupData.Exists("sup_" & cleanEmail) Then
        data = lookupData("sup_" & cleanEmail)
        FastGetSupervisorName = data(0)
    Else
        FastGetSupervisorName = "Unknown"
    End If
End Function

Function FastGetEmployer(email As String) As String
    Dim cleanEmail As String
    Dim data As Variant
    
    cleanEmail = LCase(Trim(ExtractEmailAddress(email)))
    
    If lookupData.Exists("sup_" & cleanEmail) Then
        data = lookupData("sup_" & cleanEmail)
        FastGetEmployer = data(1)
    Else
        FastGetEmployer = "Unknown"
    End If
End Function

Sub FastRebuildSentLog()
    Dim myNameSpace As Outlook.NameSpace
    Dim personalInbox As Outlook.folder
    Dim sentFolder As Outlook.folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim oMail As Outlook.mailItem
    Dim processedCount As Long
    Dim skippedCount As Long
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelSheet As Object
    Dim htmlBody As String
    Dim studentEmail As String
    Dim supervisorEmail As String
    Dim studentName As String
    Dim supervisorName As String
    Dim employerName As String
    Dim cleanStudentEmail As String
    Dim acesID As String
    Dim data As Variant
    Dim cutoffStart As Date
    Dim cutoffEnd As Date
    Dim excelFilePath As String

    ' Collect into array before writing to Excel
    Dim results() As Variant
    Dim resultCount As Long
    ReDim results(1 To 500, 1 To 8)
    resultCount = 0

    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub

    On Error GoTo ErrorHandler

    If lookupData Is Nothing Then Call LoadLookupDataIntoMemory

    cutoffStart = GetPayPeriodCutoffStart()
    cutoffEnd = GetPayPeriodCutoffEnd()

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set personalInbox = myNameSpace.GetDefaultFolder(olFolderInbox)

    If personalInbox Is Nothing Then
        MsgBox "ERROR: Cannot access default inbox", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set sentFolder = personalInbox.Folders("Sent_super")
    On Error GoTo ErrorHandler

    If sentFolder Is Nothing Then
        MsgBox "ERROR: Sent_super folder not found", vbCritical
        Exit Sub
    End If

    Set myItems = sentFolder.Items
    myItems.Sort "[ReceivedTime]", True

    processedCount = 0
    skippedCount = 0

    For Each myItem In myItems
        If TypeOf myItem Is Outlook.mailItem Then
            Set oMail = myItem

            If oMail.ReceivedTime < cutoffStart Then Exit For
            If oMail.ReceivedTime > cutoffEnd Then GoTo NextSentItem

            htmlBody = ""
            On Error Resume Next
            htmlBody = oMail.htmlBody
            On Error GoTo ErrorHandler

            If htmlBody = "" Then GoTo NextSentItem


            studentEmail = ExtractEmailFromHTMLBody(htmlBody)

            If studentEmail <> "" Then
                cleanStudentEmail = LCase(Trim(studentEmail))
                If InStr(cleanStudentEmail, "@") > 0 Then
                    acesID = Left(cleanStudentEmail, InStr(cleanStudentEmail, "@") - 1)
                Else
                    acesID = cleanStudentEmail
                End If

                studentName = "Name Not Found"
                supervisorEmail = ""
                supervisorName = "Unknown"
                employerName = "Unknown"

                If lookupData.Exists(acesID) Then
                    data = lookupData(acesID)
                    studentName = data(0)
                    supervisorEmail = data(1)
                    supervisorName = data(2)
                    employerName = data(3)
                ElseIf lookupData.Exists(cleanStudentEmail) Then
                    data = lookupData(cleanStudentEmail)
                    studentName = data(0)
                    supervisorEmail = data(1)
                    supervisorName = data(2)
                    employerName = data(3)
                Else
                    skippedCount = skippedCount + 1
                End If

                resultCount = resultCount + 1
                results(resultCount, 1) = studentName
                results(resultCount, 2) = supervisorName
                results(resultCount, 3) = supervisorEmail
                results(resultCount, 4) = employerName
                results(resultCount, 5) = oMail.subject
                results(resultCount, 6) = studentEmail
                results(resultCount, 7) = oMail.ReceivedTime
                results(resultCount, 8) = studentEmail & "|" & supervisorEmail

                processedCount = processedCount + 1
            End If

            DoEvents
        End If

NextSentItem:
    Next

    ' Write all results to Excel in one batch
    If resultCount > 0 Then
        Set excelApp = CreateObject("Excel.Application")
        excelApp.Visible = False
        excelApp.DisplayAlerts = False
        Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
        Set excelSheet = excelWorkbook.Sheets("Sent_Log")
        excelSheet.Range("A2:H" & excelSheet.rows.count).ClearContents

        ' Single range write -- all rows at once
        excelSheet.Range("A2").Resize(resultCount, 8).Value = results

        excelWorkbook.Save
        excelWorkbook.Close False
        excelApp.Quit
    End If

    MsgBox "Sent_Log: " & processedCount & " processed, " & skippedCount & " skipped", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error in FastRebuildSentLog: " & Err.Description, vbCritical
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub



Sub FastRebuildReplyLog()
    Dim myNameSpace As Outlook.NameSpace
    Dim sharedMailbox As Outlook.folder
    Dim sharedInbox As Outlook.folder
    Dim replyFolder As Outlook.folder
    Dim myItems As Outlook.Items
    Dim myItem As Object
    Dim oMail As Outlook.mailItem
    Dim cutoffDate As Date
    Dim processedCount As Long
    Dim movedCount As Long
    Dim noMatchCount As Long
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim replySheet As Object
    Dim sentLogSheet As Object
    Dim currentRow As Long
    Dim responseType As String
    Dim isReply As Boolean
    Dim sentLastRow As Long
    
    ' Variables from Sent_Log
    Dim studentName As String
    Dim supervisorName As String
    Dim employerName As String
    Dim studentEmail As String
    Dim supervisorEmail As String
    
    ' Collection to hold emails to move
    Dim emailsToMove As Collection
    
    ' Dictionary to track logged replies
    Dim loggedReplies As Object
    Dim replyKey As String
    
    ' Best match row
    Dim bestMatchRow As Long
    
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim cutoffStart As Date
    Dim cutoffEnd As Date
    cutoffStart = GetPayPeriodCutoffStart()
    cutoffEnd = GetPayPeriodCutoffEnd()
    
    ' Open Excel
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    
    ' Clear Reply_Log (keep header row)
    replySheet.Range("A2:G" & replySheet.rows.count).ClearContents
    currentRow = 2
    
    ' Get Sent_Log last row
    sentLastRow = sentLogSheet.cells(sentLogSheet.rows.count, 1).End(-4162).Row
    
    ' Initialize dictionary to track logged replies
    Set loggedReplies = CreateObject("Scripting.Dictionary")
    
    ' Setup Outlook folders
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set sharedMailbox = myNameSpace.Folders("DST-INTERNPAYROLL@district.edu")
    Set sharedInbox = sharedMailbox.Folders("Inbox")
    
    ' Create Supervisor_Replies folder if needed
    On Error Resume Next
    Set replyFolder = sharedInbox.Folders("Supervisor_Replies")
    If replyFolder Is Nothing Then
        Set replyFolder = sharedInbox.Folders.Add("Supervisor_Replies")
    End If
    On Error GoTo ErrorHandler
    
    ' Collection to store emails to move
    Set emailsToMove = New Collection
    
    Set myItems = sharedInbox.Items
    myItems.Sort "[ReceivedTime]", True
    
    processedCount = 0
    movedCount = 0
    noMatchCount = 0
    
    ' Loop through shared inbox emails
    For Each myItem In myItems
        If TypeOf myItem Is Outlook.mailItem Then
            Set oMail = myItem
            
        If oMail.ReceivedTime < cutoffStart Then Exit For
        If oMail.ReceivedTime > cutoffEnd Then GoTo NextSentItem
            
' Check if this is a supervisor reply
            isReply = (InStr(1, oMail.subject, "Timesheet", vbTextCompare) > 0 Or _
                       InStr(1, oMail.subject, "INTERN_PROGRAM", vbTextCompare) > 0 Or _
                       InStr(1, oMail.subject, "Approval", vbTextCompare) > 0) And _
                      (InStr(1, oMail.Body, "APPROVE", vbTextCompare) > 0 Or _
                       InStr(1, oMail.Body, "REJECT", vbTextCompare) > 0 Or _
                       InStr(1, oMail.Body, "CORRECTIONS", vbTextCompare) > 0)
            
            If isReply Then
                ' Get response type from email body
                responseType = DetermineResponseType(oMail.Body)
                
                ' USE THE NEW MATCHING FUNCTION
                bestMatchRow = FindBestMatchForReply(oMail, sentLogSheet, sentLastRow)
                
                ' Log if match found
                If bestMatchRow > 0 Then
                    studentName = sentLogSheet.cells(bestMatchRow, 1).Value
                    supervisorName = sentLogSheet.cells(bestMatchRow, 2).Value
                    supervisorEmail = sentLogSheet.cells(bestMatchRow, 3).Value
                    employerName = sentLogSheet.cells(bestMatchRow, 4).Value
                    studentEmail = sentLogSheet.cells(bestMatchRow, 6).Value
                    
                    ' Create unique key for this reply
                    replyKey = LCase(studentEmail) & "|" & LCase(supervisorEmail) & "|" & Format(oMail.ReceivedTime, "yyyymmddhhmmss")
                    
                    ' Only log if not already logged
                    If Not loggedReplies.Exists(replyKey) Then
                        loggedReplies.Add replyKey, True
                        
                        ' Write to Reply_Log
                        replySheet.cells(currentRow, 1).Value = studentName
                        replySheet.cells(currentRow, 2).Value = supervisorName
                        replySheet.cells(currentRow, 3).Value = supervisorEmail
                        replySheet.cells(currentRow, 4).Value = employerName
                        replySheet.cells(currentRow, 5).Value = oMail.ReceivedTime
                        replySheet.cells(currentRow, 6).Value = responseType
                        replySheet.cells(currentRow, 7).Value = studentEmail & "|" & supervisorEmail
                        
                        currentRow = currentRow + 1
                        processedCount = processedCount + 1
                    End If
                Else
                    noMatchCount = noMatchCount + 1
                End If
                
                ' Add email to move list
                emailsToMove.Add oMail
            End If
        End If
    
NextSentItem:
    Next
    
    ' Save Excel before moving emails
    excelWorkbook.Save
    
    ' Move emails to Supervisor_Replies folder
    Dim emailToMove As Object
    For Each emailToMove In emailsToMove
        On Error Resume Next
        emailToMove.Move replyFolder
        If Err.Number = 0 Then
            movedCount = movedCount + 1
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next
    
Cleanup:
    If Not excelWorkbook Is Nothing Then
        excelWorkbook.Save
        excelWorkbook.Close False
    End If
    If Not excelApp Is Nothing Then excelApp.Quit
    
    MsgBox "Reply_Log Complete!" & vbCrLf & _
           "Entries logged: " & processedCount & vbCrLf & _
           "No match found: " & noMatchCount & vbCrLf & _
           "Emails moved: " & movedCount, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error rebuilding Reply_Log: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
Sub RunDailyTimesheetProcess()
    MsgBox "Starting daily process...", vbInformation
    
    ' Load lookup data ONCE at the start
    Call LoadLookupDataIntoMemory
    
    Call MoveItems
    Call ProcessSupervisorReplies
    
    ' Clear lookup data when done
    Call ClearLookupData
    
    Call UpdateSubmissionStatus
    Call GeneratePayPeriodSummary
    
    MsgBox "Daily process complete!", vbInformation
End Sub



' ==========================================
' SUBMISSION TRACKING - AUTO-POPULATE
' ==========================================

Sub PopulateSubmissionTracking()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim sourceSheet As Object
    Dim trackingSheet As Object
    Dim lastRowSource As Long
    Dim lastRowTracking As Long
    Dim currentRow As Long
    Dim i As Long
    Dim payPeriod As String
    Dim archiveChoice As VbMsgBoxResult
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Prompt for pay period
    payPeriod = InputBox("Enter the pay period (e.g., Jan 16-31, 2026):", _
                         "Pay Period", _
                         Format(Date, "mmm") & " " & IIf(Day(Date) <= 15, "1-15", "16-" & Day(DateSerial(Year(Date), Month(Date) + 1, 0))) & ", " & Year(Date))
    
    If payPeriod = "" Then
        MsgBox "Cancelled. No changes made.", vbInformation
        Exit Sub
    End If
    
    ' Open Excel
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    
    ' Get source sheet
    Set sourceSheet = excelWorkbook.Sheets("Updated_Supervisor email")
    
    ' Get or create tracking sheet
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If trackingSheet Is Nothing Then
        Set trackingSheet = excelWorkbook.Sheets.Add(After:=excelWorkbook.Sheets(excelWorkbook.Sheets.count))
        trackingSheet.Name = "Submission_Tracking"
        
        ' Create headers
        trackingSheet.cells(1, 1).Value = "Pay Period"
        trackingSheet.cells(1, 2).Value = "Student Name"
        trackingSheet.cells(1, 3).Value = "Student Email"
        trackingSheet.cells(1, 4).Value = "Employer"
        trackingSheet.cells(1, 5).Value = "Supervisor"
        trackingSheet.cells(1, 6).Value = "Supervisor Email"
        trackingSheet.cells(1, 7).Value = "Submitted"
        trackingSheet.cells(1, 8).Value = "Submitted Date"
        trackingSheet.cells(1, 9).Value = "Approved"
        trackingSheet.cells(1, 10).Value = "Approved Date"
        trackingSheet.cells(1, 11).Value = "Status"
        
        ' Format header row
        With trackingSheet.Range("A1:K1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
    On Error GoTo ErrorHandler
    
    ' Check if data exists and ask about archiving
    lastRowTracking = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    If lastRowTracking > 1 Then
    archiveChoice = MsgBox("Existing data found in Submission_Tracking." & vbCrLf & vbCrLf & _
                           "YES = Archive to Submission_Archive and clear" & vbCrLf & _
                           "NO = Clear without archiving" & vbCrLf & _
                           "CANCEL = Stop and make no changes", _
                           vbYesNoCancel + vbQuestion, "Existing Data Found")

    If archiveChoice = vbCancel Then
        MsgBox "Cancelled. No changes made.", vbInformation
        GoTo Cleanup
    ElseIf archiveChoice = vbYes Then
        Call ArchiveSubmissionData(excelWorkbook, trackingSheet)
        lastRowTracking = 1
    ElseIf archiveChoice = vbNo Then
        ' Clear without archiving
        trackingSheet.Range("A2:K" & lastRowTracking).Clear
        lastRowTracking = 1
    End If
End If
    
    ' Get last row of source data
    lastRowSource = sourceSheet.cells(sourceSheet.rows.count, 3).End(-4162).Row
    
    If lastRowSource < 2 Then
        MsgBox "No student data found in Updated_Supervisor email sheet.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Start populating
    currentRow = lastRowTracking + 1
    
    For i = 2 To lastRowSource
        ' Skip empty rows
        If Trim(sourceSheet.cells(i, 3).Value) <> "" Then
            trackingSheet.cells(currentRow, 1).Value = payPeriod                          ' Pay Period
            trackingSheet.cells(currentRow, 2).Value = sourceSheet.cells(i, 3).Value      ' Student Name (C)
            trackingSheet.cells(currentRow, 3).Value = sourceSheet.cells(i, 2).Value      ' Student Email (B)
            trackingSheet.cells(currentRow, 4).Value = sourceSheet.cells(i, 6).Value      ' Employer (F)
            trackingSheet.cells(currentRow, 5).Value = sourceSheet.cells(i, 7).Value      ' Supervisor (G)
            trackingSheet.cells(currentRow, 6).Value = sourceSheet.cells(i, 8).Value      ' Supervisor Email (H)
            trackingSheet.cells(currentRow, 7).Value = "NO"                               ' Submitted
            trackingSheet.cells(currentRow, 8).Value = ""                                 ' Submitted Date
            trackingSheet.cells(currentRow, 9).Value = "PENDING"                          ' Approved
            trackingSheet.cells(currentRow, 10).Value = ""                                ' Approved Date
            trackingSheet.cells(currentRow, 11).Value = "Not Submitted"                   ' Status
            
            ' Color the status cell red
            trackingSheet.cells(currentRow, 11).Interior.Color = RGB(255, 200, 200)
            
            currentRow = currentRow + 1
        End If
    Next i
    
    ' Auto-fit columns
    trackingSheet.Columns("A:K").AutoFit
    
    ' Add data validation for Submitted column (G)
    On Error Resume Next
    trackingSheet.Range("G2:G" & currentRow - 1).Validation.Delete
    trackingSheet.Range("G2:G" & currentRow - 1).Validation.Add Type:=3, AlertStyle:=1, _
        Formula1:="YES,NO"
    
    ' Add data validation for Approved column (I)
    trackingSheet.Range("I2:I" & currentRow - 1).Validation.Delete
    trackingSheet.Range("I2:I" & currentRow - 1).Validation.Add Type:=3, AlertStyle:=1, _
        Formula1:="YES,NO,REJECTED,PENDING"
    On Error GoTo ErrorHandler
    
    excelWorkbook.Save
    
    MsgBox "Submission Tracking populated!" & vbCrLf & vbCrLf & _
           "Pay Period: " & payPeriod & vbCrLf & _
           "Students Added: " & (currentRow - lastRowTracking - 1), vbInformation
    
Cleanup:
    If Not excelWorkbook Is Nothing Then
        excelWorkbook.Close True
    End If
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If
    Set trackingSheet = Nothing
    Set sourceSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub




' ==========================================
' CLEAN FILE NAME (REMOVE INVALID CHARACTERS)
' ==========================================

Function CleanFileName(fileName As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    
    result = fileName
    
    ' Replace invalid characters
    result = Replace(result, "/", "-")
    result = Replace(result, "\", "-")
    result = Replace(result, ":", "-")
    result = Replace(result, "*", "")
    result = Replace(result, "?", "")
    result = Replace(result, """", "")
    result = Replace(result, "<", "")
    result = Replace(result, ">", "")
    result = Replace(result, "|", "")
    result = Replace(result, " ", "_")
    result = Replace(result, ",", "")
    
    CleanFileName = result
End Function

' ==========================================
' CLEAR LOGS FOR NEW PAY PERIOD
' Run this AFTER archiving
' ==========================================

Sub ClearLogsForNewPeriod()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim sentLogSheet As Object
    Dim replyLogSheet As Object
    Dim trackingSheet As Object
    Dim dashboardSheet As Object
    Dim nonSubmittersSheet As Object
    Dim callListSheet As Object
    Dim clearChoice As VbMsgBoxResult
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    clearChoice = MsgBox("CLEAR LOGS FOR NEW PAY PERIOD" & vbCrLf & _
                         "==============================" & vbCrLf & vbCrLf & _
                         "WARNING: This will clear the following:" & vbCrLf & _
                         "- Sent_Log (all rows except header)" & vbCrLf & _
                         "- Reply_Log (all rows except header)" & vbCrLf & _
                         "- Submission_Tracking (all rows except header)" & vbCrLf & _
                         "- Dashboard (will be rebuilt)" & vbCrLf & _
                         "- Non_Submitters (if exists)" & vbCrLf & _
                         "- Call_List (if exists)" & vbCrLf & vbCrLf & _
                         "Have you already archived the current pay period?" & vbCrLf & vbCrLf & _
                         "This cannot be undone!", vbYesNo + vbExclamation, "Confirm Clear")
    
    If clearChoice = vbNo Then
        MsgBox "Cancelled. No changes made.", vbInformation
        Exit Sub
    End If
    
    ' Double confirm
    clearChoice = MsgBox("Are you ABSOLUTELY SURE you want to clear all logs?" & vbCrLf & vbCrLf & _
                         "Type YES to confirm this action.", vbYesNo + vbCritical, "Final Confirmation")
    
    If clearChoice = vbNo Then
        MsgBox "Cancelled. No changes made.", vbInformation
        Exit Sub
    End If
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    
    ' Clear Sent_Log
    On Error Resume Next
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    If Not sentLogSheet Is Nothing Then
        Dim sentLastRow As Long
        sentLastRow = sentLogSheet.cells(sentLogSheet.rows.count, 1).End(-4162).Row
        If sentLastRow > 1 Then
            sentLogSheet.Range("A2:H" & sentLastRow).ClearContents
        End If
    End If
    
    ' Clear Reply_Log
    Set replyLogSheet = excelWorkbook.Sheets("Reply_Log")
    If Not replyLogSheet Is Nothing Then
        Dim replyLastRow As Long
        replyLastRow = replyLogSheet.cells(replyLogSheet.rows.count, 1).End(-4162).Row
        If replyLastRow > 1 Then
            replyLogSheet.Range("A2:G" & replyLastRow).ClearContents
        End If
    End If
    
    ' Clear Submission_Tracking
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If Not trackingSheet Is Nothing Then
        Dim trackLastRow As Long
        trackLastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
        If trackLastRow > 1 Then
            trackingSheet.Range("A2:K" & trackLastRow).Clear
        End If
    End If
    
    ' Clear Dashboard
    Set dashboardSheet = excelWorkbook.Sheets("Dashboard")
    If Not dashboardSheet Is Nothing Then
        dashboardSheet.cells.Clear
    End If
    
    ' Clear Non_Submitters if it exists
    Set nonSubmittersSheet = excelWorkbook.Sheets("Non_Submitters")
    If Not nonSubmittersSheet Is Nothing Then
        nonSubmittersSheet.cells.Clear
    End If
    
    ' Clear Call_List if it exists
    Set callListSheet = excelWorkbook.Sheets("Call_List")
    If Not callListSheet Is Nothing Then
        callListSheet.cells.Clear
    End If
    
    On Error GoTo ErrorHandler
    
    excelWorkbook.Save
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "All logs cleared!" & vbCrLf & vbCrLf & _
           "You are ready to start a new pay period." & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Run PopulateSubmissionTracking for the new period" & vbCrLf & _
           "2. Run your daily process as normal", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error clearing logs: " & Err.Description, vbCritical
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub


' ==========================================
' VIEW ARCHIVED PAY PERIODS
' Lists all archived files
' ==========================================

Sub ViewArchivedPayPeriods()
    Dim fso As Object
    Dim archiveFolder As String
    Dim folder As Object
    Dim file As Object
    Dim fileList As String
    Dim fileCount As Long
    
    archiveFolder = "" & Environ("USERPROFILE") & "\Desktop\ApprovalFlow_Archive\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(archiveFolder) Then
        MsgBox "No archive folder found." & vbCrLf & vbCrLf & _
               "Archives will be created at:" & vbCrLf & _
               archiveFolder, vbInformation
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(archiveFolder)
    
    fileList = ""
    fileCount = 0
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            fileCount = fileCount + 1
            fileList = fileList & fileCount & ". " & file.Name & vbCrLf & _
                       "   Modified: " & file.DateLastModified & vbCrLf & _
                       "   Size: " & Round(file.Size / 1024, 1) & " KB" & vbCrLf & vbCrLf
        End If
    Next file
    
    If fileCount = 0 Then
        MsgBox "No archived pay periods found." & vbCrLf & vbCrLf & _
               "Archive folder: " & archiveFolder, vbInformation
    Else
        MsgBox "ARCHIVED PAY PERIODS" & vbCrLf & _
               "====================" & vbCrLf & vbCrLf & _
               "Location: " & archiveFolder & vbCrLf & vbCrLf & _
               fileList, vbInformation
    End If
End Sub



' ==========================================
' UPDATE SUBMISSION STATUS FROM SENT_LOG
' ==========================================

Sub UpdateSubmissionStatus()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim sentLogSheet As Object
    Dim replyLogSheet As Object
    Dim lastRowTracking As Long
    Dim lastRowSent As Long
    Dim lastRowReply As Long
    Dim i As Long
    Dim j As Long
    Dim studentEmail As String
    Dim cleanStudentEmail As String
    Dim sentStudentEmail As String
    Dim replyStudentEmail As String
    Dim matchFound As Boolean
    Dim updatedSubmitted As Long
    Dim updatedApproved As Long
    Dim statusText As String
    Dim statusColor As Long
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If trackingSheet Is Nothing Then
        MsgBox "Submission_Tracking sheet not found. Run PopulateSubmissionTracking first.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    Set replyLogSheet = excelWorkbook.Sheets("Reply_Log")
    
    lastRowTracking = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    lastRowSent = sentLogSheet.cells(sentLogSheet.rows.count, 6).End(-4162).Row
    lastRowReply = replyLogSheet.cells(replyLogSheet.rows.count, 1).End(-4162).Row
    
    If lastRowTracking < 2 Then
        MsgBox "No data in Submission_Tracking to update.", vbInformation
        GoTo Cleanup
    End If
    
    updatedSubmitted = 0
    updatedApproved = 0
    
    ' Loop through each student in tracking sheet
    For i = 2 To lastRowTracking
        studentEmail = Trim(trackingSheet.cells(i, 3).Value)
        
        If studentEmail <> "" Then
            cleanStudentEmail = LCase(studentEmail)
            
            ' Extract just the username part for matching
            Dim acesID As String
            If InStr(cleanStudentEmail, "@") > 0 Then
                acesID = Left(cleanStudentEmail, InStr(cleanStudentEmail, "@") - 1)
            Else
                acesID = cleanStudentEmail
            End If
            
            ' Check Sent_Log for submission (column F = student email)
            If trackingSheet.cells(i, 7).Value = "NO" Then
                For j = 2 To lastRowSent
                    sentStudentEmail = LCase(Trim(sentLogSheet.cells(j, 6).Value))
                    
                    If InStr(sentStudentEmail, acesID) > 0 Then
                        trackingSheet.cells(i, 7).Value = "YES"
                        trackingSheet.cells(i, 8).Value = sentLogSheet.cells(j, 7).Value  ' Date from Sent_Log
                        trackingSheet.cells(i, 9).Value = "PENDING"
                        updatedSubmitted = updatedSubmitted + 1
                        Exit For
                    End If
                Next j
            End If
            
            ' Check Reply_Log for approval (column 1 = student name, but we match on lookup key column 7)
            If trackingSheet.cells(i, 7).Value = "YES" And _
               (trackingSheet.cells(i, 9).Value = "PENDING" Or trackingSheet.cells(i, 9).Value = "") Then
                
                For j = 2 To lastRowReply
                    ' Reply_Log column 7 contains lookup key: studentEmail|supervisorEmail
                    Dim lookupKey As String
                    lookupKey = LCase(Trim(replyLogSheet.cells(j, 7).Value))
                    
                    If InStr(lookupKey, acesID) > 0 Then
                        Dim responseType As String
                        responseType = UCase(Trim(replyLogSheet.cells(j, 6).Value))
                        
                        If InStr(responseType, "APPROVED") > 0 Then
                            trackingSheet.cells(i, 9).Value = "YES"
                        ElseIf InStr(responseType, "REJECTED") > 0 Then
                            trackingSheet.cells(i, 9).Value = "REJECTED"
                        Else
                            trackingSheet.cells(i, 9).Value = "PENDING"
                        End If
                        
                        trackingSheet.cells(i, 10).Value = replyLogSheet.cells(j, 5).Value  ' Reply date
                        updatedApproved = updatedApproved + 1
                        Exit For
                    End If
                Next j
            End If
            
            ' Update Status column based on current values
            Dim submitted As String
            Dim approved As String
            submitted = trackingSheet.cells(i, 7).Value
            approved = trackingSheet.cells(i, 9).Value
            
            If submitted = "NO" Or submitted = "" Then
                statusText = "Not Submitted"
                statusColor = RGB(255, 200, 200)  ' Red
            ElseIf approved = "YES" Then
                statusText = "Complete"
                statusColor = RGB(200, 255, 200)  ' Green
            ElseIf approved = "REJECTED" Then
                statusText = "Rejected - Resubmit"
                statusColor = RGB(255, 200, 200)  ' Red
            ElseIf approved = "PENDING" Or approved = "" Then
                statusText = "Pending Approval"
                statusColor = RGB(255, 255, 200)  ' Yellow
            Else
                statusText = "Unknown"
                statusColor = RGB(224, 224, 224)  ' Gray
            End If
            
            trackingSheet.cells(i, 11).Value = statusText
            trackingSheet.cells(i, 11).Interior.Color = statusColor
        End If
    Next i
    
    excelWorkbook.Save
    
    MsgBox "Submission Tracking Updated!" & vbCrLf & vbCrLf & _
           "Submissions found: " & updatedSubmitted & vbCrLf & _
           "Approvals found: " & updatedApproved, vbInformation
    
Cleanup:
    If Not excelWorkbook Is Nothing Then
        excelWorkbook.Close True
    End If
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If
    Set replyLogSheet = Nothing
    Set sentLogSheet = Nothing
    Set trackingSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================
' VIEW SUBMISSION TRACKING SUMMARY
' ==========================================

Sub ViewSubmissionSummary()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim totalStudents As Long
    Dim submitted As Long
    Dim notSubmitted As Long
    Dim approved As Long
    Dim rejected As Long
    Dim pending As Long
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If trackingSheet Is Nothing Then
        MsgBox "Submission_Tracking sheet not found.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    totalStudents = 0
    submitted = 0
    notSubmitted = 0
    approved = 0
    rejected = 0
    pending = 0
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            totalStudents = totalStudents + 1
            
            If trackingSheet.cells(i, 7).Value = "YES" Then
                submitted = submitted + 1
            Else
                notSubmitted = notSubmitted + 1
            End If
            
            Select Case UCase(Trim(trackingSheet.cells(i, 9).Value))
                Case "YES"
                    approved = approved + 1
                Case "REJECTED"
                    rejected = rejected + 1
                Case Else
                    If trackingSheet.cells(i, 7).Value = "YES" Then
                        pending = pending + 1
                    End If
            End Select
        End If
    Next i
    
    MsgBox "SUBMISSION TRACKING SUMMARY" & vbCrLf & _
           "==========================" & vbCrLf & vbCrLf & _
           "Pay Period: " & trackingSheet.cells(2, 1).Value & vbCrLf & vbCrLf & _
           "Total Students: " & totalStudents & vbCrLf & _
           "----------------------------" & vbCrLf & _
           "Submitted: " & submitted & vbCrLf & _
           "Not Submitted: " & notSubmitted & vbCrLf & _
           "----------------------------" & vbCrLf & _
           "Approved: " & approved & vbCrLf & _
           "Rejected: " & rejected & vbCrLf & _
           "Pending Approval: " & pending, vbInformation
    
Cleanup:
    If Not excelWorkbook Is Nothing Then
        excelWorkbook.Close False
    End If
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If
    Set trackingSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Function GetSubmissionDeadline() As Date
    Dim targetDate As Date
    Dim currentDay As Long
    Dim currentMonth As Long
    Dim currentYear As Long
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim anchorDate As Date
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Function
    
    If overridePayPeriodDate > 0 Then
        targetDate = overridePayPeriodDate
    Else
        On Error Resume Next
        Set excelApp = CreateObject("Excel.Application")
        excelApp.Visible = False
        excelApp.DisplayAlerts = False
        Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
        anchorDate = excelWorkbook.Sheets("Updated_Supervisor email").Range("J1").Value
        excelWorkbook.Close False
        excelApp.Quit
        On Error GoTo 0
        
        If IsDate(anchorDate) And anchorDate > 0 Then
            targetDate = anchorDate
        Else
            targetDate = Date
        End If
    End If
    
    currentDay = Day(targetDate)
    currentMonth = Month(targetDate)
    currentYear = Year(targetDate)
    
    If currentDay <= 15 Then
        GetSubmissionDeadline = DateSerial(currentYear, currentMonth, 15)
    Else
        GetSubmissionDeadline = DateSerial(currentYear, currentMonth + 1, 0)
    End If
End Function

Function GetPayrollDeadline() As Date
    GetPayrollDeadline = DateAdd("d", 2, GetSubmissionDeadline())
End Function

Function GetCurrentPayPeriodPhase() As String
    Dim submissionDeadline As Date
    Dim payrollDeadline As Date
    Dim today As Date
    
    today = Date
    submissionDeadline = GetSubmissionDeadline()
    payrollDeadline = GetPayrollDeadline()
    
    If today < submissionDeadline Then
        GetCurrentPayPeriodPhase = "BEFORE_DEADLINE"
    ElseIf today = submissionDeadline Then
        GetCurrentPayPeriodPhase = "DEADLINE_DAY"
    ElseIf today > submissionDeadline And today <= payrollDeadline Then
        GetCurrentPayPeriodPhase = "APPROVAL_WINDOW"
    Else
        GetCurrentPayPeriodPhase = "PAST_PAYROLL"
    End If
End Function

Function DaysUntilPayroll() As Long
    DaysUntilPayroll = DateDiff("d", Date, GetPayrollDeadline())
    If DaysUntilPayroll < 0 Then DaysUntilPayroll = 0
End Function

' ==========================================
' SMART REMINDER RUNNER
' Automatically determines what reminders to send based on date
' ==========================================

Sub RunSmartReminders()
    Dim phase As String
    Dim submissionDeadline As Date
    Dim payrollDeadline As Date
    Dim daysToPayroll As Long
    Dim runChoice As VbMsgBoxResult
    
    phase = GetCurrentPayPeriodPhase()
    submissionDeadline = GetSubmissionDeadline()
    payrollDeadline = GetPayrollDeadline()
    daysToPayroll = DaysUntilPayroll()
    
    Dim statusMsg As String
    statusMsg = "PAY PERIOD STATUS" & vbCrLf & _
                "=================" & vbCrLf & vbCrLf & _
                "Today: " & Format(Date, "mm/dd/yyyy (dddd)") & vbCrLf & _
                "Submission Deadline: " & Format(submissionDeadline, "mm/dd/yyyy") & vbCrLf & _
                "Payroll Deadline: " & Format(payrollDeadline, "mm/dd/yyyy") & vbCrLf & _
                "Days Until Payroll: " & daysToPayroll & vbCrLf & vbCrLf
    
    Select Case phase
            
        Case "DEADLINE_DAY"
            statusMsg = statusMsg & "Phase: SUBMISSION DEADLINE DAY" & vbCrLf & _
                        "Recommended: Send submission reminders to non-submitters NOW."
            
            runChoice = MsgBox(statusMsg & vbCrLf & vbCrLf & _
                        "Run submission deadline reminders?", _
                        vbYesNo + vbExclamation, "Smart Reminders")
            
            If runChoice = vbYes Then
                Call UpdateSubmissionStatus
                Call SendDeadlineDayReminders
            End If
            
        Case "APPROVAL_WINDOW"
            statusMsg = statusMsg & "Phase: APPROVAL WINDOW (2-day turnaround)" & vbCrLf & _
                        "Recommended: Send approval reminders to ALL pending supervisors."
            
            runChoice = MsgBox(statusMsg & vbCrLf & vbCrLf & _
                        "Run approval window reminders?", _
                        vbYesNo + vbExclamation, "Smart Reminders")
            
            If runChoice = vbYes Then
                Call UpdateSubmissionStatus
                Call SendApprovalWindowReminders
            End If
            
        Case "PAST_PAYROLL"
            statusMsg = statusMsg & "Phase: PAST PAYROLL DEADLINE" & vbCrLf & _
                        "Recommended: Send urgent reminders and prepare for manual processing."
            
            runChoice = MsgBox(statusMsg & vbCrLf & vbCrLf & _
                        "Run past-deadline urgent reminders?", _
                        vbYesNo + vbCritical, "Smart Reminders")
            
            If runChoice = vbYes Then
                Call UpdateSubmissionStatus
                Call SendPastDeadlineReminders
            End If
    End Select
    
    MsgBox "Reminder process complete!", vbInformation
End Sub

' ==========================================
' DEADLINE DAY REMINDERS
' Sent on the 15th or last day of month
' To students who have NOT submitted
' ==========================================

Sub SendDeadlineDayReminders()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim studentName As String
    Dim studentEmail As String
    Dim supervisorName As String
    Dim supervisorEmail As String
    Dim employerName As String
    Dim payPeriod As String
    Dim reminderCount As Long
    Dim sendChoice As VbMsgBoxResult
    Dim ol As Outlook.Application
    Dim newMail As Outlook.mailItem
    Dim signature As String
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    sendChoice = MsgBox("TODAY IS THE SUBMISSION DEADLINE" & vbCrLf & vbCrLf & _
                        "This will send URGENT reminders to all students who have not submitted." & vbCrLf & _
                        "Both the student AND supervisor will receive the email." & vbCrLf & vbCrLf & _
                        "Continue?", vbYesNo + vbExclamation, "Deadline Day Reminders")
    
    If sendChoice = vbNo Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    Set ol = New Outlook.Application
    signature = LoadEmailSignature()
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If trackingSheet Is Nothing Then
        MsgBox "Submission_Tracking sheet not found.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    payPeriod = trackingSheet.cells(2, 1).Value
    reminderCount = 0
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            If trackingSheet.cells(i, 7).Value = "NO" Or trackingSheet.cells(i, 7).Value = "" Then
                
                studentName = trackingSheet.cells(i, 2).Value
                studentEmail = trackingSheet.cells(i, 3).Value
                employerName = trackingSheet.cells(i, 4).Value
                supervisorName = trackingSheet.cells(i, 5).Value
                supervisorEmail = trackingSheet.cells(i, 6).Value
                
                Set newMail = ol.CreateItem(olMailItem)
                
                With newMail
                    .Importance = olImportanceHigh
                    .To = studentEmail
                    .CC = supervisorEmail
                    .subject = "URGENT: Timesheet Due TODAY - " & studentName & " - " & payPeriod
                    
                    .htmlBody = "<div style='background-color:#D32F2F; color:white; padding:20px; margin-bottom:20px;'>" & _
                        "<h2 style='margin:0;'>YOUR TIMESHEET IS DUE TODAY</h2>" & _
                        "<p style='font-size:16px; margin:10px 0 0 0;'>Submit by 11:59 PM to avoid payment delay</p>" & _
                        "</div>" & _
                        "<p>Dear " & GetFirstName(studentName) & ",</p>" & _
                        "<p>Our records show you have <b>not yet submitted</b> your timesheet for this pay period. " & _
                        "<b>Today is the deadline.</b></p>" & _
                        "<div style='background-color:#FFEBEE; padding:15px; margin:15px 0; border:2px solid #D32F2F;'>" & _
                        "<p style='font-size:16px; margin:5px 0;'><b>Pay Period:</b> " & payPeriod & "</p>" & _
                        "<p style='font-size:16px; margin:5px 0;'><b>Employer:</b> " & employerName & "</p>" & _
                        "<p style='font-size:16px; margin:5px 0;'><b>Deadline:</b> TODAY by 11:59 PM</p>" & _
                        "</div>" & _
                        "<p><b>Submit your timesheet now:</b></p>" & _
                        "<ol style='font-size:14px;'>" & _
                        "<li>Log into Banner/ACES</li>" & _
                        "<li>Go to Employee Self-Service > Timekeeping</li>" & _
                        "<li>Enter your hours and click Submit</li>" & _
                        "</ol>" & _
                        "<p style='background-color:#FFCDD2; padding:10px;'><b>If you do not submit today, your payment will be delayed to the next pay cycle.</b></p>" & _
                        "<p><b>" & supervisorName & "</b> (CC'd): Please follow up with your intern to ensure timely submission.</p>" & _
                        "<p>Contact me immediately if you have questions or issues submitting.</p>" & signature
                    
                    On Error Resume Next
                    .SendUsingAccount = ol.Session.Accounts("DST-INTERNPAYROLL@district.edu")
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo ErrorHandler
                    
                    .Send
                    reminderCount = reminderCount + 1
                End With
                
                Set newMail = Nothing
            End If
        End If
    Next i
    
    MsgBox "Deadline Day Reminders Sent!" & vbCrLf & vbCrLf & _
           "Reminders sent: " & reminderCount, vbInformation
    
Cleanup:
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================
' APPROVAL WINDOW REMINDERS
' Sent on Day 1 or Day 2 after submission deadline
' To ALL supervisors with pending approvals
' ==========================================

Sub SendApprovalWindowReminders()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim sentLogSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim studentName As String
    Dim studentEmail As String
    Dim supervisorName As String
    Dim supervisorEmail As String
    Dim employerName As String
    Dim payPeriod As String
    Dim submittedDate As Date
    Dim hoursRemaining As Long
    Dim reminderCount As Long
    Dim sendChoice As VbMsgBoxResult
    Dim ol As Outlook.Application
    Dim newMail As Outlook.mailItem
    Dim signature As String
    Dim payrollDeadline As Date
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    payrollDeadline = GetPayrollDeadline()
    hoursRemaining = DateDiff("h", Now, payrollDeadline)
    If hoursRemaining < 0 Then hoursRemaining = 0
    
    sendChoice = MsgBox("APPROVAL WINDOW ACTIVE" & vbCrLf & vbCrLf & _
                        "Payroll deadline: " & Format(payrollDeadline, "mm/dd/yyyy") & vbCrLf & _
                        "Hours remaining: ~" & hoursRemaining & vbCrLf & vbCrLf & _
                        "This will send reminders to ALL supervisors with pending approvals." & vbCrLf & vbCrLf & _
                        "Continue?", vbYesNo + vbExclamation, "Approval Window Reminders")
    
    If sendChoice = vbNo Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    Set ol = New Outlook.Application
    signature = LoadEmailSignature()
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If trackingSheet Is Nothing Then
        MsgBox "Submission_Tracking sheet not found.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    payPeriod = trackingSheet.cells(2, 1).Value
    reminderCount = 0
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            ' Submitted but not approved
            If (trackingSheet.cells(i, 7).Value = "YES") And _
               (trackingSheet.cells(i, 9).Value = "PENDING" Or trackingSheet.cells(i, 9).Value = "" Or trackingSheet.cells(i, 9).Value = "NO") Then
                
                studentName = trackingSheet.cells(i, 2).Value
                studentEmail = trackingSheet.cells(i, 3).Value
                employerName = trackingSheet.cells(i, 4).Value
                supervisorName = trackingSheet.cells(i, 5).Value
                supervisorEmail = trackingSheet.cells(i, 6).Value
                
                On Error Resume Next
                submittedDate = trackingSheet.cells(i, 8).Value
                If Not IsDate(submittedDate) Then
                    submittedDate = FindSubmittedDate(sentLogSheet, studentEmail)
                End If
                On Error GoTo ErrorHandler
                
                Dim urgencyColor As String
                Dim urgencyText As String
                
                If hoursRemaining <= 24 Then
                    urgencyColor = "#D32F2F"
                    urgencyText = "URGENT - Less than 24 hours remaining"
                Else
                    urgencyColor = "#FF9800"
                    urgencyText = "Action Required - " & hoursRemaining & " hours until payroll deadline"
                End If
                
                Set newMail = ol.CreateItem(olMailItem)
                
                With newMail
                    If hoursRemaining <= 24 Then
                        .Importance = olImportanceHigh
                    End If
                    .To = supervisorEmail
                    .CC = studentEmail
                    .subject = "Approval Needed by " & Format(payrollDeadline, "mm/dd") & " - " & studentName & " Timesheet"
                    
                    .htmlBody = "<div style='background-color:" & urgencyColor & "; color:white; padding:15px; margin-bottom:15px;'>" & _
                        "<p style='font-size:16px; margin:0;'><b>" & urgencyText & "</b></p>" & _
                        "</div>" & _
                        "<p>Dear " & supervisorName & ",</p>" & _
                        "<p>A timesheet from your intern requires your approval before we can process payroll.</p>" & _
                        "<table style='border-collapse:collapse; margin:15px 0;'>" & _
                        "<tr><td style='padding:8px; border:1px solid #ddd; background-color:#f5f5f5;'><b>Student:</b></td><td style='padding:8px; border:1px solid #ddd;'>" & studentName & "</td></tr>" & _
                        "<tr><td style='padding:8px; border:1px solid #ddd; background-color:#f5f5f5;'><b>Pay Period:</b></td><td style='padding:8px; border:1px solid #ddd;'>" & payPeriod & "</td></tr>" & _
                        "<tr><td style='padding:8px; border:1px solid #ddd; background-color:#f5f5f5;'><b>Submitted:</b></td><td style='padding:8px; border:1px solid #ddd;'>" & Format(submittedDate, "mm/dd/yyyy") & "</td></tr>" & _
                        "<tr><td style='padding:8px; border:1px solid #ddd; background-color:#f5f5f5;'><b>Payroll Deadline:</b></td><td style='padding:8px; border:1px solid #ddd; color:#D32F2F;'><b>" & Format(payrollDeadline, "mm/dd/yyyy") & "</b></td></tr>" & _
                        "</table>" & _
                        "<div style='background-color:#E8F5E9; padding:15px; margin:15px 0; border-left:5px solid #4CAF50;'>" & _
                        "<p style='margin:0 0 10px 0;'><b>Please reply with ONE of the following:</b></p>" & _
                        "<p style='font-size:16px; margin:5px 0;'><b>APPROVED</b> - Hours are correct</p>" & _
                        "<p style='font-size:16px; margin:5px 0;'><b>REJECTED</b> - Hours need correction</p>" & _
                        "<p style='font-size:16px; margin:5px 0;'><b>CORRECTIONS</b> - Provide corrected hours</p>" & _
                        "</div>" & _
                        "<p>If you cannot locate the original approval request email, you may reply directly to this message.</p>" & _
                        "<p>Thank you for your prompt response.</p>" & signature
                    
                    On Error Resume Next
                    .SendUsingAccount = ol.Session.Accounts("DST-INTERNPAYROLL@district.edu")
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo ErrorHandler
                    
                    .Send
                    reminderCount = reminderCount + 1
                End With
                
                Set newMail = Nothing
            End If
        End If
    Next i
    
    MsgBox "Approval Window Reminders Sent!" & vbCrLf & vbCrLf & _
           "Reminders sent: " & reminderCount & vbCrLf & _
           "Payroll deadline: " & Format(payrollDeadline, "mm/dd/yyyy"), vbInformation
    
Cleanup:
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================
' PAST DEADLINE REMINDERS
' Sent when payroll deadline has passed
' Extremely urgent tone
' ==========================================

Sub SendPastDeadlineReminders()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim sentLogSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim studentName As String
    Dim studentEmail As String
    Dim supervisorName As String
    Dim supervisorEmail As String
    Dim employerName As String
    Dim payPeriod As String
    Dim reminderCount As Long
    Dim sendChoice As VbMsgBoxResult
    Dim ol As Outlook.Application
    Dim newMail As Outlook.mailItem
    Dim signature As String
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    sendChoice = MsgBox("PAYROLL DEADLINE HAS PASSED" & vbCrLf & vbCrLf & _
                        "This will send FINAL URGENT reminders to supervisors with outstanding approvals." & vbCrLf & _
                        "These students may have delayed payment." & vbCrLf & vbCrLf & _
                        "Continue?", vbYesNo + vbCritical, "Past Deadline Reminders")
    
    If sendChoice = vbNo Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    Set ol = New Outlook.Application
    signature = LoadEmailSignature()
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If trackingSheet Is Nothing Then
        MsgBox "Submission_Tracking sheet not found.", vbExclamation
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    payPeriod = trackingSheet.cells(2, 1).Value
    reminderCount = 0
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            If (trackingSheet.cells(i, 7).Value = "YES") And _
               (trackingSheet.cells(i, 9).Value = "PENDING" Or trackingSheet.cells(i, 9).Value = "" Or trackingSheet.cells(i, 9).Value = "NO") Then
                
                studentName = trackingSheet.cells(i, 2).Value
                studentEmail = trackingSheet.cells(i, 3).Value
                employerName = trackingSheet.cells(i, 4).Value
                supervisorName = trackingSheet.cells(i, 5).Value
                supervisorEmail = trackingSheet.cells(i, 6).Value
                
                Set newMail = ol.CreateItem(olMailItem)
                
                With newMail
                    .Importance = olImportanceHigh
                    .To = supervisorEmail
                    .CC = studentEmail
                    .subject = "FINAL NOTICE: Immediate Approval Required - " & studentName
                    
                    .htmlBody = "<div style='background-color:#B71C1C; color:white; padding:20px;'>" & _
                        "<h2 style='margin:0;'>PAYROLL DEADLINE PASSED - IMMEDIATE ACTION REQUIRED</h2>" & _
                        "<p style='margin:10px 0 0 0;'>This student's timesheet is now delayed pending your approval.</p>" & _
                        "</div>" & _
                        "<p>Dear " & supervisorName & ",</p>" & _
                        "<p>The payroll deadline has passed and we still have not received your approval for the following timesheet. " & _
                        "<b>The student's payment may be delayed if not approved before our payroll department processes timesheets.</b></p>" & _
                        "<table style='border-collapse:collapse; margin:15px 0; border:2px solid #B71C1C;'>" & _
                        "<tr style='background-color:#FFEBEE;'><td style='padding:12px; border:1px solid #ddd;'><b>Student:</b></td><td style='padding:12px; border:1px solid #ddd;'>" & studentName & "</td></tr>" & _
                        "<tr><td style='padding:12px; border:1px solid #ddd;'><b>Pay Period:</b></td><td style='padding:12px; border:1px solid #ddd;'>" & payPeriod & "</td></tr>" & _
                        "<tr style='background-color:#FFEBEE;'><td style='padding:12px; border:1px solid #ddd;'><b>Employer:</b></td><td style='padding:12px; border:1px solid #ddd;'>" & employerName & "</td></tr>" & _
                        "</table>" & _
                        "<div style='background-color:#FFCDD2; padding:15px; margin:15px 0; border:2px solid #B71C1C;'>" & _
                        "<p style='margin:0; font-size:16px;'><b>Please reply IMMEDIATELY with:</b></p>" & _
                        "<p style='font-size:18px; margin:10px 0;'><b>APPROVED</b>, <b>REJECTED</b>, or <b>CORRECTIONS</b></p>" & _
                        "</div>" & _
                        "<p>If I do not receive your response by end of business today, I will need to call you directly.</p>" & _
                        "<p>Contact me at <b>[PHONE NUMBER REDACTED]</b> if you have any questions.</p>" & signature
                    
                    On Error Resume Next
                    .SendUsingAccount = ol.Session.Accounts("DST-INTERNPAYROLL@district.edu")
                    If Err.Number <> 0 Then Err.Clear
                    On Error GoTo ErrorHandler
                    
                    .Send
                    reminderCount = reminderCount + 1
                End With
                
                Set newMail = Nothing
            End If
        End If
    Next i
    
    ' Also generate call list
    Call GenerateCallList
    
    MsgBox "Past Deadline Reminders Sent!" & vbCrLf & vbCrLf & _
           "Reminders sent: " & reminderCount & vbCrLf & vbCrLf & _
           "A call list has been generated in the Call_List sheet.", vbInformation
    
Cleanup:
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================
' GENERATE CALL LIST
' Creates a printable list of supervisors to call
' ==========================================

Sub GenerateCallList()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim callSheet As Object
    Dim lastRow As Long
    Dim callRow As Long
    Dim i As Long
    Dim studentName As String
    Dim supervisorName As String
    Dim supervisorEmail As String
    Dim employerName As String
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    
    ' Create or clear Call_List sheet
    On Error Resume Next
    Set callSheet = excelWorkbook.Sheets("Call_List")
    If callSheet Is Nothing Then
        Set callSheet = excelWorkbook.Sheets.Add(After:=excelWorkbook.Sheets(excelWorkbook.Sheets.count))
        callSheet.Name = "Call_List"
    Else
        callSheet.cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Header
    With callSheet.Range("A1:F1")
        .Merge
        .Value = "SUPERVISORS TO CALL - " & Format(Now, "mm/dd/yyyy hh:mm AM/PM")
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    callSheet.cells(3, 1).Value = "Supervisor Name"
    callSheet.cells(3, 2).Value = "Supervisor Email"
    callSheet.cells(3, 3).Value = "Employer"
    callSheet.cells(3, 4).Value = "Student(s) Pending"
    callSheet.cells(3, 5).Value = "Called?"
    callSheet.cells(3, 6).Value = "Notes"
    
    With callSheet.Range("A3:F3")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    callRow = 4
    
    ' Use dictionary to group by supervisor
    Dim supervisors As Object
    Set supervisors = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            If (trackingSheet.cells(i, 7).Value = "YES") And _
               (trackingSheet.cells(i, 9).Value = "PENDING" Or trackingSheet.cells(i, 9).Value = "" Or trackingSheet.cells(i, 9).Value = "NO") Then
                
                supervisorEmail = trackingSheet.cells(i, 6).Value
                studentName = trackingSheet.cells(i, 2).Value
                
                If supervisors.Exists(supervisorEmail) Then
                    supervisors(supervisorEmail) = supervisors(supervisorEmail) & ", " & studentName
                Else
                    supervisors.Add supervisorEmail, studentName
                    
                    callSheet.cells(callRow, 1).Value = trackingSheet.cells(i, 5).Value  ' Supervisor Name
                    callSheet.cells(callRow, 2).Value = supervisorEmail
                    callSheet.cells(callRow, 3).Value = trackingSheet.cells(i, 4).Value  ' Employer
                    callSheet.cells(callRow, 5).Value = "[ ]"
                    callRow = callRow + 1
                End If
            End If
        End If
    Next i
    
    ' Go back and fill in student names
    Dim key As Variant
    callRow = 4
    For Each key In supervisors.Keys
        callSheet.cells(callRow, 4).Value = supervisors(key)
        callRow = callRow + 1
    Next key
    
    callSheet.Columns("A:F").AutoFit
    callSheet.Columns("D").columnWidth = 40
    callSheet.Columns("F").columnWidth = 30
    
    excelWorkbook.Save
    callSheet.Activate
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating call list: " & Err.Description, vbCritical
End Sub

' ==========================================
' HELPER: LOAD EMAIL SIGNATURE
' ==========================================

Function LoadEmailSignature() As String
    Dim signaturePath As String
    Dim signatureFileName As String
    Dim fso As Object
    Dim signatureFile As Object
    
    signaturePath = Environ("APPDATA") & "\Microsoft\Signatures\"
    signatureFileName = signaturePath & "Alexis (DST-INTERNPAYROLL@district.edu).htm"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    LoadEmailSignature = ""
    
    If fso.FileExists(signatureFileName) Then
        On Error Resume Next
        Set signatureFile = fso.OpenTextFile(signatureFileName, 1)
        If Err.Number = 0 Then
            LoadEmailSignature = signatureFile.ReadAll
            signatureFile.Close
        End If
        On Error GoTo 0
    End If
End Function

' ==========================================
' HELPER: GET FIRST NAME FROM "LAST, FIRST" FORMAT
' ==========================================

Function GetFirstName(fullName As String) As String
    Dim parts As Variant
    
    If InStr(fullName, ",") > 0 Then
        parts = Split(fullName, ",")
        If UBound(parts) >= 1 Then
            GetFirstName = Trim(parts(1))
        Else
            GetFirstName = fullName
        End If
    Else
        GetFirstName = fullName
    End If
End Function

' ==========================================
' PAY PERIOD SUMMARY REPORT
' Generates a comprehensive summary before archiving
' ==========================================

Sub GeneratePayPeriodSummary()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim summarySheet As Object
    Dim trackingSheet As Object
    Dim sentLogSheet As Object
    Dim replyLogSheet As Object
    Dim supervisorSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim currentRow As Long
    Dim payPeriod As String
    
    ' Counters
    Dim totalStudents As Long
    Dim totalSubmitted As Long
    Dim totalNotSubmitted As Long
    Dim totalApproved As Long
    Dim totalRejected As Long
    Dim totalPending As Long
    Dim totalEmailsSent As Long
    Dim totalRepliesReceived As Long
    
    ' For supervisor stats
    Dim supervisorStats As Object
    Dim supervisorEmail As String
    Dim supervisorName As String
    Dim employerName As String
    
    ' For employer stats
    Dim employerStats As Object
    
    ' For response time calculation
    Dim totalResponseDays As Double
    Dim responseCount As Long
    Dim avgResponseDays As Double
    
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    
    ' Get sheets
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    Set sentLogSheet = excelWorkbook.Sheets("Sent_Log")
    Set replyLogSheet = excelWorkbook.Sheets("Reply_Log")
    Set supervisorSheet = excelWorkbook.Sheets("Updated_Supervisor email")
    On Error GoTo ErrorHandler
    
    If trackingSheet Is Nothing Then
        MsgBox "Submission_Tracking sheet not found.", vbExclamation
        GoTo Cleanup
    End If
    
    ' Get pay period
    payPeriod = trackingSheet.cells(2, 1).Value
    If payPeriod = "" Then
        payPeriod = "Unknown Pay Period"
    End If
    
    ' Create or clear Summary sheet
    On Error Resume Next
    Set summarySheet = excelWorkbook.Sheets("Pay_Period_Summary")
    If summarySheet Is Nothing Then
        Set summarySheet = excelWorkbook.Sheets.Add(Before:=excelWorkbook.Sheets(1))
        summarySheet.Name = "Pay_Period_Summary"
    Else
        summarySheet.cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Initialize dictionaries for stats
    Set supervisorStats = CreateObject("Scripting.Dictionary")
    Set employerStats = CreateObject("Scripting.Dictionary")
    
    ' ==========================================
    ' GATHER STATISTICS FROM TRACKING SHEET
    ' ==========================================
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    totalStudents = 0
    totalSubmitted = 0
    totalNotSubmitted = 0
    totalApproved = 0
    totalRejected = 0
    totalPending = 0
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            totalStudents = totalStudents + 1
            
            supervisorEmail = trackingSheet.cells(i, 6).Value
            supervisorName = trackingSheet.cells(i, 5).Value
            employerName = trackingSheet.cells(i, 4).Value
            
            ' Submission stats
            If trackingSheet.cells(i, 7).Value = "YES" Then
                totalSubmitted = totalSubmitted + 1
            Else
                totalNotSubmitted = totalNotSubmitted + 1
            End If
            
            ' Approval stats
            Select Case UCase(Trim(trackingSheet.cells(i, 9).Value))
                Case "YES"
                    totalApproved = totalApproved + 1
                Case "REJECTED"
                    totalRejected = totalRejected + 1
                Case Else
                    If trackingSheet.cells(i, 7).Value = "YES" Then
                        totalPending = totalPending + 1
                    End If
            End Select
            
            ' Employer stats
            If employerName <> "" Then
                If employerStats.Exists(employerName) Then
                    employerStats(employerName) = employerStats(employerName) + 1
                Else
                    employerStats.Add employerName, 1
                End If
            End If
            
            ' Supervisor stats (track pending per supervisor)
            If supervisorEmail <> "" Then
                Dim supKey As String
                supKey = supervisorEmail & "|" & supervisorName & "|" & employerName
                
                If Not supervisorStats.Exists(supKey) Then
                    ' Array: Total, Submitted, Approved, Rejected, Pending
                    supervisorStats.Add supKey, Array(0, 0, 0, 0, 0)
                End If
                
                Dim supData As Variant
                supData = supervisorStats(supKey)
                supData(0) = supData(0) + 1  ' Total
                
                If trackingSheet.cells(i, 7).Value = "YES" Then
                    supData(1) = supData(1) + 1  ' Submitted
                End If
                
                Select Case UCase(Trim(trackingSheet.cells(i, 9).Value))
                    Case "YES"
                        supData(2) = supData(2) + 1  ' Approved
                    Case "REJECTED"
                        supData(3) = supData(3) + 1  ' Rejected
                    Case Else
                        If trackingSheet.cells(i, 7).Value = "YES" Then
                            supData(4) = supData(4) + 1  ' Pending
                        End If
                End Select
                
                supervisorStats(supKey) = supData
            End If
        End If
    Next i
    
    ' ==========================================
    ' GATHER EMAIL STATISTICS
    ' ==========================================
    
    If Not sentLogSheet Is Nothing Then
        Dim sentLastRow As Long
        sentLastRow = sentLogSheet.cells(sentLogSheet.rows.count, 1).End(-4162).Row
        If sentLastRow > 1 Then
            totalEmailsSent = sentLastRow - 1
        End If
    End If
    
    If Not replyLogSheet Is Nothing Then
        Dim replyLastRow As Long
        replyLastRow = replyLogSheet.cells(replyLogSheet.rows.count, 1).End(-4162).Row
        If replyLastRow > 1 Then
            totalRepliesReceived = replyLastRow - 1
        End If
        
        ' Calculate average response time
        totalResponseDays = 0
        responseCount = 0
        
        For i = 2 To replyLastRow
            Dim replyDate As Date
            Dim studentEmail As String
            Dim sentDate As Date
            
            On Error Resume Next
            replyDate = replyLogSheet.cells(i, 5).Value
            studentEmail = replyLogSheet.cells(i, 7).Value  ' Lookup key contains student email
            On Error GoTo ErrorHandler
            
            If IsDate(replyDate) And studentEmail <> "" Then
                ' Find matching sent date
                sentDate = FindSubmittedDateFromLog(sentLogSheet, studentEmail)
                
                If IsDate(sentDate) And sentDate > 0 Then
                    Dim daysDiff As Double
                    daysDiff = replyDate - sentDate
                    
                    If daysDiff >= 0 And daysDiff <= 14 Then
                        totalResponseDays = totalResponseDays + daysDiff
                        responseCount = responseCount + 1
                    End If
                End If
            End If
        Next i
        
        If responseCount > 0 Then
            avgResponseDays = Round(totalResponseDays / responseCount, 1)
        End If
    End If
    
    ' ==========================================
    ' BUILD SUMMARY REPORT
    ' ==========================================
    
    currentRow = 1
    
    ' Title
    With summarySheet.Range("A" & currentRow & ":H" & currentRow)
        .Merge
        .Value = "PAY PERIOD SUMMARY REPORT"
        .Font.Size = 18
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = -4108
        .RowHeight = 30
    End With
    
    currentRow = currentRow + 1
    
    ' Pay period and generation date
    summarySheet.cells(currentRow, 1).Value = "Pay Period:"
    summarySheet.cells(currentRow, 2).Value = payPeriod
    summarySheet.cells(currentRow, 1).Font.Bold = True
    
    summarySheet.cells(currentRow, 4).Value = "Generated:"
    summarySheet.cells(currentRow, 5).Value = Format(Now, "mm/dd/yyyy hh:mm AM/PM")
    summarySheet.cells(currentRow, 4).Font.Bold = True
    
    currentRow = currentRow + 2
    
    ' ==========================================
    ' SECTION 1: OVERALL STATISTICS
    ' ==========================================
    
    With summarySheet.Range("A" & currentRow & ":D" & currentRow)
        .Merge
        .Value = "OVERALL STATISTICS"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    currentRow = currentRow + 1
    
    ' Headers
    summarySheet.cells(currentRow, 1).Value = "Metric"
    summarySheet.cells(currentRow, 2).Value = "Count"
    summarySheet.cells(currentRow, 3).Value = "Percentage"
    summarySheet.cells(currentRow, 4).Value = "Status"
    
    With summarySheet.Range("A" & currentRow & ":D" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    currentRow = currentRow + 1
    
    ' Total Students
    summarySheet.cells(currentRow, 1).Value = "Total Students"
    summarySheet.cells(currentRow, 2).Value = totalStudents
    summarySheet.cells(currentRow, 3).Value = "100%"
    summarySheet.cells(currentRow, 4).Value = "-"
    currentRow = currentRow + 1
    
    ' Submitted
    summarySheet.cells(currentRow, 1).Value = "Timesheets Submitted"
    summarySheet.cells(currentRow, 2).Value = totalSubmitted
    If totalStudents > 0 Then
        summarySheet.cells(currentRow, 3).Value = Round((totalSubmitted / totalStudents) * 100, 1) & "%"
    End If
    If totalSubmitted = totalStudents Then
        summarySheet.cells(currentRow, 4).Value = "COMPLETE"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(200, 255, 200)
    Else
        summarySheet.cells(currentRow, 4).Value = totalNotSubmitted & " MISSING"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(255, 255, 200)
    End If
    currentRow = currentRow + 1
    
    ' Not Submitted
    summarySheet.cells(currentRow, 1).Value = "Not Submitted"
    summarySheet.cells(currentRow, 2).Value = totalNotSubmitted
    If totalStudents > 0 Then
        summarySheet.cells(currentRow, 3).Value = Round((totalNotSubmitted / totalStudents) * 100, 1) & "%"
    End If
    If totalNotSubmitted > 0 Then
        summarySheet.cells(currentRow, 4).Value = "ACTION NEEDED"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(255, 200, 200)
    Else
        summarySheet.cells(currentRow, 4).Value = "OK"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(200, 255, 200)
    End If
    currentRow = currentRow + 1
    
    ' Approved
    summarySheet.cells(currentRow, 1).Value = "Approved"
    summarySheet.cells(currentRow, 2).Value = totalApproved
    If totalSubmitted > 0 Then
        summarySheet.cells(currentRow, 3).Value = Round((totalApproved / totalSubmitted) * 100, 1) & "%"
    End If
    If totalApproved = totalSubmitted And totalSubmitted > 0 Then
        summarySheet.cells(currentRow, 4).Value = "COMPLETE"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(200, 255, 200)
    Else
        summarySheet.cells(currentRow, 4).Value = (totalSubmitted - totalApproved) & " PENDING"
    End If
    currentRow = currentRow + 1
    
    ' Rejected
    summarySheet.cells(currentRow, 1).Value = "Rejected"
    summarySheet.cells(currentRow, 2).Value = totalRejected
    If totalSubmitted > 0 Then
        summarySheet.cells(currentRow, 3).Value = Round((totalRejected / totalSubmitted) * 100, 1) & "%"
    End If
    If totalRejected > 0 Then
        summarySheet.cells(currentRow, 4).Value = "NEEDS RESUBMISSION"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(255, 200, 200)
    Else
        summarySheet.cells(currentRow, 4).Value = "OK"
    End If
    currentRow = currentRow + 1
    
    ' Pending Approval
    summarySheet.cells(currentRow, 1).Value = "Pending Approval"
    summarySheet.cells(currentRow, 2).Value = totalPending
    If totalSubmitted > 0 Then
        summarySheet.cells(currentRow, 3).Value = Round((totalPending / totalSubmitted) * 100, 1) & "%"
    End If
    If totalPending > 0 Then
        summarySheet.cells(currentRow, 4).Value = "AWAITING RESPONSE"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(255, 255, 200)
    Else
        summarySheet.cells(currentRow, 4).Value = "OK"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(200, 255, 200)
    End If
    currentRow = currentRow + 2
    
    ' ==========================================
    ' SECTION 2: EMAIL STATISTICS
    ' ==========================================
    
    With summarySheet.Range("A" & currentRow & ":D" & currentRow)
        .Merge
        .Value = "EMAIL STATISTICS"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Metric"
    summarySheet.cells(currentRow, 2).Value = "Value"
    
    With summarySheet.Range("A" & currentRow & ":B" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Approval Requests Sent"
    summarySheet.cells(currentRow, 2).Value = totalEmailsSent
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Supervisor Replies Received"
    summarySheet.cells(currentRow, 2).Value = totalRepliesReceived
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Response Rate"
    If totalEmailsSent > 0 Then
        summarySheet.cells(currentRow, 2).Value = Round((totalRepliesReceived / totalEmailsSent) * 100, 1) & "%"
    Else
        summarySheet.cells(currentRow, 2).Value = "N/A"
    End If
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Average Response Time"
    If avgResponseDays > 0 Then
        summarySheet.cells(currentRow, 2).Value = avgResponseDays & " days"
    Else
        summarySheet.cells(currentRow, 2).Value = "N/A"
    End If
    currentRow = currentRow + 2
    
    ' ==========================================
    ' SECTION 3: EMPLOYER BREAKDOWN
    ' ==========================================
    
    With summarySheet.Range("A" & currentRow & ":C" & currentRow)
        .Merge
        .Value = "STUDENTS BY EMPLOYER"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With
    
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Employer"
    summarySheet.cells(currentRow, 2).Value = "Student Count"
    summarySheet.cells(currentRow, 3).Value = "Percentage"
    
    With summarySheet.Range("A" & currentRow & ":C" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    currentRow = currentRow + 1
    
    ' Sort employers by count (descending)
    Dim empKey As Variant
    Dim empSorted As Object
    Set empSorted = CreateObject("Scripting.Dictionary")
    
    For Each empKey In employerStats.Keys
        empSorted.Add empKey, employerStats(empKey)
    Next empKey
    
    ' Output employer stats
    For Each empKey In empSorted.Keys
        summarySheet.cells(currentRow, 1).Value = empKey
        summarySheet.cells(currentRow, 2).Value = empSorted(empKey)
        If totalStudents > 0 Then
            summarySheet.cells(currentRow, 3).Value = Round((empSorted(empKey) / totalStudents) * 100, 1) & "%"
        End If
        currentRow = currentRow + 1
    Next empKey
    
    currentRow = currentRow + 1
    
  ' ==========================================
    ' SECTION 4: SUPERVISORS WITH PENDING APPROVALS
    ' ==========================================
    
    With summarySheet.Range("A" & currentRow & ":G" & currentRow)
        .Merge
        .Value = "SUPERVISORS WITH PENDING APPROVALS"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(255, 200, 200)
    End With
    
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Supervisor"
    summarySheet.cells(currentRow, 2).Value = "Employer"
    summarySheet.cells(currentRow, 3).Value = "Email"
    summarySheet.cells(currentRow, 4).Value = "Total"
    summarySheet.cells(currentRow, 5).Value = "Approved"
    summarySheet.cells(currentRow, 6).Value = "Pending"
    summarySheet.cells(currentRow, 7).Value = "Pending Students"
    
    With summarySheet.Range("A" & currentRow & ":G" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    currentRow = currentRow + 1
    
    Dim supKey2 As Variant
    Dim hasPending As Boolean
    hasPending = False
    
    For Each supKey2 In supervisorStats.Keys
        Dim supParts As Variant
        supParts = Split(supKey2, "|")
        
        Dim supStats As Variant
        supStats = supervisorStats(supKey2)
        
        If supStats(4) > 0 Then
            hasPending = True
            
            ' Collect pending student names for this supervisor
            Dim pendingStudents As String
            pendingStudents = ""
            Dim lastRowSec4 As Long
            lastRowSec4 = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
            
            Dim j As Long
            For j = 2 To lastRowSec4
                If Trim(trackingSheet.cells(j, 2).Value) <> "" Then
                    If trackingSheet.cells(j, 6).Value = supParts(0) Then
                        If trackingSheet.cells(j, 7).Value = "YES" Then
                            If UCase(Trim(trackingSheet.cells(j, 9).Value)) <> "YES" And _
                               UCase(Trim(trackingSheet.cells(j, 9).Value)) <> "REJECTED" Then
                                If pendingStudents = "" Then
                                    pendingStudents = trackingSheet.cells(j, 2).Value
                                Else
                                    pendingStudents = pendingStudents & ", " & trackingSheet.cells(j, 2).Value
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
            
            summarySheet.cells(currentRow, 1).Value = supParts(1)
            summarySheet.cells(currentRow, 2).Value = supParts(2)
            summarySheet.cells(currentRow, 3).Value = supParts(0)
            summarySheet.cells(currentRow, 4).Value = supStats(0)
            summarySheet.cells(currentRow, 5).Value = supStats(2)
            summarySheet.cells(currentRow, 6).Value = supStats(4)
            summarySheet.cells(currentRow, 7).Value = pendingStudents
            
            summarySheet.Range("A" & currentRow & ":G" & currentRow).Interior.Color = RGB(255, 230, 230)
            
            ' Wrap text on the student names cell in case there are many
            summarySheet.cells(currentRow, 7).WrapText = True
            
            currentRow = currentRow + 1
        End If
    Next supKey2
    
    If Not hasPending Then
        summarySheet.cells(currentRow, 1).Value = "No pending approvals - all complete!"
        summarySheet.Range("A" & currentRow & ":G" & currentRow).Merge
        summarySheet.Range("A" & currentRow & ":G" & currentRow).Interior.Color = RGB(200, 255, 200)
        currentRow = currentRow + 1
    End If
    
    currentRow = currentRow + 1
    
    ' ==========================================
    ' SECTION 5: STUDENTS WHO DID NOT SUBMIT
    ' ==========================================
    
    With summarySheet.Range("A" & currentRow & ":D" & currentRow)
        .Merge
        .Value = "STUDENTS WHO DID NOT SUBMIT"
        .Font.Size = 14
        .Font.Bold = True
        .Interior.Color = RGB(255, 200, 200)
    End With
    
    currentRow = currentRow + 1
    
    summarySheet.cells(currentRow, 1).Value = "Student Name"
    summarySheet.cells(currentRow, 2).Value = "Employer"
    summarySheet.cells(currentRow, 3).Value = "Supervisor"
    summarySheet.cells(currentRow, 4).Value = "Student Email"
    
    With summarySheet.Range("A" & currentRow & ":D" & currentRow)
        .Font.Bold = True
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    currentRow = currentRow + 1
    
    Dim hasNonSubmitters As Boolean
    hasNonSubmitters = False
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    For i = 2 To lastRow
        If Trim(trackingSheet.cells(i, 2).Value) <> "" Then
            If trackingSheet.cells(i, 7).Value = "NO" Or trackingSheet.cells(i, 7).Value = "" Then
                hasNonSubmitters = True
                summarySheet.cells(currentRow, 1).Value = trackingSheet.cells(i, 2).Value
                summarySheet.cells(currentRow, 2).Value = trackingSheet.cells(i, 4).Value
                summarySheet.cells(currentRow, 3).Value = trackingSheet.cells(i, 5).Value
                summarySheet.cells(currentRow, 4).Value = trackingSheet.cells(i, 3).Value
                
                summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(255, 230, 230)
                
                currentRow = currentRow + 1
            End If
        End If
    Next i
    
    If Not hasNonSubmitters Then
        summarySheet.cells(currentRow, 1).Value = "All students submitted - none missing!"
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Merge
        summarySheet.Range("A" & currentRow & ":D" & currentRow).Interior.Color = RGB(200, 255, 200)
        currentRow = currentRow + 1
    End If
    
    currentRow = currentRow + 2
    
    ' ==========================================
    ' FOOTER
    ' ==========================================
    
    summarySheet.cells(currentRow, 1).Value = "Report generated by ApprovalFlow v2.0"
    summarySheet.cells(currentRow, 1).Font.Italic = True
    summarySheet.cells(currentRow, 1).Font.Color = RGB(128, 128, 128)
    
    ' Auto-fit columns
    summarySheet.Columns("A:H").AutoFit
    
    ' Move summary to first position
    summarySheet.Move Before:=excelWorkbook.Sheets(1)
    
    excelWorkbook.Save
    summarySheet.Activate
    summarySheet.Range("A1").Select
    
    MsgBox "Pay Period Summary Generated!" & vbCrLf & vbCrLf & _
           "Pay Period: " & payPeriod & vbCrLf & _
           "Total Students: " & totalStudents & vbCrLf & _
           "Submitted: " & totalSubmitted & vbCrLf & _
           "Approved: " & totalApproved & vbCrLf & _
           "Pending: " & totalPending, vbInformation
    
    Exit Sub
    
Cleanup:
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating summary: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

' ==========================================
' HELPER: Find submitted date from log
' ==========================================

Function FindSubmittedDateFromLog(sentLogSheet As Object, lookupKey As String) As Date
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    
    On Error GoTo ErrorHandler
    
    FindSubmittedDateFromLog = 0
    
    If sentLogSheet Is Nothing Then Exit Function
    
    lastRow = sentLogSheet.cells(sentLogSheet.rows.count, 8).End(-4162).Row
    
    For i = lastRow To 2 Step -1
        cellValue = LCase(Trim(sentLogSheet.cells(i, 8).Value))
        
        If InStr(cellValue, LCase(Left(lookupKey, InStr(lookupKey, "|") - 1))) > 0 Then
            If IsDate(sentLogSheet.cells(i, 7).Value) Then
                FindSubmittedDateFromLog = sentLogSheet.cells(i, 7).Value
                Exit Function
            End If
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    FindSubmittedDateFromLog = 0
End Function

' ==========================================
' HELPER: FIND SUBMITTED DATE FROM SENT_LOG
' ==========================================

Function FindSubmittedDate(sentLogSheet As Object, studentEmail As String) As Date
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim acesID As String
    
    On Error GoTo ErrorHandler
    
    FindSubmittedDate = 0
    
    If sentLogSheet Is Nothing Then Exit Function
    
    ' Extract ACES ID from student email
    If InStr(studentEmail, "@") > 0 Then
        acesID = LCase(Left(studentEmail, InStr(studentEmail, "@") - 1))
    Else
        acesID = LCase(studentEmail)
    End If
    
    lastRow = sentLogSheet.cells(sentLogSheet.rows.count, 6).End(-4162).Row
    
    ' Search backwards (most recent first)
    For i = lastRow To 2 Step -1
        cellValue = LCase(Trim(sentLogSheet.cells(i, 6).Value))
        
        If InStr(cellValue, acesID) > 0 Then
            If IsDate(sentLogSheet.cells(i, 7).Value) Then
                FindSubmittedDate = sentLogSheet.cells(i, 7).Value
                Exit Function
            End If
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    FindSubmittedDate = 0
End Function


' ==========================================
' ENHANCED ARCHIVE WITH SUMMARY AND BACKUP
' ==========================================

Sub ArchivePayPeriodComplete()
    Dim archiveChoice As VbMsgBoxResult
    Dim backupSuccess As Boolean
    
    archiveChoice = MsgBox("COMPLETE PAY PERIOD ARCHIVE" & vbCrLf & _
                           "===========================" & vbCrLf & vbCrLf & _
                           "This will:" & vbCrLf & _
                           "1. Generate Pay Period Summary Report" & vbCrLf & _
                           "2. Archive workbook to local folder" & vbCrLf & vbCrLf & _
                           "Continue?", vbYesNo + vbQuestion, "Complete Archive")
    
    If archiveChoice = vbNo Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Step 1: Generate Summary
    MsgBox "Step 1 of 3: Generating Pay Period Summary...", vbInformation
    Call GeneratePayPeriodSummary
    
    ' Step 2: Local Archive
    MsgBox "Step 3 of 3: Creating local archive...", vbInformation
    Call ArchivePayPeriodWorkbook
    
    MsgBox "Complete Archive Finished!" & vbCrLf & vbCrLf & _
           "Summary: Generated" & vbCrLf & _
           "Local Archive: Complete", vbInformation
End Sub

' ==========================================
' ARCHIVE PAY PERIOD WORKBOOK
' Copies entire workbook to archive folder
' ==========================================

Sub ArchivePayPeriodWorkbook()
    Dim fso As Object
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim archiveFolder As String
    Dim archiveFileName As String
    Dim payPeriod As String
    Dim archiveChoice As VbMsgBoxResult
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    archiveFolder = "" & Environ("USERPROFILE") & "\Desktop\ApprovalFlow_Archive\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create archive folder if it does not exist
    If Not fso.FolderExists(archiveFolder) Then
        fso.CreateFolder archiveFolder
    End If
    
    ' Get pay period from workbook
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    
    On Error Resume Next
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    If Not trackingSheet Is Nothing Then
        payPeriod = trackingSheet.cells(2, 1).Value
    End If
    On Error GoTo ErrorHandler
    
    excelWorkbook.Close False
    excelApp.Quit
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    
    If payPeriod = "" Then
        payPeriod = InputBox("Enter pay period for archive file name:", "Pay Period", Format(Date, "mmm_yyyy"))
        If payPeriod = "" Then
            MsgBox "Cancelled.", vbInformation
            Exit Sub
        End If
    End If
    
    archiveFileName = "ApprovalFlow_" & CleanFileName(payPeriod) & ".xlsx"
    
    ' Check if archive already exists
    If fso.FileExists(archiveFolder & archiveFileName) Then
        archiveChoice = MsgBox("Archive already exists:" & vbCrLf & archiveFileName & vbCrLf & vbCrLf & _
                               "Overwrite?", vbYesNo + vbExclamation, "Archive Exists")
        If archiveChoice = vbNo Then
            MsgBox "Cancelled.", vbInformation
            Exit Sub
        End If
    End If
    
    ' Copy file to archive
    fso.CopyFile excelFilePath, archiveFolder & archiveFileName, True
    
    MsgBox "Archive Created!" & vbCrLf & vbCrLf & _
           "File: " & archiveFileName & vbCrLf & _
           "Location: " & archiveFolder, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error archiving workbook: " & Err.Description, vbCritical
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub



Sub StartNewPayPeriodComplete()
    Dim startChoice As VbMsgBoxResult
    
    startChoice = MsgBox("START NEW PAY PERIOD (COMPLETE)" & vbCrLf & _
                         "================================" & vbCrLf & vbCrLf & _
                         "This will run the full end-of-period workflow:" & vbCrLf & vbCrLf & _
                         "1. Generate Pay Period Summary" & vbCrLf & _
                         "2. Archive current workbook locally" & vbCrLf & _
                         "3. Clear all logs" & vbCrLf & _
                         "4. Populate Submission_Tracking for new period" & vbCrLf & vbCrLf & _
                         "Continue?", vbYesNo + vbQuestion, "Start New Pay Period")
    
    If startChoice = vbNo Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Step 1: Generate Summary
    MsgBox "Step 1 of 5: Generating Pay Period Summary...", vbInformation
    Call GeneratePayPeriodSummary
    
    ' Step 2: Local Archive
    MsgBox "Step 3 of 5: Creating local archive...", vbInformation
    Call ArchivePayPeriodWorkbook
    
    ' Step 3: Clear logs
    MsgBox "Step 4 of 5: Clearing logs for new period...", vbInformation
    Call ClearLogsForNewPeriod
    
    ' Step 4: Populate new period
    MsgBox "Step 5 of 5: Populating Submission_Tracking for new period...", vbInformation
    Call PopulateSubmissionTracking
    
    MsgBox "New Pay Period Setup Complete!" & vbCrLf & vbCrLf & _
           "Your system is ready for the new pay period.", vbInformation
End Sub
' ==========================================
' STRIP HTML TAGS FROM STRING
' ==========================================

Function StripHTML(htmlText As String) As String
    Dim result As String
    Dim inTag As Boolean
    Dim i As Long
    Dim char As String
    
    result = ""
    inTag = False
    
    For i = 1 To Len(htmlText)
        char = Mid(htmlText, i, 1)
        
        If char = "<" Then
            inTag = True
        ElseIf char = ">" Then
            inTag = False
        ElseIf Not inTag Then
            result = result & char
        End If
    Next i
    
    StripHTML = Trim(result)
End Function

Sub SetPayPeriodOverride()
    Dim inputDate As String
    inputDate = InputBox("Enter a date within the pay period you want to view:" & vbCrLf & vbCrLf & _
                         "Examples:" & vbCrLf & _
                         "  1/20/2026 for Jan 16-31" & vbCrLf & _
                         "  1/5/2026 for Jan 1-15" & vbCrLf & _
                         "  12/20/2025 for Dec 16-31", _
                         "Set Pay Period Override")
    
    If inputDate = "" Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If
    
    If IsDate(inputDate) Then
        overridePayPeriodDate = CDate(inputDate)
        MsgBox "Pay period set to: " & GetPayPeriodLabel() & vbCrLf & vbCrLf & _
               "Start: " & GetPayPeriodStartDate() & vbCrLf & _
               "End: " & GetPayPeriodEndDate(), vbInformation
    Else
        MsgBox "Invalid date. Please try again.", vbExclamation
    End If
End Sub

Sub ClearPayPeriodOverride()
    overridePayPeriodDate = 0
    MsgBox "Override cleared. Now using current date." & vbCrLf & vbCrLf & _
           "Current pay period: " & GetPayPeriodLabel(), vbInformation
End Sub


Sub DeduplicateSentLog()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim sentSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim studentEmail As String
    Dim compareEmail As String
    Dim sentDate As Date
    Dim compareDate As Date
    Dim removedCount As Long
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set sentSheet = excelWorkbook.Sheets("Sent_Log")
    
    lastRow = sentSheet.cells(sentSheet.rows.count, 6).End(-4162).Row
    removedCount = 0
    
    ' Work backwards so row deletion does not shift indices
    For i = lastRow To 2 Step -1
        studentEmail = LCase(Trim(sentSheet.cells(i, 6).Value))
        
        If studentEmail <> "" Then
            On Error Resume Next
            sentDate = sentSheet.cells(i, 7).Value
            On Error GoTo ErrorHandler
            
            ' Look for a newer entry for the same student
            For j = i + 1 To lastRow
                compareEmail = LCase(Trim(sentSheet.cells(j, 6).Value))
                
                If compareEmail = studentEmail Then
                    On Error Resume Next
                    compareDate = sentSheet.cells(j, 7).Value
                    On Error GoTo ErrorHandler
                    
                    ' If row j is newer, delete row i (current is older)
                    If IsDate(compareDate) And IsDate(sentDate) Then
                        If compareDate > sentDate Then
                            sentSheet.rows(i).Delete
                            removedCount = removedCount + 1
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    excelWorkbook.Save
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "Sent_Log deduplicated!" & vbCrLf & _
           "Duplicate rows removed: " & removedCount, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub
                
 Sub DeduplicateReplyLog()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim replySheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim lookupKey As String
    Dim compareLookupKey As String
    Dim replyDate As Date
    Dim compareDate As Date
    Dim removedCount As Long
    Dim excelFilePath As String
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    
    lastRow = replySheet.cells(replySheet.rows.count, 1).End(-4162).Row
    removedCount = 0
    
    For i = lastRow To 2 Step -1
        lookupKey = LCase(Trim(replySheet.cells(i, 7).Value))
        
        If lookupKey <> "" And Not InStr(lookupKey, "unmatched") > 0 Then
            On Error Resume Next
            replyDate = replySheet.cells(i, 5).Value
            On Error GoTo ErrorHandler
            
            For j = i + 1 To lastRow
                compareLookupKey = LCase(Trim(replySheet.cells(j, 7).Value))
                
                If compareLookupKey = lookupKey Then
                    On Error Resume Next
                    compareDate = replySheet.cells(j, 5).Value
                    On Error GoTo ErrorHandler
                    
                    If IsDate(compareDate) And IsDate(replyDate) Then
                        If compareDate > replyDate Then
                            replySheet.rows(i).Delete
                            removedCount = removedCount + 1
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If
    Next i
    
    excelWorkbook.Save
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "Reply_Log deduplicated!" & vbCrLf & _
           "Duplicate rows removed: " & removedCount, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    On Error Resume Next
    If Not excelWorkbook Is Nothing Then excelWorkbook.Close False
    If Not excelApp Is Nothing Then excelApp.Quit
End Sub
Function FindBestMatchForReply(oMail As Outlook.mailItem, sentLogSheet As Object, sentLastRow As Long) As Long
    Dim supervisorEmail As String
    Dim cleanSupervisorEmail As String
    Dim sentRow As Long
    Dim sentDate As Date
    Dim timeDiff As Double
    Dim studentNameFromSubject As String
    Dim studentEmailFromCC As String
    Dim sentStudentName As String
    Dim sentStudentEmail As String
    Dim bestMatchRow As Long
    Dim bestMatchScore As Long
    Dim currentScore As Long
    
    bestMatchRow = 0
    bestMatchScore = 0
    
    supervisorEmail = GetSMTPAddress(oMail)
    cleanSupervisorEmail = LCase(Trim(supervisorEmail))
    
    studentNameFromSubject = ExtractStudentNameFromReplySubject(oMail.subject)
    
    studentEmailFromCC = ""
    If oMail.CC <> "" Then
        studentEmailFromCC = LCase(Trim(ExtractEmailAddress(oMail.CC)))
    End If
    
    Dim studentNameFromBody As String
    studentNameFromBody = ExtractStudentNameFromBody(oMail.Body)
    
    For sentRow = 2 To sentLastRow
        currentScore = 0
        
        Dim sentSupEmail As String
        sentSupEmail = LCase(Trim(sentLogSheet.cells(sentRow, 3).Value))
        
        If InStr(1, sentSupEmail, cleanSupervisorEmail, vbTextCompare) > 0 Or _
           InStr(1, cleanSupervisorEmail, sentSupEmail, vbTextCompare) > 0 Then
            
            currentScore = 1
            
            On Error Resume Next
            sentDate = sentLogSheet.cells(sentRow, 7).Value
            On Error GoTo 0
            
            If IsDate(sentDate) Then
                timeDiff = oMail.ReceivedTime - sentDate
                
                If timeDiff > 0 And timeDiff <= 30 Then
                    
                    If timeDiff <= 3 Then
                        currentScore = currentScore + 5
                    ElseIf timeDiff <= 7 Then
                        currentScore = currentScore + 3
                    ElseIf timeDiff <= 14 Then
                        currentScore = currentScore + 1
                    End If
                    
                    sentStudentName = sentLogSheet.cells(sentRow, 1).Value
                    sentStudentEmail = LCase(Trim(sentLogSheet.cells(sentRow, 6).Value))
                    
                    If sentStudentName = "Name Not Found" Or sentStudentName = "" Then
                        currentScore = 0
                    Else
                        If studentEmailFromCC <> "" And sentStudentEmail <> "" Then
                            If InStr(1, sentStudentEmail, studentEmailFromCC, vbTextCompare) > 0 Or _
                               InStr(1, studentEmailFromCC, sentStudentEmail, vbTextCompare) > 0 Then
                                currentScore = currentScore + 20
                            End If
                        End If
                        
                        If studentNameFromSubject <> "" Then
                            If InStr(1, LCase(sentStudentName), LCase(studentNameFromSubject), vbTextCompare) > 0 Or _
                               InStr(1, LCase(studentNameFromSubject), LCase(sentStudentName), vbTextCompare) > 0 Then
                                currentScore = currentScore + 15
                            End If
                        End If
                        
                        If studentNameFromBody <> "" Then
                            If InStr(1, LCase(sentStudentName), LCase(studentNameFromBody), vbTextCompare) > 0 Or _
                               InStr(1, LCase(studentNameFromBody), LCase(sentStudentName), vbTextCompare) > 0 Then
                                currentScore = currentScore + 10
                            End If
                        End If
                    End If
                    
                    If currentScore > bestMatchScore Then
                        bestMatchScore = currentScore
                        bestMatchRow = sentRow
                    End If
                End If
            End If
        End If
    Next sentRow
    
    FindBestMatchForReply = bestMatchRow
End Function

Function GetPayPeriodCutoffStart() As Date
    Dim startDate As Date
    startDate = GetPayPeriodStartDate()
    
    ' Allow emails from 5 days before the period start to catch early submissions
    GetPayPeriodCutoffStart = DateAdd("d", -5, startDate)
End Function

Function GetPayPeriodCutoffEnd() As Date
    Dim endDate As Date
    endDate = GetPayPeriodEndDate()
    
    ' Allow emails up to 4 days after period end to catch late approvals
    GetPayPeriodCutoffEnd = DateAdd("d", 9, endDate)
End Function


Sub RefreshPayPeriodSummary()
    Call UpdateSubmissionStatus
    Call GeneratePayPeriodSummary
    MsgBox "Summary refreshed for " & GetPayPeriodLabel(), vbInformation
End Sub

Function GetWorkbookPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Try current user Desktop first - works for any successor
    Dim dynamicPath As String
    dynamicPath = Environ("USERPROFILE") & "\Desktop\Student_Sup_email.xlsx"
    
    If fso.FileExists(dynamicPath) Then
        GetWorkbookPath = dynamicPath
        Exit Function
    End If
    
   ' Fallback: prompt user to locate the file manually
    Dim fallbackPath As String
    fallbackPath = Environ("USERPROFILE") & "\OneDrive\Desktop\Student_Sup_email.xlsx"
    
    If fso.FileExists(fallbackPath) Then
        GetWorkbookPath = fallbackPath
        Exit Function
    End If
    
    MsgBox "Student_Sup_email.xlsx not found on Desktop." & vbCrLf & _
           "Please ensure the file is saved to your Desktop.", vbCritical
    GetWorkbookPath = ""
End Function

Function FastGetStudentCCEmail(employeeEmail As String) As String
    Dim cleanEmail As String
    Dim acesID As String
    Dim data As Variant
    
    cleanEmail = LCase(Trim(ExtractEmailAddress(employeeEmail)))
    
    If InStr(cleanEmail, "@") > 0 Then
        acesID = Left(cleanEmail, InStr(cleanEmail, "@") - 1)
    Else
        acesID = cleanEmail
    End If
    
    If Not lookupData Is Nothing Then
        If lookupData.Exists(acesID) Then
            data = lookupData(acesID)
            If UBound(data) >= 4 Then
                FastGetStudentCCEmail = data(4)
                Exit Function
            End If
        End If
        If lookupData.Exists(cleanEmail) Then
            data = lookupData(cleanEmail)
            If UBound(data) >= 4 Then
                FastGetStudentCCEmail = data(4)
                Exit Function
            End If
        End If
    End If
    
    FastGetStudentCCEmail = ""
End Function

' ==========================================
' EMAIL EXTRACTION FUNCTIONS
' ==========================================

Function ExtractEmailAddress(ByVal inputString As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    On Error GoTo SimpleExtract
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
    regEx.Global = False
    regEx.IgnoreCase = True
    
    If regEx.Test(inputString) Then
        Set matches = regEx.Execute(inputString)
        ExtractEmailAddress = matches(0).Value
    Else
        ExtractEmailAddress = Trim(inputString)
    End If
    Exit Function
    
SimpleExtract:
    ' Fallback: just return trimmed input
    ExtractEmailAddress = Trim(inputString)
End Function

Function GetSMTPAddress(oMail As Outlook.mailItem) As String
    Dim smtpAddress As String
    
    On Error Resume Next
    
    ' Try to get real SMTP address for internal Exchange senders
    If Left(oMail.SenderEmailAddress, 3) = "/O=" Then
        Dim exchUser As Object
        Set exchUser = oMail.Sender.GetExchangeUser()
        If Not exchUser Is Nothing Then
            smtpAddress = exchUser.PrimarySmtpAddress
        End If
    End If
    
    ' Fall back to SenderEmailAddress if not Exchange DN
    If smtpAddress = "" Then
        smtpAddress = oMail.SenderEmailAddress
    End If
    
    On Error GoTo 0
    GetSMTPAddress = LCase(Trim(smtpAddress))
End Function



Private Function ExtractEmailFromHTMLBody(htmlBody As String) As String
    ' Locates the 6th table cell content without DOM parsing.
    ' Targets the student @district.edu email Banner places there.
    Dim tdCount As Long
    Dim pos As Long
    Dim tdStart As Long
    Dim tdEnd As Long
    Dim cellText As String
    Dim i As Long

    tdCount = 0
    pos = 1

    Do
        pos = InStr(pos, htmlBody, "<td", vbTextCompare)
        If pos = 0 Then Exit Do

        tdCount = tdCount + 1
        tdStart = InStr(pos, htmlBody, ">")
        If tdStart = 0 Then Exit Do
        tdStart = tdStart + 1

        tdEnd = InStr(tdStart, htmlBody, "</td>", vbTextCompare)
        If tdEnd = 0 Then Exit Do

        If tdCount = 6 Then
            cellText = Mid(htmlBody, tdStart, tdEnd - tdStart)
            ' Strip any inner tags
            cellText = StripHTML(cellText)
            cellText = Trim(cellText)
            ' Validate it looks like an email
            If InStr(cellText, "@") > 0 And InStr(cellText, ".") > 0 Then
                ExtractEmailFromHTMLBody = cellText
            End If
            Exit Do
        End If

        pos = tdEnd + 5
    Loop

    If ExtractEmailFromHTMLBody = "" Then
        ' Fallback: find first @district.edu occurrence in plain text
        Dim plainText As String
        plainText = StripHTML(htmlBody)
        Dim atPos As Long
        atPos = InStr(1, LCase(plainText), "@district.edu", vbTextCompare)
        If atPos > 0 Then
            Dim startPos As Long
            startPos = atPos
            Do While startPos > 1
                Dim c As String
                c = Mid(plainText, startPos - 1, 1)
                If c = " " Or c = Chr(9) Or c = Chr(13) Or c = Chr(10) Then Exit Do
                startPos = startPos - 1
            Loop
            Dim endPos As Long
            endPos = atPos + 9
            Do While endPos < Len(plainText)
                c = Mid(plainText, endPos + 1, 1)
                If c = " " Or c = Chr(9) Or c = Chr(13) Or c = Chr(10) Then Exit Do
                endPos = endPos + 1
            Loop
            ExtractEmailFromHTMLBody = Trim(Mid(plainText, startPos, endPos - startPos + 1))
        End If
    End If
End Function


Sub DiagnoseUnmatchedReplies()
    Dim olNS As Outlook.NameSpace
    Dim sharedMailbox As Outlook.folder
    Dim olFolder As Outlook.folder
    Dim olMail As Object
    Dim upperBody As String
    
    Set olNS = Application.GetNamespace("MAPI")
    Set sharedMailbox = olNS.Folders("DST-INTERNPAYROLL@district.edu")
    Set olFolder = sharedMailbox.Folders("Inbox").Folders("Supervisor_Replies")
    Debug.Print olFolder.Items.count
    
    For Each olMail In olFolder.Items
        If olMail.Class = 43 Then
            upperBody = UCase(olMail.Body)
            If InStr(upperBody, "APPROVE") = 0 And InStr(upperBody, "REJECT") = 0 And _
               InStr(upperBody, "LOOKS GOOD") = 0 And InStr(upperBody, "CONFIRMED") = 0 Then
                Debug.Print "--- UNMATCHED ---"
                Debug.Print olMail.SenderName
                Debug.Print Left(olMail.Body, 200)
                Debug.Print "---"
            End If
        End If
    Next
End Sub


Sub DiagnoseSubmissionTracking()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim replySheet As Object
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    
    Dim lastRowTracking As Long
    Dim lastRowReply As Long
    lastRowTracking = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    lastRowReply = replySheet.cells(replySheet.rows.count, 1).End(-4162).Row
    
    Debug.Print "Tracking rows: " & lastRowTracking
    Debug.Print "Reply_Log rows: " & lastRowReply
    Debug.Print "Tracking col3 sample: " & trackingSheet.cells(2, 3).Value
    Debug.Print "Tracking col7 sample: " & trackingSheet.cells(2, 7).Value
    Debug.Print "Reply col7 sample: " & replySheet.cells(2, 7).Value
    Debug.Print "Tracking col9 sample: " & trackingSheet.cells(2, 9).Value
Debug.Print "Tracking col9 row3: " & trackingSheet.cells(3, 9).Value
Debug.Print "Tracking col9 row4: " & trackingSheet.cells(4, 9).Value
    excelWorkbook.Close False
    excelApp.Quit
End Sub


Sub ResetAllApprovals()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=False)
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    For i = 2 To lastRow
        If trackingSheet.cells(i, 7).Value = "YES" Then
            trackingSheet.cells(i, 9).Value = "PENDING"
            trackingSheet.cells(i, 10).Value = ""
        End If
    Next i
    
    excelWorkbook.Save
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "All approval statuses reset to PENDING. Now run UpdateSubmissionStatus.", vbInformation
End Sub

Sub DiagnoseCol9()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim yesCount As Long
    Dim pendingCount As Long
    Dim rejectedCount As Long
    Dim otherCount As Long
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    For i = 2 To lastRow
        Select Case UCase(Trim(trackingSheet.cells(i, 9).Value))
            Case "YES"
                yesCount = yesCount + 1
            Case "PENDING"
                pendingCount = pendingCount + 1
            Case "REJECTED"
                rejectedCount = rejectedCount + 1
            Case ""
                otherCount = otherCount + 1
            Case Else
                Debug.Print "Row " & i & " unexpected value: " & trackingSheet.cells(i, 9).Value
        End Select
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "YES: " & yesCount & vbCrLf & _
           "PENDING: " & pendingCount & vbCrLf & _
           "REJECTED: " & rejectedCount & vbCrLf & _
           "Blank: " & otherCount, vbInformation
End Sub

Sub DiagnoseCol7()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim yesCount As Long
    Dim noCount As Long
    Dim blankCount As Long
    Dim otherCount As Long
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    For i = 2 To lastRow
        Select Case UCase(Trim(trackingSheet.cells(i, 7).Value))
            Case "YES"
                yesCount = yesCount + 1
            Case "NO"
                noCount = noCount + 1
            Case ""
                blankCount = blankCount + 1
            Case Else
                otherCount = otherCount + 1
                Debug.Print "Row " & i & " col7 unexpected: [" & trackingSheet.cells(i, 7).Value & "]"
        End Select
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "YES: " & yesCount & vbCrLf & _
           "NO: " & noCount & vbCrLf & _
           "Blank: " & blankCount & vbCrLf & _
           "Other: " & otherCount, vbInformation
End Sub

Sub DiagnoseReplyLogKeys()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim replySheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim goodCount As Long
    Dim badCount As Long
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    
    lastRow = replySheet.cells(replySheet.rows.count, 1).End(-4162).Row
    
    For i = 2 To lastRow
        Dim keyVal As String
        keyVal = replySheet.cells(i, 7).Value
        If InStr(keyVal, "@district.edu") > 0 And InStr(keyVal, "|") > 0 Then
            goodCount = goodCount + 1
        Else
            badCount = badCount + 1
            Debug.Print "Row " & i & " bad key: [" & Left(keyVal, 100) & "]"
        End If
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "Good keys: " & goodCount & vbCrLf & _
           "Bad keys: " & badCount, vbInformation
End Sub

Sub FindUnmatchedReplies()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim replySheet As Object
    Dim lastRowTracking As Long
    Dim lastRowReply As Long
    Dim i As Long
    Dim j As Long
    Dim matched As Boolean
    Dim acesID As String
    Dim keyVal As String
    Dim unmatchedCount As Long
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    
    lastRowTracking = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    lastRowReply = replySheet.cells(replySheet.rows.count, 1).End(-4162).Row
    
    For i = 2 To lastRowReply
        keyVal = LCase(Trim(replySheet.cells(i, 7).Value))
        matched = False
        
        For j = 2 To lastRowTracking
            Dim studentEmail As String
            studentEmail = LCase(Trim(trackingSheet.cells(j, 3).Value))
            If InStr(studentEmail, "@") > 0 Then
                acesID = Left(studentEmail, InStr(studentEmail, "@") - 1)
            Else
                acesID = studentEmail
            End If
            
            If InStr(keyVal, acesID) > 0 Then
                matched = True
                Exit For
            End If
        Next j
        
        If Not matched Then
            unmatchedCount = unmatchedCount + 1
            Debug.Print "Unmatched Reply Row " & i & ": " & replySheet.cells(i, 1).Value & " | " & Left(keyVal, 80)
        End If
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "Unmatched Reply_Log entries: " & unmatchedCount, vbInformation
End Sub

Sub CountUniqueReplyStudents()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim replySheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim uniqueStudents As Object
    Dim keyVal As String
    Dim acesID As String
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set replySheet = excelWorkbook.Sheets("Reply_Log")
    Set uniqueStudents = CreateObject("Scripting.Dictionary")
    
    lastRow = replySheet.cells(replySheet.rows.count, 1).End(-4162).Row
    
    For i = 2 To lastRow
        keyVal = LCase(Trim(replySheet.cells(i, 7).Value))
        If InStr(keyVal, "@") > 0 Then
            acesID = Left(keyVal, InStr(keyVal, "@") - 1)
        Else
            acesID = keyVal
        End If
        If acesID <> "" And Not uniqueStudents.Exists(acesID) Then
            uniqueStudents.Add acesID, True
        End If
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
    
    MsgBox "Unique students in Reply_Log: " & uniqueStudents.count, vbInformation
End Sub

Sub DiagnoseReplySubjects()
    Dim myNameSpace As Outlook.NameSpace
    Dim sharedMailbox As Outlook.folder
    Dim replyFolder As Outlook.folder
    Dim myItem As Object
    Dim oMail As Outlook.mailItem
    Dim i As Long
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set sharedMailbox = myNameSpace.Folders("DST-INTERNPAYROLL@district.edu")
    Set replyFolder = sharedMailbox.Folders("Inbox").Folders("Supervisor_Replies")
    
    i = 0
    For Each myItem In replyFolder.Items
        If TypeOf myItem Is Outlook.mailItem Then
            Set oMail = myItem
            Dim parsedName As String
            parsedName = ExtractStudentNameFromReplySubject(oMail.subject)
            If parsedName = "" Then
                i = i + 1
                Debug.Print "No name parsed: " & oMail.subject
            End If
        End If
    Next
    
    MsgBox "Subjects with no parsed name: " & i
End Sub

Sub ListPendingStudents()
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim trackingSheet As Object
    Dim lastRow As Long
    Dim i As Long
    Dim excelFilePath As String
    
    excelFilePath = GetWorkbookPath()
    If excelFilePath = "" Then Exit Sub
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    Set excelWorkbook = excelApp.Workbooks.Open(excelFilePath, ReadOnly:=True)
    Set trackingSheet = excelWorkbook.Sheets("Submission_Tracking")
    
    lastRow = trackingSheet.cells(trackingSheet.rows.count, 2).End(-4162).Row
    
    For i = 2 To lastRow
        If trackingSheet.cells(i, 7).Value = "YES" And _
           UCase(Trim(trackingSheet.cells(i, 9).Value)) = "PENDING" Then
            Debug.Print trackingSheet.cells(i, 2).Value & " | " & _
                        trackingSheet.cells(i, 5).Value & " | " & _
                        trackingSheet.cells(i, 6).Value
        End If
    Next i
    
    excelWorkbook.Close False
    excelApp.Quit
End Sub

Sub DiagnoseSenderEmails()
    Dim myNameSpace As Outlook.NameSpace
    Dim sharedMailbox As Outlook.folder
    Dim replyFolder As Outlook.folder
    Dim myItem As Object
    Dim oMail As Outlook.mailItem
    
    Set myNameSpace = Application.GetNamespace("MAPI")
    Set sharedMailbox = myNameSpace.Folders("DST-INTERNPAYROLL@district.edu")
    Set replyFolder = sharedMailbox.Folders("Inbox").Folders("Supervisor_Replies")
    
    For Each myItem In replyFolder.Items
        If TypeOf myItem Is Outlook.mailItem Then
            Set oMail = myItem
            If InStr(1, oMail.subject, "Mauricio", vbTextCompare) > 0 Or _
               InStr(1, oMail.subject, "Gladis", vbTextCompare) > 0 Or _
               InStr(1, oMail.subject, "Mayli", vbTextCompare) > 0 Then
                Debug.Print "Subject: " & oMail.subject
                Debug.Print "SenderEmailAddress: " & oMail.SenderEmailAddress
                Debug.Print "SenderName: " & oMail.SenderName
                Debug.Print "---"
            End If
        End If
    Next
End Sub


