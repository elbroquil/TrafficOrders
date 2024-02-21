VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Email Form"
   ClientHeight    =   9330.001
   ClientLeft      =   140
   ClientTop       =   60
   ClientWidth     =   13860
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub OpenUserForm()

    UserForm1.Show

End Sub
Private Sub UserForm_Initialize()

    tnewspaper.AddItem "GREENWICH WEEKENDER"
    tinvoice.AddItem "Harry Potter"
    tinvoice.AddItem "William Blake"
    
End Sub
Private Sub cexplorer_Click()

'Previous version of the attachment picker. Replaced with the macro pulling automatically attachments into the emails.

    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

End Sub
Private Sub autofill_Click()
    
    'Autofill button - pulls information from the table into the userform
    
    Dim ws As Sheet1
    Set ws = ThisWorkbook.Sheets("Projects")
    
    ttitle.Value = ws.Cells(trow.Value, "A")
    tlacode.Value = ws.Cells(trow.Value, "B")
    tpclcode.Value = ws.Cells(trow.Value, "C")
    tofficer.Value = ws.Cells(trow.Value, "D")
    texpcode.Value = ws.Cells(trow.Value, "E")
    tnewspaper.Value = "Greenwich Weekender"
    tnumber.Value = ws.Cells(trow.Value, "V")
    tinvoice.Value = "Richy Udemezue"
    
    'Check if the Order is in the NoP (Drafting) or NoM (being Made) stage
    If IsEmpty(ws.Cells(trow.Value, "P")) = False Then
        tdate.Value = ws.Cells(trow.Value, "P")
        Else
            tdate.Value = ws.Cells(trow.Value, "I")
        End If
        
        If IsEmpty(ws.Cells(trow.Value, "P")) = False Then
            tnop.Value = "NoM"
        Else
            tnop.Value = "NoP"
        End If
    
    Dim OffEmail As String
    
    OffEmail = UserForm1.tofficer.Value
    OffEmail = Replace(OffEmail, " ", ".")
    OffEmail = OffEmail & "@royalgreenwich.gov.uk"
    toffemail.Value = OffEmail
    
    Dim NextThursday As Date
    NextThursday = Date + Choose(Weekday(Date, 1), 4, 3, 2, 1, 0, 5)
    tsign.Value = Format(NextThursday, "dd/MM/yyyy")
    
    tstart.Value = UserForm1.tdate.Value
    
    Dim EndDate As Date
    EndDate = Date + 48 'Assuming emails sent on Wednesday
    tend.Value = Format(EndDate, "dd/MM/yyyy")
    
End Sub
Private Sub cengineer_Click()

'Stage 1: Confirm that engineers have been contacted for Order approval.

    Dim ws As Sheet1
    Dim Yes As String
    Yes = "Y"
    
    Set ws = ThisWorkbook.Sheets("Projects")

        If UserForm1.cengineer.Value = False Then
            ws.Cells(trow.Value, "J") = ""
        Else
            ws.Cells(trow.Value, "J") = Yes
        End If
        
End Sub
Private Sub cnewspaper_Click()

'Stage 2: Confirm that the notice has been sent to the newspapers

Dim ws As Sheet1
Dim Yes As String
Yes = "Y"

Set ws = ThisWorkbook.Sheets("Projects")

    If UserForm1.cnewspaper.Value = False Then
        ws.Cells(trow.Value, "L") = ""
        ws.Cells(trow.Value, "M") = ""
        
    Else
        ws.Cells(trow.Value, "L") = Yes
        ws.Cells(trow.Value, "M") = Yes
    End If
    
End Sub
Private Sub cstakeholders_Click()

'Stage 3: Confirm that the notice has been sent to the stakeholders

Dim ws As Sheet1
Dim Yes As String
Yes = "Y"

Set ws = ThisWorkbook.Sheets("Projects")

    If UserForm1.cstakeholders.Value = False Then
        ws.Cells(trow.Value, "N") = ""
        
    Else
        ws.Cells(trow.Value, "N") = Yes
    
    End If
    
End Sub
Private Sub emailengineer_Click()

'Stage 1: Send emails to the engineers for Order confirmation.

    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    Dim FolderLocation As String
    
    FolderLocation = Application.ActiveWorkbook.Path
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(FolderLocation & "\Engineer Sign Off.oft")
    
    StrPath = FolderLocation & "\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        OutMail.To = UserForm1.toffemail.Value
        OutMail.CC = "jerome.pilley@projectcentre.co.uk"
        'OutMail.BCC = "magdalena.misiewicz@projectcentre.co.uk"
        OutMail.Subject = Replace(OutMail.Subject, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "NextThursday", UserForm1.tsign.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "NextWednesday", UserForm1.tstart.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "Officer", UserForm1.tofficer.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "6Wednesdays", UserForm1.tend.Value)
        
        'Add other attachments
        Do While Len(StrFile) > 0
            OutMail.Attachments.Add StrPath & StrFile
            StrFile = Dir
        Loop
        
        'I've commented the line below out, but it was saving the email in the project folder
        'OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\NoP for Signature - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0
    
End Sub
Private Sub email_Click()

'Stage 2: Send emails and forms to the newspapers

    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    Dim FolderLocation As String
    
    FolderLocation = Application.ActiveWorkbook.Path
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(FolderLocation & "\Greenwich Notice 2.oft")
    
    StrPath = FolderLocation & "\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        OutMail.To = "Gaynor.Granger@royalgreenwich.gov.uk; " + "Nicola.McGuire@royalgreenwich.gov.uk; " + "adapt.studio@royalgreenwich.gov.uk"
        'OutMail.CC = "magdalena.misiewicz@projectcentre.co.uk"
        OutMail.CC = "jerome.pilley@projectcentre.co.uk"
        'OutMail.BCC = ""
        OutMail.Subject = Replace(OutMail.Subject, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Newspaper%", UserForm1.tnewspaper.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Date%", UserForm1.tdate.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%ExpCode%", UserForm1.texpcode.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%LACode%", UserForm1.tlacode.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "PCLCode", UserForm1.tpclcode.Value & UserForm1.tnop.Value)
        
        'Add other attachments
        Do While Len(StrFile) > 0
            OutMail.Attachments.Add StrPath & StrFile
            StrFile = Dir
        Loop
        
        'Add booking form
        OutMail.Attachments.Add (FolderLocation & "\Forms\Greenwich Weekender Booking Form - " & UserForm1.ttitle.Value & ".docx")
        
        'I've commented the line below out, but it was saving the email in the project folder
        'OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Greenwich Weekender - " & UserForm1.ttitle.Value & ".msg"
        
        OutMail.Display
    'End With
    On Error GoTo 0
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(FolderLocation & "\PENNA Notice 2.oft")
    
    StrPath = FolderLocation & "\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        OutMail.To = "greenwichpn@penna.com"
        'OutMail.To = "magdalena.misiewicz@projectcentre.co.uk"
        'OutMail.CC = "magdalena.misiewicz@projectcentre.co.uk"
        OutMail.CC = "jerome.pilley@projectcentre.co.uk"
        'OutMail.BCC = ""
        OutMail.Subject = Replace(OutMail.Subject, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Date%", UserForm1.tdate.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "PCLCode", UserForm1.tpclcode.Value & UserForm1.tnop.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "ExpCode", UserForm1.texpcode.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "LACode", UserForm1.tlacode.Value)
        
        'Add other attachments
        Do While Len(StrFile) > 0
            OutMail.Attachments.Add StrPath & StrFile
            StrFile = Dir
        Loop
        
        'Add booking form
        OutMail.Attachments.Add (FolderLocation & "\Forms\PENNA Advertisement Order Form - " & UserForm1.ttitle.Value & ".docx")
        
        'I've commented the line below out, but it was saving the email in the project folder
        'OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\PENNA - " & UserForm1.ttitle.Value & ".msg"
        
        OutMail.Display
    'End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub
Private Sub emailstakeholder_Click()

'Stage 3: Send notice to the stakeholders.

    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    Dim FolderLocation As String
    
    FolderLocation = Application.ActiveWorkbook.Path
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(FolderLocation & "\Peter Kavanagh Consultation Email.oft")
    
    StrPath = FolderLocation & "\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        'OutMail.To = "magdalena.misiewicz@projectcentre.co.uk"
        'OutMail.CC = "magdalena.misiewicz@projectcentre.co.uk"
        'OutMail.BCC = "magdalena.misiewicz@projectcentre.co.uk"
        OutMail.Subject = Replace(OutMail.Subject, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "TelephoneNumber", UserForm1.tnumber.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "Deadline", UserForm1.tdeadline.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "PROJECTOFFICER", UserForm1.tofficer.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "PCLCode", UserForm1.tpclcode.Value)
        
        'Add other attachments
        Do While Len(StrFile) > 0
            OutMail.Attachments.Add StrPath & StrFile
            StrFile = Dir
        Loop
        
        'Add booking form - not applicable for consultations
        'OutMail.Attachments.Add ("G:\Project Centre\")
        
        'Line below commented out
        'OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Consultation - London Cab Ranks - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(FolderLocation & "\All Stakeholders Consultation Email.oft")
    
    StrPath = FolderLocation & "\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        'OutMail.To = "magdalena.misiewicz@projectcentre.co.uk"
        'OutMail.CC = "magdalena.misiewicz@projectcentre.co.uk"
        'OutMail.BCC = "magdalena.misiewicz@projectcentre.co.uk"
        OutMail.Subject = Replace(OutMail.Subject, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "TelephoneNumber", UserForm1.tnumber.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "Deadline", UserForm1.tdeadline.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "PROJECTOFFICER", UserForm1.tofficer.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "PCLCode", UserForm1.tpclcode.Value)
        
        'Add other attachments
        Do While Len(StrFile) > 0
            OutMail.Attachments.Add StrPath & StrFile
            StrFile = Dir
        Loop
        
        'Line below commented out
        'OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Consultation - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub
Private Sub creminder_Click()

'Stage 4: Send reminder emails - consultation period has passed and (no) objections were received.

    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    Dim FolderLocation As String
    
    FolderLocation = Application.ActiveWorkbook.Path
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate(FolderLocation & "\Making the Order - Reminder.oft")
    
    StrPath = FolderLocation & "\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        OutMail.To = UserForm1.toffemail.Value
        OutMail.CC = "jerome.pilley@projectcentre.co.uk"
        'OutMail.BCC = "magdalena.misiewicz@projectcentre.co.uk"
        OutMail.Subject = Replace(OutMail.Subject, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%Title%", UserForm1.ttitle.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "%date%", UserForm1.tdate.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "LACode", UserForm1.tlacode.Value)
        OutMail.HTMLBody = Replace(OutMail.HTMLBody, "Officer", UserForm1.tofficer.Value)
        
        'Add other attachments
        Do While Len(StrFile) > 0
            OutMail.Attachments.Add StrPath & StrFile
            StrFile = Dir
        Loop
        
        'I've commented out the line below, but it was saving the sent emails in the project folder.
        'OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Making the Order - Reminder - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0
    
End Sub
Private Sub generate_Click()

'Generate newspaper request forms (Greenwich Weekender and PENNA).
'The forms will save a copy of the templates in the FORMS folder, fill them in and leave Word open for review.

    Dim objWord As Object
    Dim objDoc As Object
    Dim MyDate As Object
    Dim Form As Object
    Dim objRange As Object
    Dim FolderLocation As String
    
    FolderLocation = Application.ActiveWorkbook.Path
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open(FolderLocation & "\Greenwich Weekender Traffic Booking Form - Template - Copy.docx")
    objWord.Application.DisplayAlerts = False
    objWord.Visible = True
    objDoc.SaveAs (FolderLocation & "\Greenwich Weekender Booking Form - " & UserForm1.ttitle.Value & ".docx")
    Set objRange = objDoc.Bookmarks("Date").Range
    objRange.InsertAfter (UserForm1.tdate.Value)
    
    Set objRange = objDoc.Bookmarks("Today").Range
    objRange.InsertAfter (Date)
    
    Set objRange = objDoc.Bookmarks("Officer").Range
    objRange.InsertAfter (UserForm1.tofficer.Value)
    
    Set objRange = objDoc.Bookmarks("LACode").Range
    objRange.InsertAfter (UserForm1.tlacode.Value)
    
    Set objRange = objDoc.Bookmarks("ExpCode").Range
    objRange.InsertAfter (UserForm1.texpcode.Value)
    
    Set objRange = objDoc.Bookmarks("PCLCode").Range
    objRange.InsertAfter (UserForm1.tpclcode.Value & UserForm1.tnop.Value)
    
    Set objRange = objDoc.Bookmarks("Invoice").Range
    objRange.InsertAfter (UserForm1.tinvoice.Value)
    
    objDoc.SaveAs (FolderLocation & "\Forms\Greenwich Weekender Booking Form - " & UserForm1.ttitle.Value & ".docx")
    
    'Lines below to export as PDF and not Word doc
    'objDoc.ExportAsFixedFormat OutputFileName:= _
    "G:\Project Centre\Parking Team\Magda\Greenwich\Greenwich Weekender Traffic Booking Form - "UserForm1.ttitle.Value" .pdf", _
    ExportFormat:=wdExportFormatPDF, _
    Range:=wdExportFromTo, From:=1, To:=1
    
    Set objDoc = objWord.Documents.Open(FolderLocation & "\PENNA Advertisement Order Form - Copy.docx")
    objWord.Application.DisplayAlerts = False
    objWord.Visible = True
    objDoc.SaveAs (FolderLocation & "\Forms\PENNA Advertisement Order Form - " & UserForm1.ttitle.Value & ".docx")
    'objDoc.Selection.Find.Text = "LACode"
    'objDoc.Selection.Find.ReplacementText = UserForm1.tlacode.Value
    
    'Lines below to export as PDF and not Word doc
    'objDoc.ExportAsFixedFormat OutputFileName:= _
    "G:\Project Centre\Parking Team\Magda\Greenwich\Greenwich Weekender Traffic Booking Form - "UserForm1.ttitle.Value" .pdf", _
    ExportFormat:=wdExportFormatPDF, _
    Range:=wdExportFromTo, From:=1, To:=1
    
    MsgBox "Forms saved and prefilled. Please see Word to review and close/edit."

End Sub










