VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Email Form"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   -240
   ClientWidth     =   9240
   OleObjectBlob   =   "Greenwich Projects UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub autofill_Click()
    Dim ws As Sheet1
    Set ws = ThisWorkbook.Sheets("Symology Projects")
    
    ttitle.Value = ws.Cells(trow.Value, "A")
    tlacode.Value = ws.Cells(trow.Value, "B")
    tpclcode.Value = ws.Cells(trow.Value, "C")
    tofficer.Value = ws.Cells(trow.Value, "D")
    texpcode.Value = ws.Cells(trow.Value, "E")
    tnewspaper.Value = "Greenwich Weekender"
    tnumber.Value = ws.Cells(trow.Value, "V")
    tinvoice.Value = "Richy Udemezue"
    
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
Dim ws As Sheet1
Dim Yes As String
Yes = "Y"

Set ws = ThisWorkbook.Sheets("Symology Projects")

    If UserForm1.cengineer.Value = False Then
        ws.Cells(trow.Value, "J") = ""
    Else
        ws.Cells(trow.Value, "J") = Yes
    End If
End Sub

Private Sub cexplorer_Click()
Dim fd As Office.FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
End Sub

Private Sub cnewspaper_Click()
Dim ws As Sheet1
Dim Yes As String
Yes = "Y"

    If UserForm1.CheckBox1.Value = False Then
        ws.Cells(trow.Value, "L") = ""
        ws.Cells(trow.Value, "M") = ""
        
    Else
        ws.Cells(trow.Value, "L") = Yes
        ws.Cells(trow.Value, "M") = Yes
    End If
End Sub

Private Sub creminder_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Making the Order - Reminder.oft")
    
    StrPath = "G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Attachments\"
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
        
        'Add booking form
        'OutMail.Attachments.Add ("G:\Project Centre\")
        OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Making the Order - Reminder - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0
End Sub

Private Sub cstakeholders_Click()
Dim ws As Sheet1
Dim Yes As String
Yes = "Y"

    If UserForm1.cstakeholders.Value = False Then
        ws.Cells(trow.Value, "N") = ""
        
    Else
        ws.Cells(trow.Value, "N") = Yes
    
    End If
End Sub

Private Sub emailengineer_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Engineer Sign Off.oft")
    
    StrPath = "G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Attachments\"
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
        
        'Add booking form
        'OutMail.Attachments.Add ("G:\Project Centre\")
        OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\NoP for Signature - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0
End Sub

Private Sub emailstakeholder_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Peter Kavanagh Consultation Email.oft")
    
    StrPath = "G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Attachments\"
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
        OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Consultation - London Cab Ranks - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\All Stakeholders Consultation Email.oft")
    
    StrPath = "G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Attachments\"
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
        
        'Add booking form
        'OutMail.Attachments.Add ("G:\Project Centre\")
        OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Consultation - " & UserForm1.ttitle.Value & ".msg"
        OutMail.Display
    'End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Private Sub generate_Click()

    Dim objWord As Object
    Dim objDoc As Object
    Dim MyDate As Object
    Dim Form As Object
    Dim objRange As Object
    
    Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Greenwich Weekender Traffic Booking Form - Template - Copy.docx")
    objWord.Application.DisplayAlerts = False
    objWord.Visible = True
    objDoc.SaveAs ("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Forms\Greenwich Weekender Booking Form - " & UserForm1.ttitle.Value & ".docx")
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
    
    objDoc.SaveAs ("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Forms\Greenwich Weekender Booking Form - " & UserForm1.ttitle.Value & ".docx")
    
    'objDoc.ExportAsFixedFormat OutputFileName:= _
    "G:\Project Centre\Parking Team\Magda\Greenwich\Greenwich Weekender Traffic Booking Form - "UserForm1.ttitle.Value" .pdf", _
    ExportFormat:=wdExportFormatPDF, _
    Range:=wdExportFromTo, From:=1, To:=1
    
    'Set objWord = CreateObject("Word.Application")
    Set objDoc = objWord.Documents.Open("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\PENNA Advertisement Order Form - Copy.docx")
    objWord.Application.DisplayAlerts = False
    objWord.Visible = True
    objDoc.SaveAs ("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Forms\PENNA Advertisement Order Form - " & UserForm1.ttitle.Value & ".docx")
    'objDoc.Selection.Find.Text = "LACode"
    'objDoc.Selection.Find.ReplacementText = UserForm1.tlacode.Value
    
    'objDoc.ExportAsFixedFormat OutputFileName:= _
    "G:\Project Centre\Parking Team\Magda\Greenwich\Greenwich Weekender Traffic Booking Form - "UserForm1.ttitle.Value" .pdf", _
    ExportFormat:=wdExportFormatPDF, _
    Range:=wdExportFromTo, From:=1, To:=1

End Sub

Private Sub email_Click()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim StrFile As String
    Dim StrPath As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Greenwich Notice 2.oft")
    
    StrPath = "G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Attachments\"
    StrFile = Dir(StrPath & "*.*")
    
    On Error Resume Next
    'With OutMail
        OutMail.To = "Gaynor.Granger@royalgreenwich.gov.uk; " + "Nicola.McGuire@royalgreenwich.gov.uk; " + "adapt.studio@royalgreenwich.gov.uk"
        'OutMail.To = "magdalena.misiewicz@projectcentre.co.uk"
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
        OutMail.Attachments.Add ("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Forms\Greenwich Weekender Booking Form - " & UserForm1.ttitle.Value & ".docx")
        
        OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\Greenwich Weekender - " & UserForm1.ttitle.Value & ".msg"
        
        OutMail.Display
    'End With
    On Error GoTo 0
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItemFromTemplate("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\PENNA Notice 2.oft")
    
    StrPath = "G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Attachments\"
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
        OutMail.Attachments.Add ("G:\Project Centre\Parking Team\01 Team\Magda\Greenwich\Forms\PENNA Advertisement Order Form - " & UserForm1.ttitle.Value & ".docx")
        
        OutMail.SaveAs "G:\Project Centre\Project-BST\" & UserForm1.tpclcode & " - RBG " & UserForm1.ttitle.Value & "\2 Project Delivery\2 Communications\1  Project Emails\PENNA - " & UserForm1.ttitle.Value & ".msg"
        
        OutMail.Display
    'End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Private Sub UserForm_Initialize()
    tnewspaper.AddItem "GREENWICH WEEKENDER"
    tinvoice.AddItem "Richard Cornell"
    tinvoice.AddItem "Richy Udemezue"
    'Ward.AddItem "Abbey Wood"
    'Ward.AddItem "Blackheath Westcombe"
    'Ward.AddItem "Charlton"
    'Ward.AddItem "Coldharbour and New Eltham"
    'Ward.AddItem "Eltham North"
    'Ward.AddItem "Glyndon"
    'Ward.AddItem "Greenwich West"
    'Ward.AddItem "Kidbrooke with Hornfair"
    'Ward.AddItem "Middle Park and Sutcliffe"
    'Ward.AddItem "Plumstead"
    'Ward.AddItem "Shooters Hill"
    'Ward.AddItem "Thamesmead Moorings"
    'Ward.AddItem "Woolwich Common"
    'Ward.AddItem "Woolwich Riverside"
End Sub
