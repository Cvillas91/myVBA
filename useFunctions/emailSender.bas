Attribute VB_Name = "emailSender"

Sub emailSender()

    Dim OutlookApp As Object
    Dim OutlookMessage As Object
    
    Dim strTo As String, strCC As String, strBCC As String, strSubject As String
    Dim varAttachments As Variant
    Dim i As Long
    
    strTo = Sheet("Control").Range("TO")
    strCC = Sheet("Control").Range("CC")
    strBCC = Sheet("Control").Range("BCC")
    strSubject = Sheet("Control").Range("Subject")
    varAttachments = Sheet("Control").Range("Attachments")

    i = Ubound(varAttachments, 1)

    'Optmize Code
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    'Create an instance of outlook
    On Error Resume Next
        Set OutlookApp = GetObject(class:= "Outlook.Application")
        Err.Clear 
        If OutlookApp Is Nothing Then Set OutlookApp = CreateObject(class:="Outlook.Application")   

    'Create new email message
    Set OutlookMessage = OutlookApp.CreateItem(0)

    'Create Outlook email with attachment
    On Error Resume Next
        With OutlookMessage
            .To = strTo
            .CC = strCC
            .BCC = strBCC
            .Subject = strSubject
            .HTMLBody = RangetoHTML(Range("Body"))

            For j = 1 to i

                If varAttachments(j,1) <> "" Then .Attachments.Add varAttachments(j,1)

            Next j

            .Display
        End With
    
    'Clear Memory
    Set OutlookMessage = Nothing
    Set OutlookApp = Nothing

End Sub
Function RangetoHTML(rng as Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB as Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to pass the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False 
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
        SourceType:= xlSourceRange, _
        Filename:= TempFile, _
        Sheet:=TempWB.Sheets(1).Name, _
        SourceL=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1,-2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
    "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:= False

    'Delete the htm file used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing

End function
