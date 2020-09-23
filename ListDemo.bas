Attribute VB_Name = "ListDemo"
Sub LoadListFromFile(xPath As String)
'This Sub will load any list of information
'from your hard drive into a ListBox control.
'xPath represents the file to be loaded.
Dim TheFile As Integer
Dim OurBuffer As String
If Dir(xPath) = "" Then MsgBox "File Not Found!", 16, "Error: File Not Found": Exit Sub
TheFile = FreeFile() 'This will assure that the
                     'File number is a good one
                     
Open xPath For Input As #TheFile
    Do While Not EOF(TheFile)
        Line Input #TheFile, OurBuffer
        'If the line being read is empty, do not add it
        If OurBuffer = "" Then GoTo SkipAdd
        frmMain.lstFromFile.AddItem OurBuffer
SkipAdd: 'If the item is empty, instructions jump here
        DoEvents 'Allow other proccesses to run
    Loop
Close #TheFile
End Sub
Sub FindItem(xItem As String)
'This will find a string inside the ListBox control.
'If none are found, it will display a message
For x = 0 To frmMain.lstOperations.ListCount - 1
    If UCase(xItem) = UCase(frmMain.lstOperations.List(x)) Then
    MsgBox "The string ''" & xItem & "'' was found!" & vbCrLf & "The item you searched for is index: " & x & " in the list.", 64, "Item Found!"
    Exit Sub
    End If
Next x
    MsgBox "The item was not found.", 16, "Error: Item not found in list"
End Sub
Sub SaveList(xPath As String)
'This Sub saves the contents of a ListBox control to a
'file on the hard drive
Dim TheFile, ListCount As Integer
TheFile = FreeFile() 'This assures that our file number
                     'is not already in use by another file
                     
Open xPath For Output As #TheFile
    For ListCount = 0 To frmMain.lstFromFile.ListCount - 1
        Print #TheFile, frmMain.lstFromFile.List(ListCount)
    DoEvents
    Next ListCount
Close #TheFile
        
        
End Sub
