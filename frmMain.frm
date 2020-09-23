VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "ListBox Demo - Ryan Lederman"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSelectAll2 
      Caption         =   "Select All Items"
      Height          =   375
      Left            =   3480
      TabIndex        =   23
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton btnClearSelections2 
      Caption         =   "Clear Selections"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "< Copy Selected"
      Height          =   375
      Left            =   3480
      TabIndex        =   18
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Multi Select"
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   6255
      Begin VB.CommandButton btnSelectAll1 
         Caption         =   "Select All Items"
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton btnClearSelections1 
         Caption         =   "Clear Selections"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Selected"
         Height          =   375
         Left            =   4800
         TabIndex        =   17
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton btnMoveItems 
         Caption         =   "Copy Selected >"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton btnDeleteItems 
         Caption         =   "Delete Selected"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Height          =   3100
         Left            =   3120
         TabIndex        =   14
         Top             =   140
         Width           =   30
      End
      Begin VB.ListBox lstMulti2 
         Height          =   2010
         ItemData        =   "frmMain.frx":0442
         Left            =   3360
         List            =   "frmMain.frx":0458
         MultiSelect     =   1  'Simple
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
      Begin VB.ListBox lstMulti1 
         Height          =   2010
         ItemData        =   "frmMain.frx":049B
         Left            =   120
         List            =   "frmMain.frx":04CC
         MultiSelect     =   1  'Simple
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton btnEnd 
      Cancel          =   -1  'True
      Caption         =   "Exit App"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton btnDeleteItem 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton btnFindItem 
      Caption         =   "Find"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Misc Operations"
      Height          =   2895
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton btnAddItem 
         Caption         =   "Add..."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   855
      End
      Begin VB.ListBox lstOperations 
         Height          =   2010
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2520
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Loading/Saving (\Names.txt)"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton btnSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "&Load"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   855
      End
      Begin VB.ListBox lstFromFile 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   " rlederman@mad.scientist.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddItem_Click()

TheItem = InputBox("Enter the string you wish to add to the list:", "Add string to list")

If TheItem = "" Then Exit Sub

lstOperations.AddItem TheItem
End Sub

Private Sub btnClear_Click()
If lstFromFile.ListCount = 0 Then MsgBox "The list is already empty.", 16, "List already empty": Exit Sub
retval = MsgBox("Are you sure you want to clear your list?", 36, "Confirm Clearing of ListBox")

Select Case retval

    Case vbYes
        lstFromFile.Clear
End Select
End Sub

Private Sub btnClearSelections1_Click()
If lstMulti1.SelCount = 0 Then MsgBox "No items selected. Cannot continue.", 16, "Error: No Selected Items": Exit Sub
For x = 0 To lstMulti1.ListCount - 1
    lstMulti1.Selected(x) = False
    DoEvents
Next x
End Sub

Private Sub btnClearSelections2_Click()
If lstMulti2.SelCount = 0 Then MsgBox "No items selected. Cannot continue.", 16, "Error: No Selected Items": Exit Sub
For x = 0 To lstMulti2.ListCount - 1
    lstMulti2.Selected(x) = False
    DoEvents
Next x
End Sub

Private Sub btnDeleteItem_Click()
If lstOperations.ListCount = 0 Then MsgBox "No items in list, cannot continue.", 16, "No Items Selected": Exit Sub
TheItem = lstOperations.ListIndex
If TheItem < 0 Then Exit Sub
lstOperations.RemoveItem TheItem

End Sub

Private Sub btnDeleteItems_Click()
If lstMulti1.SelCount = 0 Then MsgBox "No items selected.", 16, "Error: No items selected.": Exit Sub
On Error Resume Next
'We go through the list and remove all the items from the
'first ListBox...
Do While x < lstMulti1.ListCount
    If lstMulti1.Selected(x) = True Then 'Found a selected item
        lstMulti1.RemoveItem (x)  'Remove from first list
    Else 'This is very important. without this, it would generate
         'an error:
         x = x + 1
    End If
DoEvents
Loop
End Sub

Private Sub btnEnd_Click()
End
End Sub

Private Sub btnFindItem_Click()
If lstOperations.ListCount = 0 Then MsgBox "No items in list, cannot continue.", 16, "Error: No Items In List": Exit Sub
Dim WhatToFind As String

WhatToFind = InputBox("Please enter the string to find in the list below:", "String to find")

Call FindItem(WhatToFind)
End Sub

Private Sub btnLoad_Click()
If lstFromFile.ListCount > 0 Then lstFromFile.Clear
Call LoadListFromFile(App.Path + "\Names.txt")

End Sub

Private Sub btnMoveItems_Click()
If lstMulti1.SelCount = 0 Then MsgBox "No items selected. Cannot continue.", 16, "Error: No Selected Items": Exit Sub
'We go through the list and add all the selected items
'to the second ListBox...
For x = 0 To lstMulti1.ListCount - 1
    If lstMulti1.Selected(x) = True Then 'Found a selected item
        lstMulti2.AddItem lstMulti1.List(x) 'Add to second list
    End If
DoEvents
Next x
'Now, call the clear selections button to take away the
'selection focus on the list items
btnClearSelections1_Click
End Sub

Private Sub btnSave_Click()
If lstFromFile.ListCount = 0 Then MsgBox "Your list is empty. You cannot save an empty list.", 16, "Error: List Empty": Exit Sub

With CommonDialog
    .Flags = 4 'This disables the 'Read Only' checkbox
    .DialogTitle = "Save List As"
    .Filter = "Text Files|*.Txt|All Files|*.*"
    .FileName = "Names.txt"
    .ShowSave
    
    Call SaveList(.FileName)
End With
End Sub

Private Sub btnSelectAll1_Click()
If lstMulti1.ListCount = 0 Then MsgBox "No Items To Select.", 16, "Error": Exit Sub
For x = 0 To lstMulti1.ListCount - 1
    lstMulti1.Selected(x) = True
    DoEvents
Next x
End Sub

Private Sub btnSelectAll2_Click()
If lstMulti2.ListCount = 0 Then MsgBox "No Items To Select.", 16, "Error": Exit Sub
For x = 0 To lstMulti2.ListCount - 1
    lstMulti2.Selected(x) = True
    DoEvents
Next x
End Sub

Private Sub Command1_Click()
If lstMulti2.SelCount = 0 Then MsgBox "No items selected. Cannot continue.", 16, "Error: No Selected Items": Exit Sub
On Error Resume Next
'We go through the list and remove all the items from the
'first ListBox...
Do While x < lstMulti2.ListCount
    If lstMulti2.Selected(x) = True Then 'Found a selected item
        lstMulti2.RemoveItem (x)  'Remove from first list
    Else 'This is very important. without this, it would generate
         'an error:
         x = x + 1
    End If
DoEvents
Loop
End Sub

Private Sub Command2_Click()
If lstMulti2.SelCount = 0 Then MsgBox "No items selected. Cannot continue.", 16, "Error: No Selected Items": Exit Sub
'We go through the list and add all the selected items
'to the second ListBox...
For x = 0 To lstMulti2.ListCount - 1
    If lstMulti2.Selected(x) = True Then 'Found a selected item
        lstMulti1.AddItem lstMulti1.List(x) 'Add to second list
    End If
DoEvents
Next x
'Now, call the clear selections button to take away the
'selection focus on the list items
btnClearSelections2_Click
End Sub
