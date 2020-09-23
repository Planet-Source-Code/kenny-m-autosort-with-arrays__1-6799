VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   2790
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComments 
      Height          =   1725
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   960
      Width           =   5175
   End
   Begin VB.ComboBox cboArray 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "cboArray"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   -1320
      TabIndex        =   4
      Text            =   "cboAddress"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboComments 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "cboComments"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox cboNumber 
      Height          =   315
      Left            =   -1320
      TabIndex        =   1
      Text            =   "cboNumber"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Comments:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit Program"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' AutoSort With Array(s) _
  NOTE: This code may contain bugs... It's your job to work _
  them out. :) _
  If you need to contact me you may by: _
  E-Mail: bandit@unetworks.net _
  ICQ: 42469515 _
  Have Fun with the code :) _

  
  
'------------READ BELOW--------------
'  If you use this code or it's format you must give me some _
   credit after all I showed you how to do this. Thanks :)
'--If you don't agree delete this example right and do not go _
   father than this point

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Option Explicit
  Dim Parsed As Variant
  Dim sLine As Variant
  
Private Sub cboName_click()
'This was the ItemClick. Make sure the combo box style property _
is set to 2 - Dropdownlist, this will keep the user from typing _
in th combo box.

'Set Text Boxes so they can't be edited
cboName.Locked = False
txtComments.Locked = True

'Copy the info to the text boxes
txtComments.Text = cboComments.List(cboName.ListIndex)
End Sub
Sub FileOpen(filename As String)
On Error Resume Next
Dim sLine As String, i As Integer
Dim FileNumber As Integer
Dim ArrayedEntrys As Variant
Dim ArraysSpot As Integer

'Unlock Boxes So the program can use them
txtComments.Locked = False

    
    If filename = "" Then Exit Sub 'If no file exit now...
    

    cboName.Clear 'Delete this line if you don't want to clear it...
    cboComments.Clear 'Delete this line if you don't want to clear it...
    cboArray.Clear
    ArraysSpot = 0

    
    FileNumber = FreeFile
    Open filename For Input As FileNumber 'Open the file to read

        Do Until EOF(FileNumber) 'Go until end of file

            Line Input #FileNumber, sLine 'Get Line and set to variable sLine
            sLine = sLine
            cboArray.AddItem sLine
skip1:
        Loop
    Close #FileNumber
    
            ArrayedEntrys = cboArray.ListCount
            Do Until ArrayedEntrys = 0
            Parsed = Split(cboArray.List(ArraysSpot), "|") 'Get line of file contents and split into an array
            cboName.AddItem Parsed(0) 'Add Name
            cboComments.AddItem Parsed(1) 'Add Phone Number
            ArraysSpot = ArraysSpot + 1
            ArrayedEntrys = ArrayedEntrys - 1
        Loop
    
    'If List has more then 1 item select the first one.
    If cboName.ListCount > 0 Then cboName.ListIndex = 0
           
'lock boxes again
txtComments.Locked = True
End Sub

Private Sub mnuAbout_Click()
'frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
'End th program, you might want to add a feature to ask to save before _
exiting the program.
Unload Me
End
End Sub
Private Sub Form_Load()
On Error Resume Next
        ' Center Form
        Move ((Screen.Width - Me.Width) / 2), _
        ((Screen.Height - Me.Height) / 2)
        
FileOpen (App.Path & "\data.txt")
End Sub
