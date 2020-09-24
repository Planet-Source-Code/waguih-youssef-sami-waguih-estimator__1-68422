VERSION 5.00
Begin VB.Form frmContCat 
   Caption         =   "Categorize SubContracts"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   6540
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCategory 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   2595
      ItemData        =   "frmContCat.frx":0000
      Left            =   240
      List            =   "frmContCat.frx":0002
      TabIndex        =   6
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAlphabetize 
      Caption         =   "Alphabetize"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Category"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmContCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KK

Private Sub cmdAdd_Click()
List1.AddItem txtCategory.Text
End Sub

Private Sub cmdDelete_Click()
Dim MyPath
MyPath = App.Path
Dim Contracts
  Contracts = MyPath & "\" & "Contracts.ini"
Dim index
 index = List1.ListIndex
KK = CStr(List1.List(index))
Dim MyString
Dim NewMat As String
 Open Contracts For Input As #1

 Do While Not EOF(1)
  Line Input #1, MyString
If MyString = KK Then
  MyString = ""
 NewMat = NewMat & MyString
    
    Line Input #1, MyString
    Do While MyString <> ""
    MyString = ""
    NewMat = NewMat & MyString
    Line Input #1, MyString
    Loop
  End If
NewMat = NewMat & MyString & vbCrLf
Loop
Close #1

Dim NewFile
NewFile = MyPath & "\" & "Contracts.ini"
Open NewFile For Output As #1
Print #1, NewMat
Close #1

End Sub

Private Sub cmdOK_Click()
Dim MyPath
MyPath = App.Path
Dim Contracts
  Contracts = MyPath & "\" & "Contracts.ini"
Dim index
 index = List1.ListIndex
KK = CStr(List1.List(index))
If KK = "" Then
MsgBox "Please click on the List Below", vbOKOnly
Exit Sub
End If
Open Contracts For Append As #1
Print #1, KK
Close #1


'Me.Hide
'frmSubContract.Show
End Sub

Private Sub Form_Load()
Dim MyPath
MyPath = App.Path
Dim Contracts
  Contracts = MyPath & "\" & "Contracts.ini"
 Dim index
 index = 0
 Open Contracts For Input As #1     ' Open file for read.
  Do While Not EOF(1)
  Line Input #1, MyString
  If IsNumeric(Left(MyString, 1)) = False Then
 If MyString <> "" Then
  List1.AddItem MyString, index
index = index + 1
  End If
  End If
 
  Loop
  Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSubContract.Show
Me.Hide
End Sub

Private Sub List1_Click()
Dim index
 index = List1.ListIndex

KK = CStr(List1.List(index))
End Sub

