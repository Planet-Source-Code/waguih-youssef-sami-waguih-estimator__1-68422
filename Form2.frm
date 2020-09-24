VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMaterial 
   Caption         =   "New Material"
   ClientHeight    =   3990
   ClientLeft      =   1890
   ClientTop       =   1485
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   ScaleHeight     =   3990
   ScaleWidth      =   6540
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCategory 
      Caption         =   "Categories"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alphabetize"
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cndDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "Form2.frx":0000
      Left            =   120
      List            =   "Form2.frx":0002
      TabIndex        =   6
      Top             =   1800
      Width           =   4935
   End
   Begin VB.ComboBox cmbMaterials 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox txtCost 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   660
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyMaterial

Private Sub cmbMaterials_Click()
MyMaterial = cmbMaterials.Text
End Sub

Private Sub cmdAdd_Click()
If txtDescription.Text <> "" And txtCost.Text <> "" Then
List1.AddItem txtCost.Text * 100 & " " & txtDescription.Text
txtDescription.Text = ""
txtCost.Text = ""
End If
End Sub

Private Sub cmdCancel_Click()
Me.Hide
Form1.Show
End Sub

Private Sub cmdCategory_Click()
frmMatCat.Show
Unload frmMaterial
End Sub

Private Sub cmdHelp_Click()
Dim MyPath
MyPath = App.Path
Const HelpFinder = &HB
CommonDialog1.Action = 6 '6 means run winhlp32.exe
CommonDialog1.HelpFile = MyPath & "\" & "Charm.hlp"
CommonDialog1.HelpCommand = HelpFinder
CommonDialog1.HelpCommand = cdlHelpContents

End Sub

Private Sub cmdOK_Click()
If cmbMaterials.Text = "" Then
MsgBox "Please choose a Category", vbOKOnly
Exit Sub
End If
Dim I
Dim MyPath
MyPath = App.Path
Dim Materials
  Materials = MyPath & "\" & "Materials.ini"
Dim MyString
Dim NewMat As String
Open Materials For Input As #1
Do While Not EOF(1)
Line Input #1, MyString
NewMat = NewMat & MyString & vbCrLf
If MyString = cmbMaterials.Text Then
For I = 0 To List1.ListCount - 1
NewMat = NewMat & List1.List(I) & vbCrLf
Next
End If

Loop
Close #1

Dim NewFile
NewFile = MyPath & "\" & "Materials.ini"
Open NewFile For Output As #1
Print #1, NewMat
Close #1
End Sub

Private Sub cndDelete_Click()
Dim index
'Index = List1.NewIndex
'List1.Selected(Index) = True
index = List1.ListIndex
List1.RemoveItem index
End Sub

Private Sub Form_Load()
Dim MyString
'*************
Dim MyPath
MyPath = App.Path
Dim Materials
  Materials = MyPath & "\" & "Materials.ini"
 cmbMaterials.Clear
 Open Materials For Input As #1     ' Open file for read.
  Do While Not EOF(1)
  Line Input #1, MyString
  If IsNumeric(Left(MyString, 1)) = False Then
 If MyString <> "" Then
  cmbMaterials.AddItem MyString
  End If
  End If
  Loop
  Close #1
End Sub
