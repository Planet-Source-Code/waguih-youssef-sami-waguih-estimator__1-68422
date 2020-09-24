VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewer 
   Caption         =   "Waguih Viewer"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog COMDialog1 
      Left            =   480
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RTBox1 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Viewer.frx":0000
   End
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7095
      TabIndex        =   0
      Top             =   0
      Width           =   7155
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Text            =   "10"
         Top             =   0
         Width           =   840
      End
      Begin VB.ComboBox cboFontsScreen 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Text            =   "Times New Roman"
         Top             =   0
         Width           =   2325
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   375
         Left            =   5070
         Picture         =   "Viewer.frx":00E9
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdFind 
         Height          =   375
         Left            =   5520
         Picture         =   "Viewer.frx":061B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdPast 
         Height          =   350
         Left            =   4635
         Picture         =   "Viewer.frx":071D
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Paste"
         Top             =   0
         Width           =   350
      End
      Begin VB.CommandButton cmdCopy 
         Height          =   350
         Left            =   4215
         Picture         =   "Viewer.frx":0C4F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Copy"
         Top             =   0
         Width           =   350
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   350
         Left            =   3780
         Picture         =   "Viewer.frx":1181
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   350
      End
      Begin VB.CommandButton cmdSave 
         Height          =   350
         Left            =   3360
         Picture         =   "Viewer.frx":16B3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save Waguih"
         Top             =   0
         Width           =   350
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuPicture 
         Caption         =   "&Picture"
      End
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboFontSize_Click()
RTBox1.SelFontSize = cboFontSize.Text
End Sub

Private Sub cboFontsScreen_Click()
RTBox1.SelFontName = cboFontsScreen.Text
End Sub

Private Sub cmdCopy_Click()
Clipboard.SetText RTBox1.SelText
End Sub

Private Sub cmdFind_Click()
z = InputBox("Text to Find", "Find Dialog")
RTBox1.Find z
Dim FoundPos, FoundLine As Integer
FoundPos = RTBox1.Find(z, , , rtfWholeWord)
If FoundPos <> -1 Then
FoundLine = RTBox1.GetLineFromChar(FoundPos)
MsgBox "Word found on Line" & CStr(FoundLine)
Else
MsgBox "Word not found"
End If
RTBox1.SetFocus

End Sub

Private Sub cmdOpen_Click()
COMDialog1.Action = 1
COMDialog1.Flags = cdlOFNNoChangeDir
COMDialog1.DialogTitle = "Waguih Ask you Open...?"
RTBox1.LoadFile COMDialog1.filename
RTBox1.Font.Name = "Times New Roman"
RTBox1.Font.Size = 12
RTBox1.Refresh
End Sub

Private Sub cmdPast_Click()
RTBox1.SelText = Clipboard.GetText()
'RTBox1.Font = cboFontsScreen
'RTBox1.Font.Size = cboFontSize
RTBox1.Refresh
End Sub

Private Sub cmdPrint_Click()
frmViewer.Print cboFontsScreen.Text
End Sub

Private Sub cmdSave_Click()
COMDialog1.DialogTitle = "Waguih Ask you Save As...?"
COMDialog1.Action = 2
RTBox1.SaveFile COMDialog1.filename
COMDialog1.Flags = cdlOFNOverwritePrompt
End Sub

Private Sub Form_Activate()
Dim pos
pos = InStr(RTBox1.Text, "Company")
RTBox1.SelStart = pos - 1
RTBox1.SelLength = 15
RTBox1.SelBold = True
RTBox1.SelUnderline = True

pos = InStr(RTBox1.Text, "CLient")
RTBox1.SelStart = pos - 1
RTBox1.SelLength = 11
RTBox1.SelBold = True
RTBox1.SelUnderline = True

pos = InStr(RTBox1.Text, "I-Materials")
RTBox1.SelStart = pos - 1
RTBox1.SelLength = 11
RTBox1.SelBold = True
RTBox1.SelUnderline = True

pos = InStr(RTBox1.Text, "II-Labor")
RTBox1.SelStart = pos - 1
RTBox1.SelLength = 7
RTBox1.SelBold = True
RTBox1.SelUnderline = True

pos = InStr(RTBox1.Text, "III-Equipment")
RTBox1.SelStart = pos - 1
RTBox1.SelLength = 13
RTBox1.SelBold = True
RTBox1.SelUnderline = True

pos = InStr(RTBox1.Text, "IV-SubContracs")
RTBox1.SelStart = pos - 1
RTBox1.SelLength = 14
RTBox1.SelBold = True
RTBox1.SelUnderline = True

RTBox1.SelStart = 0

End Sub

Private Sub Form_Load()

'**************combo box with fonts
NumOfScreenFonts = Screen.FontCount
Dim I
For I = 0 To NumOfScreenFonts - 1 Step 1
cboFontsScreen.AddItem Screen.Fonts(I)
Next

'*************combo box for font size
Dim x
For x = 10 To 16
cboFontSize.AddItem (x), x - 10
Next
'---------------
COMDialog1.Filter = "All Files(*.*)|*.*|Text Files(*.txt)|*.txt"
COMDialog1.Flags = cdlOFNCreatePrompt
'frmDir.Enabled = False
'------------------------
'RTBox1.Width = frmViewer.Width - 560
'RTBox1.Height = frmViewer.Height - 1000
'*****************write report
RTBox1.Font.Name = "Times New Roman"
RTBox1.Font.Size = "12"

RTBox1.Text = "Company Details" & vbCrLf
RTBox1.Text = RTBox1.Text & Form1.txtCompany.Text & vbCrLf
RTBox1.Text = RTBox1.Text & "**************************" & vbCrLf
RTBox1.Text = RTBox1.Text & "CLient Name" & vbCrLf
RTBox1.Text = RTBox1.Text & Form1.txtCustomer & vbCrLf
RTBox1.Text = RTBox1.Text & "**************************" & vbCrLf

Dim K
Dim mstr
'****************
RTBox1.Text = RTBox1.Text & "I-Materials" & vbCrLf
For K = 0 To Form1.List(0).ListCount - 1
mstr = CStr(Form1.List(0).List(K))
RTBox1.Text = RTBox1.Text & mstr & vbCrLf
Next
RTBox1.Text = RTBox1.Text & "Total Material Cost= " & Form1.txtResult(0).Text & vbCrLf
RTBox1.Text = RTBox1.Text & "Total Material Cost After " & Form1.txtResultMargin(0).Text & "% Margin= " & Form1.txtAfterMargin(0).Text & vbCrLf & vbCrLf
'******************
RTBox1.Text = RTBox1.Text & "II-Labor" & vbCrLf
For K = 0 To Form1.List(1).ListCount - 1
mstr = CStr(Form1.List(1).List(K))
RTBox1.Text = RTBox1.Text & mstr & vbCrLf
Next
RTBox1.Text = RTBox1.Text & "Total Labor Cost= " & Form1.txtResult(1).Text & vbCrLf
RTBox1.Text = RTBox1.Text & "Total Labor Cost After " & Form1.txtResultMargin(1).Text & "% Margin= " & Form1.txtAfterMargin(1).Text & vbCrLf & vbCrLf
'*******************
RTBox1.Text = RTBox1.Text & "III-Equipment" & vbCrLf
For K = 0 To Form1.List(2).ListCount - 1
mstr = CStr(Form1.List(2).List(K))
RTBox1.Text = RTBox1.Text & mstr & vbCrLf
Next
RTBox1.Text = RTBox1.Text & "Total Equipment Cost= " & Form1.txtResult(2).Text & vbCrLf
RTBox1.Text = RTBox1.Text & "Total Equipment Cost After " & Form1.txtResultMargin(2).Text & "% Margin= " & Form1.txtAfterMargin(2).Text & vbCrLf & vbCrLf
'********************
RTBox1.Text = RTBox1.Text & "IV-SubContracs" & vbCrLf
For K = 0 To Form1.List(3).ListCount - 1
mstr = CStr(Form1.List(3).List(K))
RTBox1.Text = RTBox1.Text & mstr & vbCrLf
Next
RTBox1.Text = RTBox1.Text & "Total SubContracts Cost= " & Form1.txtResult(3).Text & vbCrLf
RTBox1.Text = RTBox1.Text & "Total SubContracts Cost After " & Form1.txtResultMargin(3).Text & "% Margin= " & Form1.txtAfterMargin(3).Text & vbCrLf & vbCrLf



End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form1.Show
End Sub

Private Sub mnuCopy_Click()
Clipboard.SetText RTBox1.SelText
End Sub

Private Sub mnuExit_Click()
'COMDialog1.Flags = cdlOFNCreatePrompt
'End
Me.Hide
Form1.Show
End Sub

Private Sub mnuOpen_Click()
COMDialog1.Action = 1
COMDialog1.Flags = cdlOFNNoChangeDir
COMDialog1.DialogTitle = "Waguih Ask you Open...?"
RTBox1.LoadFile COMDialog1.filename
End Sub

Private Sub mnuPaste_Click()
RTBox1.SelText = Clipboard.GetText()
RTBox1.Refresh
End Sub

'Private Sub mnuPicture_Click()
'frmOption.Left = (frmViewer.Left + 2000)
'frmOption.Top = (frmViewer.Top + 500)
'frmOption.Show
'frmOption.optCFF.Value = False
'frmOption.optCNF.Value = False
'
'End Sub

Private Sub mnuSave_Click()
COMDialog1.DialogTitle = "Waguih Ask you Save As...?"
COMDialog1.Action = 2
RTBox1.SaveFile COMDialog1.filename
COMDialog1.Flags = cdlOFNOverwritePrompt

End Sub

Private Sub RTBox1_DragOver(source As Control, x As Single, y As Single, State As Integer)
'frmDir.File1
End Sub

Private Sub RTBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuEdit
End If
End Sub
