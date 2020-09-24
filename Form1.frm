VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Waguih Estimating Program"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4920
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   0
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   7
      Top             =   360
      Width           =   6735
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   720
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   4575
      End
   End
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   1
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   9
      Top             =   360
      Width           =   6735
      Begin VB.CommandButton Command6 
         Caption         =   "Open Template"
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Customize Labor"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Customize Subcontracts"
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save Template"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Customize Equipment"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Customize Materials"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtCustomer 
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtCompany 
         Height          =   855
         Left            =   2520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "Form1.frx":09CA
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   2
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   20
      Top             =   360
      Width           =   6735
      Begin VB.TextBox txtTotall 
         BackColor       =   &H80000004&
         Height          =   405
         Index           =   0
         Left            =   4080
         TabIndex        =   39
         Text            =   "0"
         Top             =   3690
         Width           =   1695
      End
      Begin VB.ListBox List 
         Height          =   1230
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ListBox ListMargin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         ItemData        =   "Form1.frx":09D0
         Left            =   4680
         List            =   "Form1.frx":09D7
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox ListUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         ItemData        =   "Form1.frx":09E7
         Left            =   3600
         List            =   "Form1.frx":09EE
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "Materials Total"
         Height          =   255
         Left            =   2880
         TabIndex        =   56
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Material List"
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Margin%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   3
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   26
      Top             =   360
      Width           =   6735
      Begin VB.TextBox txtTotall 
         BackColor       =   &H80000004&
         Height          =   405
         Index           =   1
         Left            =   4560
         TabIndex        =   38
         Text            =   "0"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   33
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   32
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   30
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox ListMargin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         ItemData        =   "Form1.frx":09FD
         Left            =   4680
         List            =   "Form1.frx":0A04
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox ListUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         ItemData        =   "Form1.frx":0A14
         Left            =   3600
         List            =   "Form1.frx":0A1B
         TabIndex        =   28
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox List 
         Height          =   1425
         Index           =   1
         Left            =   480
         TabIndex        =   27
         Top             =   2280
         Width           =   5295
      End
      Begin VB.Label Label30 
         Caption         =   "Labor Total"
         Height          =   375
         Left            =   3120
         TabIndex        =   93
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Labor List:"
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Choose a Category"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Choose an Item"
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Margin%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   4
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   42
      Top             =   360
      Width           =   6735
      Begin VB.ListBox List 
         Height          =   1425
         Index           =   2
         Left            =   480
         TabIndex        =   50
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ListBox ListUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         ItemData        =   "Form1.frx":0A2A
         Left            =   3600
         List            =   "Form1.frx":0A31
         TabIndex        =   49
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox ListMargin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         ItemData        =   "Form1.frx":0A41
         Left            =   4680
         List            =   "Form1.frx":0A48
         TabIndex        =   48
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   47
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   46
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   45
         Top             =   720
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   44
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtTotall 
         BackColor       =   &H80000004&
         Height          =   405
         Index           =   2
         Left            =   4560
         TabIndex        =   43
         Text            =   "0"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label31 
         Caption         =   "Equipment Total"
         Height          =   255
         Left            =   3120
         TabIndex        =   94
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Margin%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   55
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   54
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Choose an tem"
         Height          =   375
         Left            =   360
         TabIndex        =   53
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Choose a Category"
         Height          =   255
         Left            =   360
         TabIndex        =   52
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Equipment Lst"
         Height          =   375
         Left            =   480
         TabIndex        =   51
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   5
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   78
      Top             =   360
      Width           =   6735
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   86
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   85
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   84
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   83
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox ListUnits 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         ItemData        =   "Form1.frx":0A59
         Left            =   3600
         List            =   "Form1.frx":0A60
         TabIndex        =   82
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox ListMargin 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         ItemData        =   "Form1.frx":0A70
         Left            =   4680
         List            =   "Form1.frx":0A77
         TabIndex        =   81
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox List 
         Height          =   1230
         Index           =   3
         Left            =   480
         TabIndex        =   80
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox txtTotall 
         BackColor       =   &H80000004&
         Height          =   405
         Index           =   3
         Left            =   4080
         TabIndex        =   79
         Text            =   "0"
         Top             =   3690
         Width           =   1695
      End
      Begin VB.Label Label29 
         Caption         =   "Choose Category"
         Height          =   255
         Left            =   360
         TabIndex        =   92
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Choose Item"
         Height          =   375
         Left            =   360
         TabIndex        =   91
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   90
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label26 
         Caption         =   "Margin%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   89
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "SubContract List"
         Height          =   375
         Left            =   480
         TabIndex        =   88
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Contracts Total"
         Height          =   255
         Left            =   2880
         TabIndex        =   87
         Top             =   3840
         Width           =   1095
      End
   End
   Begin VB.PictureBox Page 
      Height          =   4335
      Index           =   6
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6675
      TabIndex        =   57
      Top             =   360
      Width           =   6735
      Begin VB.CommandButton cmdViewWithout 
         Caption         =   "View Report Without Details"
         Height          =   375
         Left            =   480
         TabIndex        =   77
         Top             =   3720
         Width           =   5535
      End
      Begin VB.CommandButton cmdViewDetails 
         Caption         =   "View Report With Details"
         Height          =   375
         Left            =   480
         TabIndex        =   76
         Top             =   3240
         Width           =   5535
      End
      Begin VB.TextBox txtProject 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtAfterMargin 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   3
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtAfterMargin 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   2
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtAfterMargin 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   1
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtAfterMargin 
         BackColor       =   &H80000004&
         Height          =   405
         Index           =   0
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtResultMargin 
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   65
         Text            =   "1.25"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtResultMargin 
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   64
         Text            =   "1.25"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtResultMargin 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   63
         Text            =   "1.25"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtResultMargin 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   62
         Text            =   "1.25"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   3
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   2
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   1
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H80000004&
         Height          =   375
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Grand Total"
         Height          =   375
         Left            =   2760
         TabIndex        =   75
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "SubContracts Total"
         Height          =   375
         Left            =   240
         TabIndex        =   74
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Equipment Total"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "Labor Total"
         Height          =   375
         Left            =   240
         TabIndex        =   72
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Material Total"
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Materials"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Labor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Equipment"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Subcontracts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Result"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurentTextIndex As Integer

Private Sub cmdAdd_Click(index As Integer)

index = cmdAdd(index).index
If Combo2(index).Text <> "" Then
Dim pos, Price
pos = InStr(Combo2(index).Text, " ")
Price = Left(Combo2(index), pos - 1) / 100

Dim MyUnit
MyUnit = ListUnits(index).Text
If MyUnit <> "" Then
List(index).AddItem MyUnit & " of " & "$" & Price & " " & Mid(Combo2(index).Text, pos + 1, Len(Combo2(index).Text) - pos)

txtTotall(index).Text = txtTotall(index).Text + CSng(Price) * MyUnit
txtTotall(index).Text = Format(txtTotall(index).Text, "##.00")
txtResult(index).Text = txtTotall(index).Text
txtAfterMargin(index).Text = txtResult(index).Text * txtResultMargin(index).Text

TotalProject
End If
End If

End Sub

Private Sub cmdViewDetails_Click()
'Me.Hide
frmViewer.Show
End Sub

Private Sub Combo1_Click(index As Integer)

index = Combo1(index).index
Combo2(index).Clear
Dim MyPath
MyPath = App.Path
'**************
Dim MyFile
If index = 0 Then
 MyFile = MyPath & "\" & "Materials.ini"
 End If
If index = 1 Then
MyFile = MyPath & "\" & "Labor.ini"
End If
If index = 2 Then
MyFile = MyPath & "\" & "Equipment.ini"
End If
If index = 3 Then
MyFile = MyPath & "\" & "Contracts.ini"
End If
'**************
 Dim MyString, MyNumber
 
  Open MyFile For Input As #1     ' Open file for read.

  Do While Not EOF(1)
     
      Line Input #1, MyString
20:
   If MyString = Combo1(index).Text Then
   'cmbContract2.Clear
   GoTo 10
   End If

   Loop
10:
  Do While Not EOF(1)
  Line Input #1, MyString
  
    If MyString <> "" Then
    Combo2(index).AddItem MyString
    Else
    GoTo 20
    End If
  Loop
Close #1

End Sub

Private Sub Command1_Click()
frmMaterial.Show
End Sub

Private Sub Command2_Click()
frmEquipment.Show
End Sub

Private Sub Command4_Click()
frmSubContract.Show
End Sub

Private Sub Command5_Click()
frmLabor.Show
End Sub


Private Sub Form_Activate()
Dim MyString
Dim MyPath
MyPath = App.Path

Dim MyFile
Dim index
For index = 0 To 3
  If index = 0 Then
  MyFile = MyPath & "\" & "Materials.ini"
  End If
  If index = 1 Then
  MyFile = MyPath & "\" & "Labor.ini"
  End If
  If index = 2 Then
  MyFile = MyPath & "\" & "Equipment.ini"
  End If
  If index = 3 Then
  MyFile = MyPath & "\" & "Contracts.ini"
  End If
 
 
 Open MyFile For Input As #1     ' Open file for read.
  Do While Not EOF(1)
  Line Input #1, MyString
  If IsNumeric(Left(MyString, 1)) = False Then
 If MyString <> "" Then
  Combo1(index).AddItem MyString
  End If
  End If
  Loop
  Close #1
  Next index

End Sub

Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path & "\charm.ico")
CurentTextIndex = 0

Text1.Text = "Waguih Estimating Software V1.1"
txtCompany.Text = "Waguih High Technology Services" & vbCrLf & "11,Hussain Shfik El-Masry Str" & vbCrLf & "Cairo-Egypt"
txtCustomer.Text = "Egypt Gas"
Label1.Caption = "Company Name" & vbCr & "Contract Informations"
Label2.Caption = "Customer Name"
Label3.Caption = "choose a Category"
Label4.Caption = "Choose an Item"
Dim I
For I = 1 To 9
ListUnits(0).AddItem I, I - 1
ListUnits(1).AddItem I, I - 1
ListUnits(2).AddItem I, I - 1
ListUnits(3).AddItem I, I - 1
Next
For I = 1 To 150
ListMargin(0).AddItem I, I - 1
ListMargin(1).AddItem I, I - 1
ListMargin(2).AddItem I, I - 1
ListMargin(3).AddItem I, I - 1
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub TabStrip1_Click()
   
   Page(TabStrip1.SelectedItem.index - 1).Visible = True
   Page(CurentTextIndex).Visible = False
   ' Set mintCurFrame to new value.
   CurentTextIndex = TabStrip1.SelectedItem.index - 1
   
End Sub

Private Sub txtResultMargin_Change(index As Integer)

index = txtResultMargin(index).index
txtAfterMargin(index).Text = txtResult(index).Text * txtResultMargin(index).Text
TotalProject
End Sub

Public Sub TotalProject()
Dim a
a = Array(0, 1, 2, 3)
Dim I

For I = 0 To 3
If txtAfterMargin(I).Text <> "" Then
a(I) = CSng(txtAfterMargin(I).Text)
Else
a(I) = 0
End If
Next
txtProject.Text = a(0) + a(1) + a(2) + a(3)

End Sub

Private Sub cmdDelete_Click(index As Integer)


index = cmdDelete(index).index
Dim MyIndex
MyIndex = List(index).ListIndex
List(index).RemoveItem MyIndex
List(index).Refresh
Dim I, x, z
Dim a
Dim b
Dim NewTotal
NewTotal = 0
Dim pos, Pos2
Dim mstr
For I = 0 To List(index).ListCount - 1
mstr = CStr(List(index).List(I))
pos = InStr(mstr, " ")
a = Mid(mstr, 1, pos - 1)
Pos2 = InStr(mstr, "$")
    For x = Pos2 + 1 To Len(mstr) - Pos2 - 1
    b = Mid$(mstr, x, 1)
   
   If b <> " " Then
    z = z & b
    Else
    GoTo 10
    End If
    Next
10:
NewTotal = NewTotal + CSng(a) * CSng(Val(z))
z = ""
Next
txtTotall(index).Text = Format(NewTotal, "##.00")
txtResult(index).Text = txtTotall(index)
txtAfterMargin(index).Text = txtResult(index).Text * txtResultMargin(index).Text

TotalProject

End Sub


