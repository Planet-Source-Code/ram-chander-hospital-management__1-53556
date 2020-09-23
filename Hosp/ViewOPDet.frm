VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VewOPDet 
   Caption         =   "View Out Patient Details"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   8700
   Begin VB.TextBox Text3 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   2900
      Width           =   6975
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1800
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Text            =   "Name"
      Top             =   2900
      Width           =   3735
   End
   Begin VB.ComboBox cmbgndr 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "ViewOPDet.frx":0000
      Left            =   3720
      List            =   "ViewOPDet.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ComboBox cmbmarstat 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "ViewOPDet.frx":003C
      Left            =   5400
      List            =   "ViewOPDet.frx":004F
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox cmbcaste 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   300
      Left            =   9240
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   3620
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1605
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   11895
      Begin VB.TextBox Text10 
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   " "
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbregist 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "ViewOPDet.frx":0081
         Left            =   7920
         List            =   "ViewOPDet.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7920
         TabIndex        =   10
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   6120
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "ViewOPDet.frx":009D
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9120
      Top             =   360
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   2
         Text            =   "Date"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   1
         Text            =   "Time"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "St. Anns Jubilee Memorial Hospital"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2280
         TabIndex        =   5
         Top             =   0
         Width           =   7875
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visakhapatnam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5160
         TabIndex        =   4
         Top             =   960
         Width           =   1875
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Malkapuram"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5280
         TabIndex        =   3
         Top             =   600
         Width           =   1470
      End
   End
   Begin ComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24576003
      CurrentDate     =   38039
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1695
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   11775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   34
      Top             =   2700
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   33
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   32
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1980
      TabIndex        =   31
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   30
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   29
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7200
      TabIndex        =   28
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9120
      TabIndex        =   27
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H00000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      Height          =   1725
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   2700
      Width           =   11610
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H00000000&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   2
      Height          =   1905
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   9450
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1695
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   4785
      Width           =   9450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blood Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   195
      TabIndex        =   26
      Top             =   5025
      Width           =   1065
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   25
      Top             =   5025
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1650
      TabIndex        =   24
      Top             =   5025
      Width           =   1080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6120
      TabIndex        =   23
      Top             =   4800
      Width           =   840
   End
End
Attribute VB_Name = "VewOPDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
