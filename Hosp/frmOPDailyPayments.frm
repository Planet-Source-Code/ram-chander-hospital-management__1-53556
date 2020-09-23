VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOPDailyPayments 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Out Patients Daily Payments"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   FillColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   7
         Text            =   "Date"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   6
         Text            =   "Time"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Management"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   1440
         TabIndex        =   8
         Top             =   45
         Width           =   4050
      End
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      ToolTipText     =   "Click To Display Report"
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Click To Close"
      Top             =   2430
      Width           =   1455
   End
   Begin ComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   1950
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   37985
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "OUT PATIENTS DAILY PAYMENTS REPORT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   6765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Select The Date :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1365
      TabIndex        =   3
      Top             =   1680
      Width           =   1650
   End
End
Attribute VB_Name = "frmOPDailyPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReport_Click()
DataEnvironment1.Command1_Grouping DTPicker1.Value
OPDailyPayments.Show
End Sub
