VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmadmlst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Admission List"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmAdmLst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   6
         Text            =   "Time"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   5
         Text            =   "Date"
         Top             =   0
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
         Left            =   2400
         TabIndex        =   7
         Top             =   -75
         Width           =   4050
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Click to Display the List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      WordWrap        =   -1  'True
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin ComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   38090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Date of Admission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4020
   End
End
Attribute VB_Name = "frmadmlst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
con.BeginTrans
Set rs = con.Execute("SELECT name,docexamined,wardjoined,roomtype,bedno FROM INPATIENTS WHERE DOA=" & Format(DTPicker1.Value, mm - dd - yy))
MFG.Rows = 2
Do Until rs.EOF
    MFG.Rows = MFG.Rows + 1
    j = MFG.Rows
    i = MFG.Rows - 2
    MFG.TextMatrix(i, 0) = rs(0)
    MFG.TextMatrix(i, 1) = rs(1)
    MFG.TextMatrix(i, 2) = rs(2)
    MFG.TextMatrix(i, 3) = rs(3)
    MFG.TextMatrix(i, 4) = rs(4)
rs.MoveNext
Loop
rs.Close
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Activate()
On Error Resume Next
Call Connection.connected
    MFG.TextMatrix(i, 0) = "Patient Name"
    MFG.TextMatrix(i, 1) = "Doctor Examnied"
    MFG.TextMatrix(i, 2) = "Ward Joined"
    MFG.TextMatrix(i, 3) = "Room Type"
    MFG.TextMatrix(i, 4) = "Bed No"
    

End Sub

