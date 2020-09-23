VERSION 5.00
Begin VB.Form frmIPBillPaymentReport 
   BackColor       =   &H00C96C59&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "In Patient Bill Report"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmIPBillReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   9
         Text            =   "Date"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   8
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
         Left            =   3000
         TabIndex        =   10
         Top             =   45
         Width           =   4050
      End
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
      Height          =   345
      Left            =   4350
      TabIndex        =   3
      ToolTipText     =   "Click To Close Window"
      Top             =   2790
      Width           =   1455
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
      Height          =   345
      Left            =   2730
      TabIndex        =   2
      ToolTipText     =   "Click To Display Report"
      Top             =   2790
      Width           =   1455
   End
   Begin VB.ComboBox cmbBillNo 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   6405
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select The Patient Bill Number"
      Top             =   1980
      Width           =   1875
   End
   Begin VB.ComboBox cmbCustName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1845
      TabIndex        =   0
      Text            =   "Patient Code"
      ToolTipText     =   "Select The Patient Name"
      Top             =   1980
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      Caption         =   "Patient Bill No :"
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
      Left            =   4755
      TabIndex        =   6
      Top             =   2070
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      Caption         =   "Patient Code :"
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
      Left            =   135
      TabIndex        =   5
      Top             =   2070
      Width           =   1350
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "PATIENT BILL REPORT"
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
      Left            =   210
      TabIndex        =   4
      Top             =   1320
      Width           =   8565
   End
End
Attribute VB_Name = "frmIPBillPaymentReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCustName_DropDown()
con.BeginTrans
Set rs = con.Execute("SELECT PCODE FROM INPATIENTS")
    Do Until rs.EOF
        cmbCustName.AddItem (rs(0))
        rs.MoveNext
    Loop
End Sub
Private Sub cmbCustName_LostFocus()
On Error Resume Next
Set rs = con.Execute("Select BillId,BillNo from Bill where PCODE =" & cmbCustName.Text)
If rs.EOF = True Then
    cmbBillNo.Clear
    rs.Close
    Exit Sub
Else
    cmbBillNo.Clear
    Do While rs.EOF = False
        cmbBillNo.AddItem (rs(1))
        cmbBillNo.ItemData(i) = rs(0)
        rs.MoveNext
    Loop
End If
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdReport_Click()
If cmbBillNo.Text = "" Then
MsgBox "Sorry !!!!!!!! Please Select the Bill ", vbInformation
Else
DataEnvironment1.IPBillPayments (cmbBillNo.Text)
IPBillReport.Show
End If
End Sub
Private Sub cmOk_Click()
'Dim i As Integer
'Set rs = con.Execute("Select Name,PCODE from INPATIENTS where PCODE in (Select PCODE from Bill where BillDate=#" & Format(DTPicker1.Value, "mm/dd/yyyy") & "#)")
'If rs.EOF = True Then
'    rs.Close
'    Call Txt_Clear
'    MsgBox "No Bills On This Dates...", vbInformation + vbOKOnly
'Else
'    i = 0
'    Call Txt_Clear
'    Do While rs.EOF = False
'        cmbCustName.AddItem (rs(0))
'        cmbCustName.ItemData(i) = rs(1)
'        i = i + 1
'        rs.MoveNext
'    Loop
'    rs.Close
'End If
End Sub
Private Sub Form_Load()
'If con.State Then con.Close
Call Connection.connected
'Call Txt_Clear
'DTPicker1.Value = Now
End Sub
Private Sub Txt_Clear()
cmbCustName.Clear
cmbBillNo.Clear
End Sub
