VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmdischarge 
   Caption         =   "Discharge Form"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   9390
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   33
         Text            =   "Date"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   32
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
         Left            =   2760
         TabIndex        =   34
         Top             =   45
         Width           =   4050
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Click to Discharge"
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7320
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   25
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6720
      TabIndex        =   22
      Text            =   "Text6"
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MFG1 
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   7080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      WordWrap        =   -1  'True
      Enabled         =   0   'False
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Text            =   " "
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   7680
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      WordWrap        =   -1  'True
      Enabled         =   0   'False
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   975
      Left            =   6120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   9000
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin ComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   38090
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISCHARGE DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   360
      Left            =   4440
      TabIndex        =   27
      Top             =   1400
      Width           =   3135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Rent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4440
      TabIndex        =   26
      Top             =   4605
      Width           =   1380
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Days in Hospital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   24
      Top             =   7200
      Width           =   1965
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   23
      Top             =   7800
      Width           =   1590
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ward Joined :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   1665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   15
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Type :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosis Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   9
      Top             =   3000
      Width           =   2115
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Examined"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   1785
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discharge Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Admission :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   5
      Top             =   2280
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   3375
   End
End
Attribute VB_Name = "frmdischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nod As Variant
Dim tot, tot1, inpchrgs As Integer
Dim disdb As ADODB.Connection
Dim disrs As ADODB.Recordset
Private Sub Command1_Click()
Set disdb = New ADODB.Connection
Set disrs = New ADODB.Recordset
    disdb.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
disrs.Open "SELECT * FROM DISCHARGE", disdb
With rs
    .AddNew
        !pcode = List1.Text
        !pname = Text1.Text
        !DOA = Text3.Text
        !dod = DTPicker1.Value
        !docname = Text4.Text
        !DIAGNOSIS = Text5.Text
        !wardname = Text8.Text
        !ROOMTYPE = Text1.Text
        !bedno = Text7.Text
        !ROOMRENT = Text9.Text
        !DAYSINHOS = Text6.Text
        !totalamt = txtgrndtot.Text
    .Update
End With
End Sub

Private Sub Command3_Click()
con.BeginTrans
    Set rs = con.Execute("DELETE * FROM INPATIENTS where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM BILL where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM BillPayments where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM BillDetails where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM IPEXAMS where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM IPMEDICINE where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM IPSURGERY where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM CustomerAdvance where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM Room where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM CustomerDue where PCODE=" & List1.Text)
    rs.Close
    Set rs = con.Execute("DELETE * FROM CustomerNetValue where PCODE=" & List1.Text)
    rs.Close
End Sub

Private Sub Command4_Click()
If List1.Text = "" Then
    MsgBox "Plz Select the Patient Code", , "Report"
Else
DataEnvironment1.Discharge (List1.Text)
rptDischarge.Show
End If
End Sub
Private Sub Command5_Click()
txtgrndtot.Text = Val(tot) + Val(tot1) + Val(Text9.Text) + Val(inpchrgs)
End Sub

Private Sub Form_Activate()
Call Connection.connected
Call MFGVALUES
con.BeginTrans
    Set rs = con.Execute("SELECT PCODE FROM INPATIENTS")
    Do Until rs.EOF
        List1.AddItem (rs(0))
        rs.MoveNext
    Loop
con.CommitTrans
DTPicker1.SetFocus
End Sub
Private Sub List1_Click()
On Error Resume Next
tot = 0
tot1 = 0
inpchrgs = 0
con.BeginTrans
    Set rs = con.Execute("SELECT NAME,DOA,DOCEXAMINED,DIAGNOSIS,WARDJOINED,ROOMTYPE,BEDNO,ROOMRENT from INPATIENTS where pcode=" & List1.Text)
     Text2.Text = rs(0)
     Text3.Text = rs(1)
     Text4.Text = rs(2)
     Text5.Text = rs(3)
     Text8.Text = rs(4)
     Text1.Text = rs(5)
     Text7.Text = rs(6)
     Text9.Text = rs(7)
Text6.Text = DateDiff("d", Format(DateValue(Text3.Text), mm - dd - yy), Format(DTPicker1.Value, mm - dd - yy))
inpchrgs = Val(rs(7)) * Val(Text6.Text)
Set rs = con.Execute("SELECT * from IPMEDICINE where pcode =" & List1.Text)
MFG.Rows = 2
If rs.EOF = True Then
    MFG.Clear
    Call MFGVALUES
End If
Do Until rs.EOF
    MFG.Rows = MFG.Rows + 1
    j = MFG.Rows
    i = MFG.Rows - 2
    MFG.TextMatrix(i, 0) = rs(2)
    MFG.TextMatrix(i, 1) = rs(3)
    MFG.TextMatrix(i, 2) = rs(4)
    MFG.TextMatrix(i, 3) = rs(5)
    MFG.TextMatrix(i, 4) = rs(6)
    MFG.TextMatrix(i, 5) = rs(7)
    MFG.TextMatrix(i, 6) = rs(8)
    MFG.TextMatrix(i, 7) = rs(9)
tot = tot + Val(rs(9))
rs.MoveNext
Loop
rs.Close
Set rs = con.Execute("SELECT EXAMTYPE,EXAMDATE,EXAMDETAILS,TOTALAMOUNT FROM IPEXAMS WHERE PCODE =" & List1.Text)
MFG1.Rows = 2
If rs.EOF = True Then
    MFG1.Clear
    Call MFGVALUES
End If
Do Until rs.EOF
    MFG1.Rows = MFG1.Rows + 1
    j = MFG1.Rows
    i = MFG.Rows - 2
    MFG1.TextMatrix(i, 0) = rs(0)
    MFG1.TextMatrix(i, 1) = rs(1)
    MFG1.TextMatrix(i, 2) = rs(2)
    MFG1.TextMatrix(i, 3) = rs(3)
    MFG1.TextMatrix(i, 4) = rs(4)
tot1 = tot1 + Val(rs(4))
rs.MoveNext
Loop
rs.Close
Set rs = con.Execute("SELECT GRANDTOTAL from BILL where PCODE = " & List1.Text)
If rs.EOF = True Then
    txtgrndtot.Text = 0
Else
    txtgrndtot.Text = rs(0)
End If
con.CommitTrans
End Sub
Public Sub MFGVALUES()
MFG.TextMatrix(0, 0) = "MEDICINE NAME"
MFG.TextMatrix(0, 1) = "MEDICINE TYPE"
MFG.TextMatrix(0, 2) = "DATE OF ISSUE"
MFG.TextMatrix(0, 3) = "QUANTITY"
MFG.TextMatrix(0, 4) = "RATE PER UNIT "
MFG.TextMatrix(0, 5) = "AMOUNT"
MFG.TextMatrix(0, 6) = "DISCOUNT"
MFG.TextMatrix(0, 7) = "TOTAL AMOUNT"
MFG1.TextMatrix(0, 0) = "EXAM TYPE"
MFG1.TextMatrix(0, 1) = "EXAM DATE"
MFG1.TextMatrix(0, 2) = "EXAM DETAILS"
MFG1.TextMatrix(0, 3) = "TOTAL AMOUNT"
End Sub
