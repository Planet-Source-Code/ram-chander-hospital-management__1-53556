VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmipexm 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patients Exam Details"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9480
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   10440
      TabIndex        =   35
      Text            =   "Combo3"
      Top             =   600
      Width           =   1335
   End
   Begin ComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   8280
      TabIndex        =   32
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   24641537
      CurrentDate     =   38088
   End
   Begin VB.TextBox txtbill 
      Height          =   285
      Left            =   6360
      TabIndex        =   31
      Top             =   560
      Width           =   855
   End
   Begin VB.TextBox txtpayable 
      Height          =   300
      Left            =   8880
      TabIndex        =   27
      Top             =   5325
      Width           =   1575
   End
   Begin VB.TextBox txtdisgvn 
      Height          =   300
      Left            =   8880
      TabIndex        =   26
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
      Height          =   300
      Left            =   8880
      TabIndex        =   25
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CLOSE"
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
      Left            =   7200
      TabIndex        =   24
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create Bill"
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
      Left            =   7200
      TabIndex        =   23
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CLEAR"
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
      Left            =   8760
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
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
      Left            =   5760
      TabIndex        =   21
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Exam Details"
      Height          =   1935
      Left            =   4440
      TabIndex        =   11
      Top             =   1440
      Width           =   7335
      Begin VB.TextBox txtdis 
         Height          =   285
         Left            =   5520
         TabIndex        =   39
         Top             =   720
         Width           =   1695
      End
      Begin ComCtl2.DTPicker exmdt 
         Height          =   300
         Left            =   5520
         TabIndex        =   37
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   38088
      End
      Begin VB.TextBox txttotamt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtamt 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox exmdet 
         Height          =   495
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   38
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   19
         Top             =   1210
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3960
         TabIndex        =   18
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Result"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Test Conducted"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1650
      End
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   120
      Picture         =   "frmipexm.frx":0000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   5655
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9720
      TabIndex        =   36
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7320
      TabIndex        =   34
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   33
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Payable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   30
      Top             =   5355
      Width           =   1335
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7920
      TabIndex        =   29
      Top             =   4845
      Width           =   915
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   28
      Top             =   4350
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN PATIENTS EXAM DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   2880
      TabIndex        =   10
      Top             =   50
      Width           =   4245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Patient"
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
      Left            =   195
      TabIndex        =   9
      Top             =   1560
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Admission"
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
      Left            =   45
      TabIndex        =   8
      Top             =   2760
      Width           =   2220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
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
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment Date"
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
      Top             =   3240
      Width           =   1905
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Code"
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
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   2640
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmipexm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim meddb As ADODB.Connection
Dim medrs As ADODB.Recordset
Dim ID, bllinc As Integer
Dim txt1, txt As String
Private Sub Combo1_DropDown()
On Error Resume Next
ID = 0
con.BeginTrans
    Set rs = con.Execute("SELECT PCODE FROM INPATIENTS")
    Do Until rs.EOF
        Combo1.AddItem rs(0)
      rs.MoveNext
    Loop
rs.Close
con.CommitTrans
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo3.SetFocus
End If
End Sub
Private Sub Combo1_LostFocus()
On Error Resume Next
txt = Combo1.Text
con.BeginTrans
    Set rs = con.Execute("SELECT NAME,DOA,DOCEXAMINED,TREATMENTDATE from INPATIENTS where pcode=" & Combo1.Text)
     Text2.Text = rs(0)
     Text4.Text = rs(1)
     Text5.Text = rs(2)
     Text6.Text = rs(3)
rs.Close
Combo1.Clear
Combo1.Text = txt
Set rs = con.Execute("SELECT GRANDTOTAL from BILL where PCODE = " & Combo1.Text)
If rs.EOF = True Then
    txtgrndtot.Text = 0
Else
    txtgrndtot.Text = rs(0)
End If
con.CommitTrans
End Sub
Private Sub Combo2_DropDown()
On Error Resume Next
con.BeginTrans
 Set rs = con.Execute("SELECT LABTESTS from LABEXAMS")
    Do Until rs.EOF
      Combo2.AddItem (rs(0))
      rs.MoveNext
    Loop
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    exmdt.SetFocus
End If
End Sub
Private Sub Command1_Click()
On Error Resume Next
Set meddb = New ADODB.Connection
Set medrs = New ADODB.Recordset
meddb.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hospital.mdb"
medrs.Open "SELECT * FROM IPEXMAS ", meddb, adOpenDynamic, adLockOptimistic
        With medrs
            .AddNew
                !PCODE = Combo1.Text
                !Name = Text2.Text
                !EXAMTYPE = Combo2.Text
                !EXAMDATE = exmdt.Value
                !EXAMDETAILS = exmdet.Text
                !EXAMCOST = txtamt.Text
                !RATEPERUNIT = txtrpu.Text
                !DISCOUTGIVEN = txtdis.Text
                !TOTALAMOUNT = txttotamt.Text
            .Update
        End With
ID = ID + Val(txttotamt.Text) + Val(txtgrndtot.Text)
medrs.Close
txtgrndtot.Text = ID
End Sub
Private Sub Command5_Click()
On Error Resume Next
Set meddb = New ADODB.Connection
Set medrs = New ADODB.Recordset
meddb.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hospital.mdb"
medrs.Open "SELECT * FROM Bill", meddb, adOpenDynamic, adLockOptimistic
    With medrs
        .AddNew
            !billid = txtbill.Text
            !billno = txtbill.Text
            !BillDate = billdt.Value
            !PCODE = Combo1.Text
            !CreditYN = cmbcrdt.Text
            !GrandTotal = txtgrndtot.Text
            !Discount = txtdisgvn.Text
            !NetValue = txtpayable.Text
            .Update
    End With
End Sub

Private Sub Command6_Click()
Unload Me
End Sub
Private Sub Command3_Click()
On Error Resume Next
Set meddb = New ADODB.Connection
Set medrs = New ADODB.Recordset
meddb.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hospital.mdb"
medrs.Open "SELECT * FROM Bill", meddb, adOpenDynamic, adLockOptimistic
    With medrs
        .AddNew
            !billid = txtbill.Text
            !billno = txtbill.Text
            !BillDate = billdt.Value
            !PCODE = Combo1.Text
            !CreditYN = cmbcrdt.Text
            !GrandTotal = txtgrndtot.Text
            !Discount = txtdisgvn.Text
            !NetValue = txtpayable.Text
            .Update
    End With
End Sub

Private Sub exmdt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtamt.SetFocus
End If
End Sub
Private Sub Form_Activate()
Call Connection.connected
'Call MFGVALUES
con.BeginTrans
Set rs = con.Execute("SELECT BILLNO from BILL")
 If rs.EOF = False Then
    txtbill = Val(rs(0)) + 1
Else
    txtbill = 1
End If
'MFG.TextMatrix(i, 10) = txtGivenFree.Text
'MFG.TextMatrix(i, 11) = MedStockId
End Sub

Private Sub txtamt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdis.SetFocus
End If
End Sub
Private Sub txtdis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    exmdet.SetFocus
End If
End Sub
Private Sub txtdis_LostFocus()
If txtdis.Text = "" Then
 txtdis.Text = 0
End If
txttotamt.Text = Val(txtamt.Text) - Val(txtdis.Text)
End Sub
Private Sub txtdisgvn_LostFocus()
If txtdisgvn.Text = "" Then
    txtdisgvn.Text = 0
End If
txtpayable = Val(txtgrndtot.Text) - Val(txtdisgvn.Text)
End Sub
Private Sub txtqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtrpu.SetFocus
End If
End Sub
Private Sub txtqty_LostFocus()
If txtqty = "" Then
    txtqty.Text = 0
End If
End Sub
Private Sub txtrpu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdis.SetFocus
End If
End Sub
Private Sub txtrpu_LostFocus()
If txtrpu.Text = "" Then
    txtrpu.Text = 0
End If
txtamt.Text = Val(txtqty.Text) * Val(txtrpu.Text)
End Sub
Public Sub MFGVALUES()
MFG.TextMatrix(0, 1) = "MEDICINE NAME"
MFG.TextMatrix(0, 2) = "MEDICINE TYPE"
MFG.TextMatrix(0, 3) = "DATE OF ISSUE"
MFG.TextMatrix(0, 4) = "QUANTITY"
MFG.TextMatrix(0, 5) = "RATE PER UNIT "
MFG.TextMatrix(0, 6) = "AMOUNT"
MFG.TextMatrix(0, 7) = "DISCOUNT"
MFG.TextMatrix(0, 8) = "TOTAL AMOUNT"
End Sub
