VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmipmed 
   BackColor       =   &H00FF8080&
   Caption         =   "Patient Medical bill"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   8715
   Begin VB.CommandButton Command6 
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
      Left            =   5040
      TabIndex        =   45
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtpayable 
      Height          =   375
      Left            =   9960
      TabIndex        =   41
      Top             =   7245
      Width           =   1455
   End
   Begin VB.TextBox txtdisgvn 
      Height          =   375
      Left            =   6120
      TabIndex        =   40
      Top             =   7245
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
      Height          =   375
      Left            =   3240
      TabIndex        =   39
      Top             =   7245
      Width           =   1095
   End
   Begin VB.ComboBox cmbcrdt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmipmeds.frx":0000
      Left            =   10800
      List            =   "frmipmeds.frx":000A
      TabIndex        =   37
      Text            =   "Credit"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BILL"
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
      Left            =   3720
      TabIndex        =   36
      Top             =   7920
      Width           =   975
   End
   Begin ComCtl2.DTPicker billdt 
      Height          =   285
      Left            =   8400
      TabIndex        =   22
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24707073
      CurrentDate     =   38087
   End
   Begin VB.TextBox txtbill 
      Height          =   285
      Left            =   6720
      TabIndex        =   21
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add to the List"
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
      TabIndex        =   15
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Medicine Details"
      Height          =   2295
      Left            =   4320
      TabIndex        =   20
      Top             =   1320
      Width           =   7455
      Begin VB.TextBox txtmedname 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   200
         Width           =   3135
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Text            =   "Combo3"
         Top             =   200
         Width           =   855
      End
      Begin VB.ComboBox cmbmedtype 
         Height          =   315
         ItemData        =   "frmipmeds.frx":001C
         Left            =   1680
         List            =   "frmipmeds.frx":0032
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txttotamt 
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtdis 
         Height          =   285
         Left            =   6360
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtamt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtrpu 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   6600
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin ComCtl2.DTPicker dtissu 
         Height          =   300
         Left            =   4320
         TabIndex        =   9
         Top             =   840
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24707073
         CurrentDate     =   38087
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Type"
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
         TabIndex        =   35
         Top             =   850
         Width           =   1560
      End
      Begin VB.Label Label13 
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
         Left            =   4200
         TabIndex        =   31
         Top             =   1920
         Width           =   1380
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Given"
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
         Left            =   4680
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Left            =   2760
         TabIndex        =   29
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Per Unit"
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
         Left            =   225
         TabIndex        =   28
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date"
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
         Left            =   3120
         TabIndex        =   26
         Top             =   850
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Name"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   250
         Width           =   1635
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine ID"
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
         TabIndex        =   24
         Top             =   250
         Width           =   1245
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   1545
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   2025
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Width           =   1545
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2280
      TabIndex        =   2
      Top             =   1545
      Width           =   2025
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      MergeCells      =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount Payable"
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
      Left            =   7800
      TabIndex        =   44
      Top             =   7320
      Width           =   2130
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Given"
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
      Left            =   4440
      TabIndex        =   43
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label19 
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
      Left            =   1920
      TabIndex        =   42
      Top             =   7320
      Width           =   1245
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Terms"
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
      Left            =   9600
      TabIndex        =   38
      Top             =   615
      Width           =   1140
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN PATIENTS MEDICINE ISSUE "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   3765
      TabIndex        =   34
      Top             =   -15
      Width           =   4530
   End
   Begin VB.Label Label15 
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
      Left            =   7440
      TabIndex        =   33
      Top             =   615
      Width           =   900
   End
   Begin VB.Label Label14 
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
      Left            =   5880
      TabIndex        =   32
      Top             =   615
      Width           =   705
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
      Left            =   120
      TabIndex        =   23
      Top             =   600
      Width           =   1575
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
      Left            =   165
      TabIndex        =   19
      Top             =   2760
      Width           =   1905
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
      Left            =   480
      TabIndex        =   18
      Top             =   2160
      Width           =   1590
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
      Left            =   -150
      TabIndex        =   17
      Top             =   3360
      Width           =   2220
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
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1950
   End
End
Attribute VB_Name = "frmipmed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim meddb, ipcon As ADODB.Connection
Dim medrs, iprs As ADODB.Recordset
Dim ID, bllinc As Integer
Dim txt1, txt As String
Private Sub cmbmedtype_Change()
If KeyAscii = 13 Then
    dtissu.SetFocus
End If
End Sub
Private Sub Combo1_DropDown()
On Error Resume Next
ID = 0
con.BeginTrans
    Set rs = con.Execute("SELECT PCODE FROM INPATIENTS")
    Do Until rs.EOF
        Combo1.AddItem rs(0)
      rs.MoveNext
    Loop
    MFG.Clear
    i = MFG.Rows - 1
    MFG.TextMatrix(i, 1) = ""
    MFG.TextMatrix(i, 2) = ""
    MFG.TextMatrix(i, 3) = ""
    MFG.TextMatrix(i, 4) = ""
    MFG.TextMatrix(i, 5) = ""
    MFG.TextMatrix(i, 6) = ""
    MFG.TextMatrix(i, 7) = ""
    MFG.TextMatrix(i, 8) = ""
rs.Close
Call MFGVALUES
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
Set rs = con.Execute("SELECT * from IPMEDICINE where pcode =" & Combo1.Text)
MFG.Rows = 2
If rs.EOF = True Then
    MFG.Clear
    Call MFGVALUES
End If
Do Until rs.EOF
    MFG.Rows = MFG.Rows + 1
    j = MFG.Rows
    i = MFG.Rows - 2
    MFG.TextMatrix(i, 1) = rs(2)
    MFG.TextMatrix(i, 2) = rs(3)
    MFG.TextMatrix(i, 3) = rs(4)
    MFG.TextMatrix(i, 4) = rs(5)
    MFG.TextMatrix(i, 5) = rs(6)
    MFG.TextMatrix(i, 6) = rs(7)
    MFG.TextMatrix(i, 7) = rs(8)
    MFG.TextMatrix(i, 8) = rs(9)
rs.MoveNext
Loop
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
Private Sub Combo3_DropDown()
On Error Resume Next
con.BeginTrans
 Set rs = con.Execute("SELECT MEDID,MEDICINENAME from MEDICINE")
    Do Until rs.EOF
      Combo3.AddItem (rs(0))
      rs.MoveNext
    Loop
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbmedtype.SetFocus
End If
End Sub
Private Sub Combo3_LostFocus()
On Error Resume Next
con.BeginTrans
    Set rs = con.Execute("SELECT MEDICINENAME from MEDICINE where medid=" & Combo3.Text)
     txtmedname.Text = rs(0)
    rs.Close
con.CommitTrans
'ID = Val(Combo3.Text)
End Sub
Private Sub Command1_Click()
On Error Resume Next
Set meddb = New ADODB.Connection
Set medrs = New ADODB.Recordset
meddb.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Hospital.mdb"
medrs.Open "SELECT * FROM IPMEDICINE", meddb, adOpenDynamic, adLockOptimistic
        With medrs
            .AddNew
                !PCODE = Combo1.Text
                !Name = Text2.Text
                !MEDICINENAME = txtmedname.Text
                !MEDICINETYPE = cmbmedtype.Text
                !DATEOFISSUE = dtissu.Value
                !QUANTITY = txtqty.Text
                !RATEPERUNIT = txtrpu.Text
                !AMOUNT = txtamt.Text
                !TOTALAMOUNT = txttotamt.Text
            .Update
        End With
If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        If MFG.TextMatrix(i, 1) = txtmedname.Text Then
            MsgBox "Medicine Already Exist In The List Cannot Add Same Medicine Again.....", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next i
End If
ID = ID + Val(txttotamt.Text) + Val(txtgrndtot.Text)
MFG.Rows = MFG.Rows + 1
j = MFG.Rows
i = MFG.Rows - 2
'MFG.TextMatrix(i, 0) = i
MFG.TextMatrix(i, 1) = txtmedname.Text
MFG.TextMatrix(i, 2) = cmbmedtype.Text
MFG.TextMatrix(i, 3) = dtissu.Value
MFG.TextMatrix(i, 4) = txtqty.Text
MFG.TextMatrix(i, 5) = txtrpu.Text
MFG.TextMatrix(i, 6) = txtamt.Text
MFG.TextMatrix(i, 7) = txtdis.Text
MFG.TextMatrix(i, 8) = txttotamt.Text
'MFG.TextMatrix(i, 9) = txtTotalPrice.Text
'MFG.TextMatrix(i, 10) = txtGivenFree.Text
'MFG.TextMatrix(i, 11) = MedStockId
''Set rs = con.Execute("Select ExpDate from MedicineStock where MedicineStockId=" & MedStockId)
''MFG.TextMatrix(i, 12) = Format(rs(0), "dd-MMM-yyyy")
''Call Txt_Clear
''Call Calc_TotAmt
''cmbMedicine.SetFocus
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
Private Sub Form_Activate()
Call Connection.connected
Call MFGVALUES
    Set ipcon = New ADODB.Connection
    Set iprs = New ADODB.Recordset
        ipcon.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
        iprs.Open "SELECT BILLNO FROM BILL", ipcon, adOpenDynamic, adLockOptimistic
         If iprs.EOF = False Then
            iprs.MoveLast
            txtbill.Text = Val(iprs(0)) + 1
    Else
        txtbill.Text = 1
        End If
'MFG.TextMatrix(i, 10) = txtGivenFree.Text
'MFG.TextMatrix(i, 11) = MedStockId
End Sub
Private Sub txtdis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
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
