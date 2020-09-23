VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmopdet 
   Caption         =   "Add Out Patient Details"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   855
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   39
         Text            =   "Date"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   38
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
         Left            =   2400
         TabIndex        =   40
         Top             =   165
         Width           =   4050
      End
   End
   Begin VB.TextBox cmbdname 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3000
      TabIndex        =   36
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmddiag 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Diagnosis"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5640
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   1440
   End
   Begin MSAdodcLib.Adodc pcinc 
      Height          =   330
      Left            =   3360
      Top             =   9000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Mlk Hos\MlkHos.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Mlk Hos\MlkHos.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT PCODE from Opdet order by PCODE;"
      Caption         =   "pcinc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1200
      Top             =   9000
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Mlk Hos\MlkHos.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Mlk Hos\MlkHos.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Opdet"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Click on the Diagnosis to Enter the Medicinal Detais"
      Top             =   6360
      Width           =   1440
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3840
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   1440
   End
   Begin VB.CommandButton cmdclr 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Clear All"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2160
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   1440
   End
   Begin VB.TextBox txtadd 
      BackColor       =   &H0080C0FF&
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
      Left            =   4440
      TabIndex        =   3
      Top             =   2520
      Width           =   3735
   End
   Begin ComCtl2.DTPicker dttrmnt 
      Height          =   300
      Left            =   5280
      TabIndex        =   12
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   38076
   End
   Begin VB.ComboBox cmbdcode 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox txtbld 
      Height          =   315
      ItemData        =   "frmopdet.frx":0000
      Left            =   240
      List            =   "frmopdet.frx":001C
      TabIndex        =   10
      Top             =   4320
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9120
      Top             =   360
   End
   Begin VB.TextBox txtdiag 
      Height          =   975
      Left            =   6720
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Text            =   "frmopdet.frx":0042
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   27
      Top             =   1320
      Width           =   1575
      Begin VB.TextBox txtcode 
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Text            =   " "
         Top             =   400
         Width           =   1335
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
         TabIndex        =   28
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.TextBox txttelno 
      Height          =   300
      Left            =   8160
      TabIndex        =   9
      Top             =   3375
      Width           =   1575
   End
   Begin VB.ComboBox cmbcaste 
      Height          =   315
      ItemData        =   "frmopdet.frx":0048
      Left            =   6120
      List            =   "frmopdet.frx":0058
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ComboBox cmbmarstat 
      Height          =   315
      ItemData        =   "frmopdet.frx":007C
      Left            =   4320
      List            =   "frmopdet.frx":008F
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ComboBox cmbgndr 
      Height          =   315
      ItemData        =   "frmopdet.frx":00C1
      Left            =   2640
      List            =   "frmopdet.frx":00D1
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H0080C0FF&
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
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox txtage 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   3360
      Width           =   615
   End
   Begin ComCtl2.DTPicker dtdob 
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24576003
      CurrentDate     =   38039
   End
   Begin VB.TextBox txtdmbsve 
      Height          =   480
      Left            =   360
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   6360
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtcmdnew 
      BackColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   3840
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   6360
      Width           =   1440
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   5640
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   6360
      Width           =   1440
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Treatment"
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
      Left            =   5040
      TabIndex        =   31
      Top             =   4080
      Width           =   1560
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
      Left            =   240
      TabIndex        =   30
      Top             =   4080
      Width           =   1065
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
      Left            =   6720
      TabIndex        =   29
      Top             =   3840
      Width           =   840
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
      Left            =   1530
      TabIndex        =   26
      Top             =   4065
      Width           =   1080
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
      Left            =   3000
      TabIndex        =   25
      Top             =   4065
      Width           =   1125
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
      Left            =   8160
      TabIndex        =   24
      Top             =   3120
      Width           =   1215
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
      Left            =   6120
      TabIndex        =   23
      Top             =   3120
      Width           =   495
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
      Left            =   4320
      TabIndex        =   22
      Top             =   3120
      Width           =   1185
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
      Left            =   2640
      TabIndex        =   21
      Top             =   3120
      Width           =   330
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
      Left            =   1860
      TabIndex        =   20
      Top             =   3120
      Width           =   345
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
      Left            =   315
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
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
      Left            =   4680
      TabIndex        =   18
      Top             =   2280
      Width           =   690
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
      Left            =   315
      TabIndex        =   0
      Top             =   2280
      Width           =   1290
   End
End
Attribute VB_Name = "frmopdet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ipcon As ADODB.Connection
Dim iprs As ADODB.Recordset
Private Sub DTPicker1_Change()
'Text6.Text = CalcAge(DTPicker1.Value)
End Sub
Private Sub cmbdcode_DropDown()
 Set rs = con.Execute("SELECT DOCCODE from DOCTORS")
    Do Until rs.EOF
        cmbdcode.AddItem (rs(0))
    rs.MoveNext
Loop
End Sub
Private Sub cmbregist_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtname.SetFocus
End If
End Sub
Private Sub cmbregist_LostFocus()
If cmbregist.ListIndex = 1 Then
    frmcom.Visible = True
Else
    frmcom.Visible = False
End If
End Sub
Private Sub cmbdcode_LostFocus()
Set rs = con.Execute("SELECT DOCNAME from DOCTORS where DOCCODE=" & cmbdcode.Text)
    cmbdname.Text = rs(0)
End Sub
Private Sub cmdclr_Click()
txtcode.Text = ""
txtname.Text = ""
End Sub
Private Sub cmddiag_Click()
OPDiag.Show
End Sub
Private Sub cmdnew_Click()
    Set ipcon = New ADODB.Connection
    Set iprs = New ADODB.Recordset
        ipcon.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
        iprs.Open "SELECT PCODE FROM OPDET order by pcode", ipcon, adOpenDynamic, adLockOptimistic
         If iprs.EOF = False Then
            iprs.MoveLast
            txtcode.Text = Val(iprs(0)) + 1
    Else
        txtcode.Text = 1
        End If
End Sub
Private Sub cmdsave_Click()
On Error Resume Next
    Set ipcon = New ADODB.Connection
    Set iprs = New ADODB.Recordset
        ipcon.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
        iprs.Open "SELECT * FROM OPDET", ipcon, adOpenDynamic, adLockOptimistic
With iprs
    .AddNew
        !pcode = txtcode.Text
        !pname = txtname.Text
        !Address = txtadd.Text
        !bdate = dtdob.Value
        !caste = cmbcaste.Text
        !gender = cmbgndr.Text
        !age = txtage.Text
        !telno = txttelno.Text
        !marstat = cmbmarstat.Text
        !bgrp = txtbld.Text
        !DOCCODE = cmbdcode.Text
        !docname = cmbdname.Text
        !DIAGNOSIS = txtdiag.Text
        !tretmntdt = dttrmnt.Value
       .Update
End With
'Call ClearAll
'pcinc.Refresh
'pcinc.Recordset.MoveLast
con.CommitTrans
End Sub
Private Sub dtdob_Change()
'txtage.Text = CalcAge(dtdob.Value)
End Sub
Private Sub dtdob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbgndr.SetFocus
End If
End Sub
Private Sub Form_Activate()
Call Connection.connected
End Sub
Private Sub Form_Load()
frmopdet.WindowState = 2
txtdt.Text = Format$(Now, "DD-MM-YYYY") & Space(2)
txttme.Text = Format$(Now, "hh:mm:ss AM/PM") & Space(2)
End Sub
Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
'PopupMenu MDIForm1.main
End If
End Sub
Private Sub Timer1_Timer()
txttme.Text = Format$(Now, "hh:mm:ss AM/PM") & Space(2)
End Sub

Private Sub txtadd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    dtdob.SetFocus
End If
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbgndr.SetFocus
End If
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtadd.SetFocus
End If
End Sub
