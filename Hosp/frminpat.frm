VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frminpat 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patients Add Form"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10470
   Icon            =   "frminpat.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   10470
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   615
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   59
         Text            =   "Date"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   58
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
         Left            =   3360
         TabIndex        =   60
         Top             =   45
         Width           =   4050
      End
   End
   Begin VB.TextBox txtrmrnt 
      Height          =   325
      Left            =   6000
      TabIndex        =   16
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox cmbdname 
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Text            =   " "
      Top             =   2160
      Width           =   2175
   End
   Begin ComCtl2.DTPicker txttrmnt 
      Height          =   300
      Left            =   6000
      TabIndex        =   21
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   38086
   End
   Begin VB.TextBox txtdiag 
      Height          =   735
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "frminpat.frx":08CA
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ListBox lstname 
      Height          =   3765
      Left            =   8760
      TabIndex        =   51
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ComboBox cmbsrgry 
      Height          =   315
      ItemData        =   "frminpat.frx":08D0
      Left            =   6000
      List            =   "frminpat.frx":08DA
      TabIndex        =   20
      Text            =   "NO"
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox cmblbrcse 
      Height          =   315
      ItemData        =   "frminpat.frx":08EC
      Left            =   6000
      List            =   "frminpat.frx":08F9
      TabIndex        =   19
      Text            =   "NO"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.ComboBox cmblbr 
      Height          =   315
      ItemData        =   "frminpat.frx":0915
      Left            =   6000
      List            =   "frminpat.frx":091F
      TabIndex        =   18
      Text            =   "NO"
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ComboBox cmbanes 
      Height          =   315
      ItemData        =   "frminpat.frx":092C
      Left            =   6000
      List            =   "frminpat.frx":0936
      TabIndex        =   17
      Top             =   5400
      Width           =   2175
   End
   Begin VB.ComboBox cmbdcode 
      DataField       =   "DOCCODE"
      DataSource      =   "DocList"
      Height          =   315
      ItemData        =   "frminpat.frx":0949
      Left            =   1680
      List            =   "frminpat.frx":094B
      TabIndex        =   10
      Top             =   7440
      Width           =   735
   End
   Begin VB.ComboBox cmbrtype 
      Height          =   315
      ItemData        =   "frminpat.frx":094D
      Left            =   6000
      List            =   "frminpat.frx":094F
      TabIndex        =   14
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtbdno 
      Height          =   325
      Left            =   6000
      TabIndex        =   15
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox cmbwrd 
      Height          =   315
      ItemData        =   "frminpat.frx":0951
      Left            =   6000
      List            =   "frminpat.frx":095E
      TabIndex        =   13
      Top             =   3480
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc dbsrgry 
      Height          =   330
      Left            =   4080
      Top             =   8880
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "Surgery"
      Caption         =   "Srgry"
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
   Begin MSAdodcLib.Adodc proom 
      Height          =   330
      Left            =   2280
      Top             =   9240
      Width           =   2295
      _ExtentX        =   4048
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
      RecordSource    =   "Room"
      Caption         =   "Patient Room"
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
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ComboBox cmbcast 
      Height          =   315
      ItemData        =   "frminpat.frx":097E
      Left            =   1680
      List            =   "frminpat.frx":098E
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txttel 
      Height          =   325
      Left            =   1680
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.ComboBox cmbmar 
      Height          =   315
      ItemData        =   "frminpat.frx":09B2
      Left            =   1680
      List            =   "frminpat.frx":09C2
      TabIndex        =   9
      Top             =   6960
      Width           =   1695
   End
   Begin VB.ComboBox cmbsex 
      Height          =   315
      ItemData        =   "frminpat.frx":09EC
      Left            =   1680
      List            =   "frminpat.frx":09FC
      TabIndex        =   8
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox txtage 
      Height          =   325
      Left            =   1680
      TabIndex        =   4
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox txtadd 
      Height          =   855
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3120
      Width           =   2415
   End
   Begin VB.ComboBox txtbld 
      Height          =   315
      ItemData        =   "frminpat.frx":0A28
      Left            =   1680
      List            =   "frminpat.frx":0A44
      TabIndex        =   6
      Top             =   5520
      Width           =   855
   End
   Begin MSAdodcLib.Adodc pcinc 
      Height          =   330
      Left            =   2280
      Top             =   8880
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "SELECT PCODE from Ipdet order by PCODE;"
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
      Left            =   10200
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7080
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Height          =   480
      Left            =   10200
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   7080
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
      Left            =   10200
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6480
      Width           =   1440
   End
   Begin VB.TextBox txtcmdnew 
      BackColor       =   &H00E0E0E0&
      Height          =   480
      Left            =   10200
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   6480
      Width           =   1440
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      CausesValidation=   0   'False
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
      Left            =   10200
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Click on the Diagnosis to Enter the Medicinal Detais"
      Top             =   5880
      Width           =   1440
   End
   Begin VB.TextBox txtdmbsve 
      Height          =   480
      Left            =   10200
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   5880
      Width           =   1440
   End
   Begin MSAdodcLib.Adodc DocList 
      Height          =   330
      Left            =   360
      Top             =   8880
      Visible         =   0   'False
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
      RecordSource    =   "DocMast"
      Caption         =   "Doc List"
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   3375
      Begin ComCtl2.DTPicker txtdoa 
         Height          =   300
         Left            =   1560
         TabIndex        =   29
         Top             =   405
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   38086
      End
      Begin VB.TextBox txtcode 
         BackColor       =   &H0080C0FF&
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   28
         Top             =   400
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Admission"
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
         Left            =   1560
         TabIndex        =   27
         Top             =   180
         Width           =   1545
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
         TabIndex        =   26
         Top             =   180
         Width           =   1245
      End
   End
   Begin ComCtl2.DTPicker txtdob 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   38086
   End
   Begin VB.CommandButton cmdupdate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Update"
      CausesValidation=   0   'False
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
      Left            =   8520
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Click on the Diagnosis to Enter the Medicinal Detais"
      Top             =   6600
      Width           =   1440
   End
   Begin VB.TextBox Text2 
      Height          =   480
      Left            =   8520
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   6600
      Width           =   1440
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Rent"
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
      Left            =   4920
      TabIndex        =   56
      Top             =   5040
      Width           =   960
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN PATIENT DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3960
      TabIndex        =   55
      Top             =   1545
      Width           =   2985
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Treatment Done on "
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
      Left            =   4215
      TabIndex        =   53
      Top             =   7560
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosis Done"
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
      Left            =   4575
      TabIndex        =   52
      Top             =   2880
      Width           =   1350
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surgery Type"
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
      Left            =   4785
      TabIndex        =   50
      Top             =   7080
      Width           =   1140
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Case"
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
      Left            =   4845
      TabIndex        =   49
      Top             =   6600
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Room"
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
      Left            =   4785
      TabIndex        =   48
      Top             =   6000
      Width           =   1140
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anesthesia Type"
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
      Left            =   4500
      TabIndex        =   47
      Top             =   5520
      Width           =   1425
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Incharge"
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
      Left            =   4470
      TabIndex        =   46
      Top             =   2280
      Width           =   1395
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
      Left            =   465
      TabIndex        =   45
      Top             =   7560
      Width           =   1080
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Type"
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
      Left            =   4950
      TabIndex        =   44
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bed No"
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
      Left            =   5280
      TabIndex        =   43
      Top             =   4550
      Width           =   645
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ward Name"
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
      Left            =   4920
      TabIndex        =   42
      Top             =   3480
      Width           =   1005
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
      Left            =   1050
      TabIndex        =   41
      Top             =   5160
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
      Left            =   330
      TabIndex        =   40
      Top             =   6120
      Width           =   1215
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
      Left            =   360
      TabIndex        =   39
      Top             =   6960
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
      Left            =   1215
      TabIndex        =   38
      Top             =   6480
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
      Left            =   1200
      TabIndex        =   37
      Top             =   4680
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
      Left            =   450
      TabIndex        =   36
      Top             =   4200
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
      Left            =   855
      TabIndex        =   35
      Top             =   3120
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
      Left            =   255
      TabIndex        =   34
      Top             =   2760
      Width           =   1290
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
      Left            =   480
      TabIndex        =   33
      Top             =   5640
      Width           =   1065
   End
End
Attribute VB_Name = "frminpat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ipcon As ADODB.Connection
Dim iprs As ADODB.Recordset
Dim txt As String
Private Sub cmbanes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmblbr.SetFocus
End If
End Sub
Private Sub cmbcast_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbsex.SetFocus
    End If
End Sub
Private Sub cmbdcode_DropDown()
On Error Resume Next
con.BeginTrans
Set rs = con.Execute("SELECT DOCCODE from DOCTORS")
    Do Until rs.EOF
        cmbdcode.AddItem rs(0)
    rs.MoveNext
Loop
con.CommitTrans
End Sub
Private Sub cmbdcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbwrd.SetFocus
End If
End Sub
Private Sub cmbdcode_LostFocus()
On Error Resume Next
con.BeginTrans
    Set rs = con.Execute("SELECT DOCNAME from DOCTORS where DOCCODE=" & cmbdcode.Text)
     cmbdname.Text = rs(0)
    rs.Close
con.CommitTrans
End Sub
Private Sub cmbdname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbwrd.SetFocus
End If
End Sub
Private Sub cmblbr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmblbrcse.SetFocus
End If
End Sub
Private Sub cmblbrcse_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbsrgry.SetFocus
End If
End Sub
Private Sub cmbmar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbdcode.SetFocus
End If
End Sub
Private Sub cmbrtype_DropDown()
con.BeginTrans
    Set rs = con.Execute("SELECT ROOMTYPE FROM ROOM_MAST")
    Do Until rs.EOF
        cmbrtype.AddItem (rs(0))
            rs.MoveNext
    Loop
End Sub

Private Sub cmbrtype_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbdno.SetFocus
End If
End Sub
Private Sub cmbsex_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbld.SetFocus
End If
End Sub
Private Sub cmbsrgry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdiag.SetFocus
End If
End Sub
Private Sub cmbwrd_DropDown()
con.BeginTrans
Set rs = con.Execute("SELECT WARDNAME FROM ROOM_MAST")
    Do Until rs.EOF
        cmbwrd.AddItem (rs(0))
        rs.MoveNext
    Loop
con.CommitTrans
End Sub
Private Sub cmbwrd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbrtype.SetFocus
End If
End Sub
Private Sub cmdclr_Click()
Call ClearAll
End Sub
Private Sub cmdnew_Click()
On Error Resume Next
    Set ipcon = New ADODB.Connection
    Set iprs = New ADODB.Recordset
        ipcon.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
        iprs.Open "SELECT PCODE FROM INPATIENTS order by pcode", ipcon, adOpenDynamic, adLockOptimistic
         If iprs.EOF = False Then
            iprs.MoveLast
            txtcode.Text = Val(iprs(0)) + 1
    Else
        txtcode.Text = 1
        End If
End Sub
Private Sub cmdsave_Click()
'On Error Resume Next
  Set ipcon = New ADODB.Connection
  Set iprs = New ADODB.Recordset
   ipcon.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
  iprs.Open "SELECT * FROM INPATIENTS", ipcon, adOpenDynamic, adLockOptimistic
    With iprs
        .AddNew
        !pcode = txtcode.Text
        !DOA = txtdoa.Value
        !Name = txtname.Text
        !Address = txtadd.Text
        !DOB = txtdob.Value
        !age = txtage.Text
        !caste = cmbcast.Text
        !SEX = cmbsex.Text
        !BLOODGROUP = txtbld.Text
        !telno = txttel.Text
        !MARITALSTATUS = cmbmar.Text
        !DOCCODE = cmbdcode.Text
        !DOCEXAMINED = cmbdname.Text
        !WARDJOINED = cmbwrd.Text
        !ROOMTYPE = cmbrtype.Text
        !bedno = txtbdno.Text
        !ROOMRENT = txtrmrnt.Text
        !DIAGNOSIS = txtdiag.Text
        !TREATMENTDATE = txttrmnt.Value
        !ANESTYPE = cmbanes.Text
        !LABOURROOM = cmblbr.Text
        !LABOURCASE = cmblbrcse.Text
        !SURGERYTYPE = cmbsrgry.Text
    .Update
End With
iprs.Close
Set ipcon = New ADODB.Connection
Set iprs = New ADODB.Recordset
    ipcon.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
  iprs.Open "SELECT * FROM ROOM", ipcon, adOpenDynamic, adLockOptimistic
    With iprs
        .AddNew
            !pcode = txtcode.Text
            !Name = txtname.Text
            !wardname = cmbwrd.Text
            !RTYPE = cmbrtype.Text
            !bedno = txtbdno.Text
            !rent = txtrmrnt.Text
        .Update
    End With
    txtname.SetFocus
    lstname.AddItem (txtname.Text)
iprs.Close
End Sub
Private Sub cmdUpdate_Click()
On Error Resume Next
con.BeginTrans
    con.Execute ("Update INPATIENTS set DOA = '" & txtdoa.Value & "', NAME ='" & txtname.Text & "',ADDRESS = '" & txtadd.Text & "',DOB='" & txtdob.Value & "',AGE='" & txtage.Text & "' ,CASTE= '" & cmbcast.Text & "',SEX='" & cmbsex.Text & "' ,BLOODGROUP='" & txtbld.Text & "' ,TELNO= '" & txttel.Text & "', MARITALSTATUS= '" & cmbmar.Text & "',DOCCODE= '" & cmbdcode.Text & "',DOCEXAMINED= '" & cmbdname.Text & "' where PCODE =" & txtcode.Text)
    con.CommitTrans
    Call DisplayName
    MsgBox "Details Updated.....", vbInformation + vbOKOnly
    txtname.SetFocus
End Sub
Private Sub Form_Activate()
txtname.SetFocus
Call Connection.connected
Call DisplayName
End Sub

Private Sub lstname_Click()
On Error Resume Next
If lstname.SelCount = 0 Then
    Exit Sub
End If
Set rs = con.Execute("Select * from INPATIENTS where NAME ='" & lstname.Text & "'")
If rs.EOF = True Then
    rs.Close
Else
    txtcode.Text = rs(0)
    txtdoa.Value = rs(1)
    txtname.Text = rs(2)
    txtadd.Text = rs(3)
    txtdob.Value = rs(4)
    txtage.Text = rs(5)
    cmbcast.Text = rs(6)
    cmbsex.Text = rs(7)
    txtbld.Text = rs(8)
    txttel.Text = rs(9)
    cmbmar.Text = rs(10)
    cmbdcode.Text = rs(11)
    cmbdname.Text = rs(12)
    cmbrtype.Text = rs(13)
    txtbdno.Text = rs(14)
    txtdiag.Text = rs(15)
    'txttrmnt.Value = rs(16)
    'cmbanes.Text = rs(17)
    cmblbr.Text = rs(18)
    cmblbrcse.Text = rs(19)
    cmbsrgry.Text = rs(20)
    rs.Close
End If

End Sub

Private Sub txtadd_KeyPress(KeyAscii As Integer)
'Dim str As String
'str = ".*:,&()[]{}\abcdefghijklmnopqrstuvwxyz ABCDEFGHIJKLMNOPQRSTUVWXYZ-0123456789/# "
'If KeyAscii <> 13 Then
'    If InStr(str, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
'        KeyAscii = 0
'    End If
'End If
End Sub
Private Sub txtage_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbcast.SetFocus
End If
End Sub

Private Sub txtbdno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbanes.SetFocus
End If
End Sub
Private Sub txtbdno_LostFocus()
con.BeginTrans
    Set rs = con.Execute("SELECT BEDNO from ROOM")
     If rs(0) = txtbdno Then
        MsgBox "Sorry !!!! Bed is Occupied", , "Accepting Bed NO"
        txtbdno.Text = ""
        txtbdno.SetFocus
     End If
End Sub
Private Sub txtbld_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttel.SetFocus
End If
End Sub
Private Sub txtdiag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttrmnt.SetFocus
End If
End Sub
Private Sub txtdob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtage.SetFocus
End If
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtadd.SetFocus
End If
End Sub

Private Sub txttel_KeyPress(KeyAscii As Integer)
Dim str As String
str = "-0123456789NA"
If KeyAscii = 13 Then
    cmbmar.SetFocus
End If
If InStr(str, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
    KeyAscii = 0
End If

End Sub
Private Sub txttrmnt_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
If KeyAscii = 13 Then
    cmdSave.SetFocus
End If
End Sub
Public Sub DisplayName()
On Error Resume Next
Set rs = con.Execute("Select name from INPATIENTS")
If rs.EOF = True Then
    lstname.Clear
    rs.Close
Else
    lstname.Clear
    Do While rs.EOF = False
        lstname.AddItem (rs(0))
        rs.MoveNext
    Loop
    rs.Close
End If
End Sub
Public Sub ClearAll()
txtcode.Text = ""
txtname.Text = ""
txtadd.Text = ""
txtage.Text = ""
txttel.Text = ""
cmbdcode.Text = ""
cmbdname.Text = ""
cmbwrd.Text = ""
cmbrtype.Text = ""
txtbdno.Text = ""
txtdiag.Text = ""
txtrmrnt.Text = ""
cmbanes.Text = ""
cmblbr.Text = ""
cmblbrcse.Text = ""
cmbsrgry.Text = ""
End Sub
