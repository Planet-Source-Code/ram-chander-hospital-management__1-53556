VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOPBillPayments 
   BackColor       =   &H00FF8080&
   Caption         =   "Out Patient Bill Payments"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   ForeColor       =   &H00C0C0FF&
   Icon            =   "frmOPBillPayments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin ComCtl2.DTPicker dtpDDDate 
      Height          =   315
      Left            =   8040
      TabIndex        =   49
      Top             =   6720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   37985
   End
   Begin ComCtl2.DTPicker dtpPayDate 
      Height          =   315
      Left            =   5160
      TabIndex        =   48
      Top             =   5640
      Width           =   1900
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   37985
   End
   Begin VB.TextBox txtBalAdv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5130
      Locked          =   -1  'True
      TabIndex        =   46
      ToolTipText     =   "Customer Balance Advance"
      Top             =   6450
      Width           =   1905
   End
   Begin VB.TextBox txtCustomerAdv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   44
      ToolTipText     =   "Customer Advance Amount"
      Top             =   6060
      Width           =   1845
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      TabIndex        =   19
      ToolTipText     =   "Click To Close"
      Top             =   7110
      Width           =   1545
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2100
      TabIndex        =   17
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   7110
      Width           =   1545
   End
   Begin VB.TextBox txtBillStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   42
      ToolTipText     =   "Customer Bill Status"
      Top             =   6450
      Width           =   1845
   End
   Begin VB.TextBox txtBalAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5130
      Locked          =   -1  'True
      TabIndex        =   40
      ToolTipText     =   "Bill Balance Amount"
      Top             =   6060
      Width           =   1905
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C96C59&
      Caption         =   "Payment Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   7170
      TabIndex        =   36
      Top             =   5490
      Width           =   4545
      Begin VB.OptionButton optCash 
         Appearance      =   0  'Flat
         BackColor       =   &H00C96C59&
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   300
         TabIndex        =   10
         ToolTipText     =   "Click Here If Payment Is Cash"
         Top             =   420
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton optDD 
         Appearance      =   0  'Flat
         BackColor       =   &H00C96C59&
         Caption         =   "DD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   11
         ToolTipText     =   "Click Here If Payment By DD"
         Top             =   420
         Width           =   585
      End
      Begin VB.OptionButton optCheque 
         Appearance      =   0  'Flat
         BackColor       =   &H00C96C59&
         Caption         =   "CHEQUE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2100
         TabIndex        =   12
         ToolTipText     =   "Click Here If Payment By Cheque"
         Top             =   420
         Width           =   1095
      End
      Begin VB.OptionButton optOthers 
         Appearance      =   0  'Flat
         BackColor       =   &H00C96C59&
         Caption         =   "OTHERS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3270
         TabIndex        =   13
         ToolTipText     =   "Click Here If Payment By Others"
         Top             =   420
         Width           =   1155
      End
      Begin VB.ComboBox cmbBank 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "frmOPBillPayments.frx":0442
         Left            =   870
         List            =   "frmOPBillPayments.frx":045E
         TabIndex        =   15
         Text            =   "State Bank Of India"
         ToolTipText     =   "Select The Bank Name"
         Top             =   1650
         Width           =   3615
      End
      Begin VB.TextBox txtDDNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   870
         TabIndex        =   14
         ToolTipText     =   "Enter the DD Number"
         Top             =   840
         Width           =   3645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         Caption         =   "BANK :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   210
         TabIndex        =   39
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         Caption         =   "DATE :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   210
         TabIndex        =   38
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         Caption         =   "DD No :"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   37
         Top             =   870
         Width           =   705
      End
   End
   Begin VB.TextBox txtPayingAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1710
      TabIndex        =   9
      ToolTipText     =   "Enter The Paying Amount"
      Top             =   5670
      Width           =   1845
   End
   Begin VB.TextBox txtBal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Bill Balance Amount"
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txtPaidAmt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3780
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Total Amount Paid"
      Top             =   4920
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1965
      Left            =   210
      TabIndex        =   6
      ToolTipText     =   "Bill Payments List"
      Top             =   2850
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3466
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColor       =   128
      ForeColorFixed  =   8388608
      GridColor       =   13200473
      GridColorFixed  =   13200473
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtBillTerms 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Bill Terms"
      Top             =   1620
      Width           =   2415
   End
   Begin VB.TextBox txtBillDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Bill Date"
      Top             =   1620
      Width           =   2415
   End
   Begin VB.TextBox txtBillItems 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Bill Total Items"
      Top             =   1620
      Width           =   2415
   End
   Begin VB.TextBox txtBillAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Bill Total Amount"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtNetValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "7"
      ToolTipText     =   "Bill Net Value"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Bill Discount"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.ComboBox cmbCustName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Pateint Code"
      ToolTipText     =   "Select The Patient Name"
      Top             =   540
      Width           =   3405
   End
   Begin VB.ComboBox cmbBillNo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   8220
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select Bill Number"
      Top             =   540
      Width           =   3345
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Adv :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3765
      TabIndex        =   47
      Top             =   6510
      Width           =   1320
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Adv :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   435
      TabIndex        =   45
      Top             =   6120
      Width           =   1245
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Status :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   585
      TabIndex        =   43
      Top             =   6510
      Width           =   1095
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Amt :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3750
      TabIndex        =   41
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   35
      Top             =   5730
      Width           =   1485
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Paying Amount :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   34
      Top             =   5730
      Width           =   1590
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   11910
      Y1              =   5385
      Y2              =   5370
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6330
      TabIndex        =   33
      Top             =   5010
      Width           =   885
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Paid :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1740
      TabIndex        =   32
      Top             =   5010
      Width           =   1905
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "BILL PAYMENT DETAILS :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDAF3C&
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   2580
      Width           =   2385
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "BILL INFORMATION :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Terms :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4110
      TabIndex        =   29
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   450
      TabIndex        =   28
      Top             =   1680
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Items :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8010
      TabIndex        =   27
      Top             =   1710
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   945
      Left            =   150
      Shape           =   4  'Rounded Rectangle
      Top             =   1530
      Width           =   11565
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amt :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   510
      TabIndex        =   26
      Top             =   2100
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Value :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8010
      TabIndex        =   25
      Top             =   2100
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4245
      TabIndex        =   23
      Top             =   2100
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      Caption         =   "COMPANY :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   4305
      TabIndex        =   21
      Top             =   630
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   30
      X2              =   11970
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   30
      X2              =   11940
      Y1              =   1005
      Y2              =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Bill Number :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6210
      TabIndex        =   20
      Top             =   630
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Patient Code :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   2010
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "OUT PATIENT BILL PAYMENTS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11985
   End
End
Attribute VB_Name = "frmOPBillPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RowNo As Integer
Private Sub cmbBillNo_Click()
Dim rs1 As New ADODB.Recordset
Dim i As Integer
Dim s As Double
If cmbBillNo.Text = "" Then
    Exit Sub
End If
Set rs = con.Execute("Select * from OPBill where BillId=" & cmbBillNo.ItemData(cmbBillNo.ListIndex))
If rs.EOF = True Then
    rs.Close
Else
    txtBillDate.Text = Format(rs!BillDate, "dd-MMM-yyyy")
    txtBillTerms.Text = rs!CreditYN
    txtBillAmt.Text = rs!GrandTotal
    txtDiscount.Text = rs!Discount
    txtNetValue.Text = rs!NetValue
    Set rs1 = con.Execute("Select count(*) from OPBillDetails where BillId=" & cmbBillNo.ItemData(cmbBillNo.ListIndex))
    If rs1.EOF = True Then
        rs1.Close
    Else
        txtBillItems.Text = rs1(0)
        rs1.Close
    End If
    rs.Close
End If
Set rs = con.Execute("Select * from OPBillPayments where BillId=" & cmbBillNo.ItemData(cmbBillNo.ListIndex))
If rs.EOF = True Then
    rs.Close
    txtPaidAmt.Text = "0"
    txtBal.Text = txtBillAmt.Text
Else
    i = 1
    s = 0
    MFG.Rows = 2
    Do While rs.EOF = False
        MFG.TextMatrix(i, 0) = i
        MFG.TextMatrix(i, 1) = rs!AmountPaid
        MFG.TextMatrix(i, 2) = Format(rs!PaidDate, "dd-MMM-yyyy")
        MFG.TextMatrix(i, 3) = rs!PayType
        If IsNull(rs!DDNo) = False Then
            MFG.TextMatrix(i, 4) = rs!DDNo
        End If
        If IsNull(rs!DDDate) = False Then
            MFG.TextMatrix(i, 5) = rs!DDDate
        End If
        If IsNull(rs!Bank) = False Then
            MFG.TextMatrix(i, 6) = rs!Bank
        End If
        s = s + Val(rs!AmountPaid)
        rs.MoveNext
        i = i + 1
        MFG.Rows = MFG.Rows + 1
    Loop
    rs.Close
    txtPaidAmt.Text = s
    txtBal.Text = Round(Val(txtBillAmt.Text) - Val(txtPaidAmt.Text), 2)
End If
RowNo = MFG.Rows - 1
Set rs = con.Execute("Select CustomerAmt from OPCustomerAdvance where pcode =" & cmbCustName.Text)
If rs.EOF = True Then
    rs.Close
    txtCustomerAdv.Text = "0"
Else
    txtCustomerAdv.Text = rs(0)
    rs.Close
End If
txtPayingAmt.SetFocus
End Sub
Private Sub cmbCustName_Click()
Dim i As Integer
i = 0
If cmbCustName.Text = "" Then
    Exit Sub
End If
cmbBillNo.Clear
Set rs = con.Execute("Select BillNo,BillId from OPBill where pcode =" & cmbCustName.Text)
If rs.EOF = True Then
    rs.Close
Else
    Do While rs.EOF = False
        cmbBillNo.AddItem (rs(0))
        cmbBillNo.ItemData(i) = rs(1)
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
End If
End Sub
Private Sub cmbCustName_DropDown()
Set rs = con.Execute("SELECT PCODE from OPDET")
End Sub
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdsave_Click()
Dim str As String
Dim BillPayId As Double
If txtPayingAmt.Text = "" Then
    MsgBox "Paying Amount Not Found...", vbCritical + vbOKOnly
    txtPayingAmt.SetFocus
    Exit Sub
End If
Set rs = con.Execute("Select Max(BillPaymentId) from OPBillPayments")
If IsNull(rs(0)) = True Then
    BillPayId = 0
Else
    BillPayId = rs(0) + 1
    rs.Close
End If
If optCash.Value = True Then
    str = "CASH"
ElseIf optDD.Value = True Then
    str = "DD"
ElseIf optCheque.Value = True Then
    str = "Cheque"
Else
    str = "Others"
End If
If MsgBox("Confirm To Save Bill Information ?", vbQuestion + vbYesNo) = vbYes Then
    con.BeginTrans
    If optCash.Value = True Then
        con.Execute ("Insert into OPBillPayments values(" & BillPayId & "," & cmbBillNo.ItemData(cmbBillNo.ListIndex) & "," & Val(txtPayingAmt.Text) + (Val(txtCustomerAdv.Text) - Val(txtBalAdv.Text)) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str & "',Null,Null,Null)")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text) + (Val(txtCustomerAdv.Text) - Val(txtBalAdv.Text))
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
    Else
        If txtDDNo.Text = "" Or cmbBank.Text = "" Then
            MsgBox "DD Number or Bank Name Not Found...", vbCritical + vbOKOnly
            txtDDNo.SetFocus
            Exit Sub
        End If
        con.Execute ("Insert into OPBillPayments values(" & BillPayId & "," & cmbBillNo.ItemData(cmbBillNo.ListIndex) & "," & Val(txtPayingAmt.Text) + (Val(txtCustomerAdv.Text) - Val(txtBalAdv.Text)) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str & "','" & txtDDNo.Text & "','" & Format(dtpDDDate.Value, "mm/dd/yy") & "','" & cmbBank.Text & "')")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text) + (Val(txtCustomerAdv.Text) - Val(txtBalAdv.Text))
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str
        MFG.TextMatrix(RowNo, 4) = txtDDNo.Text
        MFG.TextMatrix(RowNo, 5) = Format(dtpDDDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 6) = cmbBank.Text
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
    End If
    con.Execute ("Update OPCustomerAdvance Set CustomerAmt=" & Val(txtBalAdv.Text) & " where pcode =" & cmbCustName.Text)
    con.CommitTrans
    Call Txt_Clear
    txtPayingAmt.SetFocus
End If
End Sub
Private Sub Txt_Clear()
Dim i As Integer
Dim s As Double
s = 0
txtPayingAmt.Text = ""
txtBalAmt.Text = ""
txtBalAdv.Text = ""
txtBillStatus.Text = ""
txtDDNo.Text = ""
Set rs = con.Execute("Select CustomerAmt from OPCustomerAdvance where pcode =" & cmbCustName.ItemData(cmbCustName.ListIndex))
If rs.EOF = True Then
    rs.Close
    txtCustomerAdv.Text = "0"
Else
    txtCustomerAdv.Text = rs(0)
    rs.Close
End If
For i = 1 To MFG.Rows - 2 Step 1
    s = s + MFG.TextMatrix(i, 1)
Next i
txtPaidAmt.Text = s
txtBalAmt.Text = Round(Val(txtBillAmt.Text) - Val(txtPaidAmt.Text), 2)
txtBal.Text = Round(Val(txtBillAmt.Text) - Val(txtPaidAmt.Text), 2)
End Sub

Private Sub Form_Load()
If con.State Then con.Close
Call Connection.connected
Call Refresh_Data
dtpDDDate.Value = Now
dtpPayDate.Value = Now
'txtDDNo.Locked = True
'cmbBank.Locked = True
dtpDDDate.Enabled = False
End Sub
Private Sub Refresh_Data()
Dim i As Integer
i = 0
cmbCustName.Clear
Set rs = con.Execute("Select PName,PCODE from OPDET where pcode in (Select Distinct pcode from OPBill)")
If rs.EOF = True Then
    rs.Close
Else
    Do While rs.EOF = False
        cmbCustName.AddItem (rs(1))
        cmbCustName.ItemData(i) = rs(1)
        rs.MoveNext
        i = i + 1
    Loop
    rs.Close
End If
MFG.Clear
MFG.ColWidth(0) = 1000
MFG.ColAlignment(0) = 4
For i = 1 To 6 Step 1
    MFG.ColWidth(i) = 2000
    MFG.ColAlignment(i) = 4
Next i
MFG.TextMatrix(0, 0) = "SL NO"
MFG.TextMatrix(0, 1) = "AMOUNT PAID"
MFG.TextMatrix(0, 2) = "PAID DATE"
MFG.TextMatrix(0, 3) = "PAY TYPE"
MFG.TextMatrix(0, 4) = "DD/CHEQUE NO"
MFG.TextMatrix(0, 5) = "DD DATE"
MFG.TextMatrix(0, 6) = "BANK"
End Sub
Private Sub optCash_Click()
If optCash.Value = True Then
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
Else
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
End If
End Sub
Private Sub optCheque_Click()
If optCheque.Value = True Then
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
Else
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
End If
End Sub
Private Sub optDD_Click()
If optDD.Value = True Then
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
Else
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
End If
End Sub
Private Sub optOthers_Click()
If optOthers.Value = True Then
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
Else
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
End If
End Sub

Private Sub txtDDNo_KeyPress(KeyAscii As Integer)
Dim str As String
str = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-/abcdefghijklmnopqrstuvwxyz"
If KeyAscii = 13 And txtPayingAmt.Text <> "" Then
    
End If
If InStr(str, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
Private Sub txtPayingAmt_KeyPress(KeyAscii As Integer)
Dim str As String
str = "0123456789."
If KeyAscii = 13 And txtPayingAmt.Text <> "" Then
    cmdSave.SetFocus
End If
If InStr(str, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub
Private Sub txtPayingAmt_LostFocus()
If txtPayingAmt.Text <> "" Then
    If Val(txtPayingAmt.Text) = 0 Then
        MsgBox "Paying Amount Cannot Be Zero...", vbInformation + vbOKOnly
        txtPayingAmt.SetFocus
        Exit Sub
    End If
    If Val(txtPayingAmt.Text) > Val(txtBal.Text) Then
        MsgBox "Paying Amount Cannot Be Greater Than Balance Amount...", vbCritical + vbOKOnly
        txtPayingAmt.Text = ""
        txtPayingAmt.SetFocus
        Exit Sub
    End If
     txtBalAmt.Text = Round((Val(txtBal.Text) - Val(txtPayingAmt.Text)), 2)
     If Val(txtCustomerAdv.Text) <> 0 Then
        If Val(txtCustomerAdv.Text) > Val(txtBalAmt.Text) Then
            'txtCustomerAdv.Text = Val(txtCustomerAdv.Text) - Val(txtBalAmt.Text)
            txtBalAdv.Text = Val(txtCustomerAdv.Text) - Val(txtBalAmt.Text)
            txtBalAmt.Text = "0"
        Else
            txtBalAmt.Text = Round(Val(txtBalAmt.Text) - Val(txtCustomerAdv.Text), 2)
            txtCustomerAdv.Text = "0"
            txtBalAdv.Text = "0"
        End If
    Else
        txtBalAdv.Text = "0"
    End If
    If Val(txtBalAmt.Text) = 0 Then
    txtBillStatus.Text = "Paid"
        Else
    txtBillStatus.Text = "Un-Paid"
    End If
End If
End Sub

