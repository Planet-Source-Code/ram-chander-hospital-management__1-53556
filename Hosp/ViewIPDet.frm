VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VewIPDet 
   Caption         =   "View Inpatient Details"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7223
      _Version        =   393216
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
      Begin VB.TextBox txttme 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   2
         Text            =   "Time"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtdt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   10680
         TabIndex        =   1
         Text            =   "Date"
         Top             =   240
         Width           =   1095
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
         TabIndex        =   5
         Top             =   600
         Width           =   1470
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
         TabIndex        =   3
         Top             =   0
         Width           =   7875
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   240
   End
End
Attribute VB_Name = "VewIPDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim db As Database
'Dim rs As Recordset
Dim cur As Integer
Private Sub Command1_Click()
rs.Requery
If rs.RecordCount = 0 Then
MsgBox "Database Empty"
Exit Sub
End If
rs.MoveLast
MSFlexGrid1.Rows = rs.RecordCount + 1
rs.MoveFirst
cur = 0
Do Until rs.EOF
cur = cur + 1
MSFlexGrid1.TextMatrix(cur, 0) = cur
For i = 1 To 3
MSFlexGrid1.TextMatrix(cur, i + 1) = rs.Fields(i)
Next i
rs.MoveNext
Loop
End Sub
Private Sub Form_Load()
VewIPDet.WindowState = 2
Set db = OpenDatabase(App.Path + "\MlkHos.mdb")
Set rs = db.OpenRecordset("Select * from Ipdet")

End Sub
Private Sub HEAD1()
'MSFlexGrid1.Clear 'Clear the datas in the flexgrid
MSFlexGrid1.Rows = 2 'Set the total number of rows in the flexgrid as 2
'X = "Slno" 'Assign Slno to x
For i = 1 To 25 'Set the title in flexgrid as Field1,Field2,.....Field25
X = X + "|                                                Field" & i
Next i 'End of for loop
main.MSFlexGrid1.FormatString = X 'Display the title
End Sub

