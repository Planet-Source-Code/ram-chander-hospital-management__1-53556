VERSION 5.00
Begin VB.Form frmlabtest 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Text            =   " "
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Text            =   " Click on New To Add "
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Columns         =   5
      Height          =   4740
      Left            =   5880
      TabIndex        =   1
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAB EXAM DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   540
      Left            =   2280
      TabIndex        =   10
      Top             =   200
      Width           =   5610
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LAB EXAM ID's"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      TabIndex        =   7
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "frmlabtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
con.BeginTrans
Set rs = con.Execute("SELECT * FROM LABEXAMS")
Set rs = con.Execute("INSERT INTO LABEXAMS VALUES ('" & Text1.Text & "', '" & Text2.Text & "')")
con.CommitTrans
List1.AddItem (Text1.Text)
End Sub
Private Sub Command2_Click()
On Error Resume Next
con.BeginTrans
Set rs = con.Execute("SELECT * FROM LABEXAMS")
Set rs = con.Execute("UPDATE LABEXAMS SET EXAMNAME= '" & Text2.Text & "'")
con.CommitTrans
End Sub
Private Sub Command3_Click()
On Error Resume Next
Set rs = con.Execute("SELECT max(LABEXMID) FROM LABEXAMS")
 If rs.EOF = False Then
    Text1.Text = Val(rs(0)) + 1
Else
    Text1.Text = 1
End If
Text2.Enabled = True
Text2.Text = ""
End Sub
Private Sub Command4_Click()
Text2.Enabled = True
End Sub
Private Sub Form_Activate()
Call Connection.connected
Set rs = con.Execute("SELECT * FROM LABEXAMS")
Do Until rs.EOF
List1.AddItem (rs(0))
rs.MoveNext
Loop
End Sub
Private Sub List1_Click()
On Error Resume Next
Set rs = con.Execute("SELECT LABEXMID,EXAMNAME FROM LABEXAMS where LABEXMID=" & List1.Text)
    Text1.Text = rs(0)
    Text2.Text = rs(1)
End Sub

