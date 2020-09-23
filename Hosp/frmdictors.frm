VERSION 5.00
Begin VB.Form frmdoctors 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Doctors Details"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmdictors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "EDIT"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "NEW"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Text            =   " "
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Text            =   " "
      Top             =   3780
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Text            =   " "
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Text            =   " "
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "UPDATE"
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Columns         =   5
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   5880
      TabIndex        =   10
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Code Number of the Doctor"
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
      Left            =   0
      TabIndex        =   16
      Top             =   720
      Width           =   3915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTORS DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   2505
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nature"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specialization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   140
      Width           =   2700
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmdoctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
Call Connection.connected
Set rs = con.Execute("SELECT * FROM DOCTORS")
  Do Until rs.EOF
    List1.AddItem (rs(0))
rs.MoveNext
Loop
Command3.SetFocus
End Sub
Private Sub Command1_Click()
On Error Resume Next
con.BeginTrans
Set rs = con.Execute("SELECT * FROM DOCTORS")
Set rs = con.Execute("INSERT INTO DOCTORS(DOCCODE,DOCNAME,SPCLZATION,NATURE,QUALIFICATION VALUES ('" & Text1.Text & "', '" & Text2.Text & "','" & Text3.Text & "','" & Text5.Text & "','" & Text4.Text & "')")
con.CommitTrans
End Sub
Private Sub Command2_Click()
On Error Resume Next
con.BeginTrans
Set rs = con.Execute("SELECT * FROM DOCTORS")
Set rs = con.Execute("UPDATE DOCTORS SET DOCNAME='" & Text2.Text & "' ,SPCLZATION='" & Text3.Text & "',NATURE='" & Text4.Text & "',QUALIFICATION='" & Text5.Text & "' where DOCCODE=" & Text1.Text)
con.CommitTrans
',SPCLZATION='" & Text3.Text & "',NATURE='" & Text4.Text & "',QUALIFICATION='" & Text5.Text & "'
End Sub
Private Sub Command3_Click()
On Error Resume Next
Set rs = con.Execute("SELECT max(CODE) FROM DOCTORS")
 If rs.EOF = False Then
    Text1.Text = Val(rs(0)) + 1
    
Else
    Text1.Text = 1
End If
Text2.Enabled = True
Text3.Enabled = True
Text5.Enabled = True
Text4.Enabled = True
End Sub
Private Sub Command4_Click()
On Error Resume Next
Text2.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text5.Enabled = True
Text4.Enabled = True
End Sub
Private Sub List1_Click()
On Error Resume Next
con.BeginTrans
 Set rs = con.Execute("SELECT DOCCODE,DOCNAME,SPCLZATION,NATURE,QUALIFICATION FROM DOCTORS where CODE=" & List1.Text)
  'Set rs = con.Execute("SELECT DOCCODE,DNAME  FROM DOCTORS where CODE=" & List1.Text)
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text5.Text = rs(4)
    Text4.Text = rs(3)
    con.CommitTrans
End Sub

