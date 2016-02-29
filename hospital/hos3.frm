VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000006&
   Caption         =   "Patient Details"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8325
   Icon            =   "hos3.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Return"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\hospital\Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "*dd/mm/yy"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Left            =   840
      TabIndex        =   12
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Patient  Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   840
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   840
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name"
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
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
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
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errorhandle2
If Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" Then
MsgBox "Please fill all fields and correctly", vbOKOnly, "Error"
Else
Data1.RecordSource = "select * from mem_mast"
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Label9.Caption
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Text3.Text
Data1.Recordset.Fields(3) = CDate(Text5.Text)
Data1.Recordset.Fields(4) = CInt(Text4.Text)
Data1.Recordset.Fields(5) = Text6.Text
Data1.Recordset.Update
Data1.Recordset.Close
Label9.Caption = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Form2.Hide
Form1.Show
Form1.SSTab1.Tab = 0
Form1.Data1.RecordSource = "select * from mem_mast"
Form1.Data1.Refresh

End If
Exit Sub
errorhandle2:
MsgBox "Please fill all fields", vbOKOnly, "Error"
End Sub

Private Sub Command2_Click()
Form1.Data1.RecordSource = "select * from mem_mast"
Form1.Data1.Refresh
Form2.Hide
Form1.Show
End Sub

Private Sub Text2_Click()
If Label9.Caption = "" Then
MsgBox "Please fill ID", vbOKOnly, "Error"
Else
Text2.SetFocus
End If
End Sub

Private Sub Text5_Clik()
If Text4.Text = "" Then
MsgBox "Please fill DOB", vbOKOnly, "Error"
Text3.SetFocus
Else
Text5.SetFocus
End If
End Sub


Private Sub Text6_Clik()
If Text5.Text = "" Then
MsgBox "Please fill Age", vbOKOnly, "Error"
Text4.SetFocus
Else
Text6.SetFocus
End If
End Sub

Private Sub Text3_Click()
If Text2.Text = "" Then
MsgBox "Please fill Name", vbOKOnly, "Error"
Text2.SetFocus
Else
Text3.SetFocus
End If

End Sub


Private Sub Text4_Click()
If Text5.Text = "" Then
MsgBox "Please fill DOB", vbOKOnly, "Error"
Text5.SetFocus
Else
Text4.SetFocus
End If
End Sub

Private Sub Text5_Click()
If Text3.Text = "" Then
MsgBox "Please fill Address", vbOKOnly, "Error"
Text3.SetFocus
Else
Text5.SetFocus
End If
End Sub

Private Sub Text6_Change()
Command1.Enabled = True
End Sub

Private Sub Text6_Click()
If Text4.Text = "" Then
MsgBox "Please fill Age", vbOKOnly, "Error"
Text4.SetFocus
Else
Text6.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
If Text6.Text = "" Then
Command1.Enabled = False
MsgBox "Please fill Sex", vbOKcancelOnly, "Error"
Else
Command1.Enabled = True
End If
End Sub
