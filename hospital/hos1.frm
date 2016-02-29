VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Hospital"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   17.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   Icon            =   "hos1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF80FF&
      Caption         =   "Go to Home"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C000&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9360
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF80FF&
      Caption         =   "Remove Patient"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "New Patient"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7800
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "D:\hospital\Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\hospital\Member.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   19500
      _Version        =   393216
      TabHeight       =   1058
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483626
      ForeColor       =   8404992
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Patient Details"
      TabPicture(0)   =   "hos1.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DBGrid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "   Home"
      TabPicture(1)   =   "hos1.frx":09F5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1"
      Tab(1).Control(1)=   "Image2"
      Tab(1).Control(2)=   "Image3"
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(4)=   "Image4"
      Tab(1).Control(5)=   "Image5"
      Tab(1).Control(6)=   "Image6"
      Tab(1).Control(7)=   "Command5"
      Tab(1).Control(8)=   "Command12"
      Tab(1).Control(9)=   "Command3"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "  Treatment Details"
      TabPicture(2)   =   "hos1.frx":0B1C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(6)=   "Label10"
      Tab(2).Control(7)=   "Label11"
      Tab(2).Control(8)=   "Label12"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label14"
      Tab(2).Control(11)=   "Label15"
      Tab(2).Control(12)=   "Label3"
      Tab(2).Control(13)=   "DBGrid3"
      Tab(2).Control(14)=   "Command11"
      Tab(2).Control(15)=   "Text1"
      Tab(2).Control(16)=   "Command8"
      Tab(2).ControlCount=   17
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000D&
         Caption         =   "Add New Patient"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -70440
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8040
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C000&
         Caption         =   "Go to Home"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -62040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   9360
         Width           =   2175
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H8000000D&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -62040
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   8040
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H8000000D&
         Caption         =   "Exitting Patient"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -66240
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   8040
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66600
         TabIndex        =   15
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C000&
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -62880
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "hos1.frx":1155
         Height          =   3135
         Left            =   -73920
         OleObjectBlob   =   "hos1.frx":1169
         TabIndex        =   2
         Top             =   5760
         Width           =   18015
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "hos1.frx":1B44
         Height          =   3495
         Left            =   2640
         OleObjectBlob   =   "hos1.frx":1B58
         TabIndex        =   1
         Top             =   2520
         Width           =   15255
      End
      Begin VB.Label Label3 
         Caption         =   "Treatment History"
         BeginProperty Font 
            Name            =   "Budmo Jiggler"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69240
         TabIndex        =   27
         Top             =   960
         Width           =   8895
      End
      Begin VB.Label Label1 
         Caption         =   "Patient LIST"
         BeginProperty Font 
            Name            =   "Budmo Jiggler"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7440
         TabIndex        =   26
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Image Image6 
         Height          =   1080
         Left            =   -61080
         Picture         =   "hos1.frx":2533
         Top             =   6720
         Width           =   1080
      End
      Begin VB.Image Image5 
         Height          =   1080
         Left            =   -65160
         Picture         =   "hos1.frx":26D5
         Top             =   6840
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   1080
         Left            =   -69360
         Picture         =   "hos1.frx":283D
         Top             =   6840
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Hospital Database Management"
         BeginProperty Font 
            Name            =   "Budmo Jiggler"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68880
         TabIndex        =   24
         Top             =   1200
         Width           =   8895
      End
      Begin VB.Image Image3 
         Height          =   1500
         Left            =   -62040
         Picture         =   "hos1.frx":299B
         Top             =   4560
         Width           =   1500
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   -68880
         Picture         =   "hos1.frx":36ED
         Top             =   4560
         Width           =   1500
      End
      Begin VB.Image Image1 
         Height          =   3525
         Left            =   -66480
         Picture         =   "hos1.frx":4442
         Top             =   2760
         Width           =   3525
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment List "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   -65520
         TabIndex        =   20
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label14 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -62280
         TabIndex        =   19
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -62280
         TabIndex        =   18
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -62280
         TabIndex        =   17
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -68520
         TabIndex        =   28
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68520
         TabIndex        =   16
         Top             =   3000
         Width           =   3135
      End
      Begin VB.Label Label9 
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
         Height          =   615
         Left            =   -63360
         TabIndex        =   14
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label8 
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
         Height          =   615
         Left            =   -63360
         TabIndex        =   13
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -63360
         TabIndex        =   12
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label6 
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
         Height          =   615
         Left            =   -70560
         TabIndex        =   11
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Height          =   615
         Left            =   -70560
         TabIndex        =   10
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Patient ID :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69360
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

c = 0
Data1.RecordSource = "select Patient_ID from mem_mast"
Data1.Refresh

Do While Not Data1.Recordset.EOF
c = c + 1
Data1.Recordset.MoveNext
Loop

If c = 0 Then
Form2.Label9.Caption = 1

Else
Form2.Label9.Caption = c + 1
End If

Form1.Hide
Form2.Show
Form2.Text2.Text = ""
Form2.Text3.Text = ""
Form2.Text4.Text = ""
Form2.Text5.Text = ""
Form2.Text6.Text = ""
Form2.Text2.SetFocus
Data1.Recordset.Close
Form1.Data1.RecordSource = "select * from mem_mast"
Form1.Data1.Refresh

End Sub

Private Sub Command10_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub Command11_Click()
On Error GoTo errorhandle:
Form1.Data1.RecordSource = "select * from mem_mast where patient_id='" + Text1.Text + "'"
Form1.Data1.Refresh
Form1.Data1.Recordset.Edit
Label10.Caption = Form1.Data1.Recordset.Fields(1)
Label11.Caption = Form1.Data1.Recordset.Fields(2)
Label12.Caption = Form1.Data1.Recordset.Fields(3)
Label13.Caption = Form1.Data1.Recordset.Fields(4)
Label14.Caption = Form1.Data1.Recordset.Fields(5)
Form1.Data1.Recordset.Update
Form1.Data1.Recordset.Close
Data3.RecordSource = "select * from Issue_mast where patient_id='" + Text1.Text + "'"
Data3.Refresh
Exit Sub
errorhandle:
MsgBox "Error occured!Wrong Patient ID", vbInformation, "Error"
Text1.Text = ""
Text1.SetFocus
End Sub
Private Sub Command12_Click()
End
End Sub
Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "Please fill Patient id", vbOKOnly, "Error"
Text1.SetFocus
Else
Form1.Hide
Form3.Show
Form3.Label10.Caption = Form1.Text1.Text
Form3.Label11.Caption = Date
Form3.Text1.Text = ""
Form3.Text2.Text = ""
Form3.Text3.Text = ""
Form3.Text4.Text = ""
Form3.Text5.Text = ""
Form3.Text6.Text = ""
End If
End Sub

Private Sub Command3_Click()
Command1 = True
End Sub

Private Sub Command4_Click()
On Error GoTo Errorhan
dp = InputBox("Enter the patient ID", "Delete Patient")
If dp = "" Then
SSTab1.Tab = 0
Else
Data1.RecordSource = "select * from mem_mast where patient_id='" + dp + "'"
Data1.Refresh
Data1.Recordset.Delete
Data1.Recordset.Close
Data1.RecordSource = "select * from mem_mast"
Data1.Refresh
Data3.RecordSource = "select * from issue_mast where patient_id='" + dp + "'"
Data3.Refresh
Do While Not Data3.Recordset.EOF
c = c + 1
Data3.Recordset.Delete
Data3.Recordset.MoveNext
Loop
If c <> 0 Then
MsgBox "Records are Deleted Successfully", vbOKOnly, "Information"
End If


Data3.RecordSource = "select * from issue_mast"
Data3.Refresh
End If
Exit Sub
Errorhan:
Data1.RecordSource = "select * from mem_mast"
Data1.Refresh
MsgBox "Please enter correct Patient id", vbOKOnly, "Error"
End Sub

Private Sub Command5_Click()
SSTab1.Tab = 2
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Then
MsgBox "Please fill Patient id", vbOKOnly, "Error"
Text1.SetFocus
Else
Form1.Hide
Form6.Show
End If
End Sub

Private Sub Command7_Click()
SSTab1.Tab = 1
End Sub

Private Sub Command8_Click()
SSTab1.Tab = 1
End Sub

Private Sub Form_Load()
SSTab1.Tab = 1
Command1.Visible = False
Command2.Visible = False
Command4.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
Data1.RecordSource = "select * from Mem_mast"
Data1.Refresh
Command1.Visible = True
Command2.Visible = False
Command4.Visible = True
Command6.Visible = False
Command7.Visible = True
Command8.Visible = False

ElseIf SSTab1.Tab = 1 Then
Command1.Visible = False
Command2.Visible = False
Command4.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
ElseIf SSTab1.Tab = 2 Then
Data3.RecordSource = "select * from Issue_mast"
Data3.Refresh
Command1.Visible = False
Command2.Visible = True
Command4.Visible = False
Command6.Visible = True
Command7.Visible = False
Command8.Visible = True
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Text1.Text = ""
Text1.SetFocus
End If
End Sub
