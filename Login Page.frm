VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   4890
   ClientTop       =   4365
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Engravers MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   975
      Left            =   2640
      Top             =   8040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
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
      BackColor       =   16777152
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "datawise_two"
      Caption         =   "DATEWISE_TWO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "USERS"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   975
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   8040
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   975
      Left            =   10680
      Top             =   6480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1720
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
      BackColor       =   16777152
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "flight_details"
      Caption         =   "Flight_Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   975
      Left            =   5280
      Top             =   6480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1720
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
      BackColor       =   16777152
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "booking_details"
      Caption         =   "Booking_Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   120
      Top             =   6480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1720
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
      BackColor       =   16777152
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\VB projects\database\flight_reservation_system.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "users"
      Caption         =   "users"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   735
      Left            =   8640
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Okay"
      Height          =   735
      Left            =   3720
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   13815
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Old User"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8400
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Administrator"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFF00&
         Caption         =   "New User"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2280
         MaskColor       =   &H00FFFF00&
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "LOGIN PAGE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   6
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As String
Dim a As Integer
Dim c As Variant
Dim f As Integer


Private Sub Command1_Click()
If Option1.Value = True Then
r = InputBox("What is your age?", "New User", "0")
a = Val(r)
If a = 0 Then
Form1.Show
ElseIf a < 21 Then
r = MsgBox("Sorry, We can't allow you to use the application!", vbOKOnly, "Exit")
Else
Form2.Text14.Text = "nu"
Form2.Text15.Text = 0
Form1.Hide
'Unload Me
Form2.Show
Form2.SetFocus
End If



ElseIf Option3.Value = True Then
r = InputBox("Enter your contact number", "Old User", "0000000000")
f = 0
Form1.Data1.RecordSource = "Select * from users"
Form1.Data1.Refresh
Form1.Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If Form1.Data1.Recordset.Fields("contact") = r Then
f = 1
Exit Do
End If
Form1.Data1.Recordset.MoveNext
Loop
If r = "0000000000" Then
Form1.Show
ElseIf f = 1 Then
c = MsgBox("Hello " & Form1.Data1.Recordset.Fields("name") & "!", vbOK, "WELCOME")
If c = vbOK Then
Form4.Text4.Text = r
Form4.Text3.Text = r
'Unload Me
Form4.Show
Form4.SetFocus
End If
Else
r = MsgBox("Sorry, data not found!", vbOKOnly, "OOPS..")
End If



'ElseIf Option2.Value = True Then
'r = InputBox("Enter the password", "Administrator", "00000")
'f = 0
'Form1.Data1.Recordset.MoveFirst
'Do While Not Form1.Data1.Recordset.EOF
'If Form1.Data1.Recordset.Fields("password") = r Then
'Form5.Text26.Text = Form1.Data1.Recordset.Fields("name")
'Form5.Text27.Text = Form1.Data1.Recordset.Fields("age")
'Form5.Text28.Text = Form1.Data1.Recordset.Fields("designation")
'Form5.Text29.Text = Form1.Data1.Recordset.Fields("contact")
'Form5.Text30.Text = Form1.Data1.Recordset.Fields("address")
'Form5.Text31.Text = Form1.Data1.Recordset.Fields("password")
'f = 1
'Exit Do
'End If
'Form1.Data1.Recordset.MoveNext
'Loop
'If f = 1 Then
'c = MsgBox("Hellow " & Form1.Data1.Recordset.Fields("name") & "!", vbOK, "WELCOME")
'If c = vbOK Then
'Form5.Text5.Text = r
'Form5.Show
'End If
'Else
'r = MsgBox("Sorry, data not found!", vbOKOnly, "OOPS..")
'End If
End If
End Sub

Private Sub Command2_Click()
r = MsgBox("Do you want to Exit?", vbQuestion + vbYesNo, "Exit")
If r = 6 Then
End
Else
Form1.Show
End If
End Sub

