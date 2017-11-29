VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00808000&
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFC0&
      DataField       =   "economy"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   13440
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFC0&
      DataField       =   "business"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   11520
      TabIndex        =   17
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFC0&
      DataField       =   "executive"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   9720
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFC0&
      DataField       =   "economy"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFC0&
      DataField       =   "business"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFC0&
      DataField       =   "executive"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   8520
      Top             =   9840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
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
      Caption         =   "datewise_two"
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
      Height          =   615
      Left            =   4440
      Top             =   9840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
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
      Caption         =   "booking_details"
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
   Begin VB.Data Data1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "booking details"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "booking_details"
      Top             =   9840
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   12840
      TabIndex        =   12
      Top             =   9960
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LOGIN PAGE"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BOOKING "
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   14415
      Begin VB.CommandButton Command3 
         Caption         =   "CHECK BOOKING DETAILS"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CANCEL BOOKING"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   12000
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   7560
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   855
         Left            =   2760
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "CONTACT NO"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9840
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "FLIGHT ID"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEW BOOKING"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "OLD USER"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As String
Dim a As Long
Dim b As Long
Dim c As Long

Private Sub Command1_Click()
Dim f As Integer
f = 0

a = 0
b = 0
c = 0
Form4.Adodc1.Recordset.MoveFirst
Do While Not Form4.Adodc1.Recordset.EOF
If Form4.Adodc1.Recordset.Fields("flight ID") = Text1.Text And Form4.Adodc1.Recordset.Fields("date") = CDate(Text2.Text) And Form4.Adodc1.Recordset.Fields("contact") = Text3.Text Then
a = Val(Text5.Text)
b = Val(Text6.Text)
c = Val(Text7.Text)
Form4.Adodc1.Recordset.Delete
Form4.Adodc1.Refresh
f = 1
Exit Do
End If
Form4.Adodc1.Recordset.MoveNext
Loop

Dim d As Long
Dim e As Long
Dim g As Long

If f = 1 Then
Form4.Adodc2.Recordset.MoveFirst
Do While Not Form4.Adodc2.Recordset.EOF
If Form4.Adodc2.Recordset.Fields("flight ID") = Text1.Text And Form4.Adodc2.Recordset.Fields("date") = CDate(Text2.Text) Then
d = Val(Text8.Text)
e = Val(Text9.Text)
g = Val(Text10.Text)
Print a
Print b
Print c
Print d
Print e
Print g
Form4.Adodc2.Recordset.Fields("executive") = a + d
Form4.Adodc2.Recordset.Fields("business") = b + e
Form4.Adodc2.Recordset.Fields(2) = c + g
'Text8.Text = d + a
'Text9.Text = e + b
'Text10.Text = g + c
Form4.Adodc2.Recordset.Update
Form4.Adodc2.Refresh
Exit Do
End If
Form4.Adodc2.Recordset.MoveNext
Loop
End If


If f = 0 Then
r = MsgBox("Matched data not found", vbOKOnly, "OOPS..")
Else
r = MsgBox("Booking Cancel Successful!", vbOKOnly, "BOOKING CANCEL")
End If
End Sub

Private Sub Command2_Click()

Form2.Text14.Text = "ou"
Form2.Text15.Text = Form4.Text4.Text

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""


Form4.Hide
Unload Me
Form2.Show
Form2.SetFocus
End Sub

Private Sub Command3_Click()
Dim f As Integer
f = 0
Form4.Data1.Recordset.MoveFirst
Do While Not Form4.Data1.Recordset.EOF
If Form4.Data1.Recordset.Fields("flight ID") = Text1.Text And Form4.Data1.Recordset.Fields("date") = Text2.Text And Form4.Data1.Recordset.Fields("contact") = Text3.Text Then
f = 1
r = MsgBox(" FLIGHT ID: " & Form4.Data1.Recordset.Fields("flight ID") & vbNewLine & " BUSINESS SEAT: " & Form4.Data1.Recordset.Fields("business") & vbNewLine & " EXECUTIVE SEAT: " & Form4.Data1.Recordset.Fields("executive") & vbNewLine & " ECONOMY SEAT: " & Form4.Data1.Recordset.Fields("economy") & vbNewLine & " TOTAL BILL: " & Form4.Data1.Recordset.Fields("total bill") & vbNewLine & " DATE: " & Form4.Data1.Recordset.Fields("date") & vbNewLine & " SOURCE: " & Form4.Data1.Recordset.Fields("source") & vbNewLine & " DESTINATION: " & Form4.Data1.Recordset.Fields("destination") & vbNewLine & " ARRIVAL TIME: " & Form4.Data1.Recordset.Fields("arrival time") & vbNewLine & " DEPARTURE TIME: " & Form4.Data1.Recordset.Fields("deperture time"), vbOKOnly, "BOOKING DETAILS")
Exit Do
End If
Form4.Data1.Recordset.MoveNext
Loop
If f = 0 Then
r = MsgBox("Matched data not found", vbOKOnly, "OOPS..")
End If
End Sub

Private Sub Command4_Click()

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""

Form4.Hide
Unload Me
Form1.Show
Form1.SetFocus
End Sub

Private Sub Command5_Click()
r = MsgBox("Do you want to Exit?", vbQuestion + vbYesNo, "Exit")
If r = 6 Then
End
Else
Form4.Show
End If
End Sub

Private Sub Form_Load()
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
End Sub
