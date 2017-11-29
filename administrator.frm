VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "booking details"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "booking_details"
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2400
      TabIndex        =   67
      Top             =   240
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "CHANGE ADMINISTRATOR DETAILS OR ADD NEW ADMINISTRATOR"
      Height          =   2055
      Left            =   360
      TabIndex        =   53
      Top             =   7680
      Width           =   11295
      Begin VB.CommandButton Command6 
         Caption         =   "ADD NEW"
         Height          =   1455
         Left            =   9840
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "UPDATE"
         Height          =   1455
         Left            =   8280
         TabIndex        =   66
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text31 
         Height          =   615
         Left            =   6720
         TabIndex        =   65
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text30 
         Height          =   615
         Left            =   3600
         TabIndex        =   63
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text29 
         Height          =   615
         Left            =   1080
         TabIndex        =   61
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text28 
         Height          =   615
         Left            =   6720
         TabIndex        =   59
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text27 
         Height          =   615
         Left            =   3600
         TabIndex        =   57
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text26 
         Height          =   615
         Left            =   1080
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "PASSWORD"
         Height          =   495
         Left            =   5400
         TabIndex        =   64
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "CONTACT"
         Height          =   495
         Left            =   2640
         TabIndex        =   62
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label30 
         Caption         =   "ADDRESS"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "DESIGNATION"
         Height          =   495
         Left            =   5400
         TabIndex        =   58
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "AGE"
         Height          =   495
         Left            =   2760
         TabIndex        =   56
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "NAME"
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LOGIN PAGE"
      Height          =   615
      Left            =   4800
      TabIndex        =   52
      Top             =   10080
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEW BOOKING"
      Height          =   615
      Left            =   600
      TabIndex        =   51
      Top             =   10080
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Caption         =   "ADD FLIGHT"
      Height          =   6495
      Left            =   5160
      TabIndex        =   10
      Top             =   960
      Width           =   9615
      Begin VB.CommandButton Command5 
         Caption         =   "UPDATE FLIGHT"
         Height          =   615
         Left            =   6600
         TabIndex        =   50
         Top             =   5520
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ADD FLIGHT"
         Height          =   615
         Left            =   3480
         TabIndex        =   49
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox Text24 
         Height          =   495
         Left            =   1560
         TabIndex        =   48
         Top             =   5640
         Width           =   1335
      End
      Begin VB.TextBox Text23 
         Height          =   495
         Left            =   7800
         TabIndex        =   46
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox Text22 
         Height          =   495
         Left            =   4560
         TabIndex        =   44
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         Height          =   495
         Left            =   1560
         TabIndex        =   42
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox Text20 
         Height          =   495
         Left            =   7800
         TabIndex        =   40
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         Height          =   495
         Left            =   4560
         TabIndex        =   38
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         Height          =   495
         Left            =   1560
         TabIndex        =   36
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         Height          =   615
         Left            =   7800
         TabIndex        =   34
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Height          =   615
         Left            =   4560
         TabIndex        =   32
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   1560
         TabIndex        =   30
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   7800
         TabIndex        =   28
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   4560
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   7800
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   4560
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   1560
         TabIndex        =   20
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   7800
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   4560
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "FRIDAY"
         Height          =   495
         Left            =   240
         TabIndex        =   47
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "THURSDAY"
         Height          =   495
         Left            =   6240
         TabIndex        =   45
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "WEDNESDAY"
         Height          =   495
         Left            =   3240
         TabIndex        =   43
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label22 
         Caption         =   "TUESDAY"
         Height          =   495
         Left            =   240
         TabIndex        =   41
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "MONDAY"
         Height          =   495
         Left            =   6240
         TabIndex        =   39
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "SUNDAY"
         Height          =   495
         Left            =   3240
         TabIndex        =   37
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "EXECUTIVE CLASS SEAT PRICE"
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "BUSINESS CLASS SEAT PRICE"
         Height          =   615
         Left            =   6240
         TabIndex        =   33
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "ECONOMY CLASS SEAT PRICE"
         Height          =   615
         Left            =   3240
         TabIndex        =   31
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "EXECUTIVE CLASS SEAT"
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "BUSINESS CLASS SEAT"
         Height          =   615
         Left            =   6240
         TabIndex        =   27
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "ECONOMY CLASS SEAT"
         Height          =   495
         Left            =   3240
         TabIndex        =   25
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "ARRIVAL TIME"
         Height          =   495
         Left            =   6240
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "DEPERTURE TIME"
         Height          =   495
         Left            =   3240
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "DESTINATION"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "SOURCE"
         Height          =   495
         Left            =   6240
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "SATURDAY"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "COMPANY"
         Height          =   495
         Left            =   3240
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "FLIGHT ID"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BOOKING"
      Height          =   6495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "CHECK BOOKING DETAILS"
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   4800
         Width           =   3735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CANCEL BOOKING"
         Height          =   855
         Left            =   360
         TabIndex        =   8
         Top             =   3480
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   2160
         TabIndex        =   7
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   2160
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "CONTACT"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "DATE"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "FLIGHT ID"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()

End Sub

Private Sub Command2_Click()
Dim f As Integer
f = 0
Form5.Data1.Recordset.MoveFirst
Do While Not Form5.Data1.Recordset.EOF
If Form5.Data1.Recordset.Fields("flight ID") = Text1.Text And Form5.Data1.Recordset.Fields("date") = Text2.Text And Form5.Data1.Recordset.Fields("contact") = Text3.Text Then
f = 1
r = MsgBox(" FLIGHT ID: " & Form5.Data1.Recordset.Fields("flight ID") & vbNewLine & " BUSINESS SEAT: " & Form5.Data1.Recordset.Fields("business") & vbNewLine & " EXECUTIVE SEAT: " & Form5.Data1.Recordset.Fields("executive") & vbNewLine & " ECONOMY SEAT: " & Form5.Data1.Recordset.Fields("economy") & vbNewLine & " TOTAL BILL: " & Form5.Data1.Recordset.Fields("total bill") & vbNewLine & " DATE: " & Form5.Data1.Recordset.Fields("date") & vbNewLine & " SOURCE: " & Form5.Data1.Recordset.Fields("source") & vbNewLine & " DESTINATION: " & Form5.Data1.Recordset.Fields("destination") & vbNewLine & " ARRIVAL TIME: " & Form5.Data1.Recordset.Fields("arrival time") & vbNewLine & " DEPARTURE TIME: " & Form5.Data1.Recordset.Fields("deperture time") & vbNewLine & " STATUS: " & Form5.Data1.Recordset.Fields("status"), vbOKOnly, "BOOKING DETAILS")
Exit Do
End If
Form5.Data1.Recordset.MoveNext
Loop
If f = 0 Then
r = MsgBox("Matched data not found", vbOKOnly, "OOPS..")
End If
End Sub

Private Sub Command3_Click()
Form2.Text14.Text = "ad"
Form2.Text15.Text = Text5.Text
Form2.Show
Form2.SetFocus
End Sub

Private Sub Command6_Click()
Form1.Adodc1.Recordset.AddNew
End Sub

Private Sub Command7_Click()
Form1.Show
Form1.SetFocus
End Sub

Private Sub Command8_Click()

Form1.Adodc1.Recordset.Update


End Sub

Private Sub Form_Load()
Set Text26.DataSource = Form1.Adodc1
Set Text27.DataSource = Form1.Adodc1
Set Text28.DataSource = Form1.Adodc1
Set Text29.DataSource = Form1.Adodc1
Set Text30.DataSource = Form1.Adodc1
Set Text31.DataSource = Form1.Adodc1
Text26.DataField = "name"
Text27.DataField = "age"
Text28.DataField = "designation"
Text29.DataField = "contact"
Text30.DataField = "address"
Text31.DataField = "password"
Form1.Adodc1.Recordset.MoveFirst
Do While Not Form1.Adodc1.Recordset.EOF
If Form1.Adodc1.Recordset.Fields("password") = Text5.Text Then
Exit Do
End If
Form1.Adodc1.Recordset.MoveNext
Loop
Form2.Refresh

End Sub
