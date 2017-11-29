VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808000&
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox Text15 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2040
      TabIndex        =   37
      Top             =   9840
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   240
      TabIndex        =   36
      Top             =   9840
      Width           =   1455
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "datewise"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "datawise_two"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   11640
      TabIndex        =   32
      Top             =   1440
      Width           =   2895
      Begin VB.CommandButton Command6 
         Caption         =   "DETAILS"
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2880
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "FLIGHTS"
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
         Left            =   600
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BOOKING"
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
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
      Height          =   855
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FLIGHT DETAILS"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   10935
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   735
         Left            =   9240
         TabIndex        =   31
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   735
         Left            =   5640
         TabIndex        =   29
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   735
         Left            =   2040
         TabIndex        =   27
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   735
         Left            =   9240
         TabIndex        =   23
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   735
         Left            =   5640
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   735
         Left            =   2040
         TabIndex        =   19
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   8280
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   3000
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   8280
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   3000
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "PRICE"
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
         Left            =   7440
         TabIndex        =   30
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "PRICE"
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
         Left            =   3720
         TabIndex        =   28
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "PRICE"
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
         Left            =   360
         TabIndex        =   26
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "ECONOMY"
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
         Left            =   7440
         TabIndex        =   22
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "BUSINESS"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "EXECUTIVE"
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
         Left            =   360
         TabIndex        =   18
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "SEAT DETAILS"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "ARRIVAL TIME"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "DEPARTURE TIME"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "COMPANY NAME"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   10935
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   8880
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         Caption         =   "DESTINATION"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         Caption         =   "SOURCE"
         BeginProperty Font 
            Name            =   "Engravers MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
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
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "FLIGHT DETAILS PAGE"
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
      Left            =   5280
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As String
Dim f As Integer
Dim d As Date
Dim i As Integer

Private Sub Command2_Click()
If Text1.Text = "" Then
r = MsgBox("Please, enter a date!", vbOKOnly, "OOPS..")
Else
If Form1.Adodc3.Recordset.BOF Then
Form1.Adodc3.Recordset.MoveLast
Else
Form1.Adodc3.Recordset.MovePrevious
End If
End If
End Sub

Private Sub Command1_Click()
r = MsgBox("Do you want to Exit?", vbQuestion + vbYesNo, "Exit")
If r = 6 Then
End
Else
Form2.Show
End If
End Sub

Private Sub Command3_Click()

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Text8.Text = ""
'Text9.Text = ""
'Text10.Text = ""
'Text11.Text = ""
'Text12.Text = ""
'Text13.Text = ""


Form2.Hide
Form1.Show
Form1.SetFocus
End Sub

Private Sub Command4_Click()
Form3.Text5 = Text4
Form3.Text6 = Text5
Form3.Text7 = Text2
Form3.Text8 = Text3
Form3.Text9 = Text6
Form3.Text10 = Text7
Form3.Text14 = Text1
Form3.Text11 = Text11
Form3.Text12 = Text12
Form3.Text13 = Text13

Form3.Text19.Text = Text8.Text
Form3.Text20.Text = Text9.Text
Form3.Text21.Text = Text10.Text
Form3.Text22.Text = Text4.Text
Form3.Text23.Text = Text1.Text

Form3.Text30.Text = Text8.Text
Form3.Text31.Text = Text9.Text
Form3.Text32.Text = Text10.Text
Form3.Text33.Text = Text4.Text
Form3.Text34.Text = Text1.Text

Form3.Text27.Text = Text14.Text
Form3.Text28.Text = Text15.Text


'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Text8.Text = ""
'Text9.Text = ""
'Text10.Text = ""
'Text11.Text = ""
'Text12.Text = ""
'Text13.Text = ""


Form2.Hide
Form3.Show
Form3.SetFocus

End Sub

Private Sub Command5_Click()
Combo1.Clear


Dim a As Date
Dim b As Date
If Text1.Text <> "" Then
a = CDate(Text1.Text)
End If
b = Date


Dim f As Integer
If Text1.Text = "" Then
r = MsgBox("Please, enter a date , source and destination!", vbOKOnly, "OOPS..")
ElseIf a < b Then
r = MsgBox("Please, enter a valid date!", vbOKOnly, "OOPS..")
Else
d = CDate(Text1.Text)
i = Weekday(d)
f = 0
Form1.Adodc3.Recordset.MoveFirst
Do While Not Form1.Adodc3.Recordset.EOF
If Form1.Adodc3.Recordset.Fields("source") = Text2.Text And Form1.Adodc3.Recordset.Fields("destination") = Text3.Text Then
If i = 1 And Form1.Adodc3.Recordset.Fields("sun") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

If i = 2 And Form1.Adodc3.Recordset.Fields("mon") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

If i = 3 And Form1.Adodc3.Recordset.Fields("tue") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

If i = 4 And Form1.Adodc3.Recordset.Fields("wed") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

If i = 5 And Form1.Adodc3.Recordset.Fields("thu") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

If i = 6 And Form1.Adodc3.Recordset.Fields("fri") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

If i = 7 And Form1.Adodc3.Recordset.Fields("sat") = 1 Then
f = 1
Combo1.AddItem Form1.Adodc3.Recordset.Fields("flight ID")
End If

End If
Form1.Adodc3.Recordset.MoveNext
Loop
If f = 0 Then
r = MsgBox("Matched data not found", vbOKOnly, "OOPS..")
End If
End If
End Sub

Private Sub Command6_Click()
If Combo1.Text = "" Then
r = MsgBox("Matched data not found", vbOKOnly, "OOPS..")
Else
Form1.Adodc3.Recordset.MoveFirst
Do While Not Form1.Adodc3.Recordset.EOF
If Form1.Adodc3.Recordset.Fields("flight ID") = Combo1.Text Then
'Text4.Text = Form1.Adodc3.Recordset.Fields("flight ID")
'Text5.Text = Form1.Adodc3.Recordset.Fields("company")
'Text6.Text = Form1.Adodc3.Recordset.Fields("deperture time")
'Text7.Text = Form1.Adodc3.Recordset.Fields("arrival time")
Text8.Text = Form1.Adodc3.Recordset.Fields("executive")
Text9.Text = Form1.Adodc3.Recordset.Fields("business")
Text10.Text = Form1.Adodc3.Recordset.Fields("economy")
'Text11.Text = Form1.Adodc3.Recordset.Fields("executive price")
'Text12.Text = Form1.Adodc3.Recordset.Fields("business price")
'Text13.Text = Form1.Adodc3.Recordset.Fields("economy price")
Exit Do
End If
Form1.Adodc3.Recordset.MoveNext
Loop

Form2.Data1.Recordset.MoveFirst
Data1.Refresh
Do While Not Form2.Data1.Recordset.EOF
If Form2.Data1.Recordset.Fields("flight ID") = Combo1.Text And Form2.Data1.Recordset.Fields("date") = CDate(Text1.Text) Then
Text8.Text = Form2.Data1.Recordset.Fields("executive")
Text9.Text = Form2.Data1.Recordset.Fields("business")
Text10.Text = Form2.Data1.Recordset.Fields("economy")
Exit Do
End If
Form2.Data1.Recordset.MoveNext
Loop

End If
End Sub

Private Sub Form_Load()
Set Text4.DataSource = Form1.Adodc3
Set Text5.DataSource = Form1.Adodc3
Set Text6.DataSource = Form1.Adodc3
Set Text7.DataSource = Form1.Adodc3
'Set Text8.DataSource = Form1.Adodc3
'Set Text9.DataSource = Form1.Adodc3
'Set Text10.DataSource = Form1.Adodc3
Set Text11.DataSource = Form1.Adodc3
Set Text12.DataSource = Form1.Adodc3
Set Text13.DataSource = Form1.Adodc3

Text4.DataField = "flight ID"
Text5.DataField = "company"
Text6.DataField = "deperture time"
Text7.DataField = "arrival time"
'Text8.DataField = "executive"
'Text9.DataField = "business"
'Text10.DataField = "economy"
Text11.DataField = "executive price"
Text12.DataField = "business price"
Text13.DataField = "economy price"

f = 0
Form1.Adodc3.Recordset.MoveFirst


'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
'Text5.Text = ""
'Text6.Text = ""
'Text7.Text = ""
'Text8.Text = ""
'Text9.Text = ""
'Text10.Text = ""
'Text11.Text = ""
'Text12.Text = ""
'Text13.Text = ""

End Sub
