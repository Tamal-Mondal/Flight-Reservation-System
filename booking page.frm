VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00808000&
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   4890
   ClientTop       =   4155
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Data Data2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "booking_details"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "booking_details"
      Top             =   10080
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8400
      Width           =   3015
   End
   Begin VB.TextBox Text34 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   5880
      TabIndex        =   60
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox Text33 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   4560
      TabIndex        =   59
      Top             =   9480
      Width           =   975
   End
   Begin VB.TextBox Text32 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   3240
      TabIndex        =   58
      Top             =   9480
      Width           =   975
   End
   Begin VB.TextBox Text31 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   1920
      TabIndex        =   57
      Top             =   9480
      Width           =   975
   End
   Begin VB.TextBox Text30 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   600
      TabIndex        =   56
      Top             =   9480
      Width           =   975
   End
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
      Height          =   855
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Data Data1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "datewise_two"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   12600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "datawise_two"
      Top             =   10080
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   3600
      Top             =   10080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      RecordSource    =   "datewise"
      Caption         =   "datewise"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   6600
      Top             =   10080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
   Begin VB.TextBox Text29 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   4320
      TabIndex        =   54
      Top             =   120
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9600
      Top             =   10080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
   Begin VB.TextBox Text28 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2400
      TabIndex        =   53
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text27 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   600
      TabIndex        =   52
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text26 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   13320
      TabIndex        =   51
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text25 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   11400
      TabIndex        =   50
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text24 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   9480
      TabIndex        =   49
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text23 
      BackColor       =   &H00FFFFC0&
      DataField       =   "date"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   13680
      TabIndex        =   48
      Top             =   9480
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      BackColor       =   &H00FFFFC0&
      DataField       =   "flight ID"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   12120
      TabIndex        =   47
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox Text21 
      BackColor       =   &H00FFFFC0&
      DataField       =   "economy"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   10560
      TabIndex        =   46
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox Text20 
      BackColor       =   &H00FFFFC0&
      DataField       =   "business"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   9000
      TabIndex        =   45
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox Text19 
      BackColor       =   &H00FFFFC0&
      DataField       =   "executive"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   7440
      TabIndex        =   44
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEAT BOOKING"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5640
      TabIndex        =   37
      Top             =   6720
      Width           =   9255
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   7920
         TabIndex        =   43
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   4920
         TabIndex        =   41
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   1920
         TabIndex        =   39
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label18 
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
         Height          =   615
         Left            =   6240
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label17 
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
         Height          =   615
         Left            =   3240
         TabIndex        =   40
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         Height          =   615
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "BOOKING CONFIRMATION"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8400
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FLIGHT DETAILS PAGE"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8400
      Width           =   2775
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BILL"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   600
      TabIndex        =   31
      Top             =   6720
      Width           =   4695
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   2640
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "TOTAL BILL"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "FLIGHT DETAILS "
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   9255
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   6840
         TabIndex        =   30
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   2280
         TabIndex        =   28
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   6840
         TabIndex        =   26
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   2280
         TabIndex        =   24
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   6840
         TabIndex        =   22
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   2280
         TabIndex        =   20
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   6840
         TabIndex        =   18
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   2280
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   6840
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label15 
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
         Height          =   615
         Left            =   4800
         TabIndex        =   29
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "ECONOMY PRICE "
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
         TabIndex        =   27
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "BUSINESS PRICE"
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
         Left            =   4800
         TabIndex        =   25
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "EXECUTIVE PRICE"
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
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         Left            =   4800
         TabIndex        =   21
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Left            =   4800
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "COMPANY"
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
         Left            =   4800
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "FLIGHT  ID"
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
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   4695
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   1920
         TabIndex        =   8
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Height          =   615
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "AGE"
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
         Left            =   240
         TabIndex        =   7
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "ADDRESS"
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
         Left            =   240
         TabIndex        =   5
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "CONTACT"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         Caption         =   "NAME"
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
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "BOOKING PAGE"
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
      Left            =   6480
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ec As Long
Dim ex As Long
Dim bu As Long
Dim fi As String
Dim da As Date
Dim r As String
Dim a As Long
Dim b As Long
Dim c As Long
Dim f As Integer
Dim na As String
Dim ad As String
Dim co As String
Dim ph As String

Private Sub Command1_Click()
Dim q As String
q = MsgBox("Do you want to Exit?", vbQuestion + vbYesNo, "Exit")
If q = 6 Then

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
'Text14.Text = ""
'Text15.Text = 0
'Text16.Text = 0
'Text17.Text = 0
'Text18.Text = 0


End
Else
Form3.Show
End If
End Sub

Private Sub Command2_Click()
na = Text1.Text
co = Text2.Text
ad = Text3.Text
Text15.Text = Val(Text11.Text) * Val(Text16.Text) + Val(Text12.Text) * Val(Text17.Text) + Val(Text13.Text) * Val(Text18.Text)
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
'Text14.Text = ""
'Text15.Text = 0
'Text16.Text = 0
'Text17.Text = 0
'Text18.Text = 0

Form3.Hide
Unload Me
Form1.Show
Form1.SetFocus
End Sub

Private Sub Command4_Click()

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
'Text14.Text = ""
'Text15.Text = 0
'Text16.Text = 0
'Text17.Text = 0
'Text18.Text = 0


Form3.Hide
Unload Me
Form2.Show
Form2.SetFocus
End Sub

Private Sub Command5_Click()

If ex >= Val(Text16.Text) And bu >= Val(Text17.Text) And ec >= Val(Text18.Text) Then

ph = Text2.Text
If Text27.Text = "ou" Then
Form3.Adodc1.Recordset.Update
Form3.Adodc1.Refresh
Else
Form3.Adodc1.Recordset.MoveNext
Form3.Adodc1.Refresh
End If


'f = 0
'Form3.Adodc2.Recordset.MoveFirst
'Do While Not Form3.Adodc2.Recordset.EOF
'If Form3.Adodc2.Recordset.Fields("flight ID") = Text5.Text And Form3.Adodc2.Recordset.Fields("date") = CDate(Text14.Text) Then
'a = Form3.Adodc2.Recordset.Fields("executive")
'b = Form3.Adodc2.Recordset.Fields("business")
'c = Form3.Adodc2.Recordset.Fields(2)
'Text19.Text = ex - Val(Text16.Text)
'Text20.Text = bu - Val(Text17.Text)
'Text21.Text = ec - Val(Text18.Text)
'Text22.Text = fi
'Text23.Text = da
'Form3.Adodc2.Recordset.Fields("executive") = a - Val(Text16.Text)
'Form3.Adodc2.Recordset.Fields("business") = b - Val(Text17.Text)
'Form3.Adodc2.Recordset.Fields(2) = c - Val(Text18.Text)
'Print Form1.Adodc4.Recordset.Fields("flight ID")
'Print Form1.Adodc4.Recordset.Fields("date")
'Print Form1.Adodc4.Recordset.Fields("executive")
'Print Form1.Adodc4.Recordset.Fields("business")
'Print Form1.Adodc4.Recordset.Fields(2)
'Form3.Adodc2.Recordset.Update
'Form3.Adodc2.Refresh
'f = 1
'Exit Do
'End If
'Form3.Adodc2.Recordset.MoveNext
'Loop
'If f = 0 Then
'Form3.Adodc2.Recordset.MoveLast
'Form3.Adodc2.Recordset.AddNew
Text19.Text = ex - Val(Text16.Text)
Text20.Text = bu - Val(Text17.Text)
Text21.Text = ec - Val(Text18.Text)
Text22.Text = fi
Text23.Text = da
'Form3.Adodc2.Recordset.Fields("executive") = ex - Val(Text16.Text)
'Form3.Adodc2.Recordset.Fields("business") = bu - Val(Text17.Text)
'Form3.Adodc2.Recordset.Fields(2) = ec - Val(Text18.Text)
'Form3.Adodc2.Recordset.Fields("flight id") = fi
'Form3.Adodc2.Recordset.Fields("date") = da
'Form1.Adodc4.Recordset.MoveFirst
If f = 0 Then
Form3.Adodc2.Recordset.MoveNext
Form3.Adodc2.Refresh
Else
a = Form3.Adodc2.Recordset.Fields("executive")
b = Form3.Adodc2.Recordset.Fields("business")
c = Form3.Adodc2.Recordset.Fields(2)
Form3.Adodc2.Recordset.Fields("executive") = a - Val(Text16.Text)
Form3.Adodc2.Recordset.Fields("business") = b - Val(Text17.Text)
Form3.Adodc2.Recordset.Fields(2) = c - Val(Text18.Text)
Form3.Adodc2.Recordset.Update
Form3.Adodc2.Refresh
End If



Form1.Adodc2.Refresh
Dim d As String
Dim e As Long
Dim g As String
Dim h As String
Dim i As String
Dim q As String
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Date
'Dim n As String
'Dim o As String
'Dim p As String
d = Text5.Text
e = Text15.Text
g = Text7.Text
h = Text8.Text
i = Text9.Text
q = Text10.Text
j = Text16.Text
k = Text17.Text
l = Text18.Text
m = CDate(Text14.Text)
'n = Text24.Text
'o = Text25.Text
'p = Text26.Text
Set Text5.DataSource = Form1.Adodc2
Set Text15.DataSource = Form1.Adodc2
Set Text7.DataSource = Form1.Adodc2
Set Text8.DataSource = Form1.Adodc2
Set Text9.DataSource = Form1.Adodc2
Set Text10.DataSource = Form1.Adodc2
Set Text16.DataSource = Form1.Adodc2
Set Text17.DataSource = Form1.Adodc2
Set Text18.DataSource = Form1.Adodc2
Set Text14.DataSource = Form1.Adodc2
Set Text24.DataSource = Form1.Adodc2
Set Text25.DataSource = Form1.Adodc2
Set Text26.DataSource = Form1.Adodc2
Text5.DataField = "flight ID"
Text15.DataField = "total bill"
Text7.DataField = "source"
Text8.DataField = "destination"
Text9.DataField = "deperture time"
Text10.DataField = "arrival time"
Text16.DataField = "executive"
Text17.DataField = "business"
Text18.DataField = "economy"
Text14.DataField = "date"
Text24.DataField = "name"
Text25.DataField = "contact"
Text26.DataField = "address"
'Text5.Text = ""
'Text15.Text = ""
'Text7.Text = ""
'Text8.Text = ""
'Text9.Text = ""
'Text10.Text = ""
'Text16.Text = ""
'Text17.Text = ""
'Text18.Text = ""
'Text14.Text = ""
'Text24.Text = ""
'Text25.Text = ""
'Text26.Text = ""
Form1.Adodc2.Recordset.AddNew
Text5.Text = d
Text15.Text = e
Text7.Text = g
Text8.Text = h
Text9.Text = i
Text10.Text = q
Text16.Text = j
Text17.Text = k
Text18.Text = l
Text14.Text = m
'Form1.Adodc2.Recordset.Fields("name").Value = Text1.Text
'Form1.Adodc2.Recordset.Fields("contact").Value = Text2.Text
'Form1.Adodc2.Recordset.Fields("address").Value = Text3.Text
Text24.Text = na
Text25.Text = co
Text26.Text = ad
Form1.Adodc2.Recordset.MoveNext
Form1.Adodc2.Refresh
'Form1.Adodc2.Recordset.AddNew


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
'Text14.Text = ""
'Text15.Text = 0
'Text16.Text = 0
'Text17.Text = 0
'Text18.Text = 0


r = MsgBox("BOOKING SUCCESSFUL!", vbOKOnly, "BOOKING")

Else
r = MsgBox("BOOKING IS NOT POSSIBLE!", vbOKOnly, "BOOKING")
End If

End Sub

Private Sub Command6_Click()
Dim flag As Integer
flag = 0
Form3.Data2.Refresh
Form3.Data2.Recordset.MoveFirst
Do While Not Form3.Data2.Recordset.EOF
If Form3.Data2.Recordset.Fields("flight ID") = fi And Form3.Data2.Recordset.Fields("date") = da And Form3.Data2.Recordset.Fields("contact") = ph Then
flag = 1
r = MsgBox(" FLIGHT ID: " & Form3.Data2.Recordset.Fields("flight ID") & vbNewLine & " BUSINESS SEAT: " & Form3.Data2.Recordset.Fields("business") & vbNewLine & " EXECUTIVE SEAT: " & Form3.Data2.Recordset.Fields("executive") & vbNewLine & " ECONOMY SEAT: " & Form3.Data2.Recordset.Fields("economy") & vbNewLine & " TOTAL BILL: " & Form3.Data2.Recordset.Fields("total bill") & vbNewLine & " DATE: " & Form3.Data2.Recordset.Fields("date") & vbNewLine & " SOURCE: " & Form3.Data2.Recordset.Fields("source") & vbNewLine & " DESTINATION: " & Form3.Data2.Recordset.Fields("destination") & vbNewLine & " ARRIVAL TIME: " & Form3.Data2.Recordset.Fields("arrival time") & vbNewLine & " DEPARTURE TIME: " & Form3.Data2.Recordset.Fields("deperture time"), vbOKOnly, "BOOKING DETAILS")
Exit Do
End If
Form3.Data2.Recordset.MoveNext
Loop
If flag = 0 Then
r = MsgBox("Matched data not found", vbOKOnly, "OOPS..")
End If
End Sub

Private Sub Form_Load()

Form3.Text27.Text = Form2.Text14.Text
Form3.Text28.Text = Form2.Text15.Text

'Dim x As String
'Dim y As String
'x = Text27.Text
'y = Text28.Text
'Text29.Text = y

ex = Val(Form2.Text8.Text)
bu = Val(Form2.Text9.Text)
ec = Val(Form2.Text10.Text)
fi = Form2.Text4
da = CDate(Form2.Text1.Text)

'Unload Form2

'Set Text19.DataSource = Form3.Adodc2
'Set Text20.DataSource = Form3.Adodc2
'Set Text21.DataSource = Form3.Adodc2
'Set Text22.DataSource = Form3.Adodc2
'Set Text23.DataSource = Form3.Adodc2
'Text19.DataField = "executive"
'Text20.DataField = "business"
'Text21.DataField = "economy"
'Text22.DataField = "flight id"
'Text23.DataField = "date"
f = 0
Form3.Adodc2.Recordset.MoveFirst
Do While Not Form3.Adodc2.Recordset.EOF
If Form3.Adodc2.Recordset.Fields("flight ID") = fi And Form3.Adodc2.Recordset.Fields("date") = da Then
f = 1
Exit Do
End If
Form3.Adodc2.Recordset.MoveNext
Loop
If f = 0 Then
Form3.Adodc2.Recordset.AddNew
End If

'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""
Set Text1.DataSource = Form3.Adodc1
Set Text2.DataSource = Form3.Adodc1
Set Text3.DataSource = Form3.Adodc1
Set Text4.DataSource = Form3.Adodc1
Text1.DataField = "name"
Text2.DataField = "contact"
Text3.DataField = "address"
Text4.DataField = "age"
'Text1.Text = ""
'Text2.Text = ""
'Text3.Text = ""
'Text4.Text = ""




If Trim(Text27.Text) = Trim("ou") Then
Form3.Adodc1.Recordset.MoveFirst
Do While Not Form3.Adodc1.Recordset.EOF
If Form3.Adodc1.Recordset.Fields(2) = Trim(Text28.Text) Then
'Print Form3.Adodc1.Recordset.Fields("name")
'Print Form3.Adodc1.Recordset.Fields("contact")
'Print Form3.Adodc1.Recordset.Fields("address")
'Print Form3.Adodc1.Recordset.Fields("age")
'Print "abc"
Exit Do
End If
Form3.Adodc1.Recordset.MoveNext
Loop
Else
Form3.Adodc1.Recordset.AddNew
'Exit Sub
End If


'Set Text19.DataSource = Form3.Adodc2
'Set Text20.DataSource = Form3.Adodc2
'Set Text21.DataSource = Form3.Adodc2
'Set Text22.DataSource = Form3.Adodc2
'Set Text23.DataSource = Form3.Adodc2
'Text19.DataField = "executive"
'Text20.DataField = "business"
'Text21.DataField = "economy"
'Text22.DataField = "flight id"
'Text23.DataField = "date"

End Sub

