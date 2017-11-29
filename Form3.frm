VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "CUSTOMER DETAILS"
      Height          =   5655
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   4695
      Begin VB.TextBox Text4 
         Height          =   855
         Left            =   1920
         TabIndex        =   9
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   975
         Left            =   1920
         TabIndex        =   7
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   1920
         TabIndex        =   5
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "AGE"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "ADDRESS"
         Height          =   735
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "CONTACT"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "NAME"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "BOOKING PAGE"
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Adodc1.Recordset.MoveNext
End Sub

Private Sub Form_Load()
Set Text1.DataSource = Form1.Adodc1
Text1.DataField = "name"
End Sub

