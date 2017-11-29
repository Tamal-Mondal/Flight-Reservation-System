VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\VB projects\database\flight_reservation_system.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   975
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "users"
      Top             =   4680
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   735
      Left            =   7080
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Okey!"
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   1095
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   9495
      Begin VB.OptionButton Option3 
         Caption         =   "Old User"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Administrator"
         Height          =   375
         Left            =   7320
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "New User"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As Variant

Private Sub Command1_Click()
If Option1.Value = True Then
r = InputBox("What is your age?", "New User", "21")
ElseIf Option3.Value = True Then
r = InputBox("Enter your contact number", "Old User", "0000000000")
ElseIf Option2.Value = True Then
r = InputBox("Enter the password", "Administrator", "00000")
End If

End Sub

Private Sub Command2_Click()
r = MsgBox("Are you want to Exit?", vbQuestion + vbYesNo, "Exit")
If r = 6 Then
End
Else
Form1.Show
End If
End Sub
