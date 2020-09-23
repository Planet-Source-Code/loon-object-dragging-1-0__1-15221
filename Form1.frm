VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object Dragging 1.0   -   by loon"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3735
      Left            =   120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   3675
      ScaleWidth      =   5595
      TabIndex        =   5
      Top             =   840
      Width           =   5655
      Begin VB.CommandButton o1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Object 1 (Button)"
         Height          =   615
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox o2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "Form1.frx":4B588
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "loon 2k software"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   3120
         Width           =   3855
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hold your mouse down on any of the objects below and move it around."
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label l2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Object Dragging 1.0
'Created by loon
'http://www.electronerdz.com/loon

Private Sub o1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim drag1
Dim drag2
If Button = 1 Then
drag1 = (X - l1.Caption)
drag2 = (Y - l2.Caption)
o1.Left = o1.Left + (drag1)
o1.Top = o1.Top + (drag2)
Exit Sub
End If
l1 = X
l2 = Y
End Sub

Private Sub o2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim drag1
Dim drag2
If Button = 1 Then
drag1 = (X - l1.Caption)
drag2 = (Y - l2.Caption)
o2.Left = o2.Left + (drag1)
o2.Top = o2.Top + (drag2)
Exit Sub
End If
l1 = X
l2 = Y

End Sub

Private Sub Command1_Click()

End Sub
