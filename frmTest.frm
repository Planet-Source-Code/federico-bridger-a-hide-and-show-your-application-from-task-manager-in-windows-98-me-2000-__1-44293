VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ShowHideApp"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show in TaskManager"
      Height          =   495
      Left            =   713
      TabIndex        =   1
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide from TaskManager"
      Height          =   495
      Left            =   713
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmTest.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHide_Click()
  
  Dim cH As New cAppHider
  
  cH.HideApplication
  
End Sub

Private Sub cmdShow_Click()

  Dim cH As New cAppHider
  
  cH.ShowApplication Me.Caption

End Sub

