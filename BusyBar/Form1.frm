VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "BusyBar Demo"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2925
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin Project1.BusyBar BusyBar2 
      Height          =   285
      Left            =   465
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   615
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   503
      Speed           =   20
      Caption         =   "Busy Please Wait..."
      ColorStart      =   255
   End
   Begin Project1.BusyBar BusyBar1 
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   503
      BackColor       =   12648447
      CaptionShow     =   -1  'True
      CaptionColor    =   255
      ColorLines      =   14737632
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   975
      TabIndex        =   0
      Top             =   990
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   BusyBar1.Enabled = Not BusyBar1.Enabled
   BusyBar1.Visible = Not BusyBar1.Visible
   BusyBar2.Enabled = Not BusyBar2.Enabled
   BusyBar2.CaptionShow = Not BusyBar2.CaptionShow
End Sub

