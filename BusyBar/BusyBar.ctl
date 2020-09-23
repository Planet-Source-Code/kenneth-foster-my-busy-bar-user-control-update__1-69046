VERSION 5.00
Begin VB.UserControl BusyBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ScaleHeight     =   1215
   ScaleWidth      =   2205
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   210
      Top             =   645
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   1125
         ScaleHeight     =   330
         ScaleWidth      =   120
         TabIndex        =   2
         Top             =   -15
         Width           =   120
      End
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   0
         ScaleHeight     =   330
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   -15
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   30
         Width           =   855
      End
   End
End
Attribute VB_Name = "BusyBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'**                                  Busy Bar
'**                               Version 1.1.0
'**                               By Ken Foster
'**                                 July 2007
'**                     Freeware--- no copyrights claimed
'*******************************************************************
'added caption and OLE_Color selection
'=============================================

Option Explicit

Const m_def_Caption = "Please Wait"
Const m_def_CaptionShow = False
Const m_def_Enabled = False
Const m_def_Speed = 30
Const m_def_ColorStart = &H202020           'off black
Const m_def_ColorEnd = &H606060     'off white
Const m_def_BackColor = vbBlack
Const m_def_CaptionColor = vbWhite
Const m_def_ColorLines = &H606060

Dim m_Caption As String
Dim m_CaptionShow As Boolean
Dim m_Enabled As Boolean
Dim m_Speed As Integer
Dim m_ColorStart As OLE_COLOR
Dim m_ColorEnd As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR
Dim m_ColorLines As OLE_COLOR

Dim tpR As Byte
Dim tpG As Byte
Dim tpB As Byte
Dim mdR As Byte
Dim mdG As Byte
Dim mdB As Byte

Private Sub UserControl_InitProperties()
    Caption = m_def_Caption
    CaptionShow = m_def_CaptionShow
    Enabled = m_def_Enabled
    Speed = m_def_Speed
    ColorStart = m_def_ColorStart
    ColorEnd = m_def_ColorEnd
    BackColor = m_def_BackColor
    CaptionColor = m_def_CaptionColor
    ColorLines = m_def_ColorLines
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Enabled = .ReadProperty("Enabled", m_def_Enabled)
        Speed = .ReadProperty("Speed", m_def_Speed)
        BackColor = .ReadProperty("BackColor", m_def_BackColor)
        Caption = .ReadProperty("Caption", m_def_Caption)
        CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
        CaptionShow = .ReadProperty("CaptionShow", m_def_CaptionShow)
        ColorStart = .ReadProperty("ColorStart", m_def_ColorStart)
        ColorEnd = .ReadProperty("ColorEnd", m_def_ColorEnd)
        ColorLines = .ReadProperty("ColorLines", m_def_ColorLines)
    End With
End Sub

Private Sub UserControl_Resize()

  'set up some parameters to keep everything in order
    Picture1.Height = 280
    Picture1.Width = UserControl.Width
    UserControl.Height = Picture1.Height
    picBar(0).Height = 20
    picBar(1).Height = 20
    picBar(0).Top = -1
    picBar(1).Top = -1
    picBar(0).Width = 6
    picBar(1).Width = 6
    picBar(0).Left = 0
    picBar(1).Left = Picture1.ScaleWidth - picBar(1).Width
    'draw the bar lines
    Picture1.Line (0, 3)-(Picture1.Width, 3), m_ColorLines
    Picture1.Line (0, 8)-(Picture1.Width, 8), m_ColorLines
    Picture1.Line (0, 13)-(Picture1.Width, 13), m_ColorLines
    'center the caption text
    Label1.Left = (Picture1.ScaleWidth / 2) - Label1.Width / 2
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", m_Enabled, m_def_Enabled
        .WriteProperty "Speed", m_Speed, m_def_Speed
        .WriteProperty "BackColor", m_BackColor, m_def_BackColor
        .WriteProperty "Caption", m_Caption, m_def_Caption
        .WriteProperty "CaptionShow", m_CaptionShow, m_def_CaptionShow
        .WriteProperty "ColorStart", m_ColorStart, m_def_ColorStart
        .WriteProperty "ColorEnd", m_ColorEnd, m_def_ColorEnd
        .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
        .WriteProperty "ColorLines", m_ColorLines, m_def_ColorLines
    End With
End Sub

Private Sub DrawGrad()                          'DrawGrad
  Dim R1 As Double
  Dim G1 As Double
  Dim B1 As Double
  Dim R2 As Double
  Dim G2 As Double
  Dim B2 As Double
  Dim R0 As Double
  Dim G0 As Double
  Dim B0 As Double
  Dim i As Integer
  Dim btG As Double
  Dim btR As Double
  Dim btB As Double
  
    btG = tpG
    btR = tpR
    btB = tpB
    With picBar(0)
        .AutoRedraw = True
        .ScaleMode = 0
        .ScaleHeight = 100
        .ScaleWidth = 1
    End With
    With picBar(1)
        .AutoRedraw = True
        .ScaleMode = 0
        .ScaleHeight = 100
        .ScaleWidth = 1
        picBar(1).Left = Picture1.ScaleWidth - picBar(1).Width
    End With
    R0 = tpR
    G0 = tpG
    B0 = tpB

    R1 = (mdR - R0) / 50
    G1 = (mdG - G0) / 50
    B1 = (mdB - B0) / 50
    R2 = (btR - mdR) / 50
    G2 = (btG - mdG) / 50
    B2 = (btB - mdB) / 50

    For i = 0 To 100
        picBar(0).Line (0, i)-(1, i), RGB(R0 * 2.55, G0 * 2.55, B0 * 2.55)
        picBar(1).Line (0, i)-(1, i), RGB(R0 * 2.55, G0 * 2.55, B0 * 2.55)
        If i < 50 Then
            R0 = R0 + R1
            G0 = G0 + G1
            B0 = B0 + B1
          Else 'NOT I...
            R0 = R0 + R2
            G0 = G0 + G2
            B0 = B0 + B2
        End If
    Next i
End Sub

Private Sub Timer1_Timer()                                                                  'Timer1
    If picBar(0).Left + picBar(0).Width > Picture1.ScaleWidth Then
        picBar(0).Left = 0
        picBar(1).Left = Picture1.ScaleWidth - picBar(1).Width
      Else
        picBar(1).Left = picBar(1).Left - 3
        picBar(0).Left = picBar(0).Left + 3
    End If
End Sub

Private Function ConvertHexToRGB(HexValue, ByRef RValue As Byte, ByRef GValue As Byte, ByRef BValue As Byte)
    Dim ConvColor As String
    ConvColor = String(6 - Len(Trim(CStr(Hex(HexValue)))), "0") + Trim(CStr(Hex(HexValue)))
    ConvColor = Right(ConvColor, 2) & Left(Right(ConvColor, 4), 2) & Left(ConvColor, 2)
    RValue = CByte("&H" + Left(ConvColor, 2))
    GValue = CByte("&H" + Left(Right(ConvColor, 4), 2))
    BValue = CByte("&H" + Right(ConvColor, 2))
End Function

Public Property Get BackColor() As OLE_COLOR                                'BackColor
   Let BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
   Let m_BackColor = NewBackColor
   PropertyChanged "BackColor"
   Picture1.BackColor = m_BackColor
   UserControl_Resize
End Property

Public Property Get Caption() As String                                            'Caption
    Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
    m_Caption = NewCaption
    Label1.Caption = m_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Public Property Get CaptionShow() As Boolean                                 'CaptionShow
    CaptionShow = m_CaptionShow
End Property

Public Property Let CaptionShow(NewCaptionShow As Boolean)
    m_CaptionShow = NewCaptionShow
    Label1.Visible = m_CaptionShow
    PropertyChanged "CaptionShow"
End Property

Public Property Get CaptionColor() As OLE_COLOR                          'CaptionColor
   Let CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal NewCaptionColor As OLE_COLOR)
   Let m_CaptionColor = NewCaptionColor
   PropertyChanged "CaptionColor"
   Label1.ForeColor = m_CaptionColor
   UserControl_Resize
End Property
Public Property Get ColorStart() As OLE_COLOR                                   'ColorStart
   Let ColorStart = m_ColorStart
End Property

Public Property Let ColorStart(ByVal NewColorStart As OLE_COLOR)
   Let m_ColorStart = NewColorStart
   PropertyChanged "ColorStart"
   ConvertHexToRGB ColorStart, tpR, tpG, tpB
   DrawGrad
End Property

Public Property Get ColorEnd() As OLE_COLOR                                     'ColorEnd
   Let ColorEnd = m_ColorEnd
End Property

Public Property Let ColorEnd(ByVal NewColorEnd As OLE_COLOR)
   Let m_ColorEnd = NewColorEnd
   PropertyChanged "ColorEnd"
   ConvertHexToRGB ColorEnd, mdR, mdG, mdB
   DrawGrad
End Property

Public Property Get ColorLines() As OLE_COLOR                                    'ColorLines
   Let ColorLines = m_ColorLines
End Property

Public Property Let ColorLines(ByVal NewColorLines As OLE_COLOR)
   Let m_ColorLines = NewColorLines
   PropertyChanged "ColorLines"
   UserControl_Resize
End Property

Public Property Get Enabled() As Boolean                                           'Enabled
    Enabled = m_Enabled
End Property

Public Property Let Enabled(NewEnabled As Boolean)
    m_Enabled = NewEnabled
    Timer1.Enabled = m_Enabled
    If m_Enabled = False Then
       picBar(0).Left = 0
       picBar(1).Left = Picture1.ScaleWidth - picBar(1).Width
    End If
    PropertyChanged "Enabled"
End Property

Public Property Get Speed() As Integer                                              'Speed
    Speed = m_Speed
End Property

Public Property Let Speed(NewSpeed As Integer)
    m_Speed = NewSpeed
    Timer1.Interval = m_Speed
    PropertyChanged "Speed"
End Property

Private Sub UserControl_Terminate()                                                     'Terminate
    Timer1.Enabled = False
End Sub
