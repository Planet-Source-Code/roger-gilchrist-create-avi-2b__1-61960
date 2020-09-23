VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProgressBar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Tracked this down to
'__merlin__
'ProgressBar2Class (8 DrawDirections, XOR Caption,Time2End Display)
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=23431&lngWId=1
'
'all I have done is a bit of Code Fixer tidying up
'Version 2
'added Left, Centre & Right options for the caption of the prog bar
'
Private PicBoxObj            As PictureBox
Private PrintCharTemp        As String
Private x1Temp               As Long
Private y1Temp               As Long
Public Enum eDrawDirection
  Left2Right = 0
  Right2Left = 1
  Top2Bottom = 2
  Bottom2Top = 3
  Left2RightReverse = 4
  Right2LeftReverse = 5
  Top2Bottomreverse = 6
  Bottom2TopReverse = 7
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Left2Right, Right2Left, Top2Bottom, Bottom2Top, Left2RightReverse, Right2LeftReverse, Top2Bottomreverse, Bottom2TopReverse
#End If
Public Enum eCaptionMode
  ShowPercentChange = 0
  ShowCaptionC = 1
  ShowCaptionL = 2
  ShowCaptionR = 3
  ShowNothing = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private ShowPercentChange, ShowCaptionC, ShowCaptionL, ShowCaptionR, ShowNothing
#End If
Private m_DrawDirection      As eDrawDirection
Private m_Min                As Double
Private m_Max                As Double
Private m_Value              As Double
Private m_DoEvents           As Boolean
Private m_Caption            As String
Private m_CaptionMode        As eCaptionMode
Private m_Timer              As Double


Private Sub CalcParam(ByRef dblMin As Double, _
                      ByRef dblMax As Double, _
                      ByRef dblValue As Double)

  dblMax = m_Max - m_Min
  dblValue = m_Value - m_Min
  dblMin = 0

End Sub

Public Property Get Caption() As String

  Caption = m_Caption

End Property

Public Property Let Caption(ByVal Char As String)

  m_Caption = Char

End Property

Public Property Get CaptionMode() As eCaptionMode

  CaptionMode = m_CaptionMode

End Property

Public Property Let CaptionMode(ByVal Mode As eCaptionMode)

  m_CaptionMode = Mode

End Property

Public Property Get CurrentTime() As String

  CurrentTime = Time2String(Int(Timer - m_Timer))

End Property

Public Property Get DoEventsByChange() As Boolean

  DoEventsByChange = m_DoEvents

End Property

Public Property Let DoEventsByChange(ByVal blnValue As Boolean)

  m_DoEvents = blnValue

End Property

Public Property Get DrawDirection() As eDrawDirection

  DrawDirection = m_DrawDirection

End Property

Public Property Let DrawDirection(ByVal drwValue As eDrawDirection)

  m_DrawDirection = drwValue

End Property

Public Property Get Max() As Double

  Max = m_Max

End Property

Public Property Let Max(ByVal dblValue As Double)

  m_Max = dblValue

End Property

Public Property Get Min() As Double

  Min = m_Min

End Property

Public Property Let Min(ByVal dblValue As Double)

  m_Min = dblValue

End Property

Public Property Get PictureBoxObjekt() As PictureBox

  Set PictureBoxObjekt = PicBoxObj

End Property

Public Property Set PictureBoxObjekt(PropVal As PictureBox)

  Set PicBoxObj = PropVal

End Property

Public Sub SetParamFast(ByVal dblMin As Double, _
                        ByVal dblMax As Double, _
                        ByVal DrawDirection As eDrawDirection, _
                        ByVal DoEventsByChange As Boolean, _
                        ByVal CaptionMode As eCaptionMode)

  m_Min = dblMin
  m_Max = dblMax
  m_DrawDirection = DrawDirection
  m_DoEvents = DoEventsByChange
  m_CaptionMode = CaptionMode

End Sub

Public Property Let SetPictureBox(ByVal PBO As Variant)

  Set PicBoxObj = PBO
  With PicBoxObj
    .AutoRedraw = True
    .ScaleMode = 3
    .BackColor = vbWhite
  End With 'PictureBoxObjekt

End Property

Public Sub ShowBar()


  Dim Change     As Boolean
  Dim StrCaption As String
  Dim LMax       As Double
  Dim Value      As Double
  Dim x          As Long
  Dim y          As Long
  Dim x1         As Long
  Dim y1         As Long

'Dim Min       As Double
  CalcParam 0, LMax, Value
  If Value Then
    If LMax Then
      Select Case m_DrawDirection
'Left2Right
       Case 0
        x1 = PicBoxObj.ScaleWidth / 100 * (Value / LMax * 100)
        y1 = PicBoxObj.ScaleHeight
        If x1Temp <> x1 Then
          Change = True
        End If
        x1Temp = x1
'Right2Left
       Case 1
        x = PicBoxObj.ScaleWidth
        y = PicBoxObj.ScaleHeight
        x1 = x - (PicBoxObj.ScaleWidth / 100 * (Value / LMax * 100))
        If x1Temp <> x1 Then
          Change = True
        End If
        x1Temp = x1
'Top2Bottom
       Case 2
        x1 = PicBoxObj.ScaleWidth
        y1 = PicBoxObj.ScaleHeight / 100 * (Value / LMax * 100)
        If y1Temp <> y1 Then
          Change = True
        End If
        y1Temp = y1
'Bottom2Top
       Case 3
        x = PicBoxObj.ScaleWidth
        y = PicBoxObj.ScaleHeight
        y1 = y - (PicBoxObj.ScaleHeight / 100 * (Value / LMax * 100))
        If y1Temp <> y1 Then
          Change = True
        End If
        y1Temp = y1
'Left2RightReverse
       Case 4
        x = PicBoxObj.ScaleWidth
        y = PicBoxObj.ScaleHeight
        x1 = PicBoxObj.ScaleWidth / 100 * (Value / LMax * 100)
        If x1Temp <> x1 Then
          Change = True
        End If
        x1Temp = x1
'Right2LeftReverse
       Case 5
        x1 = PicBoxObj.ScaleWidth - (PicBoxObj.ScaleWidth / 100 * (Value / LMax * 100))
        y1 = PicBoxObj.ScaleHeight
        If x1Temp <> x1 Then
          Change = True
        End If
        x1Temp = x1
'Top2BottomReverse
       Case 6
        x = PicBoxObj.ScaleWidth
        y = PicBoxObj.ScaleHeight
        y1 = (PicBoxObj.ScaleHeight / 100 * (Value / LMax * 100))
        If y1Temp <> y1 Then
          Change = True
        End If
        y1Temp = y1
'Bottom2TopReverse
       Case 7
        x1 = PicBoxObj.ScaleWidth
        y1 = PicBoxObj.ScaleHeight - (PicBoxObj.ScaleHeight / 100 * (Value / LMax * 100))
        If y1Temp <> y1 Then
          Change = True
        End If
        y1Temp = y1
      End Select
'--------------------------
      Select Case m_CaptionMode
       Case ShowPercentChange
        StrCaption = Int(Value / LMax * 100) & " %"
        If StrCaption <> PrintCharTemp Then
          Change = True
        End If
        PrintCharTemp = StrCaption
       Case ShowCaptionC, ShowCaptionR, ShowCaptionL
        StrCaption = m_Caption
        If StrCaption <> PrintCharTemp Then
          Change = True
        End If
        PrintCharTemp = StrCaption
      End Select
      If Change Then
        If m_CaptionMode <> ShowNothing Then
          With PicBoxObj
            .Cls
            Select Case m_CaptionMode ' alternate positions
             Case ShowPercentChange, ShowCaptionC
              .CurrentX = (.ScaleWidth / 2) - (.TextWidth(StrCaption) / 2)
             Case ShowCaptionL
              .CurrentX = 0
             Case ShowCaptionR
              .CurrentX = .ScaleWidth - .TextWidth(StrCaption)
            End Select
            .CurrentY = (.ScaleHeight / 2) - (.TextHeight(StrCaption) / 2)
            .DrawMode = 10 'vbMaskPenNot '13
            PicBoxObj.Print StrCaption
            .DrawMode = 10
            PicBoxObj.Line (x, y)-(x1, y1), vbRed, BF
          End With 'PictureBoxObjekt
         Else
          With PicBoxObj
            .Cls
            .DrawMode = 13
            PicBoxObj.Line (x, y)-(x1, y1), , BF
          End With 'PictureBoxObjekt
        End If
        If m_DoEvents Then
          DoEvents
        End If
      End If
    End If
  End If

End Sub

Public Sub StartTimer()

  m_Timer = Timer

End Sub

Public Property Get Time2End() As String

  Dim Max   As Double
  Dim Min   As Double
  Dim Value As Double
  Dim Temp  As Long

  CalcParam Min, Max, Value
  On Error Resume Next
  Temp = Int(Max / Value * (Timer - m_Timer) - (Timer - m_Timer))
  If LenB(Time2String(Temp)) Then
    Time2End = Time2String(Temp)
  End If
  On Error GoTo 0

End Property

Private Function Time2String(ByVal Seconds As Long) As String

  Dim sTemp As String
  Dim lTemp As Long

  If Seconds >= 31536000 Then
    sTemp = sTemp & Int(Seconds / 31536000) & " Year"
    sTemp = IIf(Int(Seconds / 31536000) > 1, sTemp & "s ", sTemp & " ")
    Seconds = Seconds Mod 31536000
    lTemp = lTemp + 1
  End If
  If Seconds >= 86400 Then
    sTemp = sTemp & Int(Seconds / 86400) & " Day"
    sTemp = IIf(Int(Seconds / 86400) > 1, sTemp & "s ", sTemp & " ")
    Seconds = Seconds Mod 86400
    lTemp = lTemp + 1
  End If
  If Seconds >= 3600 And lTemp < 2 Then
    sTemp = sTemp & Int(Seconds / 3600) & " Hour"
    sTemp = IIf(Int(Seconds / 3600) > 1, sTemp & "s ", sTemp & " ")
    Seconds = Seconds Mod 3600
    lTemp = lTemp + 1
  End If
  If Seconds >= 60 And lTemp < 2 Then
    sTemp = sTemp & Int(Seconds / 60) & " Minute"
    sTemp = IIf(Int(Seconds / 60) > 1, sTemp & "s ", sTemp & " ")
    Seconds = Seconds Mod 60
    lTemp = lTemp + 1
  End If
  If Seconds >= 1 And lTemp < 2 Then
    sTemp = sTemp & Seconds & " Seconds"
  End If
  Time2String = Trim$(sTemp)

End Function

Public Property Get Value() As Double

  Value = m_Value

End Property

Public Property Let Value(ByVal dblValue As Double)

  m_Value = dblValue

End Property

':)Code Fixer V4.0.0 (Saturday, 30 July 2005 22:45:36) 38 + 329 = 367 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|3333202222222222222222222222222|1112222|2221222|222222222233|111111111111|1222222222220|333333|

