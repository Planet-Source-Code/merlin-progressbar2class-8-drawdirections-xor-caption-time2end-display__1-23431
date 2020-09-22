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
'Class:     ProgressBar2 Beta
'Author:    Merlin
'Function:  PB with 8 DrawDirections, XOR Caption and Real Time2End Display
'Import:    PictureBox Object
'
'PS: If you like it please rate for me.

Option Explicit

Public PictureBoxObjekt As PictureBox

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

Public Enum eCaptionMode
    ShowPercentChange = 0
    ShowCaption = 1
    ShowNothing = 2
End Enum


Private m_DrawDirection As eDrawDirection
Private m_Min As Double
Private m_Max As Double
Private m_Value As Double
Private m_DoEvents As Boolean
Private m_Caption As String
Private m_CaptionMode As eCaptionMode

Private m_Timer As Double

Private Function Time2String(ByVal Seconds As Long) As String
Dim sTemp As String, lTemp As Long

    
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
    
Time2String = Trim(sTemp)
End Function

Public Property Get CurrentTime() As String
    CurrentTime = Time2String(Int(Timer - m_Timer))
End Property

Public Property Get Time2End() As String
Dim Max As Double, Min As Double, Value As Double, Temp As Long
CalcParam Min, Max, Value
    
On Error Resume Next
    
    Temp = Int(Max / Value * (Timer - m_Timer) - (Timer - m_Timer))
    If Time2String(Temp) <> "" Then Time2End = Time2String(Temp)
    

End Property



Public Property Let SetPictureBox(ByVal PBO As Variant)
    Set PictureBoxObjekt = PBO
    PictureBoxObjekt.AutoRedraw = True
    PictureBoxObjekt.ScaleMode = 3
    PictureBoxObjekt.BackColor = vbWhite
End Property

Public Property Let DrawDirection(ByVal Value As eDrawDirection)
    m_DrawDirection = Value
End Property

Public Property Get DrawDirection() As eDrawDirection
    DrawDirection = m_DrawDirection
End Property

Public Property Let DoEventsByChange(ByVal Value As Boolean)
     m_DoEvents = Value
End Property

Public Property Get DoEventsByChange() As Boolean
    DoEventsByChange = m_DoEvents
End Property

Public Property Get Min() As Double
    Min = m_Min
End Property

Public Property Let Min(ByVal Value As Double)
    m_Min = Value
End Property

Public Property Get Max() As Double
    Max = m_Max
End Property

Public Property Let Max(ByVal Value As Double)
    m_Max = Value
End Property

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal Value As Double)
    m_Value = Value
End Property

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

Public Sub StartTimer()
    m_Timer = Timer
End Sub


Public Sub SetParamFast(ByVal Min As Double, ByVal Max As Double, ByVal DrawDirection As eDrawDirection, ByVal DoEventsByChange As Boolean, ByVal CaptionMode As eCaptionMode)
    m_Min = Min
    m_Max = Max
    m_DrawDirection = DrawDirection
    m_DoEvents = DoEventsByChange
    m_CaptionMode = CaptionMode
End Sub

Private Sub CalcParam(ByRef Min As Double, ByRef Max As Double, ByRef Value As Double)
    Max = m_Max - m_Min
    Value = m_Value - m_Min
    Min = 0
End Sub




Public Sub ShowBar()
Dim Change As Boolean


'Caption
Dim PrintChar As String
Static PrintCharTemp As String

'CalcParam
Dim Max As Double, Min As Double, Value As Double
CalcParam Min, Max, Value

'DrawDirection
Dim x As Long, y As Long, x1 As Long, y1 As Long
Static xTemp As Long, yTemp As Long, x1Temp As Long, y1Temp As Long



Select Case m_DrawDirection
        
    'Left2Right
    Case Is = 0
        x = 0
        y = 0
        x1 = PictureBoxObjekt.ScaleWidth / 100 * (Value / Max * 100)
        y1 = PictureBoxObjekt.ScaleHeight
        If x1Temp <> x1 Then Change = True
        x1Temp = x1
    
    
    'Right2Left
    Case Is = 1
        x = PictureBoxObjekt.ScaleWidth
        y = PictureBoxObjekt.ScaleHeight
        x1 = x - (PictureBoxObjekt.ScaleWidth / 100 * (Value / Max * 100))
        y1 = 0
        If x1Temp <> x1 Then Change = True
        x1Temp = x1

    
    'Top2Bottom
    Case Is = 2
        x = 0
        y = 0
        x1 = PictureBoxObjekt.ScaleWidth
        y1 = PictureBoxObjekt.ScaleHeight / 100 * (Value / Max * 100)
        If y1Temp <> y1 Then Change = True
        y1Temp = y1

    
    'Bottom2Top
    Case Is = 3
        x = PictureBoxObjekt.ScaleWidth
        y = PictureBoxObjekt.ScaleHeight
        x1 = 0
        y1 = y - (PictureBoxObjekt.ScaleHeight / 100 * (Value / Max * 100))
        If y1Temp <> y1 Then Change = True
        y1Temp = y1


    'Left2RightReverse
    Case Is = 4
        x = PictureBoxObjekt.ScaleWidth
        y = PictureBoxObjekt.ScaleHeight
        x1 = 0 + (PictureBoxObjekt.ScaleWidth / 100 * (Value / Max * 100))
        y1 = 0
        If x1Temp <> x1 Then Change = True
        x1Temp = x1
    
    
    
    'Right2LeftReverse
    Case Is = 5
        x = 0
        y = 0
        x1 = PictureBoxObjekt.ScaleWidth - (PictureBoxObjekt.ScaleWidth / 100 * (Value / Max * 100))
        y1 = PictureBoxObjekt.ScaleHeight
        If x1Temp <> x1 Then Change = True
        x1Temp = x1

    
    'Top2BottomReverse
    Case Is = 6
        x = PictureBoxObjekt.ScaleWidth
        y = PictureBoxObjekt.ScaleHeight
        x1 = 0
        y1 = 0 + (PictureBoxObjekt.ScaleHeight / 100 * (Value / Max * 100))
        If y1Temp <> y1 Then Change = True
        y1Temp = y1

    
    'Bottom2TopReverse
    Case Is = 7
        x = 0
        y = 0
        x1 = PictureBoxObjekt.ScaleWidth
        y1 = PictureBoxObjekt.ScaleHeight - (PictureBoxObjekt.ScaleHeight / 100 * (Value / Max * 100))
        If y1Temp <> y1 Then Change = True
        y1Temp = y1


End Select




Select Case m_CaptionMode

    Case Is = 0
        PrintChar = Int(Value / Max * 100) & " %"
        If PrintChar <> PrintCharTemp Then Change = True
        PrintCharTemp = PrintChar
    
    Case Is = 1
        PrintChar = m_Caption
        If PrintChar <> PrintCharTemp Then Change = True
        PrintCharTemp = PrintChar

End Select



If Change = True Then
        
    If m_CaptionMode <> ShowNothing Then
        PictureBoxObjekt.Cls
        PictureBoxObjekt.CurrentX = (PictureBoxObjekt.ScaleWidth / 2) - (PictureBoxObjekt.TextWidth(PrintChar) / 2)
        PictureBoxObjekt.CurrentY = (PictureBoxObjekt.ScaleHeight / 2) - (PictureBoxObjekt.TextHeight(PrintChar) / 2)
        PictureBoxObjekt.DrawMode = 13
        PictureBoxObjekt.Print PrintChar
        PictureBoxObjekt.DrawMode = 10
        PictureBoxObjekt.Line (x, y)-(x1, y1), , BF
    Else
        PictureBoxObjekt.Cls
        PictureBoxObjekt.DrawMode = 13
        PictureBoxObjekt.Line (x, y)-(x1, y1), , BF
    End If
    
    If m_DoEvents = True Then DoEvents

End If



End Sub


