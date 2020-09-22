Attribute VB_Name = "Module1"
Public cBar As New ClsProgressBar2
Public cBar2 As New ClsProgressBar2

Public Sub main()
Form1.Show



cBar.SetPictureBox = Form1.Picture1
cBar.SetParamFast 1, 100000, Right2LeftReverse, True, ShowCaption
'cBar.PictureBoxObjekt.ForeColor = RGB(64, 138, 116)
cBar.PictureBoxObjekt.ForeColor = RGB(138, 64, 93)

cBar2.SetPictureBox = Form1.Picture2
cBar2.SetParamFast 0, 16 * cBar.Max, Left2Right, False, ShowPercentChange
cBar2.PictureBoxObjekt.ForeColor = RGB(64, 94, 138)


Dim lDrawDirection As Long

cBar2.StartTimer
For x = 0 To 15
    
    lDrawDirection = lDrawDirection + 1
    If lDrawDirection = 8 Then lDrawDirection = 0
    cBar.DrawDirection = lDrawDirection
    
    Select Case lDrawDirection
    
        Case Is = 0
            cBar.Caption = "DrawDirection: Left2Right"
        
        Case Is = 1
            cBar.Caption = "DrawDirection: Right2Left"
        
        Case Is = 2
            cBar.Caption = "DrawDirection: Top2Bottom"
        
        Case Is = 3
            cBar.Caption = "DrawDirection: Bottom2Top"
        
        Case Is = 4
            cBar.Caption = "DrawDirection: Left2RightReverse"
        
        Case Is = 5
            cBar.Caption = "DrawDirection: Right2LeftReverse"
        
        Case Is = 6
            cBar.Caption = "DrawDirection: Top2Bottomreverse"
        
        Case Is = 7
            cBar.Caption = "DrawDirection: Bottom2TopReverse"
        
    End Select
    

    For y = 1 To cBar.Max
        cBar.Value = y
        cBar.ShowBar
    
        cBar2.Value = x * cBar.Max + y
        cBar2.ShowBar
    Next y
    
Next x

Form1.Timer1 = False

    
End Sub
