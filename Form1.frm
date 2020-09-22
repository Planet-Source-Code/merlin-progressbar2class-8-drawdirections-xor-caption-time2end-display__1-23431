VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "ProgressBar2 Class Beta by Merlin - Demo"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4200
   StartUpPosition =   1  'Fenstermitte
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   244
      TabIndex        =   0
      Top             =   360
      Width           =   3720
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
    Set cBar = Nothing
    Set cBar2 = Nothing
    End
End Sub

Private Sub Timer1_Timer()

Form1.Frame1.Caption = "Progress Ready in " & cBar2.Time2End
Form1.Frame2.Caption = "Elapsed Time: " & cBar2.CurrentTime

End Sub
