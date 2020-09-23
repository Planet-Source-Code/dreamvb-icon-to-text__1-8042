VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cool-Art 1.0"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&About"
      Height          =   405
      Left            =   4890
      TabIndex        =   10
      Top             =   4095
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   5145
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   9
      Top             =   3210
      Width           =   720
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2745
      TabIndex        =   8
      Text            =   "Ascii art by DreamVb"
      Top             =   4110
      Width           =   1830
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   870
      TabIndex        =   6
      Top             =   3600
      Width           =   3765
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   735
      TabIndex        =   4
      Top             =   3165
      Width           =   3900
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "....."
      Height          =   330
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3180
      Width           =   405
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   45
      Width           =   5715
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Convet and Save"
      Height          =   420
      Left            =   90
      TabIndex        =   0
      Top             =   4050
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Picture Title"
      Height          =   195
      Left            =   1830
      TabIndex        =   7
      Top             =   4155
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SaveText"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   3660
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Picture"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   3210
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim X, Y As Integer
Dim X1, Y2 As Integer
Dim TColour As Long
 Dim TextDraw As String
 X = Picture1.ScaleWidth - 1
 Y = Picture1.ScaleHeight - 1
  For X1 = 1 To X Step 1
    For Y2 = 1 To Y Step 1
     DoEvents
      TColour = Picture1.Point(Y2, X1)
       If TColour <= RGB(0, 0, 0) Then
        TextDraw = TextDraw + "#"
        ElseIf TColour <= RGB(180, 180, 180) Then
         TextDraw = TextDraw & "."
         ElseIf TColour <= RGB(250, 250, 250) Then
          TextDraw = TextDraw + "."
          Else
          TextDraw = TextDraw + " "
          End If
     Next
     TextDraw = TextDraw + vbNewLine
      DoEvents
       If Len(TextDraw) > 5000 Then
        Text1.Text = Text1.Text & TextDraw
        TextDraw = ""
    End If
      Next
      Text1.Text = TextDraw & vbNewLine & Text4
      SavePictureText Text3
       MsgBox "All done saved to " & Text3.Text
       
End Sub

Private Sub Command2_Click()
Text2 = Main.OpenFile
 Picture1.Picture = LoadPicture(Text2)
  If Len(Text2) = 0 Then
   Text3 = ""
   Else
   Text3 = Left(Text2, Len(Text2) - 3) + ".txt"
   End If
   
End Sub

Sub SavePictureText(Filename As String)
 Open Filename For Output As #1
  Print #1, Text1;
  Close #1
  
End Sub

Private Sub Command3_Click()
Form2.Show

End Sub

