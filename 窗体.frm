VERSION 5.00
Begin VB.Form 窗体 
   Caption         =   "高精度计算器"
   ClientHeight    =   4725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   Icon            =   "窗体.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5565
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "除法选项"
      Height          =   615
      Left            =   1560
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
      Begin VB.OptionButton Option6 
         Caption         =   "除高精"
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "除低精"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.OptionButton Option4 
      Caption         =   "/"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      Caption         =   "*"
      Height          =   180
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "-"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "+"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "窗体.frx":257B1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   1335
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "高精度计算器"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num1(1 To 1000) As Integer, num2(1 To 1000) As Integer, p As Integer, out(1 To 1000) As Integer
Private Sub storage(str1, str2, num1, num2)
    For i = 1 To Len(str1)
        num1(i) = Val(Mid(str1, Len(str1) - i + 1, 1))
    Next i
    For i = 1 To Len(str2)
        num2(i) = Val(Mid(str2, Len(str2) - i + 1, 1))
    Next i
End Sub
Private Function LPlus(str1 As String, str2 As String) As String
    p = 0
    LPlus = ""
    If Len(str1) < Len(str2) Then
        Dim temp As String
        temp = str1
        str1 = str2
        str2 = temp
    End If
    Call storage(str1, str2, num1, num2)
    For i = 1 To Len(str1)
        out(i) = num1(i) + num2(i) + p
        If out(i) >= 10 Then
            out(i) = out(i) - 10
            p = 1
        Else
            p = 0
        End If
    Next i
    If p = 1 Then LPlus = "1"
    For i = 1 To Len(str1)
        LPlus = LPlus + CStr(out(Len(str1) - i + 1))
    Next i
End Function
Private Function LMinus(str1 As String, str2 As String) As String
    Dim s As Integer
    p = 0
    LMinus = ""
    If Len(str1) < Len(str2) Or (Len(str1) = Len(str2) And str1 < str2) Then
        LMinus = "-"
        Dim temp As String
        temp = str1
        str1 = str2
        str2 = temp
    End If
    Call storage(str1, str2, num1, num2)
    For i = 1 To Len(str1)
        out(i) = num1(i) - num2(i) + p
        If out(i) < 0 Then
            out(i) = out(i) + 10
            p = -1
        Else
            p = 0
        End If
    Next i
    s = Len(str1)
    For i = Len(str1) To 1 Step -1
        If out(i) = 0 Then
            s = s - 1
        Else
            Exit For
        End If
    Next i
    For i = 1 To s
        LMinus = LMinus + CStr(out(s - i + 1))
    Next i
End Function
Function LMultiply(str1 As String, str2 As String) As String
    
End Function
Private Sub Command1_Click()
    If Option1 = True Then Text3.Text = LPlus(Text1.Text, Text2.Text)
    If Option2 = True Then Text3.Text = LMinus(Text1.Text, Text2.Text)
End Sub

Private Sub Option1_Click()
    Frame1.Visible = False
End Sub

Private Sub Option2_Click()
    Frame1.Visible = False
End Sub

Private Sub Option3_Click()
    Frame1.Visible = False
End Sub

Private Sub Option4_Click()
    Frame1.Visible = True
End Sub
