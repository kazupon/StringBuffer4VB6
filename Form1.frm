VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "String Paformance"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Frame Frame2 
      Caption         =   "StringBuffer"
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   3735
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   960
         TabIndex        =   11
         Text            =   "hello world."
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   960
         TabIndex        =   10
         Text            =   "1024"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Run"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "String"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Count"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&& "
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Run"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Text            =   "1024"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Text            =   "hello world."
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Count"
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "String"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   450
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.Text2 = CStr(2 ^ 14)
  Me.Text5 = CStr(2 ^ 14)
End Sub


Private Sub Command1_Click()
  
  ' 計測開始
  Dim l_start As Currency
  Call QueryPerformanceCounter(l_start)
  
  
  ' VB6 '&' による文字列連結
  
  Dim l_count As Long
  Dim l_string As String
  l_count = CLng(Text2.Text)
  l_string = Text1.Text
  
  Dim l_index As Long
  For l_index = 0 To l_count
    l_string = l_string & Text1.Text
  Next
  Debug.Print "VB6 : " & l_string
  
  
  ' 計測終了
  Dim l_end As Currency
  Call QueryPerformanceCounter(l_end)
  
  
  ' 計測結果
  Dim l_measure As Currency
    Dim l_freq As Currency
  Call QueryPerformanceFrequency(l_freq)
  l_measure = (l_end - l_start) / l_freq
  Text3.Text = CStr(l_measure)
  
End Sub

Private Sub Command2_Click()
  
  ' 計測開始
  Dim l_start As Currency
  Call QueryPerformanceCounter(l_start)
  
  
  ' StringBuffer による文字列連結
  
  Dim l_count As Long
  Dim l_string As String
  l_count = CLng(Text5.Text)
  
  Dim l_buffer As StringBuffer
  Set l_buffer = New StringBuffer
  Dim l_index As Long
  For l_index = 0 To l_count
    Call l_buffer.Append(Text4.Text)
  Next
  Debug.Print "StringBuffer : " & l_buffer.ToString()
  
  
  ' 計測終了
  Dim l_end As Currency
  Call QueryPerformanceCounter(l_end)
  
  
  ' 計測結果
  Dim l_measure As Currency
    Dim l_freq As Currency
  Call QueryPerformanceFrequency(l_freq)
  l_measure = (l_end - l_start) / l_freq
  Text6.Text = CStr(l_measure)
  
End Sub

