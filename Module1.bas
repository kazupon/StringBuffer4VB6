Attribute VB_Name = "Module1"
Option Explicit

' Win32 API 関数定義
Public Declare Function QueryPerformanceCounter Lib "Kernel32" (X As Currency) As Boolean
Public Declare Function QueryPerformanceFrequency Lib "Kernel32" (X As Currency) As Boolean
