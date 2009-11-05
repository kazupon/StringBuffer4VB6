VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -----------------------------------------------------------------------------
' クラス名： StringBuffer
' -----------------------------------------------------------------------------


' -----------------------------------------------------------------------------
' 実装インターフェイス


' -----------------------------------------------------------------------------


' -----------------------------------------------------------------------------
' フィールド

Private Const DEFAULT_BUFFER_SIZE As Long = 1024  ' デフォルトのバッファサイズ
Private m_buffer As String  ' 文字列用バッファ
Private m_length As Long  ' 文字列の長さ
Private m_bufferCapacityIncrement As Long ' 文字列バッファの増加量
Private m_bufferCapacity As Long ' 文字列用バッファの容量

' -----------------------------------------------------------------------------


' -----------------------------------------------------------------------------
' メソッド or プロパティ

' コンストラクタ
Private Sub Class_Initialize()
  Call Initialize
End Sub


' 指定した文字列の長さで文字列用バッファサイズを初期化するメソッド
Public Sub Initialize(Optional ByVal p_InitialCapacity As Long = DEFAULT_BUFFER_SIZE)
  
  ' 文字列の長さを初期化
  m_length = 0
  
  ' 同時に確保する文字数
  m_bufferCapacityIncrement = p_InitialCapacity
  
  ' 1未満ならデフォルト値で初期化する
  If m_bufferCapacityIncrement < 1 Then
    m_bufferCapacityIncrement = DEFAULT_BUFFER_SIZE
  End If
  
  ' 文字列用バッファの初期化
  m_buffer = String$(m_bufferCapacityIncrement, 0)
  ' 文字列用バッファの容量を初期化
  m_bufferCapacity = m_bufferCapacityIncrement
  
End Sub


' 文字列バッファに指定した文字列を追加するメソッド
Public Sub Append(ByRef p_string As String)
  
  Dim l_appendStringLength As Long ' 追加する文字列長
  
  l_appendStringLength = Len(p_string)
  
  ' 現在確保しているバッファ容量が足りない場合は確保する
  If m_bufferCapacity < (m_length + l_appendStringLength) Then
    m_bufferCapacityIncrement = m_bufferCapacityIncrement * 2
    m_buffer = m_buffer & String$(m_bufferCapacityIncrement, 0)
    m_bufferCapacity = m_bufferCapacity + m_bufferCapacityIncrement
  End If
  
  ' 文字列を追加
  Mid$(m_buffer, m_length + 1, l_appendStringLength) = p_string
  
  ' 文字列長の設定
  m_length = m_length + l_appendStringLength
  
End Sub

' 文字列バッファに指定した文字列郡を追加するメソッド
Public Sub Appends(ParamArray p_strings() As Variant)
  
  If IsMissing(p_strings) Then
    GoTo Finally
  End If
  
  Dim l_string As Variant
  For Each l_string In p_strings
    Call Me.Append(CStr(l_string))
  Next
  
Finally:
  
End Sub

' 文字列バッファに指定した文字列を改行付きで追加するメソッド
Public Sub AppendLine(Optional ByRef p_string As String = "")

  If Not p_string = vbNullString Then
    Call Me.Append(p_string)
  End If
  
  Call Me.Append(vbCrLf)
  
End Sub

' 文字列バッファに指定した文字列郡を最後に改行付きで追加するメソッド
Public Sub AppendsLine(ParamArray p_strings() As Variant)
  
  If IsMissing(p_strings) Then
    GoTo Finally
  End If
  
  Dim l_string As Variant
  For Each l_string In p_strings
    Call Me.Append(CStr(l_string))
  Next
  Call Me.Append(vbCrLf)
  
Finally:
  
End Sub

' 文字列用のバッファをクリアするメソッド
Public Sub Clear()

  m_buffer = vbNullString
  m_length = 0
  m_bufferCapacity = 0
  
  Call Me.Initialize
  
End Sub


' 文字列の取得
Public Property Get ToString() As String
  ToString = Left$(m_buffer, m_length)
End Property


' 文字列バッファの容量を取得するプロパティ
Public Property Get Capacity() As Long
  Capacity = m_bufferCapacity
End Property

' -----------------------------------------------------------------------------
