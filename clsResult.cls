VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsResult
Option Explicit

Private m_value As Variant
Private m_error As String
Private m_isSuccess As Boolean

' Только Public методы и свойства

Public Property Get value() As Variant
    If Not m_isSuccess Then Err.Raise 513, , "Ошибка: попытка чтения Value при ошибке"
    value = m_value
End Property

Public Property Get Error() As String
    Error = m_error
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = m_isSuccess
End Property

' Внутренняя инициализация — но НЕ вызывается извне!
' Мы сделаем её Public, но договоримся её не использовать напрямую
Public Sub InitOk(ByVal value As Variant)
    m_value = value
    m_error = ""
    m_isSuccess = True
End Sub

Public Sub InitErr(ByVal errorMsg As String)
    m_value = Empty
    m_error = errorMsg
    m_isSuccess = False
End Sub

' Bind остаётся тот же
Public Function Bind(ByVal fnName As String) As clsResult
    If Not m_isSuccess Then
        Set Bind = Me
        Exit Function
    End If

    On Error GoTo Crash
    Dim nextResult As Object
    Set nextResult = Application.Run(fnName, m_value)

    If TypeName(nextResult) <> "clsResult" Then
        Set Bind = ResultErr("Bind: функция '" & fnName & "' не вернула clsResult")
        Exit Function
    End If

    Set Bind = nextResult
    Exit Function
Crash:
    Set Bind = ResultErr("Bind: ошибка в '" & fnName & "' — " & Err.Description)
End Function

Public Function GetOrElse(ByVal defaultValue As Variant) As Variant
    If m_isSuccess Then
        GetOrElse = m_value
    Else
        GetOrElse = defaultValue
    End If
End Function
