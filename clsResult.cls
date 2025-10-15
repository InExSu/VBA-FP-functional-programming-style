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

' Only public methods and properties

Public Property Get value() As Variant
    If Not m_isSuccess Then Err.Raise 513, , "Error: attempt to read Value when in error state"
    value = m_value
End Property

Public Property Get Error() As String
    Error = m_error
End Property

Public Property Get IsSuccess() As Boolean
    IsSuccess = m_isSuccess
End Property

' Internal initialization — not intended for direct external use.
' Made Public only for factory functions; do not call directly.
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

' Bind chains operations; skips execution if already in error state.
Public Function Bind(ByVal fnName As String) As clsResult
    If Not m_isSuccess Then
        Set Bind = Me
        Exit Function
    End If

    On Error GoTo Crash
    Dim nextResult As Object
    Set nextResult = Application.Run(fnName, m_value)

    If TypeName(nextResult) <> "clsResult" Then
        Set Bind = ResultErr("Bind: function '" & fnName & "' did not return clsResult")
        Exit Function
    End If

    Set Bind = nextResult
    Exit Function

Crash:
    Set Bind = ResultErr("Bind: error in '" & fnName & "' — " & Err.Description)
End Function

' Returns the value if success, otherwise returns the provided default.
Public Function GetOrElse(ByVal defaultValue As Variant) As Variant
    If m_isSuccess Then
        GetOrElse = m_value
    Else
        GetOrElse = defaultValue
    End If
End Function
