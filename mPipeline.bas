Attribute VB_Name = "mPipeline"
Option Explicit

' Тестовая цепочка
Public Sub RunChain()

    Dim res As clsResult
    Set res = ResultOk(5) _
        .Bind("MyFunc1") _
        .Bind("MyFunc2")
    
    If res.IsSuccess Then
        Debug.Print "? " & res.value
    Else
        Debug.Print "? " & res.Error
    End If
    
End Sub

' Функция 1: проверяет, что число положительное, и умножает на 10
Public Function MyFunc1(value As Variant) As clsResult
    If Not IsNumeric(value) Then
        Set MyFunc1 = ResultErr("MyFunc1: вход не является числом")
        Exit Function
    End If
    
    Dim num As Double
    num = CDbl(value)
    
    If num <= 0 Then
        Set MyFunc1 = ResultErr("MyFunc1: число должно быть > 0")
        Exit Function
    End If
    
    Set MyFunc1 = ResultOk(num * 10)
End Function


' Функция 2: преобразует число в строку с префиксом "Result: "
Public Function MyFunc2(value As Variant) As clsResult
    If Not IsNumeric(value) Then
        Set MyFunc2 = ResultErr("MyFunc2: ожидается число")
        Exit Function
    End If
    
    Set MyFunc2 = ResultOk("Result: " & CStr(CDbl(value)))
End Function

'=== Фабрики ===
' Фабрика успеха
Public Function ResultOk(ByVal value As Variant) As clsResult
    Dim r As New clsResult
    r.InitOk value
    Set ResultOk = r
End Function

' Фабрика ошибки
Public Function ResultErr(ByVal errorMsg As String) As clsResult
    Dim r As New clsResult
    r.InitErr errorMsg
    Set ResultErr = r
End Function

