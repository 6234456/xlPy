 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     implement Python-API with VBA
'@author                                   Qiou Yang
'@license                                  MIT
'@dependency                               Lists, Nodes, TreeSets
'@lastUpdate                               30.12.2019
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' TypeError 9

Private d As Dicts
Private l As Lists


Private Sub Class_Initialize()
    Set l = New Lists
    Set d = New Dicts
End Sub

Private Sub Class_Terminate()
    Set l = Nothing
    Set d = Nothing
End Sub

Public Function abs_(ByVal num) As Double
    abs_ = Abs(num)
End Function


Public Function all(ByVal iter) As Boolean

    If Not isIterable(iter) Then
        Err.Raise 9, , "TypeError: object '" & TypeName(iter) & "' is not iterable"
    End If
    
    Dim i
    Dim res As Boolean
    res = True
    
    For Each i In iter
        If Not bool(i) Then
            res = False
            Exit For
        End If
    Next i
    
    all = res
    
End Function

Public Function any_(ByVal iter) As Boolean

    If Not isIterable(iter) Then
        Err.Raise 9, , "TypeError:" & TypeName(iter) & " is not iterable"
    End If
    
    Dim i
    Dim res As Boolean
    res = False
    
    For Each i In iter
        If Not bool(i) Then
            res = True
            Exit For
        End If
    Next i
    
    any_ = res
    
End Function

Public Function bin(ByVal num As Integer) As String
    If num > 0 Then
        Dim res As String
        res = ""
        
        Do While num > 0
            res = (num Mod 2) & res
            num = (num - (num Mod 2)) / 2
        Loop
        
        bin = "0b" & res
    ElseIf num = 0 Then
        bin = "0b0"
    Else
        bin = "-0b" & bin(-num)
    End If
End Function

Public Function bool(ByVal x) As Boolean

    bool = True
    
    If IsObject(x) Then
        bool = x Is Nothing
    Else
        If IsNull(x) Or x = 0 Or x = False Then
            bool = False
        End If
    End If

End Function

Public Function divmod(ByVal a As Integer, ByVal b As Integer) As Variant
    divmod = Array((a - (a Mod b)) / b, a Mod b)
End Function

Public Function enumerate(ByVal iter, Optional ByVal start As Integer = 0) As Dicts
    
    If Not isIterable(iter) Then
        Err.Raise 9, , "TypeError:" & TypeName(iter) & " is not iterable"
    End If
    
    If IsArray(iter) Or TypeName(iter) = "Collection" Then
        Set enumerate = enumerate(l.fromArray(iter))
    ElseIf TypeName(iter) = "Lists" Then
        Set enumerate = iter.toMap
    ElseIf TypeName(iter) = "Dicts" Then
        Set enumerate = enumerate(iter.keysArr)
    Else
        Err.Raise 10, , "Method enumerate unimplemented for object '" & TypeName(iter) & "'"
    End If
    
End Function

Public Function len_(ByVal val As Variant) As Integer
    
    

End Function

Public Function toCharArr(ByVal s As String) As Lists
   Set toCharArr = l.fromString(s)
End Function

Public Function print_(ByVal val As Variant)
    Debug.Print toString(val)
End Function

Private Function toString(ByVal arr As Variant) As String
    Dim d As New Dicts
    toString = d.x_toString(arr)
    Set d = Nothing
End Function

Private Function str2Arr(ByVal s As String) As Variant

    

End Function

Public Function isIterable(ByVal v As Variant) As Boolean
    On Error GoTo errHdl
    
    If IsArray(v) Or TypeName(v) = "Dicts" Or TypeName(v) = "Lists" Or TypeName(v) = "Collection" Or TypeName(v) = "TreeSets" Then
        isIterable = True
    Else
        Dim i

        For Each i In v
            isIterable = True
            Exit Function
        Next
    End If
errHdl:
End Function

Public Function ascii(ByVal v As String) As Integer
    ascii = Asc(v)
End Function


Public Function chr_(ByVal v As Integer) As Integer
    chr_ = Chr(v)
End Function

Public Function range_(ByVal param1, Optional ByVal param2, Optional ByVal param3 As Integer = 1) As Lists
    
    ' last element
    Dim e
    ' !! isMissing with parameter without pre-defined type
    If IsMissing(param2) Then
        Set range_ = l.fromSerial(0, param1)
        e = param1
    Else
        Set range_ = l.fromSerial(param1, param2, param3)
        e = param2
    End If
    
    If l.length > 0 Then
        If l.last = e Then
            Set range_ = l.dropLast(1)
        End If
    End If

End Function
