 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'@desc                                     implement Python-API with VBA
'@author                                   Qiou Yang
'@license                                  MIT
'@dependency                               Lists, Nodes, TreeSets
'@lastUpdate                               25.01.2020
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' TypeError 9
' ValueError 5

Private d As Dicts
Private l As Lists
Private fso As Object
Private pSep As String ' the path separator of the system


Private Sub Class_Initialize()
    Set l = New Lists
    Set d = New Dicts
    Set fso = CreateObject("scripting.filesystemobject")
    pSep = "\"
    ChDir ThisWorkbook.path
End Sub

Private Sub Class_Terminate()
    Set l = Nothing
    Set d = Nothing
    Set fso = Nothing
End Sub

Public Property Get sep() As String
    sep = pSep
End Property

Public Property Let sep(ByVal v As String)
    pSep = v
End Property

Public Property Let wd(ByVal v As String)
    ChDir v
End Property

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
    Dim tmp As String
    tmp = TypeName(val)
    
    If IsArray(val) Then
        len_ = UBound(val) - LBound(val) + 1
    ElseIf tmp = "Lists" Then
        len_ = val.length
    ElseIf tmp = "Dicts" Or tmp = "Collection" Then
        len_ = val.Count
    ElseIf tmp = "String" Then
        len_ = Len(val)
    Else
        Err.Raise 9, , "TypeError: object of type '" & tmp & "' has no len()"
    End If
End Function

Public Function toCharArr(ByVal s As String) As Lists
   Set toCharArr = l.fromString(s)
End Function

Public Function print_(ByVal val As Variant)
    Debug.Print repr(val)
End Function

Public Function eval(ByVal s As String)
    If IsObject(d.fromString(s)) Then
        Set eval = d.fromString(s)
    Else
        eval = d.fromString(s)
    End If
End Function

Public Function repr(ByVal arr As Variant) As String
    repr = d.x_toString(arr)
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

Public Function Reversed(ByRef val As Variant) As Lists
    Dim tmp As String
    tmp = TypeName(val)
    
    If IsArray(val) Or tmp = "Collection" Then
        Set Reversed = l.fromArray(val).reverse()
    ElseIf tmp = "Lists" Then
        Set Reversed = val.reverse()
    ElseIf tmp = "Dicts" Then
        Set Reversed = Reversed(d.keysArr)
    ElseIf tmp = "String" Then
        Set Reversed = Reversed(l.fromString(val))
    Else
        Err.Raise 9, , "TypeError: object of type '" & tmp & "' has no reversed()"
    End If
End Function

' list all the files
Public Function walk(ByVal path As String, Optional ByVal recursive As Boolean = False) As Lists

    If fso.FolderExists(path) Then
        Set walk = l.fromArray(fso.getfolder(path).Files)
        
        If recursive Then
            Dim e
            For Each e In fso.getfolder(path).SubFolders
                Set walk = walk.addAll(Me.walk(fso.buildpath(path, e.Name), True))
            Next e
        End If
    ElseIf fso.FileExists(path) Then
        Set walk = l.fromArray(Array(fso.getfile(path)))
    Else
        Set walk = l.fromArray(Array())
    End If
   
End Function

' list the directory structure
Public Function dir_(ByVal path As String)
    Dim d2 As New Dicts
    
    
    If fso.FolderExists(path) Then
        Dim l2 As Lists
        Set l2 = Me.walk(path)
        
        Dim e
        For Each e In fso.getfolder(path).SubFolders
            l2.add Me.dir_(fso.buildpath(path, e.Name))
        Next e
        
        d2.add path, l2
    
        Set dir_ = d2
    ElseIf fso.FileExists(path) Then
        Set dir_ = Me.walk(path)
    Else
        Set dir_ = d2
    End If
    
    Set d2 = Nothing
End Function

' create directory recursively
Public Function mkdir_(ByVal path As String)
    Dim path_ As String
    path_ = fso.GetAbsolutePathName(path)
    If Not fso.FolderExists(path_) Then
        Dim tmpS As String
        tmpS = l.fromArray(Split(path_, pSep)).dropLast(1).join(pSep)
            
        If Not fso.FolderExists(tmpS) Then
            mkdir_ tmpS
        End If
        
        fso.CreateFolder path
    End If
End Function

' return a TextStream Object
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/textstream-object
' binary use Write() and Read()
Public Function open_(ByVal file As String, Optional ByVal mode As String = "r") As Object
    If InStr(mode, "b") > 0 And InStr(mode, "t") > 0 Then
        Err.Raise 5, , "ValueError: can't have text and binary mode at once"
    End If
    
    If Len(mode) > 3 Then
        Err.Raise 5, , "ValueError: invalid mode: '" & mode & "'"
    End If
    
    Dim formatNumber As Integer
    formatNumber = IIf(InStr(mode, "b"), -2, -1)
    
    Dim hasMode1 As Integer
    hasMode1 = 0
    
    Dim hasMode2 As Integer
    hasMode2 = 0
    
    Dim hasMode3 As Integer
    hasMode3 = 0
    
    Dim e
    For Each e In Me.toCharArr(mode).toArray
       If e = "a" Or e = "w" Or e = "x" Or e = "r" Then
            hasMode1 = hasMode1 + 1
       ElseIf e = "b" Or e = "t" Then
            hasMode2 = hasMode2 + 1
       ElseIf e = "+" Then
            hasMode3 = hasMode3 + 1
       Else
            Err.Raise 5, , "ValueError: invalid mode: '" & mode & "'"
       End If
    Next e
    
    If hasMode1 <> 1 Then
         Err.Raise 5, , "must have exactly one of create/read/write/append mode"
    End If
    
    If hasMode2 > 1 Or hasMode3 > 1 Then
        Err.Raise 5, , "ValueError: invalid mode: '" & mode & "'"
    End If
    
    If InStr(mode, "w") > 0 Or InStr(mode, "x") > 0 Or InStr(mode, "+") > 0 Then
        Dim path_ As String
        path_ = fso.GetAbsolutePathName(file)
        
        If Not fso.FileExists(path_) Then
            Dim tmpS As String
            tmpS = l.fromArray(Split(path_, pSep)).dropLast(1).join(pSep)
            mkdir_ tmpS
            
            Set open_ = fso.CreateTextFile(path_, Unicode:=1)
        ElseIf InStr(mode, "x") > 0 Then
            Err.Raise 17, , "FileExistsError: File exists: '" & path_ & "'"
        ElseIf InStr(mode, "w") > 0 Or InStr(mode, "+") > 0 Then
            Set open_ = fso.getfile(path_).OpenAsTextStream(iomode:=2, Format:=formatNumber)
        End If
    ElseIf InStr(mode, "r") > 0 Then
        Set open_ = fso.getfile(file).OpenAsTextStream(iomode:=1, Format:=formatNumber)
    ElseIf InStr(mode, "a") > 0 Then
        Set open_ = fso.getfile(file).OpenAsTextStream(iomode:=8, Format:=formatNumber)
    End If
    
End Function
