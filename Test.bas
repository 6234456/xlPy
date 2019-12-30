Sub test()
    
    Dim py As New xlPy
    Dim d As New Dicts
    Dim l As New Lists
    
    Debug.Assert py.bool("")
    Debug.Assert py.bool(-1)
    Debug.Assert Not py.bool(d)
    Debug.Assert Not py.bool(Null)
    
   ' Debug.Assert Not py.all(Array(1, 2, 3, 4, Array()))
    ' Debug.Print py.all(1)  ' throw TypeError
    Debug.Assert Not py.isIterable(1)
    
    py.print_ py.range_(-5)
    py.print_ py.range_(-10, -5)
    
    py.print_ py.range_(-1, 6, 2)
    
    py.print_ py.enumerate(py.toCharArr("Qiou Yang"))

End Sub
