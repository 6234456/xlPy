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
    
    py.print_ py.len_("qiou")
    py.print_ py.len_(d.fromString("[1, 2,3, 4, [], {}]"))
    
    py.print_ py.eval("[1,2,3, {'12': 23, '234': 'qiou'}, 0]")
    
    py.print_ py.Reversed("qiou.eu").join("")
    
  '  py.print_ py.walk("D:\Books\FSharp\Expert FSahrp 3.0.pdf", 0)
    
  '  py.print_ py.dir_("D:\Books")
    
   ' py.print_ py.mkdir_("D:\Books\1\2")
   
    With py.open_("demo.txt", "a")
        .Write "World"
        .Close
    End With
    
   ' py.print_ py.request("GET", "http://wenshu.court.gov.cn/website/wenshu/js/wenshulist1.js")
   
   With py.open_("C:\Users\User\Downloads\data.txt", "rb")
        
       ' Debug.Print Mid$(Trim(.readall), 147)
        ' ## parseJSON between the closing } and ] can not contains blank
        py.eval(.readall).p

       .Close
   End With

End Sub
