
Sub test2()

    a = Array(Split("a,i,u,e,o", ","))
    
    jpstr1 = _
        Array( _
            Array("l", "ぁ", "ぃ", "ぅ", "ぇ", "ぉ"), _
            Array("", "あ", "い", "う", "え", "お"), _
            Array("k", "か", "き", "く", "け", "こ"), _
            Array("c", "か", "き", "く", "け", "こ"), _
            Array("g", "が", "ぎ", "ぐ", "げ", "ご"), _
            Array("s", "さ", "し", "す", "せ", "そ"), _
            Array("z", "ざ", "じ", "ず", "ぜ", "ぞ"), _
            Array("t", "た", "ち", "つ", "て", "と"), _
            Array("d", "だ", "ぢ", "づ", "で", "ど"), _
            Array("n", "な", "に", "ぬ", "ね", "の"), _
            Array("h", "は", "ひ", "ふ", "へ", "ほ"), _
            Array("b", "ば", "び", "ぶ", "べ", "ぼ"), _
            Array("p", "ぱ", "ぴ", "ぷ", "ぺ", "ぽ"), _
            Array("m", "ま", "み", "む", "め", "も"), _
            Array("y", "や", "", "ゆ", "", "よ"), _
            Array("ly", "ゃ", "", "ゅ", "", "ょ"), _
            Array("r", "ら", "り", "る", "れ", "ろ"), _
            Array("w", "ら", "り", "る", "れ", "ろ"), _
            Array("q", "ら", "り", "る", "れ", "ろ"), _
            Array("j", "ら", "り", "る", "れ", "ろ"), _
            Array("f", "ら", "り", "る", "れ", "ろ"), _
            Array("v", "ら", "り", "る", "れ", "ろ") _
        )
        
        'w q j f v
'            Array("", "わ", "ゐ", "ゑ", "を", "ん")
'            Array("", "っ", "ゎ", "ヴ", "ヵ", "ヶ")
    jparr2 = Split(",,,,,,ゃ,ぃ,ゅ,ぇ,ょ,ゃ,ぃ,ゅ,ぇ,ょ,ぁ,ぃ,ぅ,ぇ,ぉ", ",")
    rmarr1 = Split(",l,a,k,c,g,s,z,t,d,n,h,b,p,m,y,ly,r,w,q,j,f,v,*", ",")
    rmarr2 = Split(",a,i,u,e,o,ya,yi,yu,ye,yo,ha,hi,hu,he,ho,wa,wi,wu,we,wo", ",")
    
    indexArr = Split(",1,2,3,4,5,2,2,2,2,2,2,4,2,4,2,3,3,5,3,3", ",")
    
    Set romadict = CreateObject("Scripting.Dictionary")
    For i = 0 To 16
        Set d = CreateObject("Scripting.Dictionary")
        For j = 0 To UBound(jparr2)
            If (i = 1 Or i = 2 Or i = 15 Or i = 16) And j >= 6 Then
                d.Add rmarr2(j), ""
            Else
                d.Add rmarr2(j), jparr1(i, indexArr(j)) + jparr2(j)
            End If
        Next
        romadict.Add rmarr1(i), d
    Next
    
    'w q j f v
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(18, 1)
    d.Add rmarr2(7), jparr1(18, 2)
    d.Add rmarr2(9), jparr1(18, 3)
    d.Add rmarr2(5), jparr1(18, 4)
    d.Add rmarr2(2), jparr1(2, 3) + jparr2(17)
    d.Add rmarr2(4), jparr1(2, 3) + jparr2(19)
    d.Add rmarr2(15), jparr1(2, 3) + jparr2(20)
    romadict.Add "w", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(3, 3) + jparr2(16)
    d.Add rmarr2(2), jparr1(3, 3) + jparr2(17)
    d.Add rmarr2(3), jparr1(3, 3) + jparr2(0)
    d.Add rmarr2(4), jparr1(3, 3) + jparr2(19)
    d.Add rmarr2(5), jparr1(3, 3) + jparr2(20)
    romadict.Add "q", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(7, 2) + jparr2(6)
    d.Add rmarr2(2), jparr1(7, 2) + jparr2(0)
    d.Add rmarr2(3), jparr1(7, 2) + jparr2(8)
    d.Add rmarr2(4), jparr1(7, 2) + jparr2(19)
    d.Add rmarr2(5), jparr1(7, 2) + jparr2(10)
    romadict.Add "j", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(11, 3) + jparr2(16)
    d.Add rmarr2(2), jparr1(11, 3) + jparr2(17)
    d.Add rmarr2(3), jparr1(11, 3) + jparr2(0)
    d.Add rmarr2(4), jparr1(11, 3) + jparr2(19)
    d.Add rmarr2(5), jparr1(11, 3) + jparr2(20)
    romadict.Add "f", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(19, 3) + jparr2(16)
    d.Add rmarr2(2), jparr1(19, 3) + jparr2(17)
    d.Add rmarr2(3), jparr1(19, 3) + jparr2(0)
    d.Add rmarr2(4), jparr1(19, 3) + jparr2(19)
    d.Add rmarr2(5), jparr1(19, 3) + jparr2(20)
    romadict.Add "v", d
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nn", jparr1(18, 5)
    d.Add "ltu", jparr1(19, 1)
    d.Add "lwa", jparr1(19, 2)
    d.Add "lka", jparr1(19, 4)
    d.Add "lke", jparr1(19, 5)
    romadict.Add "*", d
        
End Sub
                        
Sub test()
    
    jpstr1 = "ぁぃぅぇぉあいうえおかきくけこかきくけこがぎぐげごさしすせそざじずぜぞたちつてとだぢづでどなにぬねのはひふへほばびぶべぼぱぴぷぺぽまみむめもや　ゆ　よゃ　ゅ　ょらりるれろわゐゑをんっゎヴヵヶ"
    ReDim jparr1(19, 5)
    For i = 0 To 94
        jparr1(i \ 5 + 1, i Mod 5 + 1) = Mid(jpstr1, i + 1, 1)
    Next
    jparr2 = Split(",,,,,,ゃ,ぃ,ゅ,ぇ,ょ,ゃ,ぃ,ゅ,ぇ,ょ,ぁ,ぃ,ぅ,ぇ,ぉ", ",")
    rmarr1 = Split(",l,a,k,c,g,s,z,t,d,n,h,b,p,m,y,ly,r,w,q,j,f,v,*", ",")
    rmarr2 = Split(",a,i,u,e,o,ya,yi,yu,ye,yo,ha,hi,hu,he,ho,wa,wi,wu,we,wo", ",")
    
    indexArr = Split(",1,2,3,4,5,2,2,2,2,2,2,4,2,4,2,3,3,5,3,3", ",")

'    y>iya ii iyu ie iyo
'    h>iya ei iyu ee iyo
'    w> ua ui  ou ue  uo
    
    Set romadict = CreateObject("Scripting.Dictionary")
    For i = 1 To 17
        Set d = CreateObject("Scripting.Dictionary")
        For j = 1 To UBound(jparr2)
            If (i = 1 Or i = 2 Or i = 15 Or i = 16) And j >= 6 Then
                d.Add rmarr2(j), ""
            Else
                d.Add rmarr2(j), jparr1(i, indexArr(j)) + jparr2(j)
            End If
        Next
        romadict.Add rmarr1(i), d
    Next
    
    'w q j f v
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(18, 1)
    d.Add rmarr2(7), jparr1(18, 2)
    d.Add rmarr2(9), jparr1(18, 3)
    d.Add rmarr2(5), jparr1(18, 4)
    d.Add rmarr2(2), jparr1(2, 3) + jparr2(17)
    d.Add rmarr2(4), jparr1(2, 3) + jparr2(19)
    d.Add rmarr2(15), jparr1(2, 3) + jparr2(20)
    romadict.Add "w", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(3, 3) + jparr2(16)
    d.Add rmarr2(2), jparr1(3, 3) + jparr2(17)
    d.Add rmarr2(3), jparr1(3, 3) + jparr2(0)
    d.Add rmarr2(4), jparr1(3, 3) + jparr2(19)
    d.Add rmarr2(5), jparr1(3, 3) + jparr2(20)
    romadict.Add "q", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(7, 2) + jparr2(6)
    d.Add rmarr2(2), jparr1(7, 2) + jparr2(0)
    d.Add rmarr2(3), jparr1(7, 2) + jparr2(8)
    d.Add rmarr2(4), jparr1(7, 2) + jparr2(19)
    d.Add rmarr2(5), jparr1(7, 2) + jparr2(10)
    romadict.Add "j", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(11, 3) + jparr2(16)
    d.Add rmarr2(2), jparr1(11, 3) + jparr2(17)
    d.Add rmarr2(3), jparr1(11, 3) + jparr2(0)
    d.Add rmarr2(4), jparr1(11, 3) + jparr2(19)
    d.Add rmarr2(5), jparr1(11, 3) + jparr2(20)
    romadict.Add "f", d
    Set d = CreateObject("Scripting.Dictionary")
    d.Add rmarr2(1), jparr1(19, 3) + jparr2(16)
    d.Add rmarr2(2), jparr1(19, 3) + jparr2(17)
    d.Add rmarr2(3), jparr1(19, 3) + jparr2(0)
    d.Add rmarr2(4), jparr1(19, 3) + jparr2(19)
    d.Add rmarr2(5), jparr1(19, 3) + jparr2(20)
    romadict.Add "v", d
    
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nn", jparr1(18, 5)
    d.Add "ltu", jparr1(19, 1)
    d.Add "lwa", jparr1(19, 2)
    d.Add "lka", jparr1(19, 4)
    d.Add "lke", jparr1(19, 5)
    romadict.Add "*", d
    
        extraKeys = romadict("*").keys
        extraItems = romadict("*").items
    
    ReDim tmparr(500)
    i = 0
    For Each k1 In romadict
        For Each k2 In romadict(k1)
            If romadict(k1)(k2) <> "" And romadict(k1)(k2) <> "　" Then
                i = i + 1
                tmparr(i) = IIf(k1 = "a" Or k1 = "*", "", k1) + k2 + "=" + romadict(k1)(k2)
            End If
        Next
    Next
    
    strSet = Join(tmparr, vbCrLf)
    
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True '複数行入力可
        .Text = strSet
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With

    End
    
    With Workbooks("タスク管理.xlsm").Worksheets("Sheet1")
    
    'fill table frame
        For i = 1 To UBound(rmarr1)
            .Cells(2 + i, 2) = rmarr1(i)
        Next
        For i = 1 To UBound(rmarr2)
            .Cells(2, 2 + i) = rmarr2(i)
        Next
        
        'fill table value
        On Error Resume Next
        For i = 1 To UBound(rmarr1) - 1
            For j = 1 To UBound(rmarr2)
                .Cells(2 + i, 2 + j) = romadict(rmarr1(i))(rmarr2(j))
            Next
        Next
        extraKeys = romadict("*").keys
        extraItems = romadict("*").items
        For j = LBound(extraKeys) To UBound(extraKeys)
            .Cells(2 + i, 3 + j) = extraKeys(j)
            .Cells(3 + i, 3 + j) = extraItems(j)
        Next
        On Error GoTo 0
        
    End With
    'https://www.detblog.com/windows-10-%E3%81%AE-ms-ime-%E3%81%A7%E3%83%AD%E3%83%BC%E3%83%9E%E5%AD%97%E5%A4%89%E6%8F%9B%E8%A1%A8%E3%82%92%E3%82%AB%E3%82%B9%E3%82%BF%E3%83%9E%E3%82%A4%E3%82%BA%E3%81%99%E3%82%8B-azik/
    'http://jgrammar.life.coocan.jp/ja/tools/imekeys.htm
    
End Sub
