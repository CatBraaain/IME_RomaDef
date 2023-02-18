Sub makeRomaDef()

    Set jpFirstChrArrDict = CreateObject("Scripting.Dictionary")
    jpFirstChrArrDict.Add "l", Split(",ぁ,ぃ,ぅ,ぇ,ぉ", ",")
    jpFirstChrArrDict.Add "", Split(",あ,い,う,え,お", ",")
    jpFirstChrArrDict.Add "k", Split(",か,き,く,け,こ", ",")
    jpFirstChrArrDict.Add "c", Split(",か,き,く,け,こ", ",")
    jpFirstChrArrDict.Add "g", Split(",が,ぎ,ぐ,げ,ご", ",")
    jpFirstChrArrDict.Add "s", Split(",さ,し,す,せ,そ", ",")
    jpFirstChrArrDict.Add "z", Split(",ざ,じ,ず,ぜ,ぞ", ",")
    jpFirstChrArrDict.Add "t", Split(",た,ち,つ,て,と", ",")
    jpFirstChrArrDict.Add "d", Split(",だ,ぢ,づ,で,ど", ",")
    jpFirstChrArrDict.Add "n", Split(",な,に,ぬ,ね,の", ",")
    jpFirstChrArrDict.Add "h", Split(",は,ひ,ふ,へ,ほ", ",")
    jpFirstChrArrDict.Add "b", Split(",ば,び,ぶ,べ,ぼ", ",")
    jpFirstChrArrDict.Add "p", Split(",ぱ,ぴ,ぷ,ぺ,ぽ", ",")
    jpFirstChrArrDict.Add "m", Split(",ま,み,む,め,も", ",")
    jpFirstChrArrDict.Add "y", Split(",や,,ゆ,いぇ,よ", ",")
    jpFirstChrArrDict.Add "ly", Split(",ゃ,,ゅ,,ょ", ",")
    jpFirstChrArrDict.Add "r", Split(",ら,り,る,れ,ろ", ",")
    jpFirstChrArrDict.Add "w", Split(",わ,うぃ,,うぇ,を", ",")
    jpFirstChrArrDict.Add "q", Split(",くぁ,くぃ,く,くぇ,くぉ", ",")
    jpFirstChrArrDict.Add "j", Split(",じゃ,じ,じゅ,じぇ,じょ", ",")
    jpFirstChrArrDict.Add "f", Split(",ふぁ,ふぃ,ふ,ふぇ,ふぉ", ",")
    jpFirstChrArrDict.Add "v", Split(",ヴぁ,ヴぃ,ヴ,ヴぇ,ヴぉ", ",")
    
    enFirstChrArr = jpFirstChrArrDict.Keys
    enSecondChrArr = Split("a,i,u,e,o,ya,yi,yu,ye,yo,ha,hi,hu,he,ho,wa,wi,wu,we,wo", ",")
    
    jpSecondChrArr = Split(",,,,,ゃ,ぃ,ゅ,ぇ,ょ,ゃ,ぃ,ゅ,ぇ,ょ,ぁ,ぃ,ぅ,ぇ,ぉ", ",")
    
    jpFirstChrIndexArr = Split("1,2,3,4,5,2,2,2,2,2,2,4,2,4,2,3,3,5,3,3", ",")
    
    Set romaDict = CreateObject("Scripting.Dictionary")
    For Each enFirstChr In jpFirstChrArrDict
        Set jpChrDict = CreateObject("Scripting.Dictionary")
        For i = 0 To 4
            jpFirstChrArr = jpFirstChrArrDict(enFirstChr)
            jpChrDict.Add enSecondChrArr(i), jpFirstChrArr(jpFirstChrIndexArr(i)) + jpSecondChrArr(i)
        Next
        If InStr("あぁやゃわくぁじゃふぁヴぁ", jpFirstChrArrDict(enFirstChr)(1)) = 0 Then
            For i = 5 To 19
                jpFirstChrArr = jpFirstChrArrDict(enFirstChr)
                jpChrDict.Add enSecondChrArr(i), jpFirstChrArr(jpFirstChrIndexArr(i)) + jpSecondChrArr(i)
            Next
        End If
        romaDict.Add enFirstChr, jpChrDict
    Next
    
'    jpUniqueChrArr = Split("うぉ,ん, ゐ, ゑ, っ, ゎ, ヵ, ヶ", ",")
    Set jpUniqueChrDict = CreateObject("Scripting.Dictionary")
    jpUniqueChrDict.Add "who", "うぉ"
    jpUniqueChrDict.Add "nn", "ん"
    jpUniqueChrDict.Add "wyi", "ゐ"
    jpUniqueChrDict.Add "wye", "ゑ"
    jpUniqueChrDict.Add "ltu", "っ"
    jpUniqueChrDict.Add "lwa", "ゎ"
    jpUniqueChrDict.Add "lka", "ヵ"
    jpUniqueChrDict.Add "lke", "ヶ"
    romaDict.Add "*", jpUniqueChrDict
    
    
    'コピー
    ReDim tmparr(310)
    i = 0
    For Each k1 In romaDict
        For Each k2 In romaDict(k1)
            If romaDict(k1)(k2) <> "" Then
                i = i + 1
                tmparr(i) = IIf(k1 = "*", "", k1) + k2 + "=" + romaDict(k1)(k2)
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
    
    
    '表示
    With ActiveSheet
        'fill table frame
        For i = 0 To UBound(enFirstChrArr)
            .Cells(3 + i, 2) = enFirstChrArr(i)
        Next
        For i = 0 To UBound(enSecondChrArr)
            .Cells(2, 3 + i) = enSecondChrArr(i)
        Next
        
        'fill table value
        On Error Resume Next
        For i = 0 To UBound(enFirstChrArr)
            For j = 0 To UBound(enSecondChrArr)
                .Cells(3 + i, 3 + j) = romaDict(enFirstChrArr(i))(enSecondChrArr(j))
            Next
        Next
        On Error GoTo 0
        
        uniqueKeys = jpUniqueChrDict.Keys
        uniqueItems = jpUniqueChrDict.Items
        For j = 0 To UBound(uniqueKeys)
            .Cells(4 + i + j, 2) = uniqueKeys(j)
            .Cells(4 + i + j, 3) = uniqueItems(j)
        Next
        
    End With
    
    'https://www.detblog.com/windows-10-%E3%81%AE-ms-ime-%E3%81%A7%E3%83%AD%E3%83%BC%E3%83%9E%E5%AD%97%E5%A4%89%E6%8F%9B%E8%A1%A8%E3%82%92%E3%82%AB%E3%82%B9%E3%82%BF%E3%83%9E%E3%82%A4%E3%82%BA%E3%81%99%E3%82%8B-azik/
    'http://jgrammar.life.coocan.jp/ja/tools/imekeys.htm
    
End Sub
