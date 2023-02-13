
Sub test()

    chrstr = "ぁぃぅぇぉあいうえおかきくけこがぎぐげごさしすせそざじずぜぞたちつてとだぢづでどなにぬねのはひふへほばびぶべぼぱぴぷぺぽまみむめもや　ゆ　よゃ　ゅ　ょらりるれろわゐゑをんっゎヴヵヶ"
    ReDim chrarr(90)
    For i = 1 To 90
        chrarr(i) = Mid(chrstr, i, 1)
    Next
    schrarr = Split(",ぁ,ぃ,ぅ,ぇ,ぉ,ゃ,ゅ,ょ", ",")
    
    consoarr = Split(",l,,[k|c],g,s,z,t,d,n,h,b,p,m,y,ly,r,w,", ",")
    vowelarr = Split(",a,i,u,e,o", ",")
    
    
    For i = 1 To 18
        For j = 1 To 5
            chrindex = (i - 1) * 5 + j
            mychr = chrarr(chrindex)
            
            Cells(chrindex * 2, 2) = mychr
            For k = 1 To 8
                Cells(chrindex * 2, k + 2) = mychr & schrarr(k)
            Next
            
            Cells(chrindex * 2 + 1, 2) = consoarr(i) & vowelarr(j)
            For k = 1 To 8
                Cells(chrindex * 2 + 1, k + 2) = "-"
            Next
            If j = 2 Then
                Cells(chrindex * 2 + 1, 4) = consoarr(i) & "yi"
                Cells(chrindex * 2 + 1, 6) = consoarr(i) & "ye"
                Cells(chrindex * 2 + 1, 8) = consoarr(i) & "ya," & consoarr(i) & "ha"
                Cells(chrindex * 2 + 1, 9) = consoarr(i) & "yu," & consoarr(i) & "hu"
                Cells(chrindex * 2 + 1, 10) = consoarr(i) & "yo," & consoarr(i) & "ho"
            End If
            If j = 3 Then
                Cells(chrindex * 2 + 1, 3) = consoarr(i) & "wa"
                Cells(chrindex * 2 + 1, 4) = consoarr(i) & "wi"
                Cells(chrindex * 2 + 1, 6) = consoarr(i) & "we"
                Cells(chrindex * 2 + 1, 7) = consoarr(i) & "wo"
            End If
            If j = 4 Then
                Cells(chrindex * 2 + 1, 4) = consoarr(i) & "hi"
                Cells(chrindex * 2 + 1, 6) = consoarr(i) & "he"
            End If
            If j = 5 Then
                Cells(chrindex * 2 + 1, 5) = consoarr(i) & "wu"
            End If
            
        Next
    Next
    
'    WFJ対応
    
'    y>iya ii iyu ie iyo
'    h>iya ei iyu ee iyo
'    w> ua ui  ou ue  uo

    
    
End Sub
