# excel_macro

```vb
Sub セル内改行連番()
'// https://vbabeginner.net/vbaでセル内の改行ごとに連番を付ける/
    Dim r           As Range        '// 対象セル
    Dim s           As String       '// セル文字列
    Dim v                           '// 改行分割後の配列
    Dim sDiv                        '// 改行分割後の配列の要素
    Dim sConnect    As String       '// 再連結後文字列
    Dim i, ii       As Long         '// ループカウンタ
    Dim reg         As RegExp       '// 正規表現クラス
    Dim vSpace                      '// 空白分割後の配列
    Dim sDivSpace                   '// 空白分割後の配列の要素
    Dim reMatch
    
    ii = 1
    '// 選択セル範囲をループ
    For Each r In Selection
        s = r.Value
        
        '// セル内の改行文字で分割
        v = Split(s, vbLf)
        sConnect = ""
        i = 1
        
        '// セル内の改行文字の分割数ループ
        For Each sDiv In v
            '// 半角スペースが含まれる場合
            If (InStr(1, sDiv, " ") > 0) Then
                '// 半角スペースで分割
                vSpace = Split(sDiv, " ")
                Set reg = New RegExp
        '        reg.Pattern = "[0-9]+\."
                reg.Pattern = "[0-9]+"
                reg.Global = True
                
                Set reMatch = reg.Execute(vSpace(0))
                If (i = 1 And ii = 1 And reMatch.Count > 0) Then ii = Val(reMatch(0).Value)
                '// すでに番号がついている場合の連番振り直し用の処理
                '// 半角スペースで区切られた一番左の要素の数値＋.を除去
                vSpace(0) = reg.Replace(vSpace(0), "")
                sDiv = ""
                
                '// 連番を除去したあとに再度連結
                For Each sDivSpace In vSpace
                    sDiv = sDiv & " " & sDivSpace
                Next
            End If
            
            '// 左右の空白を除去
            sDiv = Trim(sDiv)
            
            '// 初回ループでない場合はセル内改行を付与
            If (i <> 1) Then
                sConnect = sConnect & vbLf
            End If
            
            '// 現在行に文字列が設定されている場合
            If (sDiv <> "") Then
                '// 連番を付与
                sConnect = sConnect & CStr(ii) & " " & sDiv
'//                sConnect = sConnect & CStr(i) & ". " & sDiv
                '// 連番を１増やす
                i = i + 1
                ii = ii + 1
            End If
        Next
        
        '// セル文字列を再設定
        r.Value = sConnect
    Next
End Sub

Sub セル内改行連番削除()
'// https://vbabeginner.net/vbaでセル内の改行ごとに連番を付ける/
    Dim r           As Range        '// 対象セル
    Dim s           As String       '// セル文字列
    Dim v                           '// 改行分割後の配列
    Dim sDiv                        '// 改行分割後の配列の要素
    Dim sConnect    As String       '// 再連結後文字列
    Dim i, ii       As Long         '// ループカウンタ
    Dim reg         As RegExp       '// 正規表現クラス
    Dim vSpace                      '// 空白分割後の配列
    Dim sDivSpace                   '// 空白分割後の配列の要素
    Dim reMatch
    
    ii = 1
    '// 選択セル範囲をループ
    For Each r In Selection
        s = r.Value
        
        '// セル内の改行文字で分割
        v = Split(s, vbLf)
        sConnect = ""
        i = 1
        
        '// セル内の改行文字の分割数ループ
        For Each sDiv In v
            '// 半角スペースが含まれる場合
            If (InStr(1, sDiv, " ") > 0) Then
                '// 半角スペースで分割
                vSpace = Split(sDiv, " ")
                Set reg = New RegExp
        '        reg.Pattern = "[0-9]+\."
                reg.Pattern = "[0-9]+"
                reg.Global = True
                
                Set reMatch = reg.Execute(vSpace(0))
                If (i = 1 And ii = 1 And reMatch.Count > 0) Then ii = Val(reMatch(0).Value)
                '// すでに番号がついている場合の連番振り直し用の処理
                '// 半角スペースで区切られた一番左の要素の数値＋.を除去
                vSpace(0) = reg.Replace(vSpace(0), "")
                sDiv = ""
                
                '// 連番を除去したあとに再度連結
                For Each sDivSpace In vSpace
                    sDiv = sDiv & " " & sDivSpace
                Next
            End If
            
            '// 左右の空白を除去
            sDiv = Trim(sDiv)
            
            '// 初回ループでない場合はセル内改行を付与
            If (i <> 1) Then
                sConnect = sConnect & vbLf
            End If
            
            '// 現在行に文字列が設定されている場合
            If (sDiv <> "") Then
                '// 連番を付与
'                sConnect = sConnect & CStr(ii) & " " & sDiv
'//                sConnect = sConnect & CStr(i) & ". " & sDiv
                '// 連番を１増やす
                i = i + 1
'                ii = ii + 1
            
                sConnect = sConnect & sDiv
            End If
        Next
        
        '// セル文字列を再設定
        r.Value = sConnect
    Next
End Sub

```

