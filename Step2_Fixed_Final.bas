' ===============================================
' ステップ2：文字列削除処理（★最終修正版★）
' 修正日：2025/06/18 - 全角括弧に対応！
' ===============================================

Sub Step2_文字列削除処理()
    ' -----------------------------------------------
    ' ★最終修正：実際のデータは全角括弧（）やった！
    ' （内作）（別注）（全ﾈｼﾞ）（非在庫品）の文字列を削除
    ' 行は削除せず、文字列だけ ""に置換するで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim originalValue As String  ' ★追加：変更前の値を保存
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ2: 文字列削除開始（最終修正版） ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            ' 各セルをチェックして文字列削除
            For i = 1 To lastRow
                For col = 1 To 5  ' A~E列（B列削除後なのでA,C,D,E,Fになってる）
                    originalValue = CStr(.Cells(i, col).Value)  ' ★追加：元の値を保存
                    cellValue = originalValue
                    
                    ' ★最終修正：全角括弧（）に変更
                    ' （内作）を削除
                    If InStr(cellValue, "（内作）") > 0 Then
                        cellValue = Replace(cellValue, "（内作）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' （別注）を削除
                    If InStr(cellValue, "（別注）") > 0 Then
                        cellValue = Replace(cellValue, "（別注）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' （全ﾈｼﾞ）を削除（全角括弧＋半角カタカナ）
                    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then
                        cellValue = Replace(cellValue, "（全ﾈｼﾞ）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' ★新規追加：（非在庫品）を削除
                    If InStr(cellValue, "（非在庫品）") > 0 Then
                        cellValue = Replace(cellValue, "（非在庫品）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' ★追加安全策：半角括弧版も念のため対応
                    ' (内作)を削除
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' (別注)を削除
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' (全ﾈｼﾞ)を削除（半角括弧＋半角カタカナ）
                    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ﾈｼﾞ)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' (非在庫品)を削除
                    If InStr(cellValue, "(非在庫品)") > 0 Then
                        cellValue = Replace(cellValue, "(非在庫品)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' セルの値を更新
                    .Cells(i, col).Value = cellValue
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個の文字列を削除"
    End If
    
    ' Table002の処理
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        replaceCount = 0
        With ws2
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            ' 各セルをチェックして文字列削除
            For i = 1 To lastRow
                For col = 1 To 5  ' A~E列
                    originalValue = CStr(.Cells(i, col).Value)  ' ★追加：元の値を保存
                    cellValue = originalValue
                    
                    ' ★最終修正：全角括弧（）に変更
                    ' （内作）を削除
                    If InStr(cellValue, "（内作）") > 0 Then
                        cellValue = Replace(cellValue, "（内作）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' （別注）を削除
                    If InStr(cellValue, "（別注）") > 0 Then
                        cellValue = Replace(cellValue, "（別注）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' （全ﾈｼﾞ）を削除（全角括弧＋半角カタカナ）
                    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then
                        cellValue = Replace(cellValue, "（全ﾈｼﾞ）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' ★新規追加：（非在庫品）を削除
                    If InStr(cellValue, "（非在庫品）") > 0 Then
                        cellValue = Replace(cellValue, "（非在庫品）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' ★追加安全策：半角括弧版も念のため対応
                    ' (内作)を削除
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' (別注)を削除
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' (全ﾈｼﾞ)を削除（半角括弧＋半角カタカナ）
                    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ﾈｼﾞ)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' (非在庫品)を削除
                    If InStr(cellValue, "(非在庫品)") > 0 Then
                        cellValue = Replace(cellValue, "(非在庫品)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue  ' ★追加：削除ログ
                    End If
                    
                    ' セルの値を更新
                    .Cells(i, col).Value = cellValue
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個の文字列を削除"
    End If
    
    Debug.Print "=== ステップ2完了（最終修正版） ==="
    
End Sub