' ===============================================
' 診断用コード：セルの文字列を詳しく調べる
' 作成者：関西のおばちゃん
' 作成日：2025/06/18
' 概要：(内作)(別注)(全ﾈｼﾞ)(非在庫品)が削除されない原因を調べる
' 使い方：削除したい文字が入ってるセルを選択してから実行
' ===============================================

Option Explicit

' ===============================================
' メイン診断処理
' ===============================================

Sub 文字列診断_詳細版()
    ' -----------------------------------------------
    ' 選択したセルの文字列を1文字ずつ詳しく調べるで♪
    ' 削除対象の文字が本当に入ってるかチェックや！
    ' -----------------------------------------------
    
    Dim cellValue As String
    Dim i As Integer
    Dim char As String
    Dim asciiCode As Integer
    
    ' 選択されたセルの値を取得
    If Selection.Cells.Count > 1 Then
        MsgBox "1つのセルだけ選択してや♪", vbExclamation, "エラー"
        Exit Sub
    End If
    
    cellValue = CStr(Selection.Value)
    
    If Len(cellValue) = 0 Then
        MsgBox "選択したセルは空やで〜", vbInformation, "診断結果"
        Exit Sub
    End If
    
    Debug.Print "=========================================="
    Debug.Print "【文字列診断結果】"
    Debug.Print "セル位置: " & Selection.Address
    Debug.Print "セル内容: [" & cellValue & "]"
    Debug.Print "文字数: " & Len(cellValue)
    Debug.Print "=========================================="
    
    ' 1文字ずつ詳しく分析
    For i = 1 To Len(cellValue)
        char = Mid(cellValue, i, 1)
        asciiCode = Asc(char)
        
        Debug.Print i & "文字目: [" & char & "] ASCII: " & asciiCode & " HEX: " & Hex(asciiCode)
        
        ' 特殊文字の判定
        Select Case asciiCode
            Case 40: Debug.Print "  → 半角開き括弧 ("
            Case 41: Debug.Print "  → 半角閉じ括弧 )"
            Case 65288: Debug.Print "  → 全角開き括弧 （"
            Case 65289: Debug.Print "  → 全角閉じ括弧 ）"
            Case 12402: Debug.Print "  → 全角カタカナ ネ"
            Case 65437: Debug.Print "  → 半角カタカナ ﾈ"
            Case 12472: Debug.Print "  → 全角カタカナ ジ"
            Case 65438: Debug.Print "  → 半角カタカナ ｼ"
            Case 65440: Debug.Print "  → 半角カタカナ ﾞ（濁点）"
        End Select
    Next i
    
    Debug.Print "=========================================="
    
    ' 削除対象文字列があるかチェック
    Call Check削除対象文字列(cellValue)
    
    MsgBox "診断完了！イミディエイトウィンドウ（Ctrl+G）で結果を確認してや♪", vbInformation, "診断完了"
    
End Sub

' ===============================================
' 削除対象文字列のチェック
' ===============================================

Sub Check削除対象文字列(cellValue As String)
    ' -----------------------------------------------
    ' 削除対象の文字列があるかチェックして報告
    ' -----------------------------------------------
    
    Debug.Print "【削除対象文字列チェック】"
    
    ' (内作) - 半角括弧
    If InStr(cellValue, "(内作)") > 0 Then
        Debug.Print "✓ (内作) が見つかりました！（半角括弧）"
    Else
        Debug.Print "✗ (内作) は見つかりません（半角括弧）"
    End If
    
    ' （内作） - 全角括弧
    If InStr(cellValue, "（内作）") > 0 Then
        Debug.Print "✓ （内作） が見つかりました！（全角括弧）"
    Else
        Debug.Print "✗ （内作） は見つかりません（全角括弧）"
    End If
    
    ' (別注) - 半角括弧
    If InStr(cellValue, "(別注)") > 0 Then
        Debug.Print "✓ (別注) が見つかりました！（半角括弧）"
    Else
        Debug.Print "✗ (別注) は見つかりません（半角括弧）"
    End If
    
    ' （別注） - 全角括弧
    If InStr(cellValue, "（別注）") > 0 Then
        Debug.Print "✓ （別注） が見つかりました！（全角括弧）"
    Else
        Debug.Print "✗ （別注） は見つかりません（全角括弧）"
    End If
    
    ' (全ﾈｼﾞ) - 半角括弧＋半角カタカナ
    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then
        Debug.Print "✓ (全ﾈｼﾞ) が見つかりました！（半角括弧＋半角カタカナ）"
    Else
        Debug.Print "✗ (全ﾈｼﾞ) は見つかりません（半角括弧＋半角カタカナ）"
    End If
    
    ' （全ﾈｼﾞ） - 全角括弧＋半角カタカナ
    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then
        Debug.Print "✓ （全ﾈｼﾞ） が見つかりました！（全角括弧＋半角カタカナ）"
    Else
        Debug.Print "✗ （全ﾈｼﾞ） は見つかりません（全角括弧＋半角カタカナ）"
    End If
    
    ' (全ネジ) - 半角括弧＋全角カタカナ
    If InStr(cellValue, "(全ネジ)") > 0 Then
        Debug.Print "✓ (全ネジ) が見つかりました！（半角括弧＋全角カタカナ）"
    Else
        Debug.Print "✗ (全ネジ) は見つかりません（半角括弧＋全角カタカナ）"
    End If
    
    ' （全ネジ） - 全角括弧＋全角カタカナ
    If InStr(cellValue, "（全ネジ）") > 0 Then
        Debug.Print "✓ （全ネジ） が見つかりました！（全角括弧＋全角カタカナ）"
    Else
        Debug.Print "✗ （全ネジ） は見つかりません（全角括弧＋全角カタカナ）"
    End If
    
    ' (非在庫品) - 半角括弧
    If InStr(cellValue, "(非在庫品)") > 0 Then
        Debug.Print "✓ (非在庫品) が見つかりました！（半角括弧）"
    Else
        Debug.Print "✗ (非在庫品) は見つかりません（半角括弧）"
    End If
    
    ' （非在庫品） - 全角括弧
    If InStr(cellValue, "（非在庫品）") > 0 Then
        Debug.Print "✓ （非在庫品） が見つかりました！（全角括弧）"
    Else
        Debug.Print "✗ （非在庫品） は見つかりません（全角括弧）"
    End If
    
End Sub

' ===============================================
' 範囲診断：シート全体の削除対象文字列を探す
' ===============================================

Sub 範囲診断_削除対象文字列検索()
    ' -----------------------------------------------
    ' Table001とTable002の全体から削除対象文字列を探すで♪
    ' どこにどんな文字列があるか一覧表示や！
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim foundCount As Integer
    
    Debug.Print "=========================================="
    Debug.Print "【範囲診断：削除対象文字列検索】"
    Debug.Print "=========================================="
    
    foundCount = 0
    
    ' Table001の検索
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        Debug.Print "【" & ws1.Name & " の検索結果】"
        lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
        
        For i = 1 To lastRow
            For col = 1 To 5
                cellValue = CStr(ws1.Cells(i, col).Value)
                
                ' 削除対象文字列があるかチェック
                If InStr(cellValue, "(内作)") > 0 Or _
                   InStr(cellValue, "（内作）") > 0 Or _
                   InStr(cellValue, "(別注)") > 0 Or _
                   InStr(cellValue, "（別注）") > 0 Or _
                   InStr(cellValue, "(全ﾈｼﾞ)") > 0 Or _
                   InStr(cellValue, "（全ﾈｼﾞ）") > 0 Or _
                   InStr(cellValue, "(全ネジ)") > 0 Or _
                   InStr(cellValue, "（全ネジ）") > 0 Or _
                   InStr(cellValue, "(非在庫品)") > 0 Or _
                   InStr(cellValue, "（非在庫品）") > 0 Then
                   
                    Debug.Print "発見！ " & ws1.Cells(i, col).Address & ": [" & cellValue & "]"
                    foundCount = foundCount + 1
                End If
            Next col
        Next i
    End If
    
    ' Table002の検索（シート名変更前）
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        Debug.Print "【" & ws2.Name & " の検索結果】"
        lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
        
        For i = 1 To lastRow
            For col = 1 To 5
                cellValue = CStr(ws2.Cells(i, col).Value)
                
                ' 削除対象文字列があるかチェック
                If InStr(cellValue, "(内作)") > 0 Or _
                   InStr(cellValue, "（内作）") > 0 Or _
                   InStr(cellValue, "(別注)") > 0 Or _
                   InStr(cellValue, "（別注）") > 0 Or _
                   InStr(cellValue, "(全ﾈｼﾞ)") > 0 Or _
                   InStr(cellValue, "（全ﾈｼﾞ）") > 0 Or _
                   InStr(cellValue, "(全ネジ)") > 0 Or _
                   InStr(cellValue, "（全ネジ）") > 0 Or _
                   InStr(cellValue, "(非在庫品)") > 0 Or _
                   InStr(cellValue, "（非在庫品）") > 0 Then
                   
                    Debug.Print "発見！ " & ws2.Cells(i, col).Address & ": [" & cellValue & "]"
                    foundCount = foundCount + 1
                End If
            Next col
        Next i
    End If
    
    ' 「原価リスト」シートの検索（シート名変更後）
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("原価リスト")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        Debug.Print "【原価リスト の検索結果】"
        lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
        
        For i = 1 To lastRow
            For col = 1 To 5
                cellValue = CStr(ws1.Cells(i, col).Value)
                
                ' 削除対象文字列があるかチェック
                If InStr(cellValue, "(内作)") > 0 Or _
                   InStr(cellValue, "（内作）") > 0 Or _
                   InStr(cellValue, "(別注)") > 0 Or _
                   InStr(cellValue, "（別注）") > 0 Or _
                   InStr(cellValue, "(全ﾈｼﾞ)") > 0 Or _
                   InStr(cellValue, "（全ﾈｼﾞ）") > 0 Or _
                   InStr(cellValue, "(全ネジ)") > 0 Or _
                   InStr(cellValue, "（全ネジ）") > 0 Or _
                   InStr(cellValue, "(非在庫品)") > 0 Or _
                   InStr(cellValue, "（非在庫品）") > 0 Then
                   
                    Debug.Print "発見！ " & ws1.Cells(i, col).Address & ": [" & cellValue & "]"
                    foundCount = foundCount + 1
                End If
            Next col
        Next i
    End If
    
    Debug.Print "=========================================="
    Debug.Print "合計 " & foundCount & " 個の削除対象文字列が見つかりました"
    Debug.Print "=========================================="
    
    MsgBox "範囲診断完了！" & vbCrLf & "合計 " & foundCount & " 個の削除対象文字列が見つかりました" & vbCrLf & vbCrLf & "詳細はイミディエイトウィンドウ（Ctrl+G）で確認してや♪", vbInformation, "診断完了"
    
End Sub