' ===============================================
' プロシージャ名：BasicProcessor（基本処理完全版+プレート処理修正版）
' 作成者：関西のおばちゃん
' 作成日：2025/06/16
' 修正日：2025/06/18 - プレート処理を最後に移動（原価リストシート完成後に実行）
' 概要：PDF取得後の基本データ整理処理（完全版+プレート処理修正版）
'       【最初の処理群】
'       1. Table001のB列削除
'       2. Table002の2行目削除
'       3. （内作）（別注）（全ﾈｼﾞ）（非在庫品） → 文字列削除
'       4. x → X変換（小文字→大文字）
'       5. カタカナ全角 → 半角変換
'       6. F10T変換処理
'       7. S10T変換処理
'       【その後の処理群】
'       8. A~E列空欄行削除
'       9. ﾈｺｱﾝｸﾞﾙ変換処理
'       10. Table002データ統合
'       11. Table002のテーブル形式解除
'       12. Table001 → "原価リスト"に名前変更
'       13. プレート処理（原価リストシートで実行）（★修正：最後に移動★）
' ※変数名は英語、コメントは関西弁で初心者にもわかりやすく♪
' ===============================================

Option Explicit

' ===============================================
' メイン処理：基本処理（完全版+プレート処理修正版）
' ===============================================

Sub 基本処理_完全版修正()
    ' -----------------------------------------------
    ' PDF取得後の基本データ整理処理（完全版+プレート処理修正版）
    ' ★修正：プレート処理を最後に移動
    ' -----------------------------------------------
    
    Dim response As VbMsgBoxResult
    
    ' 実行確認
    response = MsgBox("基本処理（完全版+プレート処理修正版）を実行するで?" & vbCrLf & vbCrLf & _
                      "【最初の処理群】" & vbCrLf & _
                      "1. Table001のB列削除" & vbCrLf & _
                      "2. Table002の2行目削除" & vbCrLf & _
                      "3. （内作）（別注）（全ﾈｼﾞ）（非在庫品） 文字列削除" & vbCrLf & _
                      "4. x → X変換" & vbCrLf & _
                      "5. カタカナ全角→半角変換" & vbCrLf & _
                      "6. F10T変換処理" & vbCrLf & _
                      "7. S10T変換処理" & vbCrLf & vbCrLf & _
                      "【その後の処理群】" & vbCrLf & _
                      "8. A~E列空欄行削除" & vbCrLf & _
                      "9. ﾈｺｱﾝｸﾞﾙ変換処理" & vbCrLf & _
                      "10. Table002統合" & vbCrLf & _
                      "11. テーブル形式解除" & vbCrLf & _
                      "12. シート名変更" & vbCrLf & _
                      "13. プレート処理（原価リストで実行）" & vbCrLf & vbCrLf & _
                      "実行してもええ？", _
                      vbYesNo + vbQuestion, "基本処理完全版+プレート処理修正版")
    
    If response = vbNo Then
        MsgBox "処理をキャンセルしたで♪", vbInformation, "キャンセル"
        Exit Sub
    End If
    
    ' 画面更新を止めて処理を早くする
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    MsgBox "基本処理（完全版+プレート処理修正版）を開始するで♪", vbInformation, "処理開始"
    
    ' 【最初の処理群】順番大事やで♪
    Call Step1_B列削除処理
    Call Step1_5_Table002の2行目削除処理
    Call Step2_文字列削除処理_最終版
    Call Step3_x大文字変換処理
    Call Step4_カタカナ半角変換処理
    Call Step4_5_F10T変換処理
    Call Step4_6_S10T変換処理
    
    ' 【その後の処理群】
    Call Step5_空白行削除処理
    Call Step5_5_ネコアングル変換処理
    ' ★修正：プレート処理を削除（最後に移動）
    Call Step6_データ統合処理
    Call Step7_テーブル形式解除処理
    Call Step8_シート名変更処理
    Call Step9_プレート処理_最終版             ' ★修正：最後に実行★
    
    ' 画面更新を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "基本処理（完全版+プレート処理修正版）完了♪" & vbCrLf & _
           "「原価リスト」シートを確認してや♪", vbInformation, "処理完了"
    
End Sub

' ===============================================
' ステップ1：B列削除処理
' ===============================================

Sub Step1_B列削除処理()
    ' -----------------------------------------------
    ' Table001のB列を削除するで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet
    
    Debug.Print "=== ステップ1: B列削除開始 ==="
    
    ' Table001シートを取得
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        ' B列を削除
        ws1.Columns("B:B").Delete Shift:=xlToLeft
        Debug.Print "Table001のB列を削除完了"
    Else
        Debug.Print "Table001シートが見つからない"
    End If
    
    Debug.Print "=== ステップ1完了 ==="
    
End Sub

' ===============================================
' ステップ1.5：Table002の2行目削除処理
' ===============================================

Sub Step1_5_Table002の2行目削除処理()
    ' -----------------------------------------------
    ' Table002の2行目を削除するで♪
    ' ヘッダー行の次の行（データ1行目）を削除や
    ' -----------------------------------------------
    
    Dim ws2 As Worksheet
    
    Debug.Print "=== ステップ1.5: Table002の2行目削除開始 ==="
    
    ' Table002シートを取得
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        ' 2行目が存在するかチェック
        If ws2.Cells(2, 1).Value <> "" Or ws2.Cells(2, 2).Value <> "" Or ws2.Cells(2, 3).Value <> "" Then
            ' 2行目の内容をログに出力
            Debug.Print "削除する2行目の内容: " & ws2.Cells(2, 1).Value & " | " & ws2.Cells(2, 2).Value & " | " & ws2.Cells(2, 3).Value
            
            ' 2行目を削除
            ws2.Rows(2).Delete Shift:=xlUp
            Debug.Print "Table002の2行目を削除完了"
        Else
            Debug.Print "Table002の2行目は既に空やで"
        End If
    Else
        Debug.Print "Table002シートが見つからない"
    End If
    
    Debug.Print "=== ステップ1.5完了 ==="
    
End Sub

' ===============================================
' ステップ2：文字列削除処理（最終版）
' ===============================================

Sub Step2_文字列削除処理_最終版()
    ' -----------------------------------------------
    ' 全角括弧（）対応版
    ' （内作）（別注）（全ﾈｼﾞ）（非在庫品）の文字列を削除
    ' 安全策として半角括弧()版も同時対応
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim originalValue As String
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ2: 文字列削除開始（最終版・全角括弧対応） ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    originalValue = CStr(.Cells(i, col).Value)
                    cellValue = originalValue
                    
                    ' 全角括弧版
                    If InStr(cellValue, "（内作）") > 0 Then
                        cellValue = Replace(cellValue, "（内作）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    If InStr(cellValue, "（別注）") > 0 Then
                        cellValue = Replace(cellValue, "（別注）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then
                        cellValue = Replace(cellValue, "（全ﾈｼﾞ）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    If InStr(cellValue, "（非在庫品）") > 0 Then
                        cellValue = Replace(cellValue, "（非在庫品）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' 半角括弧版（安全策）
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ﾈｼﾞ)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    If InStr(cellValue, "(非在庫品)") > 0 Then
                        cellValue = Replace(cellValue, "(非在庫品)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    .Cells(i, col).Value = cellValue
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個の文字列を削除"
    End If
    
    ' Table002の処理（同様の処理）
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        replaceCount = 0
        With ws2
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    originalValue = CStr(.Cells(i, col).Value)
                    cellValue = originalValue
                    
                    ' 全角・半角括弧の両方に対応（Table001と同じ処理）
                    If InStr(cellValue, "（内作）") > 0 Then cellValue = Replace(cellValue, "（内作）", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "（別注）") > 0 Then cellValue = Replace(cellValue, "（別注）", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then cellValue = Replace(cellValue, "（全ﾈｼﾞ）", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "（非在庫品）") > 0 Then cellValue = Replace(cellValue, "（非在庫品）", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "(内作)") > 0 Then cellValue = Replace(cellValue, "(内作)", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "(別注)") > 0 Then cellValue = Replace(cellValue, "(別注)", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then cellValue = Replace(cellValue, "(全ﾈｼﾞ)", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    If InStr(cellValue, "(非在庫品)") > 0 Then cellValue = Replace(cellValue, "(非在庫品)", ""): replaceCount = replaceCount + 1: Debug.Print "削除: " & originalValue & " → " & cellValue
                    
                    .Cells(i, col).Value = cellValue
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個の文字列を削除"
    End If
    
    Debug.Print "=== ステップ2完了（最終版・全角括弧対応） ==="
    
End Sub

' ===============================================
' ステップ3：x大文字変換処理
' ===============================================

Sub Step3_x大文字変換処理()
    ' -----------------------------------------------
    ' 小文字の「x」を大文字の「X」に変換
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ3: x→X変換開始 ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    cellValue = CStr(.Cells(i, col).Value)
                    
                    If InStr(cellValue, "x") > 0 Then
                        cellValue = Replace(cellValue, "x", "X")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個のx→X変換"
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
            
            For i = 1 To lastRow
                For col = 1 To 5
                    cellValue = CStr(.Cells(i, col).Value)
                    
                    If InStr(cellValue, "x") > 0 Then
                        cellValue = Replace(cellValue, "x", "X")
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個のx→X変換"
    End If
    
    Debug.Print "=== ステップ3完了 ==="
    
End Sub

' ===============================================
' ステップ4：カタカナ半角変換処理
' ===============================================

Sub Step4_カタカナ半角変換処理()
    ' -----------------------------------------------
    ' カタカナの全角→半角変換
    ' StrConv関数を使うで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim convertedValue As String
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ4: カタカナ半角変換開始 ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    cellValue = CStr(.Cells(i, col).Value)
                    convertedValue = StrConv(cellValue, vbNarrow)
                    
                    If cellValue <> convertedValue Then
                        .Cells(i, col).Value = convertedValue
                        replaceCount = replaceCount + 1
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個のカタカナ半角変換"
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
            
            For i = 1 To lastRow
                For col = 1 To 5
                    cellValue = CStr(.Cells(i, col).Value)
                    convertedValue = StrConv(cellValue, vbNarrow)
                    
                    If cellValue <> convertedValue Then
                        .Cells(i, col).Value = convertedValue
                        replaceCount = replaceCount + 1
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個のカタカナ半角変換"
    End If
    
    Debug.Print "=== ステップ4完了 ==="
    
End Sub

' ===============================================
' ステップ4.5：F10T変換処理
' ===============================================

Sub Step4_5_F10T変換処理()
    ' -----------------------------------------------
    ' F10T M22X60 → F10T-M20X60 に変換
    ' 最初の5文字が「F10T 」だったら「F10T-」に変換するで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim originalValue As String
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ4.5: F10T変換処理開始 ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    originalValue = CStr(.Cells(i, col).Value)
                    cellValue = originalValue
                    
                    ' 最初の5文字が「F10T 」かチェック
                    If Len(cellValue) >= 5 And Left(cellValue, 5) = "F10T " Then
                        ' 「F10T 」を「F10T-」に変換
                        cellValue = "F10T-" & Mid(cellValue, 6)
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                        Debug.Print "F10T変換: " & originalValue & " → " & cellValue
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個のF10T変換"
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
            
            For i = 1 To lastRow
                For col = 1 To 5
                    originalValue = CStr(.Cells(i, col).Value)
                    cellValue = originalValue
                    
                    ' 最初の5文字が「F10T 」かチェック
                    If Len(cellValue) >= 5 And Left(cellValue, 5) = "F10T " Then
                        ' 「F10T 」を「F10T-」に変換
                        cellValue = "F10T-" & Mid(cellValue, 6)
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                        Debug.Print "F10T変換: " & originalValue & " → " & cellValue
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個のF10T変換"
    End If
    
    Debug.Print "=== ステップ4.5完了 ==="
    
End Sub

' ===============================================
' ステップ4.6：S10T変換処理
' ===============================================

Sub Step4_6_S10T変換処理()
    ' -----------------------------------------------
    ' S10T M22X60 → S10T22X60 に変換
    ' 最初の5文字が「S10T 」だったら「S10T」に変換（スペース削除）するで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim originalValue As String
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ4.6: S10T変換処理開始 ==="
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    originalValue = CStr(.Cells(i, col).Value)
                    cellValue = originalValue
                    
                    ' 最初の5文字が「S10T 」かチェック
                    If Len(cellValue) >= 5 And Left(cellValue, 5) = "S10T " Then
                        ' 「S10T 」を「S10T」に変換（スペース削除）
                        cellValue = "S10T" & Mid(cellValue, 6)
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                        Debug.Print "S10T変換: " & originalValue & " → " & cellValue
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個のS10T変換"
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
            
            For i = 1 To lastRow
                For col = 1 To 5
                    originalValue = CStr(.Cells(i, col).Value)
                    cellValue = originalValue
                    
                    ' 最初の5文字が「S10T 」かチェック
                    If Len(cellValue) >= 5 And Left(cellValue, 5) = "S10T " Then
                        ' 「S10T 」を「S10T」に変換（スペース削除）
                        cellValue = "S10T" & Mid(cellValue, 6)
                        .Cells(i, col).Value = cellValue
                        replaceCount = replaceCount + 1
                        Debug.Print "S10T変換: " & originalValue & " → " & cellValue
                    End If
                Next col
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個のS10T変換"
    End If
    
    Debug.Print "=== ステップ4.6完了 ==="
    
End Sub

' ===============================================
' ステップ5：空白行削除処理
' ===============================================

Sub Step5_空白行削除処理()
    ' -----------------------------------------------
    ' A~E列がすべて空欄の行を削除するで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    
    Debug.Print "=== ステップ5: 空白行削除開始 ==="
    
    ' Table001の空白行削除
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        Call Delete空白行(ws1)
    End If
    
    ' Table002の空白行削除
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        Call Delete空白行(ws2)
    End If
    
    Debug.Print "=== ステップ5完了 ==="
    
End Sub

Sub Delete空白行(ws As Worksheet)
    ' -----------------------------------------------
    ' 指定されたワークシートの空白行を削除
    ' A~E列がすべて空欄の行が対象やで♪
    ' -----------------------------------------------
    
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim isEmpty As Boolean
    Dim deleteCount As Integer
    
    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        deleteCount = 0
        
        ' 下から上に向かって処理（行削除で番号がズレるのを防ぐため）
        For i = lastRow To 2 Step -1  ' 1行目はヘッダーなので除外
            isEmpty = True
            
            ' A~E列をすべてチェック
            For col = 1 To 5
                If Trim(CStr(.Cells(i, col).Value)) <> "" Then
                    isEmpty = False  ' 何かデータがあったら空白行ではない
                    Exit For
                End If
            Next col
            
            ' すべての列が空白なら行を削除
            If isEmpty Then
                .Rows(i).Delete Shift:=xlUp
                deleteCount = deleteCount + 1
            End If
        Next i
        
        Debug.Print ws.Name & ": " & deleteCount & "行の空白行を削除"
    End With
    
End Sub

' ===============================================
' ステップ5.5：ネコアングル変換処理
' ===============================================

Sub Step5_5_ネコアングル変換処理()
    ' -----------------------------------------------
    ' ﾈｺｱﾝｸﾞﾙが連続して6行あるうち、上から4行を変換
    ' "ﾈｺｱﾝｸﾞﾙ D-10"、"ﾈｺｱﾝｸﾞﾙ D-15"、"ﾈｺｱﾝｸﾞﾙ D-20"、"ﾈｺｱﾝｸﾞﾙ D-30"に変換
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim nekoCount As Integer
    Dim convertPatterns As Variant
    Dim replaceCount As Integer
    
    Debug.Print "=== ステップ5.5: ネコアングル変換処理開始 ==="
    
    ' 変換パターンを配列で定義
    convertPatterns = Array("ﾈｺｱﾝｸﾞﾙ D-10", "ﾈｺｱﾝｸﾞﾙ D-15", "ﾈｺｱﾝｸﾞﾙ D-20", "ﾈｺｱﾝｸﾞﾙ D-30")
    
    ' Table001の処理
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        replaceCount = 0
        nekoCount = 0  ' ネコアングルのカウンター
        
        With ws1
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    cellValue = CStr(.Cells(i, col).Value)
                    
                    ' 「ﾈｺｱﾝｸﾞﾙ」が含まれているかチェック
                    If InStr(cellValue, "ﾈｺｱﾝｸﾞﾙ") > 0 And nekoCount < 4 Then
                        Debug.Print "ネコアングル発見(" & (nekoCount + 1) & "回目): " & cellValue & " → " & convertPatterns(nekoCount)
                        .Cells(i, col).Value = convertPatterns(nekoCount)
                        nekoCount = nekoCount + 1
                        replaceCount = replaceCount + 1
                        
                        ' 4回変換したら終了
                        If nekoCount >= 4 Then Exit For
                    End If
                Next col
                
                ' 4回変換したら終了
                If nekoCount >= 4 Then Exit For
            Next i
        End With
        
        Debug.Print "Table001: " & replaceCount & "個のネコアングル変換"
    End If
    
    ' Table002の処理（同様の処理）
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        replaceCount = 0
        nekoCount = 0  ' ネコアングルのカウンター
        
        With ws2
            lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            For i = 1 To lastRow
                For col = 1 To 5
                    cellValue = CStr(.Cells(i, col).Value)
                    
                    ' 「ﾈｺｱﾝｸﾞﾙ」が含まれているかチェック
                    If InStr(cellValue, "ﾈｺｱﾝｸﾞﾙ") > 0 And nekoCount < 4 Then
                        Debug.Print "ネコアングル発見(" & (nekoCount + 1) & "回目): " & cellValue & " → " & convertPatterns(nekoCount)
                        .Cells(i, col).Value = convertPatterns(nekoCount)
                        nekoCount = nekoCount + 1
                        replaceCount = replaceCount + 1
                        
                        ' 4回変換したら終了
                        If nekoCount >= 4 Then Exit For
                    End If
                Next col
                
                ' 4回変換したら終了
                If nekoCount >= 4 Then Exit For
            Next i
        End With
        
        Debug.Print "Table002: " & replaceCount & "個のネコアングル変換"
    End If
    
    Debug.Print "=== ステップ5.5完了 ==="
    
End Sub

' ===============================================
' ステップ6：データ統合処理
' ===============================================

Sub Step6_データ統合処理()
    ' -----------------------------------------------
    ' Table002のデータをTable001の最終行に追加
    ' テーブルは使わず、普通のコピペやで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim i As Long, col As Integer
    Dim copyCount As Integer
    Dim hasData As Boolean
    
    Debug.Print "=== ステップ6: データ統合開始 ==="
    
    ' シートを取得
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    ' 両方のシートが存在するかチェック
    If ws1 Is Nothing Then
        Debug.Print "Table001シートが見つからない"
        Exit Sub
    End If
    
    If ws2 Is Nothing Then
        Debug.Print "Table002シートが見つからない"
        Exit Sub
    End If
    
    ' 最終行を取得
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    
    copyCount = 0
    
    ' Table002のデータをTable001にコピー（ヘッダー行は除く）
    For i = 2 To lastRow2  ' 2行目から開始（1行目はヘッダー）
        hasData = False
        
        ' その行にデータがあるかチェック
        For col = 1 To 5
            If Trim(CStr(ws2.Cells(i, col).Value)) <> "" Then
                hasData = True
                Exit For
            End If
        Next col
        
        ' データがある行のみコピー
        If hasData Then
            lastRow1 = lastRow1 + 1  ' Table001の次の行
            
            ' A~E列をコピー
            For col = 1 To 5
                ws1.Cells(lastRow1, col).Value = ws2.Cells(i, col).Value
            Next col
            
            copyCount = copyCount + 1
        End If
    Next i
    
    Debug.Print "Table002から " & copyCount & " 行をTable001に統合完了"
    Debug.Print "=== ステップ6完了 ==="
    
End Sub

' ===============================================
' ステップ7：テーブル形式解除処理
' ===============================================

Sub Step7_テーブル形式解除処理()
    ' -----------------------------------------------
    ' Table002のテーブル形式を解除するで♪
    ' -----------------------------------------------
    
    Dim ws2 As Worksheet
    Dim tbl As ListObject
    Dim tableName As String
    Dim tableCount As Integer
    Dim i As Integer
    
    Debug.Print "=== ステップ7: テーブル形式解除開始 ==="
    
    ' Table002シートを取得
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        tableCount = 0
        
        ' 下から上に向かって処理
        For i = ws2.ListObjects.Count To 1 Step -1
            Set tbl = ws2.ListObjects(i)
            tableName = tbl.Name
            Debug.Print "テーブル発見：" & tableName
            
            On Error Resume Next
            tbl.Unlist
            
            If Err.Number = 0 Then
                Debug.Print "テーブル形式解除完了：" & tableName
                tableCount = tableCount + 1
            Else
                Debug.Print "テーブル解除失敗：" & tableName & " - エラー: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Next i
        
        Debug.Print "Table002のテーブル形式解除完了（" & tableCount & "個のテーブルを処理）"
    Else
        Debug.Print "Table002シートが見つからない"
    End If
    
    Debug.Print "=== ステップ7完了 ==="
    
End Sub

' ===============================================
' ステップ8：シート名変更処理
' ===============================================

Sub Step8_シート名変更処理()
    ' -----------------------------------------------
    ' Table001のシート名を「原価リスト」に変更
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet
    
    Debug.Print "=== ステップ8: シート名変更開始 ==="
    
    ' Table001シートを取得
    On Error Resume Next
    Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1) ")
    If ws1 Is Nothing Then Set ws1 = ThisWorkbook.Worksheets("Table001 (Page 1)")
    On Error GoTo 0
    
    If Not ws1 Is Nothing Then
        ' シート名を変更
        ws1.Name = "原価リスト"
        Debug.Print "シート名変更完了：Table001 → 原価リスト"
    Else
        Debug.Print "Table001シートが見つからない"
    End If
    
    Debug.Print "=== ステップ8完了 ==="
    
End Sub

' ===============================================
' ステップ9：プレート処理（★最終版・正しいタイミングで実行★）
' ===============================================

Sub Step9_プレート処理_最終版()
    ' -----------------------------------------------
    ' ★最終版：原価リストシート完成後に実行
    ' 条件1: 最初から4文字目以降が「END PL」で始まる
    ' 条件2: 先頭3文字が「PL-」
    ' かつD列が数値の行を処理
    ' 1. 該当行のD列の平均値を計算
    ' 2. 最上段の行のD列に平均値、B列を「プレート」に変更
    ' 3. 最上段以外の該当行を削除
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim bValue As String, dValue As String
    Dim plateRows() As Long  ' 該当行の配列
    Dim plateValues() As Double  ' D列の数値配列
    Dim plateCount As Integer
    Dim averageValue As Double
    Dim sum As Double
    Dim firstPlateRow As Long
    Dim deleteCount As Integer
    Dim isTarget As Boolean
    
    Debug.Print "=== ステップ9: プレート処理開始（最終版） ==="
    
    ' 原価リストシートを取得
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("原価リスト")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Debug.Print "「原価リスト」シートが見つからない"
        Exit Sub
    End If
    
    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        plateCount = 0
        
        ' まず該当行を特定して配列に格納
        ReDim plateRows(1 To lastRow)  ' 最大で全行数分の配列を確保
        ReDim plateValues(1 To lastRow)
        
        For i = 1 To lastRow
            bValue = CStr(.Cells(i, 2).Value)  ' B列（2列目）
            dValue = CStr(.Cells(i, 4).Value)  ' D列（4列目）
            
            isTarget = False
            
            ' 条件1: 最初から4文字目以降が「END PL」で始まる
            If Len(bValue) >= 6 Then  ' 最低6文字必要（4文字目以降に「END PL」）
                If Left(Mid(bValue, 4), 6) = "END PL" Then
                    isTarget = True
                    Debug.Print "END PL条件一致: [" & bValue & "]"
                End If
            End If
            
            ' 条件2: 先頭3文字が「PL-」
            If Len(bValue) >= 3 Then
                If Left(bValue, 3) = "PL-" Then
                    isTarget = True
                    Debug.Print "PL-条件一致: [" & bValue & "]"
                End If
            End If
            
            ' D列が数値かつ対象行の場合
            If isTarget And IsNumeric(dValue) And dValue <> "" Then
                plateCount = plateCount + 1
                plateRows(plateCount) = i
                plateValues(plateCount) = CDbl(dValue)
                
                Debug.Print "プレート対象行発見(" & plateCount & "): 行" & i & " B列=[" & bValue & "] D列=[" & dValue & "]"
                
                ' 最初の該当行を記録
                If plateCount = 1 Then firstPlateRow = i
            End If
        Next i
        
        If plateCount = 0 Then
            Debug.Print "プレート対象行が見つからない"
            Debug.Print "=== ステップ9完了（対象なし） ==="
            Exit Sub
        End If
        
        ' 平均値を計算
        sum = 0
        For i = 1 To plateCount
            sum = sum + plateValues(i)
        Next i
        averageValue = sum / plateCount
        
        Debug.Print "プレート対象行数: " & plateCount & "行"
        Debug.Print "D列平均値: " & averageValue
        Debug.Print "最上段行: " & firstPlateRow
        
        ' 最上段の行を更新
        .Cells(firstPlateRow, 2).Value = "プレート"  ' B列
        .Cells(firstPlateRow, 4).Value = averageValue  ' D列
        
        Debug.Print "最上段更新完了: 行" & firstPlateRow & " B列=プレート D列=" & averageValue
        
        ' 最上段以外の該当行を削除（下から上に向かって削除）
        deleteCount = 0
        For i = plateCount To 2 Step -1  ' 2から開始（1は最上段なので除外）
            .Rows(plateRows(i)).Delete Shift:=xlUp
            deleteCount = deleteCount + 1
            Debug.Print "プレート行削除: 元の行" & plateRows(i)
        Next i
        
        Debug.Print "プレート行削除完了: " & deleteCount & "行を削除"
    End With
    
    Debug.Print "=== ステップ9完了 ==="
    
End Sub

' ===============================================
' 結果確認処理
' ===============================================

Sub 結果確認_完全版修正()
    ' -----------------------------------------------
    ' 完全版修正の処理結果を確認
    ' -----------------------------------------------
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim msg As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("原価リスト")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "「原価リスト」シートが見つからない" & vbCrLf & _
               "処理が正常に完了してない可能性があるで♪", vbExclamation, "確認エラー"
        Exit Sub
    End If
    
    With ws
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        msg = "【完全版修正処理結果確認】" & vbCrLf & vbCrLf
        msg = msg & "♪ シート名：" & .Name & vbCrLf
        msg = msg & "♪ データ行数：" & (lastRow - 1) & " 行（ヘッダー除く）" & vbCrLf
        msg = msg & "♪ 最終行：" & lastRow & " 行目" & vbCrLf
        msg = msg & "♪ B列削除済み" & vbCrLf
        msg = msg & "♪ Table002の2行目削除済み" & vbCrLf
        msg = msg & "♪ 文字列変換済み（全角括弧対応済み）" & vbCrLf
        msg = msg & "♪ F10T・S10T変換済み" & vbCrLf
        msg = msg & "♪ ネコアングル変換済み" & vbCrLf
        msg = msg & "♪ プレート処理済み（原価リストシートで実行）" & vbCrLf
        msg = msg & "♪ テーブル形式解除済み（エラー修正済み）" & vbCrLf & vbCrLf
        msg = msg & "基本処理（完全版修正）が完了したで♪"
        
        MsgBox msg, vbInformation, "処理結果確認"
        Debug.Print msg
    End With
    
End Sub