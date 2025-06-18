' ===============================================
' プロシージャ名：BasicProcessor（基本処理最終修正版）- 全角括弧対応
' 作成者：関西のおばちゃん
' 作成日：2025/06/16
' 修正日：2025/06/18 - 全角括弧（）に対応！実際のデータ分析結果に基づく修正
' 概要：PDF取得後の基本データ整理処理（正しい順番で）
'       【最初の4項目】
'       1. Table001のB列削除
'       2. （内作）（別注）（全ﾈｼﾞ）（非在庫品） → 文字列削除（★全角括弧対応★）
'       3. x → X変換（小文字→大文字）
'       4. カタカナ全角 → 半角変換
'       【その後の処理】
'       5. A~E列空欄行削除
'       6. Table002データ統合
'       7. Table002のテーブル形式解除（★エラー修正済み★）
'       8. Table001 → "原価リスト"に名前変更
' ※変数名は英語、コメントは関西弁で初心者にもわかりやすく♪
' ===============================================

Option Explicit

' ===============================================
' メイン処理：基本処理（最終修正版）
' ===============================================

Sub 基本処理_最終修正版()
    ' -----------------------------------------------
    ' PDF取得後の基本データ整理処理（正しい順番で）
    ' -----------------------------------------------
    
    Dim response As VbMsgBoxResult
    
    ' 実行確認
    response = MsgBox("基本処理（最終修正版）を実行するで?" & vbCrLf & vbCrLf & _
                      "【最初の4項目】" & vbCrLf & _
                      "1. Table001のB列削除" & vbCrLf & _
                      "2. （内作）（別注）（全ﾈｼﾞ）（非在庫品） 文字列削除" & vbCrLf & _
                      "3. x → X変換" & vbCrLf & _
                      "4. カタカナ全角→半角変換" & vbCrLf & vbCrLf & _
                      "【その後の処理】" & vbCrLf & _
                      "5. A~E列空欄行削除" & vbCrLf & _
                      "6. Table002統合" & vbCrLf & _
                      "7. テーブル形式解除" & vbCrLf & _
                      "8. シート名変更" & vbCrLf & vbCrLf & _
                      "実行してもええ？", _
                      vbYesNo + vbQuestion, "基本処理最終修正版")
    
    If response = vbNo Then
        MsgBox "処理をキャンセルしたで♪", vbInformation, "キャンセル"
        Exit Sub
    End If
    
    ' 画面更新を止めて処理を早くする
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    MsgBox "基本処理（最終修正版）を開始するで♪", vbInformation, "処理開始"
    
    ' 【最初の4項目】順番大事やで♪
    Call Step1_B列削除処理
    Call Step2_文字列削除処理_最終版  ' ★最終修正★
    Call Step3_x大文字変換処理
    Call Step4_カタカナ半角変換処理
    
    ' 【その後の処理】
    Call Step5_空白行削除処理
    Call Step6_データ統合処理
    Call Step7_テーブル形式解除処理  ' ★エラー修正済み★
    Call Step8_シート名変更処理
    
    ' 画面更新を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "基本処理（最終修正版）完了♪" & vbCrLf & _
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
' ステップ2：文字列削除処理（★最終版★）
' ===============================================

Sub Step2_文字列削除処理_最終版()
    ' -----------------------------------------------
    ' ★最終修正：実際のデータ分析結果に基づき全角括弧（）に対応！
    ' （内作）（別注）（全ﾈｼﾞ）（非在庫品）の文字列を削除
    ' 安全策として半角括弧()版も同時対応
    ' 行は削除せず、文字列だけ ""に置換するで♪
    ' -----------------------------------------------
    
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow As Long
    Dim i As Long, col As Integer
    Dim cellValue As String
    Dim originalValue As String  ' ★追加：変更前の値を保存
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
            
            ' 各セルをチェックして文字列削除
            For i = 1 To lastRow
                For col = 1 To 5  ' A~E列（B列削除後なのでA,C,D,E,Fになってる）
                    originalValue = CStr(.Cells(i, col).Value)  ' ★追加：元の値を保存
                    cellValue = originalValue
                    
                    ' ★最終修正：メイン対応は全角括弧（）
                    ' （内作）を削除
                    If InStr(cellValue, "（内作）") > 0 Then
                        cellValue = Replace(cellValue, "（内作）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' （別注）を削除
                    If InStr(cellValue, "（別注）") > 0 Then
                        cellValue = Replace(cellValue, "（別注）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' （全ﾈｼﾞ）を削除（全角括弧＋半角カタカナ）
                    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then
                        cellValue = Replace(cellValue, "（全ﾈｼﾞ）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' ★新規追加：（非在庫品）を削除
                    If InStr(cellValue, "（非在庫品）") > 0 Then
                        cellValue = Replace(cellValue, "（非在庫品）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' ★安全策：半角括弧版も念のため対応
                    ' (内作)を削除
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' (別注)を削除
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' (全ﾈｼﾞ)を削除（半角括弧＋半角カタカナ）
                    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ﾈｼﾞ)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' (非在庫品)を削除
                    If InStr(cellValue, "(非在庫品)") > 0 Then
                        cellValue = Replace(cellValue, "(非在庫品)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
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
                    
                    ' ★最終修正：メイン対応は全角括弧（）
                    ' （内作）を削除
                    If InStr(cellValue, "（内作）") > 0 Then
                        cellValue = Replace(cellValue, "（内作）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' （別注）を削除
                    If InStr(cellValue, "（別注）") > 0 Then
                        cellValue = Replace(cellValue, "（別注）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' （全ﾈｼﾞ）を削除（全角括弧＋半角カタカナ）
                    If InStr(cellValue, "（全ﾈｼﾞ）") > 0 Then
                        cellValue = Replace(cellValue, "（全ﾈｼﾞ）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' ★新規追加：（非在庫品）を削除
                    If InStr(cellValue, "（非在庫品）") > 0 Then
                        cellValue = Replace(cellValue, "（非在庫品）", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' ★安全策：半角括弧版も念のため対応
                    ' (内作)を削除
                    If InStr(cellValue, "(内作)") > 0 Then
                        cellValue = Replace(cellValue, "(内作)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' (別注)を削除
                    If InStr(cellValue, "(別注)") > 0 Then
                        cellValue = Replace(cellValue, "(別注)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' (全ﾈｼﾞ)を削除（半角括弧＋半角カタカナ）
                    If InStr(cellValue, "(全ﾈｼﾞ)") > 0 Then
                        cellValue = Replace(cellValue, "(全ﾈｼﾞ)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' (非在庫品)を削除
                    If InStr(cellValue, "(非在庫品)") > 0 Then
                        cellValue = Replace(cellValue, "(非在庫品)", "")
                        replaceCount = replaceCount + 1
                        Debug.Print "削除: " & originalValue & " → " & cellValue
                    End If
                    
                    ' セルの値を更新
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
                    
                    ' 小文字のxを大文字のXに変換
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
                    
                    ' 小文字のxを大文字のXに変換
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
                    
                    ' カタカナを半角に変換（StrConv関数使用）
                    convertedValue = StrConv(cellValue, vbNarrow)
                    
                    ' 変換前と後で違いがあれば更新
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
                    
                    ' カタカナを半角に変換
                    convertedValue = StrConv(cellValue, vbNarrow)
                    
                    ' 変換前と後で違いがあれば更新
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
' ステップ7：テーブル形式解除処理（★修正版★）
' ===============================================

Sub Step7_テーブル形式解除処理()
    ' -----------------------------------------------
    ' Table002のテーブル形式を解除するで♪
    ' ★エラー修正：テーブル名を先に保存してからUnlistするで♪
    ' -----------------------------------------------
    
    Dim ws2 As Worksheet
    Dim tbl As ListObject
    Dim tableName As String  ' ★追加：テーブル名を保存する変数
    Dim tableCount As Integer  ' ★追加：処理したテーブル数をカウント
    Dim i As Integer  ' ★追加：ループ用変数を明示的に宣言
    
    Debug.Print "=== ステップ7: テーブル形式解除開始 ==="
    
    ' Table002シートを取得
    On Error Resume Next
    Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1) ")
    If ws2 Is Nothing Then Set ws2 = ThisWorkbook.Worksheets("Table002 (Page 1)")
    On Error GoTo 0
    
    If Not ws2 Is Nothing Then
        tableCount = 0  ' ★追加：カウンター初期化
        
        ' シート内のすべてのテーブルを解除
        ' ★修正：下から上に向かって処理（テーブル削除でインデックスがズレるのを防ぐ）
        For i = ws2.ListObjects.Count To 1 Step -1
            Set tbl = ws2.ListObjects(i)  ' ★修正：インデックス指定で取得
            
            ' ★修正：テーブル名を先に保存
            tableName = tbl.Name
            Debug.Print "テーブル発見：" & tableName
            
            ' エラーハンドリング追加
            On Error Resume Next
            ' テーブルを通常の範囲に変換
            tbl.Unlist
            
            ' エラーチェック
            If Err.Number = 0 Then
                ' ★修正：保存した名前を使用（tbl.Nameは参照できない）
                Debug.Print "テーブル形式解除完了：" & tableName
                tableCount = tableCount + 1  ' ★追加：成功カウント
            Else
                Debug.Print "テーブル解除失敗：" & tableName & " - エラー: " & Err.Description
                Err.Clear  ' エラーをクリア
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
' 結果確認処理
' ===============================================

Sub 結果確認_最終版()
    ' -----------------------------------------------
    ' 最終版の処理結果を確認
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
        
        msg = "【最終版処理結果確認】" & vbCrLf & vbCrLf
        msg = msg & "♪ シート名：" & .Name & vbCrLf
        msg = msg & "♪ データ行数：" & (lastRow - 1) & " 行（ヘッダー除く）" & vbCrLf
        msg = msg & "♪ 最終行：" & lastRow & " 行目" & vbCrLf
        msg = msg & "♪ B列削除済み" & vbCrLf
        msg = msg & "♪ 文字列変換済み（全角括弧対応済み）" & vbCrLf
        msg = msg & "♪ テーブル形式解除済み（エラー修正済み）" & vbCrLf & vbCrLf
        msg = msg & "基本処理（最終版）が完了したで♪"
        
        MsgBox msg, vbInformation, "処理結果確認"
        Debug.Print msg
    End With
    
End Sub