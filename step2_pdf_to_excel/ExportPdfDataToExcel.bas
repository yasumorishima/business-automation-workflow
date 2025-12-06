' *******************************************************************************
' Excel書き出しVBAマクロ（高速化/情報処理・保存してクローズ版）
' Power Queryの更新完了を待ち、新規件数を表示し、ファイルを保存して閉じます。
' ★機能: 指定されたシート名（=クエリ名）と保存先を設定して実行してください。
' *******************************************************************************

Sub ExportPdfDataToJson()
    
    ' --- 設定 ---
    ' Power Queryが出力したデータがあるシート名（クエリ名と同じである必要があります）
    Const TARGET_SHEET_NAME As String = "CombinedPDFData"
    
    ' Excelファイル名保存したいファイル名のフルパス
    Const OUTPUT_FOLDER_PATH As String = "C:\Users\YourName\Documents\ExcelOutput\"
    
    ' データが開始する行番号（ヘッダーを除いて2行目から）
    Const START_ROW As Long = 2
    
    ' --- 変数宣言 ---
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filePath As String
    Dim fileName As String
    Dim extractedText As String
    Dim fso As Object
    Dim newFileCount As Long
    Dim outputFileName As String
    Dim i As Long  ' メインループ用カウンタ
    
    Dim qt As Object ' クエリオブジェクト
    Dim queryName As String
    
    Dim wbNew As Workbook  ' 新規Excelブック用
    Dim wsNew As Worksheet ' 新規Excelシート用
    
    ' 新規保存処理のための変数
    Dim rowArray As Variant  ' 改行で分割した行データの配列
    Dim colArray As Variant  ' タブで分割した列データの配列
    Dim rowIndex As Long     ' 新しいシートの行番号
    Dim rowData As Variant   ' rowArrayをループするための変数
    
    ' 画面更新とイベントを一時停止（高速化）
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' エラーハンドリングの開始
    On Error GoTo ErrorHandler
    
    ' 処理対象シートの設定
    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)
    
    ' 1. Power Queryの更新（高速化と待機処理）
    queryName = TARGET_SHEET_NAME
    
    On Error Resume Next
    If ws.ListObjects.Count > 0 Then
        Set qt = ws.ListObjects(queryName)
    ElseIf ws.QueryTables.Count > 0 Then
        Set qt = ws.QueryTables(queryName)
    End If
    On Error GoTo ErrorHandler
    
    If Not qt Is Nothing Then
        ' 待機処理】バックグラウンドでのクエリ実行を停止し、完了するまで待機させる
        If TypeName(qt) = "ListObject" Then
            qt.QueryTable.BackgroundQuery = False
            qt.Refresh
        ElseIf TypeName(qt) = "QueryTable" Then
            qt.BackgroundQuery = False
            qt.Refresh
        End If
    Else
        ThisWorkbook.RefreshAll
    End If
    
    ' 2. データの読み込みと書き出し準備
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(OUTPUT_FOLDER_PATH) Then
        fso.CreateFolder OUTPUT_FOLDER_PATH
    End If
    
    newFileCount = 0
    
    ' --- データ行のループ処理 ---
    For i = START_ROW To lastRow
        
        fileName = Trim(ws.Cells(i, "A").Value)
        extractedText = ws.Cells(i, "B").Value
        outputFileName = Replace(fileName, ".pdf", ".xlsx")
        filePath = OUTPUT_FOLDER_PATH & Application.PathSeparator & outputFileName
        
        ' 抽出テキストが空の場合はスキップ
        If extractedText = "" Then
            GoTo NextLoop
        End If
        
        If fso.FileExists(filePath) Then
            ' 既存ファイルはスキップ
        Else
            ' 4. 新規Excelブックの作成
            Set wbNew = Application.Workbooks.Add
            Set wsNew = wbNew.Sheets(1)
            
            ' ★★★ 指令されたテキストをセルに展開する処理（Split関数による空定版）★★★
            ' 1. テキストを改行コード（Chr(10)）で行に分割
            rowArray = Split(extractedText, Chr(10)) ' Chr(10) は Power Query の #(lf)
            
            rowIndex = 1
            ' 各行別をループし、タブコード（Chr(9)）で列に分割して書き込む
            For Each rowData In rowArray
                
                ' 空のテキスト行はスキップ
                If Trim(CStr(rowData)) <> "" Then
                    ' 各行をタブで分割 (は Power Query の #(tab)
                    colArray = Split(CStr(rowData), Chr(9))
                    
                    ' 行末に配列を書き込み（一括書き込み）
                    wsNew.Cells(rowIndex, 1).Resize(1, UBound(colArray) - LBound(colArray) + 1).Value = colArray
                    rowIndex = rowIndex + 1
                End If
            Next rowData
            
            ' ★★★ 原閉処理 終了 ★★★
            
            ' 5. ファイルを保存
            wbNew.SaveAs fileName:=filePath, FileFormat:=xlOpenXMLWorkbook
            
            ' 新しいブックを閉じる
            wbNew.Close SaveChanges:=False
            
            Set wsNew = Nothing
            Set wbNew = Nothing
            
            newFileCount = newFileCount + 1
        End If
        
NextLoop: ' スキップ用のラベル
    Next i
    
    ' 6. 処理完了メッセージを表示
    MsgBox "PDFデータ_のExcelファイル書き出しが完了しました。" & vbCrLf & vbCrLf & _
           "[新規処理件数: " & newFileCount & " 件]" & vbCrLf & _
           "[保存先: " & OUTPUT_FOLDER_PATH, vbInformation
    
    GoTo CleanUp
    
' --- エラー処理 ---
ErrorHandler:
    ' エラー時のクリーンアップ処理
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    If Not wbNew Is Nothing Then
        wbNew.Close SaveChanges:=False
    End If
    
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "説明: " & Err.Description, vbCritical
    
' --- クリーンアップ ---
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    Set ws = Nothing
    Set fso = Nothing
    Set qt = Nothing
    Set wsNew = Nothing
    Set wbNew = Nothing
    On Error GoTo 0
    
    ' 7. 処理完了後、自動でファイルを閉じる（変更を保存）
    ThisWorkbook.Close SaveChanges:=True
    
End Sub

' --------------------------------------------------------------------------------
' [文字列のエスケープ関数]
' --------------------------------------------------------------------------------
Function EscapeJsonString(ByVal text As String) As String
    EscapeJsonString = text
End Function