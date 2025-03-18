Sub finalCheckMacro()
    Dim range1 As Range, range2 As Range
    Dim cell1 As Range, cell2 As Range
    Dim mismatchDetails As String
    Dim val1 As Variant, val2 As Variant
    Dim row1 As Long, row2 As Long
    Dim startRow1 As Long, endRow1 As Long
    Dim keyValue As String
    Dim Sheet1 As Worksheet, Sheet2 As Worksheet
    Dim P2Value As String
    Dim sheetList As String, sheetIndex As Integer
    Dim lastRow2 As Long
    Dim A0NoCol As Long
    Dim A0NoValue As String

    ' 「やるやら」シートを確認して設定
    On Error Resume Next
    Set Sheet1 = ThisWorkbook.Sheets("やるやら")
    On Error GoTo 0
    If Sheet1 Is Nothing Then
        MsgBox "「やるやら」シートが見つかりません。処理を終了します。", vbExclamation
        Exit Sub
    End If

    ' シート名リストの作成
    sheetList = "シート名リスト:" & vbCrLf
    For rowIdx = 1 To ThisWorkbook.Sheets.Count
        sheetList = sheetList & rowIdx & ". " & ThisWorkbook.Sheets(rowIdx).Name & vbCrLf
    Next rowIdx

    ' ユーザーにシート番号を入力させる
    sheetIndex = Application.InputBox("比較するシートを選択してください（番号を入力）:" & vbCrLf & sheetList, "シート選択", Type:=1)

    ' 入力チェック
    If sheetIndex < 1 Or sheetIndex > ThisWorkbook.Sheets.Count Then
        MsgBox "正しいシート番号を入力してください。", vbExclamation
        Exit Sub
    End If

    ' シート2を設定
    Set Sheet2 = ThisWorkbook.Sheets(sheetIndex)
    
    ' 「A0 No.」ラベルの列番号を取得
    A0NoCol = Application.Match("A0 No.", Sheet2.Rows(1), 0)
    
    ' 「A0 No.」が見つかった場合、その列の2行目の値の左から4文字を取得
    If Not IsError(A0NoCol) Then
        A0NoValue = Left(Sheet2.Cells(2, A0NoCol).Value, 4)
    Else
        MsgBox """A0 No.""" & " ラベルが見つかりませんでした。"
        Exit Sub
    End If
    
    ' keyValueに格納
    keyValue = Trim(A0NoValue)

    ' シート1のA列でキー値が始まる行を検索
    startRow1 = Sheet1.Columns(1).Find(What:=keyValue, LookIn:=xlValues, LookAt:=xlWhole).Row

    ' シート1のA列でキー値が終わる行を検索
    endRow1 = Sheet1.Columns(1).Find(What:=keyValue & "*", LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlPrevious).Row

    ' エラー処理
    If startRow1 = 0 Or endRow1 = 0 Then
        MsgBox "キー値（" & keyValue & "）が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' シート1（やるやら）で比較する列を選択
    Set range1 = Application.InputBox("比較する列をやるやらシートで選択してください（例: =やるやら!$B:$B）。", Type:=8)
    If range1 Is Nothing Or range1.Worksheet.Name <> "やるやら" Then
        MsgBox "やるやらシートの列が正しく選択されていません。処理を終了します。", vbExclamation
        Exit Sub
    End If

    ' シート2で比較する列を選択
    Set range2 = Application.InputBox("比較する列を他のシートで選択してください（例: =" & Sheet2.Name & "!$P:$P）。", Type:=8)
    If range2 Is Nothing Then
        MsgBox "シート2の列が選択されませんでした。処理を終了します。", vbExclamation
        Exit Sub
    End If

    ' 初期化
    mismatchDetails = ""

    ' 比較処理（1対1の行比較）
    row2 = 2 ' シート2の開始行
    For row1 = startRow1 To endRow1
        If row2 > Sheet2.Cells(Sheet2.Rows.Count, range2.Column).End(xlUp).Row Then Exit For

        ' セルを取得
        Set cell1 = Sheet1.Cells(row1, range1.Column)
        Set cell2 = Sheet2.Cells(row2, range2.Column)

        ' 値を取得して比較
        val1 = Trim(CStr(cell1.Value))
        val2 = Trim(CStr(cell2.Value))

        ' 比較条件
        If val1 <> val2 Then
            cell1.Interior.Color = RGB(255, 0, 0)
            cell2.Interior.Color = RGB(255, 0, 0)
            mismatchDetails = mismatchDetails & "シート1行 " & row1 & " / シート2行 " & row2 & ": 値が一致しません (Cell1: [" & val1 & "], Cell2: [" & val2 & "])" & vbCrLf
        End If

        ' シート2の次の行へ
        row2 = row2 + 1
    Next row1

    ' 結果を表示
    If mismatchDetails = "" Then
        MsgBox "すべて一致しました！", vbInformation
    Else
        MsgBox "以下の不一致が見つかりました:" & vbCrLf & mismatchDetails, vbExclamation
        ' 不一致行を新しいシートに書き出す
        WriteMismatchToNewSheet mismatchDetails
    End If
End Sub

Sub WriteMismatchToNewSheet(MismatchRows As String)
    Dim NewSheet As Worksheet
    Dim Lines As Variant
    Dim RowIndex As Long

    ' 新しいシートを追加
    Set NewSheet = ThisWorkbook.Sheets.Add
    NewSheet.Name = "不一致行(最終チェックマクロ)"

    ' ヘッダーを書き込む
    NewSheet.Cells(1, 1).Value = "不一致行の詳細"

    ' MismatchRows を改行で分割して配列に格納
    Lines = Split(MismatchRows, vbCrLf)

    ' 不一致情報を書き込む
    For RowIndex = LBound(Lines) To UBound(Lines)
        If Lines(RowIndex) <> "" Then
            NewSheet.Cells(RowIndex + 2, 1).Value = Lines(RowIndex)
        End If
    Next RowIndex
End Sub





