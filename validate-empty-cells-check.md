# validate-empty-cells-check
未入力のセルがある場合は背景色を黄色に、入力済みのセルは背景色を初期化する。

```vba
Sub Button1_Click()
    '' 開始列番号
    Dim beginColNo As Long
    beginColNo = 1
    '' 終了列番号
    Dim endColNo As Long
    endColNo = 13
    
    '' A列の最終行を取得します。
    '' （A列に必ず値が入る想定で作成します。）
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '' 未入力セル数
    Dim emptyCellCoount As Long
    emptyCellCoount = 0
    
    '' 行ごとに処理します。
    Dim rowNo As Long
    For rowNo = 1 To lastRow
        '' 行の開始列（beginColNo）から終了列（endColNo）までの範囲を定義します。
        Dim rng As range
        Set rng = range(Cells(rowNo, beginColNo), Cells(rowNo, endColNo))
        
        ''開始列（beginColNo）から終了列（endColNo）のセルを1列ずつチェック
        For colNo = beginColNo To endColNo
            '' 未入力のセルは色を黄色に。入力済みの列は背景色をクリア。
            If Cells(rowNo, colNo).Value = "" Then
                Cells(rowNo, colNo).Interior.Color = RGB(255, 255, 0)
                emptyCellCoount = emptyCellCoount + 1
            Else
                Cells(rowNo, colNo).Interior.ColorIndex = 0
            End If
        Next colNo
    Next rowNo
    
    If emptyCellCoount > 0 Then
        MsgBox ("未入力のセルが存在します。")
    Else
        '' 未入力のセルなし
    End If
End Sub
```
