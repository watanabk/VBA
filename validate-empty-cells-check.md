# validate-empty-cells-check
未入力のセルがある場合は背景色を黄色に、入力済みのセルは背景色を初期化する。

1. 「開発」タブ > コントロール > 挿入 > ボタン（フォームコントロール）
1. 適当な位置に配置
1. マクロ名を `Button1_Click()` に変更
1. Microsoft Visual Basic for Applications が開くので
1. `標準モジュール` の `Module1` に以下のプログラムをコピペ

```vba
Sub Button1_Click()
    '' 判定開始列番号
    '' D列（4）から判定を行うようにする。
    Dim beginColNo As Long
    beginColNo = 4
    '' 判定終了列番号（M列 = 13）
    Dim endColNo As Long
    endColNo = 13
    '' チェック処理を開始する行数
    '' 1行目はヘッダが設定されるため、2行目からチェック処理を行うようにする。
    Dim checkStartRowNo As Long
    checkStartRowNo = 2
    
    '' A列（1）の最終行を求めます。
    '' （A列には行数を設定する予定のため、必ず値が設定されている。）
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    '' 未入力セル数をカウントカウントするための変数
    Dim emptyCellCoount As Long
    emptyCellCoount = 0
    
    '' 1行ごとに処理します。
    Dim rowNo As Long
    For rowNo = checkStartRowNo To lastRow
        '' 判定開始列（beginColNo）から判定終了列（endColNo）のセルを1列ずつチェック
        Dim rng As range
        Set rng = range(Cells(rowNo, beginColNo), Cells(rowNo, endColNo))
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
