# vba-validate-input-only-either
B列とC列のどちらか一方のセルしか入力できないようにする入力チェック

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Select Case Target.Column
        Case 2 '' B列（2）が変更されたらC列（3）と比較
          Call validateInputOnlyEither(Target, 3)
        Case 3 '' C列（3）が変更されたらB列（2）と比較
          Call validateInputOnlyEither(Target, 2)
  End Select
End Sub

Function validateInputOnlyEither(ByVal Target As Range, ByVal compareColumnNo As Integer)
    If Target.Value = "" Then
    Else
      '' 比較対象列の値が未入力（空白）かどうか判定
      If Cells(Target.Row, compareColumnNo).Value = "" Then
          '' 未入力（空白）の場合は何もしない
      Else
          '' 入力済みの場合はエラーとしてC列の入力情報を初期化
          MsgBox ("どちらか1行しか入力できません")
          Cells(Target.Row, Target.Column).Value = ""
      End If
    End If
End Function
```
