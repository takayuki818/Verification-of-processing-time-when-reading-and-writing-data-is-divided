Attribute VB_Name = "Module1"
Option Explicit
Sub 連続検証()
    Dim 分割数 As Long
    With Sheets("配列分割テスト")
        For 分割数 = 1 To 10
            If Int(100000 / 分割数) = 100000 / 分割数 Then
                .Range("分割数") = 分割数
                Call 配列分割読み書き検証
            End If
        Next
    End With
    MsgBox "検証完了"
End Sub
Sub 配列分割読み書き検証()
    Dim 分割数 As Long, 回 As Long, 行 As Long, 列 As Long, 値 As Long
    Dim 始時 As Date, 終時 As Date
    With Sheets("配列分割テスト")
        分割数 = .Range("分割数")
        If Int(100000 / 分割数) <> 100000 / 分割数 Then
            MsgBox "除算の余りが生じています" & vbCrLf & "10万を割って余りの出ない分割数を入力してください"
            Exit Sub
        End If
        始時 = Now
        For 回 = 1 To 分割数
            ReDim 配列(1 To 100000 / 分割数, 1 To 100)
            For 行 = 1 To 100000 / 分割数
                For 列 = 1 To 100
                    値 = 値 + 1
                    配列(行, 列) = 値
                Next
            Next
            Range(.Cells(7 + 100000 / 分割数 * (回 - 1), 4), .Cells(7 + 100000 / 分割数 * 回 - 1, 103)) = 配列
            Erase 配列
        Next
        .Cells(7, 4).Resize(100000, 100).ClearContents
        終時 = Now
        .Range("処理時間") = (終時 - 始時) * 24 * 60 * 60
        .Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) = 分割数
        .Cells(Rows.Count, 1).End(xlUp).Offset(0, 1) = .Range("処理時間")
    End With
End Sub
