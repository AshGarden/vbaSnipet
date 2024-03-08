Sub Concatenater()

    ' Define variables. 変数を定義します。
    Dim threshold As Long
    Dim rng As Range
    Dim colNum As Integer
    Dim rowNum As Integer
    Dim lastRow As Integer
    Dim data As Variant
    Dim i As Long
    Dim j As Long
    Dim cell As Range

    ' Pause Excel updates. Excelの更新を一時停止します。
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Set error handling. エラーハンドリングを設定します。
    On Error GoTo ErrorHandler

    ' Set the threshold. This value is used to separate the processing when the number of cells in the selection range is less than or equal to this value and when it is more than this value. 
    ' The value was estimated by Fermi estimation, considering factors such as ① having multiple business systems open, ② using in power-saving mode, ③ the standard specs of a laptop provided for general office work, etc., so that the work can be carried out without delay.
    ' 閾値を設定します。この値は、選択範囲のセル数がこの値以下の場合と以上の場合で処理を分けるためのものです。
    ' この値は、①業務用システムを複数開く、②省電力モードでの使用、③一般的な事務作業を目的として貸与されるノートパソコンの標準的スペック等を考慮して、業務を滞りなく遂行できる値をフェルミ推定しました。
    threshold = 10000

    ' Set the selected range. This macro operates on the selected range.
    ' 選択範囲を設定します。このマクロは選択範囲に対して動作します。
    Set rng = Selection
    ' Store the column number of the selected column in a variable.
    ' 選択した列の列番号を変数に収めます。
    colNum = rng.Column
    ' Store the row number of the selected row in a variable.
    ' 選択した行の行番号を変数に収めます。
    rowNum = rng.row
    ' Get the last row of the selected range.
    ' 選択範囲の最終行を取得します。
    lastRow = rng.row + rng.Rows.count - 1

    ' If the selected range is not two columns, display an error message and abort the process.
    ' 選択範囲が二列でない場合、エラーメッセージを表示して処理を中断します。
    If rng.Columns.count <> 2 Then
        MsgBox "ちゃんと二列を選択してください。"
        GoTo CleanExit
    End If

    ' If the number of cells in the selected range is less than the threshold, process each cell individually.
    ' 選択範囲のセル数が閾値以下の場合、各セルを個別に処理します。
    If rng.Cells.count < threshold Then
        data = rng.Value
        For i = 1 To rng.Rows.count
            For j = 1 To rng.Columns.count - 1
                ' If the next cell is not empty, concatenate the value of the current cell and the value of the next cell.
                ' 隣のセルが空でない場合、現在のセルの値と隣のセルの値を結合します。
                If Len(Trim(rng.Cells(i, j + 1).Value)) > 0 Then
                    ' Before concatenating, remove the spaces at the beginning and end of the left and right cells.
                    ' 結合する前に、左右のセルの先頭と末尾のスペースを削除します。
                    rng.Cells(i, j).Value = Trim(rng.Cells(i, j).Value)
                    rng.Cells(i, j + 1).Value = Trim(rng.Cells(i, j + 1).Value)
                    ' Concatenate the values with the trailing spaces removed.
                    ' 末尾のスペースを削除した値を結合します。
                    rng.Cells(i, j).Value = rng.Cells(i, j).Value & " " & rng.Cells(i, j + 1).Value
                End If
            Next j
        Next i
        ' Delete the column on the right after concatenation.
        ' 結合後の右側の列を削除します。
        Range(Cells(rowNum, colNum + 1), Cells(lastRow, colNum + 1)).delete Shift:=xlToLeft
    ' If the number of cells in the selected range is greater than the threshold, use an array to speed up the processing.
    ' 選択範囲のセル数が閾値より大きい場合、配列を使用して処理を高速化します。
    Else
        data = rng.Value
        For i = 1 To UBound(data, 1)
            ' Loop up to UBound(data, 2) - 1 to avoid exceeding the last column.
            ' 最後の列を超えないようにするため、UBound(data, 2) - 1までループします。
            For j = 1 To UBound(data, 2) - 1
                ' If the next cell is also not empty, concatenate the value of the current cell and the value of the next cell.
                ' 隣のセルも空でない場合、現在のセルの値と隣のセルの値を結合します。
                If Len(Trim(data(i, j + 1))) > 0 Then
                    ' Before concatenating, remove the spaces at the beginning and end of the left and right cells.
                    ' 結合する前に、左右のセルの先頭と末尾のスペースを削除します。
                    data(i, j) = Trim(data(i, j))
                    data(i, j + 1) = Trim(data(i, j + 1))
                    ' Concatenate the values with the trailing spaces removed.
                    ' 末尾のスペースを削除した値を結合します。
                    data(i, j) = data(i, j) & " " & data(i, j + 1)
                End If
            Next j
        Next i
        ' Delete the column on the right after concatenation.
        ' 結合後の右側の列を削除します。
        Columns(colNum + 1).delete Shift:=xlToLeft
        ' Return the values of the array to the original selected range.
        ' 配列の値を元の選択範囲に戻します。
        rng.Value = data
    End If
    
    ' Adjust the width of the specified column.
    ' 指定した列の幅を調整します。
    Columns(colNum).EntireColumn.AutoFit
    
    ' End error handling.
    ' エラーハンドリングを終了します。
    GoTo CleanExit

ErrorHandler:
    ' Branch the process according to the error number.
    ' エラー番号に応じて処理を分岐
    Select Case Err.Number
        Case 9 ' Index is out of range.
            MsgBox "選択範囲が正しくありません。"
            ' Perform alternative processing as necessary.
        Case Else
            MsgBox "エラーが発生しました: " & Err.Description
    End Select

CleanExit:
    ' Resume Excel updates.
    ' Excelの更新を再開します。
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
