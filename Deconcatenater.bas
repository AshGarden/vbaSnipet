スニペット: 選択した一列の記載内容を二列に分割する(final revision: 2024/03/08)
    
Sub DeConcatenater()

    ' Define variables. 変数を定義します。
    Dim rng As Range
    Dim cell As Range
    Dim colNum As Integer
    Dim rowNum As Integer
    Dim lastRow As Integer
    Dim insertScope As Range
    Dim elements() As String
    Dim firstElement As String
    Dim leftElements As String
    Dim spacePos As Integer
    
    ' Pause Excel updates. Excelの更新を一時停止します。
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Set the selection range. This macro operates on the selection range. 選択範囲を設定します。このマクロは選択範囲に対して動作します。
    Set rng = Selection

    ' Set error handling. エラーハンドリングを設定します。
    On Error GoTo ErrorHandler
    
    ' Set the threshold. This value is used to separate the processing when the number of cells in the selection range is less than or equal to this value and when it is more than this value. 
    ' The value was estimated by Fermi estimation, considering factors such as ① having multiple business systems open, ② using in power-saving mode, ③ the standard specs of a laptop provided for general office work, etc., so that the work can be carried out without delay.
    ' 閾値を設定します。この値は、選択範囲のセル数がこの値以下の場合と以上の場合で処理を分けるためのものです。
    ' この値は、①業務用システムを複数開く、②省電力モードでの使用、③一般的な事務作業を目的として貸与されるノートパソコンの標準的スペック等を考慮して、業務を滞りなく遂行できる値をフェルミ推定しました。

    threshold = 10000

    ' If the selection range is not a single column, display an error message and abort the process. 選択範囲が一列でない場合、エラーメッセージを表示して処理を中断します。
    If rng.Columns.count <> 1 Then
        MsgBox "ちゃんと一列を選択してください。"
        Exit Sub
    End If

    ' Store the row number of the selected row in a variable. 選択した行の行番号を変数に収めます。
    colNum = rng.Column
    ' If the number of cells in the selection range is less than the threshold, process each cell individually. 選択範囲のセル数が閾値以下の場合、各セルを個別に処理します。
    If rng.Cells.count < threshold Then


    ' Store the row number of the selected row in a variable. 選択した行の行番号を変数に収めます。
    rowNum = rng.row
    ' Get the last row of the selection range. 選択範囲の最終行を取得します。
    lastRow = rng.row + rng.Rows.count - 1

    Set insertScope = Range(Cells(rowNum, colNum + 1), Cells(lastRow, colNum + 1))
    insertScope.Insert (xlShiftToRight)
            
    For Each cell In rng
        ' Split the contents of the cell by space. セルの内容をスペースで分割します。
        spacePos = InStr(cell.Value, " ")
        If spacePos > 0 Then
            firstElement = Left(cell.Value, spacePos - 1)
            leftElements = Mid(cell.Value, spacePos + 1)
        Else
            firstElement = cell.Value
            leftElements = ""
        End If
        
        ' Leave the first element in the original cell. 最初の要素を元のセルに残します。
        cell.Value = firstElement
        
        ' Move the remaining elements to the cell to the right. 残りの要素を右隣のセルに移動します。
        cell.Offset(0, 1).Value = leftElements
    Next cell


    ' If the number of cells in the selection range is greater than the threshold, use an array to speed up processing. 選択範囲のセル数が閾値より大きい場合、配列を使用して処理を高速化します。
    Else

        Columns(colNum + 1).Insert (xlShiftToRight)
        data = rng.Value
        ReDim result(1 To UBound(data, 1), 1 To 2)
        For i = 1 To UBound(data, 1)
            ' Split the contents of the cell by space. セルの内容をスペースで分割します。
            spacePos = InStr(data(i, 1), " ")
            If spacePos > 0 Then
                firstElement = Left(data(i, 1), spacePos - 1)
                leftElements = Mid(data(i, 1), spacePos + 1)
            Else
                firstElement = data(i, 1)
                leftElements = ""
            End If
            
            ' Leave the first element in the original cell. 最初の要素を元のセルに残します。
            result(i, 1) = firstElement
            
            ' Move the remaining elements to the cell to the right. 残りの要素を右隣のセルに移動します。
            result(i, 2) = leftElements
        Next i
    
        ' Write the results to the original range and the newly inserted column. 結果を元の範囲と新しく挿入した列に書き込みます。
        rng.Value = Application.Index(result, , 1)
        rng.Offset(0, 1).Value = Application.Index(result, , 2)
    
    End If

    ' Adjust the width of the split column. 分割した列の幅を調整します。
    rng.EntireColumn.AutoFit
    rng.Offset(0, 1).EntireColumn.AutoFit
    
    ' End error handling. エラーハンドリングを終了します。
    GoTo CleanExit

ErrorHandler:
    ' Branch processing according to the error number. エラー番号に応じて処理を分岐
    Select Case Err.Number
        Case 9 ' The index is out of range. インデックスが範囲を超えています
            MsgBox "選択範囲が正しくありません。"
            ' Perform alternative processing as necessary. 必要に応じて代替の処理を行う
        Case Else
            MsgBox "エラーが発生しました: " & Err.Description
    End Select

CleanExit:
    ' Resume Excel updates. Excelの更新を再開します。
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub
