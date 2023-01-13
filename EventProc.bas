Option Explicit

 
Public Sub 正方形長方形1_Click()
On Error GoTo 正方形長方形1_Click_Err

    Dim arrCSVFilePaths() As Variant
    
    'CSVファイル群を取得
    arrCSVFilePaths = GetFilePaths(ThisWorkbook.Path, CSV_EXTENSION)
    
    '画面の更新をオフにする
    Application.ScreenUpdating = False
    
    'CSVファイルをExcelファイルとして出力
    If CSVFilesToExcelFile(arrCSVFilePaths, DATA_SHEET_NAME) = False Then
        Call MsgBox(EXCEL_FILE_OUTPUT_FAILED, vbInformation, CONFIRM)
        GoTo 正方形長方形1_Click_Exit
    End If
    
正方形長方形1_Click_Err:

正方形長方形1_Click_Exit:
    '画面の更新をオンにする
    Application.ScreenUpdating = True
End Sub
