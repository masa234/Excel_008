Option Explicit

'定数
Public Const DATA_SHEET_NAME = "データ"
Public Const CSV_EXTENSION = "csv"
Public Const EXCEL_FILE_OUTPUT_FAILED = "Excelファイルの出力に失敗しました。"
Public Const CONFIRM = "確認"


'【概要】シート名が存在するか？
'【作成日】2023/01/10
Public Function IsExistsSheet(ByVal objWb As Excel.Workbook, _
                            ByVal strSheetName As String) As Boolean
On Error GoTo IsExistsSheet_Err
    
    IsExistsSheet = False
    
    Dim objWs As Excel.Worksheet
    
    '引数のシート名でシートオブジェクトを参照する
    '存在しない場合、エラーが発生する
    Set objWs = objWb.Worksheets(strSheetName)
    
    IsExistsSheet = True
    
IsExistsSheet_Err:

IsExistsSheet_Exit:
    Set objWs = Nothing
End Function


'【概要】配列をExcelのシートとして展開する
'【作成日】2023/01/10
Public Function ArrToExcelSheet(ByVal arrOutput As Variant, _
                        ByVal objWb As Excel.Workbook, _
                        ByVal strSheetName As String) As Boolean
On Error GoTo ArrToExcelSheet_Err

    Dim lngArrIdx As Long
    Dim lngCurrentRow As Long
    
    ArrToExcelSheet = False
    
    '行を初期化
    lngCurrentRow = 1

    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrOutput)
        '出力
        objWb.Worksheets(strSheetName).Cells(lngCurrentRow, 1).Value = arrOutput(lngArrIdx)
        '行をカウントアップ
        lngCurrentRow = lngCurrentRow + 1
    Next lngArrIdx
    
    ArrToExcelSheet = True

ArrToExcelSheet_Err:

ArrToExcelSheet_Exit:

End Function


'【概要】特定のディレクトリの特定の拡張子のファイル群を取得する
'【作成日】2023/01/10
Public Function GetFilePaths(ByVal strDirectoryPath As String, _
                        ByVal strExtensionName As String) As Variant
On Error GoTo GetFilePaths_Err

    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    Dim objFso As FileSystemObject
    Dim objFile As File
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    '配列の要素番号初期化
    lngArrIdx = 0
    
    With objFso
        'ファイルの数だけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            '拡張子が指定のものだった場合
            If .GetExtensionName(objFile.Name) = strExtensionName Then
                '配列再宣言
                ReDim Preserve arrRet(lngArrIdx)
                '配列格納
                arrRet(lngArrIdx) = objFile.Path
                '配列の要素番号を1つ進める
                lngArrIdx = lngArrIdx + 1
            End If
        Next objFile
    End With
    
    GetFilePaths = arrRet
    
GetFilePaths_Err:

GetFilePaths_Exit:
    Set objFso = Nothing
    Set objFile = Nothing
End Function

