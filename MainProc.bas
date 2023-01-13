Option Explicit


'【概要】シート名が存在するか？
'【作成日】2023/01/10
'TODO:　関数名改善の余地あり
Public Function GetCSVToArr(ByVal strCSVFilePath As String) As Variant
On Error GoTo GetCSVToArr_Err

    Dim lngFreeFile As Long
    Dim lngArrIdx As Long
    Dim strLine As String
    Dim arrRet() As Variant
    
    'フリーファイルを取得する
    lngFreeFile = FreeFile
       
    'CSVファイルを開く
    Open strCSVFilePath For Input As #lngFreeFile

    '配列の要素番号初期化
    lngArrIdx = 0
    
    '終端まで繰り返す
    Do Until EOF(lngFreeFile)
        '1行読み込み
        Line Input #lngFreeFile, strLine
        '配列再宣言
        ReDim Preserve arrRet(lngArrIdx)
        '配列格納
        arrRet(lngArrIdx) = strLine
        '配列の要素番号を1つ進める
        lngArrIdx = lngArrIdx + 1
    Loop
    
    GetCSVToArr = arrRet
    
GetCSVToArr_Err:

GetCSVToArr_Exit:

End Function


'【概要】シート名が存在するか？
'【作成日】2023/01/10
Public Function GetSheetNameWithSeqNumber(ByVal objWb As Excel.Workbook, _
                                    ByVal strBaseSheetName As String) As String
On Error GoTo GetSheetNameWithSeqNumber_Err
    
    Dim lngCount As Long
    Dim strSheetName As String
    
    '100回繰り返す
    For lngCount = 1 To 100
        'シート名設定
        strSheetName = strBaseSheetName & "_" & CStr(lngCount)
        'シートが存在しない場合、処理終了
        If IsExistsSheet(objWb, strSheetName) = False Then
            GetSheetNameWithSeqNumber = strSheetName
            GoTo GetSheetNameWithSeqNumber_Exit
        End If
    Next lngCount
    
GetSheetNameWithSeqNumber_Err:

GetSheetNameWithSeqNumber_Exit:

End Function


'【概要】配列をExcelのシートとして展開する
'【作成日】2023/01/10
Public Function CSVFilesToExcelFile(ByVal arrCSVFilePaths As Variant, _
                                ByVal strBaseSheetName As String) As Boolean
On Error GoTo CSVFilesToExcelFile_Err

    Dim lngArrIdx As Long
    Dim strSheetName As String
    Dim strFilePath As String
    Dim arrData() As Variant
    Dim objWb As Excel.Workbook
    
    CSVFilesToExcelFile = False
    
    'Excelのブック作成
    Set objWb = Workbooks.Add
    
    'シート名を設定
    strSheetName = strBaseSheetName
    ActiveSheet.Name = strSheetName
    
    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrCSVFilePaths)
        'CSVファイルパス
        strFilePath = arrCSVFilePaths(lngArrIdx)
        'CSVファイルを配列に格納
        arrData = GetCSVToArr(strFilePath)
        '配列をExcelのシートとして展開
        If ArrToExcelSheet(arrData, objWb, strSheetName) = False Then
            GoTo CSVFilesToExcelFile_Exit
        End If
        '次のシート名取得
        strSheetName = GetSheetNameWithSeqNumber(objWb, strBaseSheetName)
        '最終はシート作成しない
        If lngArrIdx <> UBound(arrCSVFilePaths) Then
            'シート追加
            objWb.Worksheets.Add
            ActiveSheet.Name = strSheetName
        End If
    Next lngArrIdx

    CSVFilesToExcelFile = True

CSVFilesToExcelFile_Err:

CSVFilesToExcelFile_Exit:
    Set objWb = Nothing
End Function
