Attribute VB_Name = "Module2"
Sub データ表作成()
    Dim ws As Worksheet
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim storeDict As Object
    Dim CommodityDict As Object
    Dim colorDict As Object
    Dim sizeList As Object
    Dim store As String
    Dim Commodity As String
    Dim color As String
    Dim size As String
    Dim rowIndex As Long
    Dim key As Variant
    Dim colorKey As Variant
    Dim sizeKey As Variant
    
    ' データがあるシート（ここではSheet1を例に使用）
    Set ws = ThisWorkbook.Sheets("Sheet1") ' シート名を適宜変更してください
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 重複を排除するために Dictionary を使用
    Set storeDict = CreateObject("Scripting.Dictionary")
    Set CommodityDict = CreateObject("Scripting.Dictionary")
    
    ' データの読み込みと格納
    For i = 2 To lastRow
        store = ws.Cells(i, 13).value ' 13列目(M列)に店舗名
        Commodity = ws.Cells(i, 8).value ' 8列目(H列)に商品番号
        color = ws.Cells(i, 11).value ' 11列目(K列)に色
        size = ws.Cells(i, 10).value ' 10列目(J列)にサイズ（空白を含む）
        
        ' 店舗の重複を排除して格納
        If Not storeDict.exists(store) Then
            storeDict.Add store, Nothing
        End If
        
        ' 商品番号の重複を排除して格納
        If Not CommodityDict.exists(Commodity) Then
            Set colorDict = CreateObject("Scripting.Dictionary")
            CommodityDict.Add Commodity, colorDict
        Else
            Set colorDict = CommodityDict(Commodity)
        End If
        
        ' 色のDictionaryにサイズのリストが存在しない場合は、新規作成
        If Not colorDict.exists(color) Then
            Set sizeList = CreateObject("System.Collections.ArrayList")
            sizeList.Add size
            colorDict.Add color, sizeList
        Else
            Set sizeList = colorDict(color)
            ' サイズが既にリストに存在しない場合のみ追加
            If Not sizeList.Contains(size) Then
                sizeList.Add size
            End If
        End If
        
    Next i
    
    ' resultシートが存在するか確認
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Sheets("result")
    On Error GoTo 0
    
    ' resultシートが存在しない場合は作成
    If dataSheet Is Nothing Then
        Set dataSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dataSheet.Name = "result"
    Else
        ' dataシートが存在する場合はクリア
        dataSheet.Cells.Clear
    End If
    
    ' 商品情報のヘッダーを設定
    With dataSheet.Range("B1:B2")
        .Merge
        .value = "商品情報"
    End With
    
    ' 店舗、商品番号、色、サイズをdataシートに出力
    rowIndex = 3
    
    ' 店舗
    dataSheet.Cells(rowIndex, 2).value = "店舗"
    For Each key In storeDict.keys
        rowIndex = rowIndex + 1
        dataSheet.Cells(rowIndex, 2).value = key
    Next key
    
    ' 商品番号、色、サイズ
    rowIndex = 3
    For Each key In CommodityDict.keys
        Set colorDict = CommodityDict(key)
        For Each colorKey In colorDict.keys
            Set sizeList = colorDict(colorKey)
            For Each sizeKey In sizeList
                dataSheet.Cells(1, rowIndex).value = key
                dataSheet.Cells(2, rowIndex).value = colorKey
                dataSheet.Cells(3, rowIndex).value = sizeKey ' 空白を含むサイズをそのまま出力
                rowIndex = rowIndex + 1
            Next sizeKey
        Next colorKey
    Next key
    
    MsgBox "データは 'result' シートに出力されました。", vbInformation
End Sub

