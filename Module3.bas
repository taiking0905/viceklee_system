Attribute VB_Name = "Module3"
Sub データの中身一覧表示()
    Dim storeDict As Object
    Dim CommodityDict As Object
    Dim colorDict As Object
    Dim sizeDict As Object
    Dim commodityColorDict As Object
    Dim ws As Worksheet
    Dim resultSheet As Worksheet
    Dim lastRow As Long
    Dim store As String
    Dim Commodity As String
    Dim color As String
    Dim size As String
    Dim quantity As Long
    Dim i As Long
    
    ' datalistシートが存在するか確認
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("datalist")
    On Error GoTo 0
    
    ' datalistシートが存在しない場合は作成
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultSheet.Name = "datalist"
    End If
    
    ' データがあるシート（ここではSheet1を例に使用）
    Set ws = ThisWorkbook.Sheets("Sheet1") ' シート名を適宜変更してください
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row ' 13列目(M列)を対象にする
    
    ' storeDict は、店舗名をキーとし、商品番号のキーをサブに持ち、色、サイズと数を読み取る
    Set storeDict = CreateObject("Scripting.Dictionary")
    
    ' commodityColorDict は、商品名をキーとし、その商品に関連する色、サイズと数量のDictionaryを持つ
    Set commodityColorDict = CreateObject("Scripting.Dictionary")
    
    ' データの読み込みと格納
    For i = 2 To lastRow
        store = ws.Cells(i, 13).value ' 13列目(M列)に店舗名
        Commodity = ws.Cells(i, 8).value ' 8列目(H列)に商品名
        color = ws.Cells(i, 11).value ' 11列目(K列)に色
        size = ws.Cells(i, 10).value ' 10列目(J列)にサイズ
        quantity = ws.Cells(i, 14).value ' 14列目(O列)に出荷数
        
        ' 店舗のDictionaryに商品名のDictionaryが存在しない場合は、新規作成
        If Not storeDict.exists(store) Then
            Set CommodityDict = CreateObject("Scripting.Dictionary")
            storeDict.Add store, CommodityDict
        Else
            Set CommodityDict = storeDict(store)
        End If
        
        ' 商品のDictionaryに色のDictionaryが存在しない場合は、新規作成
        If Not CommodityDict.exists(Commodity) Then
            Set colorDict = CreateObject("Scripting.Dictionary")
            CommodityDict.Add Commodity, colorDict
        Else
            Set colorDict = CommodityDict(Commodity)
        End If
        
        ' 色のDictionaryにサイズのDictionaryを追加
        If Not colorDict.exists(color) Then
            Set sizeDict = CreateObject("Scripting.Dictionary")
            colorDict.Add color, sizeDict
        Else
            Set sizeDict = colorDict(color)
        End If
        
        ' サイズのDictionaryに出荷数を追加または更新
        If sizeDict.exists(size) Then
            sizeDict(size) = sizeDict(size) + quantity
        Else
            sizeDict.Add size, quantity
        End If
        
        ' commodityColorDict にも同様にデータを追加または更新
        If Not commodityColorDict.exists(Commodity) Then
            Set colorDict = CreateObject("Scripting.Dictionary")
            commodityColorDict.Add Commodity, colorDict
        Else
            Set colorDict = commodityColorDict(Commodity)
            If Not colorDict.exists(color) Then
                Set sizeDict = CreateObject("Scripting.Dictionary")
                colorDict.Add color, sizeDict
            Else
                Set sizeDict = colorDict(color)
            End If
            
            If sizeDict.exists(size) Then
                sizeDict(size) = sizeDict(size) + quantity
            Else
                sizeDict.Add size, quantity
            End If
        End If
    Next i
    
    ' すべてのキーと値をループ処理して結果シートに出力
    Dim sKey As Variant
    Dim cKey As Variant
    Dim tKey As Variant
    Dim sizeKey As Variant
    Dim rowIndex As Long
    
    rowIndex = 1
    resultSheet.Cells(rowIndex, 1).value = "店舗"
    resultSheet.Cells(rowIndex, 2).value = "商品"
    resultSheet.Cells(rowIndex, 3).value = "色"
    resultSheet.Cells(rowIndex, 4).value = "サイズ"
    resultSheet.Cells(rowIndex, 5).value = "出荷数"
    
    For Each sKey In storeDict.keys
        Set CommodityDict = storeDict(sKey)
        For Each cKey In CommodityDict.keys
            Set colorDict = CommodityDict(cKey)
            For Each tKey In colorDict.keys
                Set sizeDict = colorDict(tKey)
                For Each sizeKey In sizeDict.keys
                    rowIndex = rowIndex + 1
                    resultSheet.Cells(rowIndex, 1).value = sKey
                    resultSheet.Cells(rowIndex, 2).value = cKey
                    resultSheet.Cells(rowIndex, 3).value = tKey
                    resultSheet.Cells(rowIndex, 4).value = sizeKey
                    resultSheet.Cells(rowIndex, 5).value = sizeDict(sizeKey)
                Next sizeKey
            Next tKey
        Next cKey
    Next sKey
    
    MsgBox "結果は 'datalist' シートに出力されました。", vbInformation
End Sub
