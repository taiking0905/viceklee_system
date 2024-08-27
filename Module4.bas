Attribute VB_Name = "Module4"
Sub データ作成()
    Dim storeDict As Object
    Dim CommodityDict As Object
    Dim colorDict As Object
    Dim sizeDict As Object
    Dim commodityColorDict As Object
    Dim ws As Worksheet
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim dataRow As Long
    Dim dataCol As Long
    Dim store As String
    Dim Commodity As String
    Dim color As String
    Dim size As String
    Dim quantity As Long
    Dim i As Long
    Dim j As Long
    
    ' データがあるシート（ここではSheet1を例に使用）
    Set ws = ThisWorkbook.Sheets("Sheet1") ' シート名を適宜変更してください
    Set dataSheet = ThisWorkbook.Sheets("result") ' シート名を適宜変更してください
    
    ' 最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row ' 13列目(M列)を対象にする
    
    ' 最終行を取得 (2列目に基づいて)
    dataRow = dataSheet.Cells(dataSheet.Rows.Count, 2).End(xlUp).Row
    
    ' 最終列を取得 (2行目に基づいて)
    dataCol = dataSheet.Cells(2, dataSheet.Columns.Count).End(xlToLeft).Column
    
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
    Next i
    
    ' 'data'シートに quantity を表示
    For i = 4 To dataRow
        For j = 3 To dataCol
            store = dataSheet.Cells(i, 2).value ' 2列目(B列)に店舗名
            Commodity = dataSheet.Cells(1, j).value ' 1行目に商品名
            color = dataSheet.Cells(2, j).value ' 2行目に色
            size = dataSheet.Cells(3, j).value ' 3行目にサイズ
            
            ' quantity の取得と表示
            If storeDict.exists(store) Then
                Set CommodityDict = storeDict(store)
                If CommodityDict.exists(Commodity) Then
                    Set colorDict = CommodityDict(Commodity)
                    If colorDict.exists(color) Then
                        Set sizeDict = colorDict(color)
                        If sizeDict.exists(size) Then
                            dataSheet.Cells(i, j).value = sizeDict(size)
                        End If

                    End If
                End If
            End If
        Next j
    Next i
    
    MsgBox "結果は 'result' シートに出力されました。", vbInformation
End Sub

