Attribute VB_Name = "Module3"
Sub �f�[�^�̒��g�ꗗ�\��()
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
    
    ' datalist�V�[�g�����݂��邩�m�F
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Sheets("datalist")
    On Error GoTo 0
    
    ' datalist�V�[�g�����݂��Ȃ��ꍇ�͍쐬
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        resultSheet.Name = "datalist"
    End If
    
    ' �f�[�^������V�[�g�i�����ł�Sheet1���Ɏg�p�j
    Set ws = ThisWorkbook.Sheets("Sheet1") ' �V�[�g����K�X�ύX���Ă�������
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row ' 13���(M��)��Ώۂɂ���
    
    ' storeDict �́A�X�ܖ����L�[�Ƃ��A���i�ԍ��̃L�[���T�u�Ɏ����A�F�A�T�C�Y�Ɛ���ǂݎ��
    Set storeDict = CreateObject("Scripting.Dictionary")
    
    ' commodityColorDict �́A���i�����L�[�Ƃ��A���̏��i�Ɋ֘A����F�A�T�C�Y�Ɛ��ʂ�Dictionary������
    Set commodityColorDict = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^�̓ǂݍ��݂Ɗi�[
    For i = 2 To lastRow
        store = ws.Cells(i, 13).value ' 13���(M��)�ɓX�ܖ�
        Commodity = ws.Cells(i, 8).value ' 8���(H��)�ɏ��i��
        color = ws.Cells(i, 11).value ' 11���(K��)�ɐF
        size = ws.Cells(i, 10).value ' 10���(J��)�ɃT�C�Y
        quantity = ws.Cells(i, 14).value ' 14���(O��)�ɏo�א�
        
        ' �X�܂�Dictionary�ɏ��i����Dictionary�����݂��Ȃ��ꍇ�́A�V�K�쐬
        If Not storeDict.exists(store) Then
            Set CommodityDict = CreateObject("Scripting.Dictionary")
            storeDict.Add store, CommodityDict
        Else
            Set CommodityDict = storeDict(store)
        End If
        
        ' ���i��Dictionary�ɐF��Dictionary�����݂��Ȃ��ꍇ�́A�V�K�쐬
        If Not CommodityDict.exists(Commodity) Then
            Set colorDict = CreateObject("Scripting.Dictionary")
            CommodityDict.Add Commodity, colorDict
        Else
            Set colorDict = CommodityDict(Commodity)
        End If
        
        ' �F��Dictionary�ɃT�C�Y��Dictionary��ǉ�
        If Not colorDict.exists(color) Then
            Set sizeDict = CreateObject("Scripting.Dictionary")
            colorDict.Add color, sizeDict
        Else
            Set sizeDict = colorDict(color)
        End If
        
        ' �T�C�Y��Dictionary�ɏo�א���ǉ��܂��͍X�V
        If sizeDict.exists(size) Then
            sizeDict(size) = sizeDict(size) + quantity
        Else
            sizeDict.Add size, quantity
        End If
        
        ' commodityColorDict �ɂ����l�Ƀf�[�^��ǉ��܂��͍X�V
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
    
    ' ���ׂẴL�[�ƒl�����[�v�������Č��ʃV�[�g�ɏo��
    Dim sKey As Variant
    Dim cKey As Variant
    Dim tKey As Variant
    Dim sizeKey As Variant
    Dim rowIndex As Long
    
    rowIndex = 1
    resultSheet.Cells(rowIndex, 1).value = "�X��"
    resultSheet.Cells(rowIndex, 2).value = "���i"
    resultSheet.Cells(rowIndex, 3).value = "�F"
    resultSheet.Cells(rowIndex, 4).value = "�T�C�Y"
    resultSheet.Cells(rowIndex, 5).value = "�o�א�"
    
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
    
    MsgBox "���ʂ� 'datalist' �V�[�g�ɏo�͂���܂����B", vbInformation
End Sub
