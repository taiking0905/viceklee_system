Attribute VB_Name = "Module4"
Sub �f�[�^�쐬()
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
    
    ' �f�[�^������V�[�g�i�����ł�Sheet1���Ɏg�p�j
    Set ws = ThisWorkbook.Sheets("Sheet1") ' �V�[�g����K�X�ύX���Ă�������
    Set dataSheet = ThisWorkbook.Sheets("result") ' �V�[�g����K�X�ύX���Ă�������
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row ' 13���(M��)��Ώۂɂ���
    
    ' �ŏI�s���擾 (2��ڂɊ�Â���)
    dataRow = dataSheet.Cells(dataSheet.Rows.Count, 2).End(xlUp).Row
    
    ' �ŏI����擾 (2�s�ڂɊ�Â���)
    dataCol = dataSheet.Cells(2, dataSheet.Columns.Count).End(xlToLeft).Column
    
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
    Next i
    
    ' 'data'�V�[�g�� quantity ��\��
    For i = 4 To dataRow
        For j = 3 To dataCol
            store = dataSheet.Cells(i, 2).value ' 2���(B��)�ɓX�ܖ�
            Commodity = dataSheet.Cells(1, j).value ' 1�s�ڂɏ��i��
            color = dataSheet.Cells(2, j).value ' 2�s�ڂɐF
            size = dataSheet.Cells(3, j).value ' 3�s�ڂɃT�C�Y
            
            ' quantity �̎擾�ƕ\��
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
    
    MsgBox "���ʂ� 'result' �V�[�g�ɏo�͂���܂����B", vbInformation
End Sub

