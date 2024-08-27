Attribute VB_Name = "Module2"
Sub �f�[�^�\�쐬()
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
    
    ' �f�[�^������V�[�g�i�����ł�Sheet1���Ɏg�p�j
    Set ws = ThisWorkbook.Sheets("Sheet1") ' �V�[�g����K�X�ύX���Ă�������
    
    ' �ŏI�s���擾
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' �d����r�����邽�߂� Dictionary ���g�p
    Set storeDict = CreateObject("Scripting.Dictionary")
    Set CommodityDict = CreateObject("Scripting.Dictionary")
    
    ' �f�[�^�̓ǂݍ��݂Ɗi�[
    For i = 2 To lastRow
        store = ws.Cells(i, 13).value ' 13���(M��)�ɓX�ܖ�
        Commodity = ws.Cells(i, 8).value ' 8���(H��)�ɏ��i�ԍ�
        color = ws.Cells(i, 11).value ' 11���(K��)�ɐF
        size = ws.Cells(i, 10).value ' 10���(J��)�ɃT�C�Y�i�󔒂��܂ށj
        
        ' �X�܂̏d����r�����Ċi�[
        If Not storeDict.exists(store) Then
            storeDict.Add store, Nothing
        End If
        
        ' ���i�ԍ��̏d����r�����Ċi�[
        If Not CommodityDict.exists(Commodity) Then
            Set colorDict = CreateObject("Scripting.Dictionary")
            CommodityDict.Add Commodity, colorDict
        Else
            Set colorDict = CommodityDict(Commodity)
        End If
        
        ' �F��Dictionary�ɃT�C�Y�̃��X�g�����݂��Ȃ��ꍇ�́A�V�K�쐬
        If Not colorDict.exists(color) Then
            Set sizeList = CreateObject("System.Collections.ArrayList")
            sizeList.Add size
            colorDict.Add color, sizeList
        Else
            Set sizeList = colorDict(color)
            ' �T�C�Y�����Ƀ��X�g�ɑ��݂��Ȃ��ꍇ�̂ݒǉ�
            If Not sizeList.Contains(size) Then
                sizeList.Add size
            End If
        End If
        
    Next i
    
    ' result�V�[�g�����݂��邩�m�F
    On Error Resume Next
    Set dataSheet = ThisWorkbook.Sheets("result")
    On Error GoTo 0
    
    ' result�V�[�g�����݂��Ȃ��ꍇ�͍쐬
    If dataSheet Is Nothing Then
        Set dataSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        dataSheet.Name = "result"
    Else
        ' data�V�[�g�����݂���ꍇ�̓N���A
        dataSheet.Cells.Clear
    End If
    
    ' ���i���̃w�b�_�[��ݒ�
    With dataSheet.Range("B1:B2")
        .Merge
        .value = "���i���"
    End With
    
    ' �X�܁A���i�ԍ��A�F�A�T�C�Y��data�V�[�g�ɏo��
    rowIndex = 3
    
    ' �X��
    dataSheet.Cells(rowIndex, 2).value = "�X��"
    For Each key In storeDict.keys
        rowIndex = rowIndex + 1
        dataSheet.Cells(rowIndex, 2).value = key
    Next key
    
    ' ���i�ԍ��A�F�A�T�C�Y
    rowIndex = 3
    For Each key In CommodityDict.keys
        Set colorDict = CommodityDict(key)
        For Each colorKey In colorDict.keys
            Set sizeList = colorDict(colorKey)
            For Each sizeKey In sizeList
                dataSheet.Cells(1, rowIndex).value = key
                dataSheet.Cells(2, rowIndex).value = colorKey
                dataSheet.Cells(3, rowIndex).value = sizeKey ' �󔒂��܂ރT�C�Y�����̂܂܏o��
                rowIndex = rowIndex + 1
            Next sizeKey
        Next colorKey
    Next key
    
    MsgBox "�f�[�^�� 'result' �V�[�g�ɏo�͂���܂����B", vbInformation
End Sub

