Attribute VB_Name = "Module1"
Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

Public flg_stop As Boolean      ' �Ď��t���O

Sub AutoCapture()
    Dim book1 As Workbook
    Dim CB As Variant
    Dim OffsetY As Long
    
    flg_stop = False
    
    ' �{��
    Dim bairitu As Double
    bairitu = Val(Range("bairitu").Value) / 100
    
    ' �I�t�Z�b�g
    Dim offset As Long
    offset = Val(Range("bairitu").Value) * 0.6
    
    ' ------------------------------
    ' �V�K�V�[�g�쐬
    ' ------------------------------
    ' �����V�[�g�폜
    Dim sheetNum As Long
    sheetNum = 1
    For Each ws In Worksheets
        If ws.Name = "�G�r�f���X" + CStr(sheetNum) Then
            sheetNum = sheetNum + 1
'            ' �폜�͂�߂�
'            Application.desplayalerts = False
'            ws.Delete
'            Application.desplayalerts = True
'            Exit For
        End If
    Next ws
        
    ' �V�[�g�쐬
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "�G�r�f���X" + CStr(sheetNum)
    
    ' ------------------------------
    ' �Ď�
    ' ------------------------------
    Do While True
        CB = Application.ClipboardFormats
        If flg_stop = True Then GoTo Quit
        If CB(1) <> -1 Then
            For i = 1 To UBound(CB)
                If CB(i) = xlClipboardFormatBitmap Then
                    Workbooks("�L���v�`���\��t���}�N��.xlsm").Activate
                    ActiveSheet.Paste Destination:=Range("B2").offset(OffsetY, 0)
                    OffsetY = OffsetY + offset
                    
                    ' �N���b�v�{�[�h����ɂ���
                    OpenClipboard
                    EmptyClipboard
                    CloseClipboard
                    
                    ' �{���ύX
                    Selection.ShapeRange.ScaleHeight bairitu, msoFalse, msoScaleFromTopLeft
                    
                    ' �ۑ�
                    ActiveWorkbook.Save
                    
                    ' �I��
                    Cells(OffsetY, 2).Select
                End If
            Next i
        End If
        DoEvents
    Loop
        
Quit:
    MsgBox "��~���܂���", vbInformation
    ActiveSheet.Cells(1, 1).ClearContents
End Sub

Sub StopCapture()
    flg_stop = True
    
    ' �ۑ�
    ActiveWorkbook.Save
    
    ' �N���b�v�{�[�h����ɂ���
    OpenClipboard
    EmptyClipboard
    CloseClipboard
End Sub
