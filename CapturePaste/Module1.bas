Attribute VB_Name = "Module1"
Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

Public flg_stop As Boolean      ' 監視フラグ

Sub AutoCapture()
    Dim book1 As Workbook
    Dim CB As Variant
    Dim OffsetY As Long
    
    flg_stop = False
    
    ' 倍率
    Dim bairitu As Double
    bairitu = Val(Range("bairitu").Value) / 100
    
    ' オフセット
    Dim offset As Long
    offset = Val(Range("bairitu").Value) * 0.6
    
    ' ------------------------------
    ' 新規シート作成
    ' ------------------------------
    ' 既存シート削除
    Dim sheetNum As Long
    sheetNum = 1
    For Each ws In Worksheets
        If ws.Name = "エビデンス" + CStr(sheetNum) Then
            sheetNum = sheetNum + 1
'            ' 削除はやめる
'            Application.desplayalerts = False
'            ws.Delete
'            Application.desplayalerts = True
'            Exit For
        End If
    Next ws
        
    ' シート作成
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "エビデンス" + CStr(sheetNum)
    
    ' ------------------------------
    ' 監視
    ' ------------------------------
    Do While True
        CB = Application.ClipboardFormats
        If flg_stop = True Then GoTo Quit
        If CB(1) <> -1 Then
            For i = 1 To UBound(CB)
                If CB(i) = xlClipboardFormatBitmap Then
                    Workbooks("キャプチャ貼り付けマクロ.xlsm").Activate
                    ActiveSheet.Paste Destination:=Range("B2").offset(OffsetY, 0)
                    OffsetY = OffsetY + offset
                    
                    ' クリップボードを空にする
                    OpenClipboard
                    EmptyClipboard
                    CloseClipboard
                    
                    ' 倍率変更
                    Selection.ShapeRange.ScaleHeight bairitu, msoFalse, msoScaleFromTopLeft
                    
                    ' 保存
                    ActiveWorkbook.Save
                    
                    ' 選択
                    Cells(OffsetY, 2).Select
                End If
            Next i
        End If
        DoEvents
    Loop
        
Quit:
    MsgBox "停止しました", vbInformation
    ActiveSheet.Cells(1, 1).ClearContents
End Sub

Sub StopCapture()
    flg_stop = True
    
    ' 保存
    ActiveWorkbook.Save
    
    ' クリップボードを空にする
    OpenClipboard
    EmptyClipboard
    CloseClipboard
End Sub
