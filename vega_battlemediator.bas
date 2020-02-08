

Sub sheetDelete()
'参数
Dim keySheetName As String
Dim key1 As String
Dim key2 As String
Dim i As Integer

Dim test As String
test = "Sheet1"


'赋值
keySheetName = "battlemediator"
key1 = "addImage"
key2 = "movieclip"


i = 0

'【step1】删除表名不为keySheetName的表
Dim sht As Worksheet
For Each sht In ThisWorkbook.Worksheets
    a = Split(sht.Name, "-")
    if a (0) <> keySheetName then
        'Application.DisplayAlerts = False
        sht.Delete
    end if
Next sht

MsgBox "end1"

'【step2】删除不包含键值的表
For Each sht In ThisWorkbook.Worksheets
    i = 1
    isKeyHave = False
    With sht
        Do While .Cells(1, i).Value <> ""
            If .Cells(1, i).Value = key1 Or .Cells(1, i).Value = key2 Then
                isKeyHave = True
                Exit Do
            End If
            i = i + 1
        Loop
        If isKeyHave = False Then
            sht.Delete
        End If
    End With
Next sht


MsgBox "end2"

End Sub








