Attribute VB_Name = "excel2wordTable"
Option Explicit
Option Base 1

Sub Excel2WordTable()
    Dim myDoc As Document
    Dim excel_file As String
    Dim excelObj As Object
    Dim destTable As Table
    Dim sheet_index As Integer
    Dim xls_r0 As Long, xls_r1 As Long, xls_c0 As Long, xls_c1 As Long
    Dim doc_r0 As Long, doc_c0 As Long
    Dim i As Integer, j As Integer
    Dim dest As String, text As String
    Set myDoc = ActiveDocument
    Set destTable = Nothing
    dest = Selection.Sentences(1).text
    For i = 1 To myDoc.Tables.Count
       Dim src As String
       src = myDoc.Tables(i).cell(1, 1).Range.text
       If (src = dest) Then
            Set destTable = myDoc.Tables(i)
            Exit For
       End If
    Next
    If (destTable Is Nothing) Then
        MsgBox "请将光标移到word表格第一个单元格，以使该表格成为目标表格", vbExclamation
        Exit Sub
    End If
    Excel2WordDialog.Show
    If (Excel2WordDialog.CheckActFlag.Value = False) Then
         MsgBox "已取消,请重新进行设置", vbExclamation
         Exit Sub
    End If
    excel_file = Excel2WordDialog.TextSrcFile.Value
    If (excel_file = "") Then
         MsgBox "未选择正确的源文件", vbExclamation
         Exit Sub
    End If
    sheet_index = Excel2WordDialog.TextSheetIndex.Value
    xls_r0 = Excel2WordDialog.TextSRC_R0.Value
    xls_c0 = Excel2WordDialog.TextSRC_C0.Value
    xls_r1 = Excel2WordDialog.TextSRC_R1.Value
    xls_c1 = Excel2WordDialog.TextSRC_C1.Value
    doc_r0 = Excel2WordDialog.TextDEST_R0.Value
    doc_c0 = Excel2WordDialog.TextDEST_C0.Value
    'If (destTable.Rows.Count - doc_r0 < xls_r1 - xls_r0 And destTable.Columns.Count - doc_c0 < xls_c1 - xls_c0) Then
    '    MsgBox "插入位置的目标表格行数和列数均小于数据源行数和列数", vbExclamation
    '    Exit Sub
    'ElseIf (destTable.Rows.Count - doc_r0 < xls_r1 - xls_r0) Then
     '   MsgBox "插入位置的目标表格行数小于数据源行数", vbExclamation
     '   Exit Sub
    If (destTable.Columns.Count - doc_c0 < xls_c1 - xls_c0) Then
        MsgBox "插入位置的目标表格列数小于数据源列数", vbExclamation
        Exit Sub
    End If
    For i = destTable.Rows.Count - doc_r0 + 1 To xls_r1 - xls_r0
        destTable.Rows.Add
    Next
    Set excelObj = GetObject(excel_file)
    For i = 0 To xls_r1 - xls_r0
      For j = 0 To xls_c1 - xls_c0
        text = excelObj.Sheets(sheet_index).Cells(xls_r0 + i, xls_c0 + j).text
        destTable.cell(doc_r0 + i, doc_c0 + j).Range.text = text    '=Format(cell.Text, cell.NumberFormatLocal
      Next
    Next
    Set excelObj = Nothing
    Set destTable = Nothing
    Set myDoc = Nothing
End Sub

