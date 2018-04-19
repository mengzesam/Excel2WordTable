VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Excel2WordDialog 
   Caption         =   "设置对话框"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3780
   OleObjectBlob   =   "Excel2WordDialog.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Excel2WordDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub BUTTON_CANCEL_Click()
    Me.CheckActFlag.Value = False
    Me.Hide
End Sub

Private Sub BUTTON_OK_Click()
    Me.CheckActFlag.Value = True
    Me.Hide
End Sub

Private Sub BUTTON_SetSource_Click()
    Dim excel_file As String
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "excel file", "*.xls;*.xlsx;*.xlsxm"
        .Filters.Add "csv file", "*.csv"
        If .Show = -1 Then  '-1 ok 0 cancel
          excel_file = .SelectedItems(1)
          If Dir(excel_file) = "" Then
            Me.TextSrcFile.Value = ""
            MsgBox excel_file & "文件不存在,请重新选择", vbExclamation
            Exit Sub
          End If
          Me.TextSrcFile.Value = excel_file
        Else
         Me.TextSrcFile.Value = ""
         MsgBox "未选文件,请重新选择", vbExclamation
         Exit Sub
        End If
    End With
End Sub
