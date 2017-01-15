
Imports Microsoft.Office.Interop
Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim app As New Excel.Application 'app 是操作 Excel 的變數
        Dim worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
        Dim workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
        Dim path As String

        Dim del(4000), curNo As Integer

        'Dim spaceN, tabN, t, stdN As Integer


        path = Application.StartupPath + "\studentList.xlsx"
        'path = Application.StartupPath + "\1.xlsx"
        workbook = app.Workbooks.Open(path) '開啟一張已存在的 Excel 檔案
        worksheet = workbook.Sheets("all")
        'worksheet = workbook.Sheets("1")
        'worksheet.Rows(1).delete
        Debug.Print(Strings.Left(worksheet.Cells(1, 2).value, 1))

        curNo = 3642
        Dim strr As String = ""
        Dim st As String = "1"
        Dim tst As String = ""
        For i = 3642 To 4128
            st = worksheet.Cells(i, 2).value
            If st <> strr Then
                tst += """" & st & ""","
                strr = st
            End If
            Debug.Print(tst)

        Next



        workbook.Save()
        workbook.Close()
        app.Quit()
    End Sub
End Class
