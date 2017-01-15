Imports Microsoft.Office.Interop

Public Class Form1

    Private Structure student
        Dim name As String
        Dim id As String
        Dim classRoom As String
        Dim department As String
    End Structure


    Const COLOR_LOW = 220
    Const COLOR_HIGH = 255
    Private myrgb() As Integer = {COLOR_HIGH, COLOR_LOW, COLOR_LOW}
    Private n As Integer = 0
    Private colorNo(,) As Integer = {{COLOR_HIGH, COLOR_LOW, COLOR_LOW}, {COLOR_HIGH, COLOR_LOW, COLOR_HIGH}, {COLOR_LOW, COLOR_LOW, COLOR_HIGH}, {COLOR_LOW, COLOR_HIGH, COLOR_HIGH}, {COLOR_LOW, COLOR_HIGH, COLOR_LOW}, {COLOR_HIGH, COLOR_HIGH, COLOR_LOW}}

    Private tableName() As String = {"工學院", "管理學院", "電資&人文學院"}
    Private departName(,) As String = {{"技化材四甲", "四化材四乙", "碩化材二甲", "技模四甲", "四模四甲", "碩模二甲", "模具系應科碩二甲", "模具應科碩專二甲", "模具系博模三甲", "技工四甲", "四工四甲", "碩工二甲", "四土四甲", "四土四乙", "土木系碩士班二甲", "土木系碩專二甲", "土木系博士班四甲", "土木系博士班延修", "四機四甲", "四機四乙", "機械系碩士班二甲", "機械系碩專班二甲", "機械系博四甲", "機械系博四甲"} _
                                    , {"四會四甲", "碩會二甲", "四金四甲", "金融系碩士班二甲", "金融系碩專班二甲", "金融系碩專班二乙", "四國企四甲", "碩國企二甲", "國企系博士班二甲", "四財管四甲", "碩財管二甲", "四企四甲", "碩企二甲", "企管高階碩專二甲", "四觀四甲", "四觀四乙", "觀光系碩士班二甲", "觀光系碩專班二甲", "四資四甲", "四資四乙", "碩資二甲", "碩春外製管二甲", "碩秋外製管二甲", "碩秋外製管二乙"} _
                                    , {"四資工四甲", "碩資工二甲", "四電四甲", "四電四乙", "四電四丙", "碩電二甲", "電機工程系 博士班博電四甲", "四子四甲", "四子四乙", "四子四丙", "碩子二甲", "電子系博子三甲", "電子系博子延修", "四外四甲", "應外系碩士班二甲", "四人四甲", "碩人二甲", "四文創四甲", "四文創延修", "文創系碩士班二甲", "碩春外資電二甲", "碩秋外資電二甲", "光半導產碩秋二甲", "碩光電二甲"} _
                                    , {"會四甲", "金四甲", "國企四甲", "土四甲", "機四甲", "電四甲", "子四甲", "模四甲", "工四甲", "外四甲", "外延修", "企四甲", "企四乙", "觀四甲", "會四甲", "金四甲", "國企四甲", "土四甲", "機四甲", "電四甲", "子四甲", "模四甲", "工四甲", "外四甲"}}
    Private dist() As Integer = {1, 1423, 2563, 3642, 4129}
    Private departCh As Integer = 0
    Private stud(4000) As student
    Private indexN As Integer = 0
    Private winnerList(4000) As Integer

    Private t2c As Integer = 0
    Private t2mod As Integer = 1

    Private bt1mod As Integer = 1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If bt1mod = 1 Then
            clrLable()
            Debug.Print(Application.StartupPath)
            Button1.Enabled = False
            Dim app As New Excel.Application 'app 是操作 Excel 的變數
            Dim worksheet As Excel.Worksheet 'Worksheet 代表的是 Excel 工作表
            Dim workbook As Excel.Workbook 'Workbook 代表的是一個 Excel 本體
            Dim path As String
            Dim winner As Integer
            Dim rep As Boolean = False

            'Dim spaceN, tabN, t, stdN As Integer


            path = Application.StartupPath + "\studentList.xlsx"
            workbook = app.Workbooks.Open(path) '開啟一張已存在的 Excel 檔案
            worksheet = workbook.Sheets("all")

            'Dim cls As String = ""
            'Dim cour As String = ""
            'Dim sss As String = "{"""
            'sss = departName(1, 23)
            'For i = 2700 To 3641
            '    cour = worksheet.Cells(i, 2).value
            '    If cls <> cour Then
            '        cls = cour
            '        sss += cls + ""","""
            '    End If
            'Next
            'Debug.Print(sss)


            Randomize()
            Debug.Print(winnerList(0))
            While rep = False
                winner = Int(Rnd() * (dist(departCh + 1) - dist(departCh))) + dist(departCh)
                rep = True
                For i = 0 To indexN
                    If winnerList(i) = winner Then
                        rep = False
                        Exit For
                    End If

                Next

            End While

            winnerList(indexN) = winner

            stud(indexN).department = worksheet.Cells(winner, 1).value
            stud(indexN).classRoom = worksheet.Cells(winner, 2).value
            stud(indexN).id = worksheet.Cells(winner, 3).value
            stud(indexN).name = worksheet.Cells(winner, 4).value
            ListBox1.Items.Add(stud(indexN).id + vbTab + stud(indexN).classRoom + vbTab + stud(indexN).name)


            indexN += 1

            'Dim exapp As New Excel.Application
            'Dim exworksh As Excel.Worksheet
            'Dim exworkbo As Excel.Workbook
            'exworkbo = exapp.Workbooks.Open(Application.StartupPath + "\studentList.xlsx")
            'exworksh = exworkbo.Sheets("all")

            'stdN = 0
            'For i = 0 To 2

            '    worksheet = workbook.Worksheets(tableName(i))
            '    spaceN = 0
            '    tabN = 0
            '    n = 1
            '    t = 1
            '    While tabN <> 2
            '        If worksheet.Cells(n, t).value = "" Then
            '            spaceN += 1
            '        Else
            '            spaceN = 0
            '            tabN = 0
            '            stud(stdN).department = tableName(i)
            '            stud(stdN).classRoom = worksheet.Cells(n, t).value
            '            stud(stdN).id = worksheet.Cells(n, t + 1).Value
            '            stud(stdN).name = worksheet.Cells(n, t + 2).value

            '            exworksh.Cells(stdN + 1, 1) = stud(stdN).department
            '            exworksh.Cells(stdN + 1, 2) = stud(stdN).classRoom
            '            exworksh.Cells(stdN + 1, 3) = stud(stdN).id
            '            exworksh.Cells(stdN + 1, 4) = stud(stdN).name

            '            ListBox1.Items.Add(stud(stdN).name)
            '            stdN += 1
            '        End If
            '        If spaceN >= 2 Then
            '            tabN += 1
            '            t += 4
            '            n = 0
            '        End If
            '        n += 1
            '    End While

            'Next
            workbook.Close() '關閉檔案
            app.Quit() '結束操作
            'exworkbo.Save()
            'exworkbo.Close()
            'exapp.Quit()
            Timer2.Interval = 10
            Timer2.Enabled = True
            Button1.Text = "停止"
            bt1mod = 2
            Button1.Enabled = True
        Else
            Button1.Enabled = False
            Button1.Text = "開始抽獎"
            bt1mod = 1
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        For i = 0 To 2
            If myrgb(i) > colorNo(n, i) Then myrgb(i) -= 1
            If myrgb(i) < colorNo(n, i) Then myrgb(i) += 1
        Next
        Button1.BackColor = ColorTranslator.FromOle(RGB(myrgb(0), myrgb(1), myrgb(2)))
        Label1.ForeColor = ColorTranslator.FromOle(RGB(myrgb(0), myrgb(1), myrgb(2)))
        Label2.ForeColor = ColorTranslator.FromOle(RGB(myrgb(0), myrgb(1), myrgb(2)))
        Label3.ForeColor = ColorTranslator.FromOle(RGB(myrgb(0), myrgb(1), myrgb(2)))
        Label4.ForeColor = ColorTranslator.FromOle(RGB(myrgb(0), myrgb(1), myrgb(2)))
        For i = 0 To 2
            If myrgb(i) <> colorNo(n, i) Then Exit Sub
        Next
        n += 1
        n = n Mod 6
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.Click
        If RadioButton4.Checked Then
            Label2.Text = RadioButton4.Text
            departCh = 3
            clrLable()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.Click
        If RadioButton1.Checked Then
            Label2.Text = RadioButton1.Text
            departCh = 0
            clrLable()
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.Click
        If RadioButton3.Checked Then
            Label2.Text = RadioButton3.Text
            departCh = 2
            clrLable()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.Click
        If RadioButton2.Checked Then
            Label2.Text = RadioButton2.Text
            departCh = 1
            clrLable()
        End If
    End Sub



    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim strlen As Integer = Strings.Len(stud(indexN - 1).id)
        Dim stdIdStep = 100 'x10ms
        If t2mod = 1 Then
            Label1.Text = departName(departCh, t2c Mod 24)
            t2c += 1
            t2c = t2c Mod 24
            If bt1mod = 1 Then
                t2c = 0
                t2mod = 2
                Label1.Text = stud(indexN - 1).classRoom
                Timer2.Interval = 500
            End If
        Else
            t2c += 1
            If t2c > strlen - 2 Then
                Timer2.Interval = 10
                Label3.Text = Strings.Left(stud(indexN - 1).id, strlen - 2 + (t2c - strlen + 2) \ (stdIdStep + 1)) & (t2c Mod 9)
                If t2c >= strlen + stdIdStep * 2 Then
                    t2c = 0
                    t2mod = 1
                    'Label3.Text = stud(indexN - 1).id
                    Label3.Text = stud(indexN - 1).id
                    Label4.Text = stud(indexN - 1).name
                    Timer2.Enabled = False
                    Timer2.Interval = 10
                    Button1.Enabled = True
                End If

            Else
                Label3.Text = Strings.Left(stud(indexN - 1).id, t2c)
            End If

        End If

    End Sub

    Private deltaF As Integer = 839

    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        TabControl1.Height = Me.Height - 39
        TabControl1.Width = Me.Width - 16
        Label1.Width = TabControl1.Width - 22
        Label2.Width = Label1.Width
        Label3.Width = Label1.Width
        Label4.Width = Label1.Width
        Button1.Width = Label1.Width
        PictureBox1.Location = New Point((Label1.Width - PictureBox1.Width) \ 2, TabControl1.Height - 94)


        Dim pt As Single
        pt = autoFont(Label1.Width / 16)
        If pt > autoFont(Label1.Height / 1.3) Then pt = autoFont(Label1.Height / 1.3)
        Label1.Font = New Font("微軟正黑體", pt)
        Label2.Font = New Font("微軟正黑體", pt)
        Label3.Font = New Font("微軟正黑體", pt)
        Label4.Font = New Font("微軟正黑體", pt)
        Button1.Font = New Font("微軟正黑體", pt)

        Dim delT As Integer
        Dim p As Point
        delT = Me.Height - deltaF
        If Math.Abs(delT) >= 4 Then
            delT \= 4
            Label1.Height += delT
            Label2.Height += delT
            Label3.Height += delT
            Label4.Height += delT
            p = New Point(0, 6 + Label1.Height)
            Label1.Location = Label2.Location + p
            Label3.Location = Label1.Location + p
            Label4.Location = Label3.Location + p
            deltaF = Me.Height
        End If
        Button1.Location = New Point(8, Label4.Location.Y + 6 + Label1.Height)
        Button1.Height = Label1.Height

    End Sub

    Private Sub clrLable()
        Label1.Text = ""
        Label3.Text = ""
        Label4.Text = ""
    End Sub

    Private Function autoFont(ByVal px As Single) As Single
        Dim g As Graphics = Me.CreateGraphics   'calc pt px
        Dim pt As Single
        pt = px * 72 / g.DpiX
        g.Dispose()
        Return pt
    End Function
End Class
