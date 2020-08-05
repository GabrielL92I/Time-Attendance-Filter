Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows
Imports Microsoft.Office.Interop.Excel
Public Class Merged_List
    Private ReadOnly Required As String = "Time Out"
    Dim OrigLst As New List(Of String)
    Dim xlApp As Application
    Dim xlBook As Workbook
    Dim xlSheet As Worksheet



    Dim st, sm As String
    Private ReadOnly x As String

    Dim totalTime1 As String





    Private Sub OriginButton3_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles OriginButton3.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Dim pth1 As String
    Private Sub OriginButton3_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles OriginButton3.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
            For Each path In files
                pth1 = path
            Next
        End If
        Dim shkurt1 As String = pth1.Substring(pth1.LastIndexOf("\")).Replace("\", "")
        Label25.Text = Label9.Text
        Label9.Text = shkurt1

        Label9.ForeColor = Color.ForestGreen

    End Sub





    Private Sub Merged_List_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form1.OriginCheckBox2.Checked = True Then

            OriginCheckBox2.Checked = True

            TextBox2.Text = Form1.TextBox2.Text
        Else

            OriginCheckBox2.Checked = False

            TextBox2.Text = ""
        End If



        Dim countsun As Integer = 0
        Dim countsat As Integer = 0
        Dim nonholiday As Integer = 0
        Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
        For i = 0 To totalDays
            Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
            If Weekday = DayOfWeek.Saturday Then
                countsat += 1
            End If
            If Weekday = DayOfWeek.Sunday Then
                countsun += 1
            End If
            If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                nonholiday += 1
            End If
        Next
        Label21.Text = "(" & nonholiday & ")"


        If CInt(DateTimePicker1.Value.ToString.Split("/"c)(1)) > 1 Or CInt(DateTimePicker2.Value.ToString.Split("/"c)(1)) < 30 Then
            OriginRadioButton1.Checked = True
            OriginRadioButton2.Enabled = False
            OriginRadioButton1.Enabled = True
        Else
            OriginRadioButton2.Checked = True
            OriginRadioButton1.Enabled = False
            OriginRadioButton2.Enabled = True
        End If
        Dim currentTime As Date = Date.Now
        If currentTime.Month = 1 Then
            Me.ComboBox1.SelectedIndex = 0
        ElseIf currentTime.Month = 2 Then
            Me.ComboBox1.SelectedIndex = 1
        ElseIf currentTime.Month = 3 Then
            Me.ComboBox1.SelectedIndex = 2
        ElseIf currentTime.Month = 4 Then
            Me.ComboBox1.SelectedIndex = 3
        ElseIf currentTime.Month = 5 Then
            Me.ComboBox1.SelectedIndex = 4
        ElseIf currentTime.Month = 6 Then
            Me.ComboBox1.SelectedIndex = 5
        ElseIf currentTime.Month = 7 Then
            Me.ComboBox1.SelectedIndex = 6
        ElseIf currentTime.Month = 8 Then
            Me.ComboBox1.SelectedIndex = 7
        ElseIf currentTime.Month = 9 Then
            Me.ComboBox1.SelectedIndex = 8
        ElseIf currentTime.Month = 10 Then
            Me.ComboBox1.SelectedIndex = 9
        ElseIf currentTime.Month = 11 Then
            Me.ComboBox1.SelectedIndex = 10
        ElseIf currentTime.Month = 12 Then
            Me.ComboBox1.SelectedIndex = 11
        End If

        ListBox5.Sorted = True


        If My.Settings.list.Count = 0 Then

        Else
            Dim strings1(My.Settings.list.Count - 1) As String
            My.Settings.list.CopyTo(strings1, 0)
            If strings1.Contains("Test") Then

            Else
                ListBox5.Items.AddRange(strings1)
            End If

        End If





        Label17.Text = "(" & ListBox5.Items.Count.ToString & ")"
        Label17.ForeColor = Color.ForestGreen
        'Label5.Text = Label17.Text
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "yyyy-MM-dd"
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "yyyy-MM-dd"


        DateTimePicker1.Value = Form1.DateTimePicker1.Value
        DateTimePicker2.Value = Form1.DateTimePicker2.Value
    End Sub

    Private Sub OriginButton3_Click(sender As Object, e As EventArgs) Handles OriginButton3.Click
        Dim open As New OpenFileDialog With {
       .Filter = "Excel File(*.xls,.xlsx*)|*.xls*"
   }
        If open.ShowDialog() = DialogResult.OK Then
            Label9.Text = (open.FileName)
            Dim strFilename As String
            strFilename = Trim(Label9.Text)

            Dim shkurt1 As String = strFilename.Substring(strFilename.LastIndexOf("\")).Replace("\", "")
            Label25.Text = Label9.Text
            Label9.Text = shkurt1

            Label9.ForeColor = Color.ForestGreen
        End If
    End Sub
    Dim tsSum As TimeSpan
    Private Sub OriginButton1_Click(sender As Object, e As EventArgs) Handles OriginButton1.Click

        ListBox4.Items.Clear()

        ListBox6.Items.Clear()

        For Each itm1 In Form1.ListBox4.Items









            For Each itm2 In Second_List.ListBox4.Items

                If (itm1.ToString.Split(" "c)(0) + itm1.ToString.Split(" "c)(1) + itm1.ToString.Split(" "c)(2)) = (itm2.ToString.Split(" "c)(0) + itm2.ToString.Split(" "c)(1) + itm2.ToString.Split(" "c)(2)) Then

                    If itm1.ToString.Split(" "c)(3).ToString.Contains("Absent") And itm2.ToString.Split(" "c)(3).ToString.Contains("Absent") Then
                        Dim ts1 As TimeSpan = TimeSpan.Parse("00:00:00")
                        Dim ts2 As TimeSpan = TimeSpan.Parse("00:00:00")
                        tsSum = ts1 + ts2
                    ElseIf itm1.ToString.Split(" "c)(3).ToString.Contains("Absent") Then

                        Dim ts1 As TimeSpan = TimeSpan.Parse("00:00:00")
                        Dim ts2 As TimeSpan = TimeSpan.Parse(itm2.ToString.Split(" "c)(3))
                        tsSum = ts1 + ts2
                    ElseIf itm2.ToString.Split(" "c)(3).ToString.Contains("Absent") Then


                        Dim ts1 As TimeSpan = TimeSpan.Parse(itm1.ToString.Split(" "c)(3))
                        Dim ts2 As TimeSpan = TimeSpan.Parse("00:00:00")
                        tsSum = ts1 + ts2
                    Else
                        Dim ts1 As TimeSpan = TimeSpan.Parse(itm1.ToString.Split(" "c)(3))
                        Dim ts2 As TimeSpan = TimeSpan.Parse(itm2.ToString.Split(" "c)(3))
                        tsSum = ts1 + ts2
                    End If




                    If tsSum.ToString = "00:00:00" Then
                        Me.ListBox4.Items.Add(itm1.ToString.Split(" "c)(0) & " " & itm1.ToString.Split(" "c)(1) & " " & itm1.ToString.Split(" "c)(2) & " " & "Absent")
                    Else
                        Me.ListBox4.Items.Add(itm1.ToString.Split(" "c)(0) & " " & itm1.ToString.Split(" "c)(1) & " " & itm1.ToString.Split(" "c)(2) & " " & tsSum.ToString)
                    End If



                End If

            Next

        Next



        If Form1.OriginCheckBox2.Checked = True Then
            ' Dim x, y As TimeSpan
            'Dim z As TimeSpan = TimeSpan.Parse("00:01:00")
            Dim totalHours1 As Integer
            Dim totalMinutes1 As Integer
            Dim totalseconds1 As Integer
            Dim minutesx, mm As Integer
            For Each oItem In ListBox4.Items
                Dim emer As String
                emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                If emer.Contains(Form1.TextBox2.Text) And emer.Length = Form1.TextBox2.Text.Length Then
                    If emer.Contains(Form1.TextBox2.Text) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
                        totalHours1 += oItem.Split(" "c)(3).Split(":")(0)
                        totalMinutes1 += oItem.Split(" "c)(3).Split(":")(1)
                        totalseconds1 += oItem.Split(" "c)(3).Split(":")(2)

                        'x += TimeSpan.Parse(oItem.Split(" "c)(3).ToString)



                    End If
                End If
            Next oItem
            Dim sec = totalseconds1 Mod 60

            Dim remainder1 = totalMinutes1 Mod 60




            minutesx = totalMinutes1 \ 60 + (totalseconds1 \ 60)
            mm = (remainder1 + (totalseconds1 \ 60)) Mod 60




            totalHours1 += totalMinutes1 \ 60 + minutesx \ 60












            totalTime1 = totalHours1.ToString("d2") & ":" & mm.ToString("d2") & ":" & sec.ToString("d2")


            Dim countsun As Integer = 0
            Dim countsat As Integer = 0
            Dim nonholiday As Integer = 0
            Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
            For i = 0 To totalDays
                Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
                If Weekday = DayOfWeek.Saturday Then
                    countsat += 1
                End If
                If Weekday = DayOfWeek.Sunday Then
                    countsun += 1
                End If
                If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                    nonholiday += 1
                End If
            Next



            Dim average
            If Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")) = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd")) Then

                ListBox6.Items.Add(Form1.TextBox2.Text & " " & totalTime1.ToString)
            Else
                If totalTime1.ToString = "00:00:00" Then
                    average = "00:00"
                Else
                    Dim result As Integer = totalTime1.ToString.Split(":")(0) * 60 + totalTime1.ToString.ToString.Split(":")(1) + (totalTime1.ToString.ToString.Split(":")(2) \ 60)
                    Dim days As Integer = result \ nonholiday

                    Dim hours As Integer = days \ 60
                    Dim minutes As Integer = days - (hours * 60)
                    Dim timeElapsed As String = CType(hours.ToString("d2"), String) & ":" & CType(minutes.ToString("d2"), String)
                    average = timeElapsed
                End If
                ListBox6.Items.Add(Form1.TextBox2.Text & " " & totalTime1.ToString & " " & "(" & average & ")")

            End If










        Else

            Dim nn(ListBox5.Items.Count) As String
            ListBox5.Items.CopyTo(nn, 0)
            Dim totalHours1 As Integer
            Dim totalMinutes1 As Integer
            Dim totalseconds1 As Integer
            Dim minutesx, mm As Integer
            For index As Integer = 0 To ListBox5.Items.Count - 1
                For Each oItem In ListBox4.Items
                    Dim emer As String
                    emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                    If emer.Contains(nn(index)) And emer.Length = nn(index).Length Then
                        If emer.Contains(nn(index)) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
                            ' x += TimeSpan.Parse(oItem.Split(" "c)(3).ToString)
                            totalHours1 += oItem.Split(" "c)(3).Split(":")(0)
                            totalMinutes1 += oItem.Split(" "c)(3).Split(":")(1)
                            totalseconds1 += oItem.Split(" "c)(3).Split(":")(2)
                        End If
                    End If
                Next oItem



                Dim sec = totalseconds1 Mod 60
                Dim remainder1 = totalMinutes1 Mod 60
                minutesx = totalMinutes1 \ 60 + (totalseconds1 \ 60)
                mm = (remainder1 + (totalseconds1 \ 60)) Mod 60
                totalseconds1 = 0
                totalHours1 += totalMinutes1 \ 60 + minutesx \ 60
                totalTime1 = totalHours1.ToString("d2") & ":" & mm.ToString("d2") & ":" & sec.ToString("d2")

                totalHours1 = 0
                totalMinutes1 = 0
                remainder1 = 0
                minutesx = 0
                sec = 0
                Dim countsun As Integer = 0
                Dim countsat As Integer = 0
                Dim nonholiday As Integer = 0
                Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
                For i = 0 To totalDays
                    Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
                    If Weekday = DayOfWeek.Saturday Then
                        countsat += 1
                    End If
                    If Weekday = DayOfWeek.Sunday Then
                        countsun += 1
                    End If
                    If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                        nonholiday += 1
                    End If
                Next
                Dim average
                If Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")) = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd")) Then

                    ListBox6.Items.Add(nn(index) & " " & totalTime1.ToString)
                Else
                    If totalTime1.ToString = "00:00:00" Then
                        average = "00:00"
                    Else
                        Dim result As Integer = totalTime1.ToString.Split(":")(0) * 60 + totalTime1.ToString.Split(":")(1) + (totalTime1.ToString.Split(":")(2) \ 60)
                        Dim days As Integer = result / nonholiday

                        Dim hours As Integer = days \ 60
                        Dim minutes As Integer = days - (hours * 60)







                        Dim timeElapsed As String = CType(hours.ToString("d2"), String) & ":" & CType(minutes.ToString("d2"), String)
                        average = timeElapsed
                    End If
                    ListBox6.Items.Add(nn(index) & " " & totalTime1.ToString & " " & "(" & average & ")")

                End If


            Next
        End If









    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If DateTimePicker1.Value > DateTimePicker2.Value Then
            DateTimePicker1.Value = Date.Now
        End If
        If CInt(DateTimePicker1.Value.ToString.Split("/"c)(1)) > 1 Or CInt(DateTimePicker2.Value.ToString.Split("/"c)(1)) < 30 Then
            OriginRadioButton1.Checked = True
            OriginRadioButton2.Enabled = False
            OriginRadioButton1.Enabled = True
        Else
            OriginRadioButton2.Checked = True
            OriginRadioButton1.Enabled = False
            OriginRadioButton2.Enabled = True
        End If
        Dim countsun As Integer = 0
        Dim countsat As Integer = 0
        Dim nonholiday As Integer = 0
        Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
        For i = 0 To totalDays
            Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
            If Weekday = DayOfWeek.Saturday Then
                countsat += 1
            End If
            If Weekday = DayOfWeek.Sunday Then
                countsun += 1
            End If
            If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                nonholiday += 1
            End If
        Next
        Label21.Text = "(" & nonholiday & ")"
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        If DateTimePicker2.Value < DateTimePicker1.Value Then
            DateTimePicker2.Value = Date.Now
        End If
        If CInt(DateTimePicker1.Value.ToString.Split("/"c)(1)) > 1 Or CInt(DateTimePicker2.Value.ToString.Split("/"c)(1)) < 30 Then
            OriginRadioButton1.Checked = True
            OriginRadioButton2.Enabled = False
            OriginRadioButton1.Enabled = True
        Else
            OriginRadioButton2.Checked = True
            OriginRadioButton1.Enabled = False
            OriginRadioButton2.Enabled = True
        End If
        Dim countsun As Integer = 0
        Dim countsat As Integer = 0
        Dim nonholiday As Integer = 0
        Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
        For i = 0 To totalDays
            Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
            If Weekday = DayOfWeek.Saturday Then
                countsat += 1
            End If
            If Weekday = DayOfWeek.Sunday Then
                countsun += 1
            End If
            If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                nonholiday += 1
            End If
        Next
        Label21.Text = "(" & nonholiday & ")"
    End Sub
    Private Sub Merged_List_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Second_List.WindowState = FormWindowState.Normal

        Second_List.OriginCheckBox4.Checked = False
    End Sub
    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        Me.Close()

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub OriginButton2_Click(sender As Object, e As EventArgs) Handles OriginButton2.Click
        SaveFileDialog1.Filter = "TXT Files (*.txt*)|*.txt"
        If SaveFileDialog1.ShowDialog = Forms.DialogResult.OK _
       Then
            Dim SW As IO.StreamWriter = IO.File.CreateText(SaveFileDialog1.FileName)
            For Each S As String In ListBox4.Items
                SW.WriteLine(S)
            Next
            SW.Close()
            MsgBox("List Exported!", MsgBoxStyle.Information)

        End If
    End Sub

    Private Sub OriginButton4_Click(sender As Object, e As EventArgs) Handles OriginButton4.Click
        SaveFileDialog1.Filter = "TXT Files (*.txt*)|*.txt"
        If SaveFileDialog1.ShowDialog = Forms.DialogResult.OK _
       Then
            Dim SW As IO.StreamWriter = IO.File.CreateText(SaveFileDialog1.FileName)
            For Each S As String In ListBox6.Items
                SW.WriteLine(S)
            Next
            SW.Close()
            MsgBox("List Exported!", MsgBoxStyle.Information)

        End If
    End Sub

    Private Sub OriginButton6_Click(sender As Object, e As EventArgs) Handles OriginButton6.Click
        If Label9.Text = "No excel file loaded..." Then
            MsgBox("Load excel file first!", MsgBoxStyle.Information)
        Else

            If Me.OriginCheckBox2.Checked = True Then

                Dim nr As Integer
                If OriginRadioButton1.Checked = True Then
                    Dim nn(ListBox5.Items.Count) As String
                    ListBox5.Items.CopyTo(nn, 0)
                    Dim oItem As String
                    Dim OffS1 As Integer = 2
                    Dim OffS2 As Integer = 0
                    xlApp = GetObject("", "Excel.Application")
                    xlBook = xlApp.Workbooks.Open(Label25.Text)
                    xlSheet = xlBook.Worksheets(ComboBox1.SelectedItem)
                    xlApp.Visible = True
                    Dim totalHours1 As Integer
                    Dim totalMinutes1 As Integer
                    Dim totalseconds1 As Integer
                    Dim minutesx, mm As Integer
                    For Each oItem In Me.ListBox4.Items
                        Dim emer As String
                        emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                        If emer.Contains(TextBox2.Text) And emer.Length = TextBox2.Text.Length Then
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).ColumnWidth = 23
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(0, OffS2).Value = "Hours ↓"
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Value = TextBox2.Text
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Interior.ColorIndex = 15
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                            'xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = oItem.Split(" "c)(2) & " " & oItem.Split(" "c)(3)
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = oItem.Split(" "c)(3)

                            xlBook.Sheets(ComboBox1.SelectedItem).Range("B1").Offset(OffS1, 0).Value = oItem.Split(" "c)(2).Split("-"c)(1) & "/" & oItem.Split(" "c)(2).Split("-"c)(2) & "/" & oItem.Split(" "c)(2).Split("-"c)(0)
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("B1").Offset(OffS1, 0).HorizontalAlignment = 2
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).HorizontalAlignment = 2
                            If emer.Contains(TextBox2.Text) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
                                totalHours1 += oItem.Split(" "c)(3).Split(":")(0)
                                totalMinutes1 += oItem.Split(" "c)(3).Split(":")(1)
                                totalseconds1 += oItem.Split(" "c)(3).Split(":")(2)
                                'x += TimeSpan.Parse(oItem.Split(" "c)(3).ToString)
                            End If
                            OffS1 = OffS1 + 1
                        End If
                    Next oItem
                    nr = OffS1
                    'xlBook.Sheets(ComboBox1.SelectedItem).Rows("34").delete()
                    Dim sec = totalseconds1 Mod 60

                    Dim remainder1 = totalMinutes1 Mod 60




                    minutesx = totalMinutes1 \ 60 + (totalseconds1 \ 60)
                    mm = (remainder1 + (totalseconds1 \ 60)) Mod 60




                    totalHours1 += totalMinutes1 \ 60 + minutesx \ 60












                    totalTime1 = totalHours1.ToString("d2") & ":" & mm.ToString("d2") & ":" & sec.ToString("d2")
                    totalHours1 = 0
                    totalMinutes1 = 0



                    Dim countsun As Integer = 0
                    Dim countsat As Integer = 0
                    Dim nonholiday As Integer = 0
                    Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
                    For i = 0 To totalDays
                        Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
                        If Weekday = DayOfWeek.Saturday Then
                            countsat += 1
                        End If
                        If Weekday = DayOfWeek.Sunday Then
                            countsun += 1
                        End If
                        If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                            nonholiday += 1
                        End If
                    Next
                    Dim average
                    If Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")) = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd")) Then
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString
                        xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"

                    Else
                        If totalTime1.ToString = "00:00:00" Then
                            average = "00:00"
                        Else
                            Dim result As Integer = totalTime1.ToString.Split(":")(0) * 60 + totalTime1.ToString.Split(":")(1) + (totalTime1.ToString.Split(":")(2) \ 60)
                            Dim days As Integer = result / nonholiday

                            Dim hours As Integer = days \ 60
                            Dim minutes As Integer = days - (hours * 60)
                            Dim timeElapsed As String = CType(hours.ToString("d2"), String) & ":" & CType(minutes.ToString("d2"), String)
                            average = timeElapsed
                        End If
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " " & "(" & average & ")"
                        xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                    End If





                    xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Interior.ColorIndex = 50
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                    OffS2 = OffS2 + 1
                    OffS1 = 2

                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Value = "Name"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Borders.LineStyle = XlLineStyle.xlContinuous

                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(1, 2).Value = "Month"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Range(xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0), xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 1)).Merge
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).Value = "Total hours + Average hours"
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).HorizontalAlignment = 3
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 1).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Save()
                    xlBook.Close(True)
                    xlApp.Quit()
                    xlSheet = Nothing
                    xlBook = Nothing
                    Dim proc As Process
                    For Each proc In Process.GetProcessesByName("EXCEL")
                        proc.Kill()
                    Next
                    MsgBox("Export Done!", MsgBoxStyle.Information)
                Else
                    Dim nn(ListBox5.Items.Count) As String
                    ListBox5.Items.CopyTo(nn, 0)
                    Dim oItem As String
                    Dim OffS1 As Integer = 2
                    Dim OffS2 As Integer = 0
                    xlApp = GetObject("", "Excel.Application")
                    xlBook = xlApp.Workbooks.Open(Label25.Text)
                    xlSheet = xlBook.Worksheets(ComboBox1.SelectedItem)
                    xlApp.Visible = True
                    Dim totalHours1 As Integer
                    Dim totalMinutes1 As Integer
                    Dim totalseconds1 As Integer
                    Dim minutesx, mm As Integer
                    For Each oItem In ListBox4.Items
                        Dim emer As String
                        emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                        If emer.Contains(TextBox2.Text) And emer.Length = TextBox2.Text.Length Then
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).ColumnWidth = 23
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(0, OffS2).Value = "Hours ↓"
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Value = TextBox2.Text
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Interior.ColorIndex = 15
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = oItem.Split(" "c)(3)
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).HorizontalAlignment = 2
                            If emer.Contains(TextBox2.Text) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
                                totalHours1 += oItem.Split(" "c)(3).Split(":")(0)
                                totalMinutes1 += oItem.Split(" "c)(3).Split(":")(1)
                                totalseconds1 += oItem.Split(" "c)(3).Split(":")(2)
                                'x += TimeSpan.Parse(oItem.Split(" "c)(3).ToString)
                            End If
                            OffS1 = OffS1 + 1
                        End If
                    Next oItem
                    'xlBook.Sheets(ComboBox1.SelectedItem).Rows("34").delete()
                    Dim sec = totalseconds1 Mod 60
                    Dim remainder1 = totalMinutes1 Mod 60
                    minutesx = totalMinutes1 \ 60 + (totalseconds1 \ 60)
                    mm = (remainder1 + (totalseconds1 \ 60)) Mod 60
                    totalHours1 += totalMinutes1 \ 60 + minutesx \ 60
                    totalTime1 = totalHours1.ToString("d2") & ":" & mm.ToString("d2") & ":" & sec.ToString("d2")
                    totalHours1 = 0
                    totalMinutes1 = 0

                    '

                    Dim countsun As Integer = 0
                    Dim countsat As Integer = 0
                    Dim nonholiday As Integer = 0
                    Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
                    For i = 0 To totalDays
                        Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
                        If Weekday = DayOfWeek.Saturday Then
                            countsat += 1
                        End If
                        If Weekday = DayOfWeek.Sunday Then
                            countsun += 1
                        End If
                        If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                            nonholiday += 1
                        End If
                    Next
                    Dim average
                    If Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")) = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd")) Then
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString

                        xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                    Else
                        If totalTime1.ToString = "00:00:00" Then
                            average = "00:00"
                        Else
                            Dim result As Integer = totalTime1.ToString.Split(":")(0) * 60 + totalTime1.ToString.Split(":")(1) + (totalTime1.ToString.Split(":")(2) \ 60)
                            Dim days As Integer = result / nonholiday

                            Dim hours As Integer = days \ 60
                            Dim minutes As Integer = days - (hours * 60)
                            Dim timeElapsed As String = CType(hours.ToString("d2"), String) & ":" & CType(minutes.ToString("d2"), String)
                            average = timeElapsed
                        End If
                        xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " " & "(" & average & ")"
                    End If



                    xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Interior.ColorIndex = 50
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                    OffS2 = OffS2 + 1
                    OffS1 = 2

                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Value = "Name"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Borders.LineStyle = XlLineStyle.xlContinuous

                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Range(xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1), xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 2)).Merge
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).Value = "Total hours + Average hours"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).HorizontalAlignment = 3
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 2).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Save()
                    xlBook.Close(True)
                    xlApp.Quit()
                    xlSheet = Nothing
                    xlBook = Nothing
                    Dim proc As Process
                    For Each proc In Process.GetProcessesByName("EXCEL")
                        proc.Kill()
                    Next
                    MsgBox("Export Done!", MsgBoxStyle.Information)
                End If




            Else
                Dim x As TimeSpan
                Dim nr As Integer
                If Me.OriginRadioButton1.Checked = True Then
                    Dim nn(Me.ListBox5.Items.Count) As String
                    Me.ListBox5.Items.CopyTo(nn, 0)
                    Dim oItem As String
                    Dim OffS1 As Integer = 2
                    Dim OffS2 As Integer = 0
                    xlApp = GetObject("", "Excel.Application")
                    xlBook = xlApp.Workbooks.Open(Me.Label25.Text)
                    xlSheet = xlBook.Worksheets(Me.ComboBox1.SelectedItem)
                    xlApp.Visible = True
                    Dim totalHours1 As Integer
                    Dim totalMinutes1 As Integer
                    Dim totalseconds1 As Integer
                    Dim minutesx, mm As Integer
                    For index As Integer = 0 To Me.ListBox5.Items.Count - 1
                        For Each oItem In Me.ListBox4.Items
                            Dim emer As String
                            emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                            If emer.Contains(nn(index)) And emer.Length = nn(index).Length Then
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).ColumnWidth = 23
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(0, OffS2).Value = "Hours ↓"
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Value = nn(index)
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Interior.ColorIndex = 15
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                                'xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = oItem.Split(" "c)(2) & " " & oItem.Split(" "c)(3)
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = oItem.Split(" "c)(3)

                                xlBook.Sheets(ComboBox1.SelectedItem).Range("B1").Offset(OffS1, 0).Value = oItem.Split(" "c)(2).Split("-"c)(1) & "/" & oItem.Split(" "c)(2).Split("-"c)(2) & "/" & oItem.Split(" "c)(2).Split("-"c)(0)
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("B1").Offset(OffS1, 0).HorizontalAlignment = 2

                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).HorizontalAlignment = 2
                                If emer.Contains(nn(index)) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
                                    totalHours1 += oItem.Split(" "c)(3).Split(":")(0)
                                    totalMinutes1 += oItem.Split(" "c)(3).Split(":")(1)
                                    totalseconds1 += oItem.Split(" "c)(3).Split(":")(2)
                                    ' x += TimeSpan.Parse(oItem.Split(" "c)(3).ToString)
                                End If
                                OffS1 = OffS1 + 1
                            End If
                        Next oItem
                        nr = OffS1

                        'xlBook.Sheets(ComboBox1.SelectedItem).Rows("34").delete()
                        Dim sec = totalseconds1 Mod 60
                        Dim remainder1 = totalMinutes1 Mod 60
                        minutesx = totalMinutes1 \ 60 + (totalseconds1 \ 60)
                        mm = (remainder1 + (totalseconds1 \ 60)) Mod 60
                        totalseconds1 = 0
                        totalHours1 += totalMinutes1 \ 60 + minutesx \ 60
                        totalTime1 = totalHours1.ToString("d2") & ":" & mm.ToString("d2") & ":" & sec.ToString("d2")
                        totalHours1 = 0
                        totalMinutes1 = 0



                        Dim countsun As Integer = 0
                        Dim countsat As Integer = 0
                        Dim nonholiday As Integer = 0
                        Dim totalDays = (Me.DateTimePicker2.Value - Me.DateTimePicker1.Value).Days
                        For i = 0 To totalDays
                            Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
                            If Weekday = DayOfWeek.Saturday Then
                                countsat += 1
                            End If
                            If Weekday = DayOfWeek.Sunday Then
                                countsun += 1
                            End If
                            If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                                nonholiday += 1
                            End If
                        Next
                        Dim average
                        If Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")) = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd")) Then
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString
                            xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"

                        Else
                            If totalTime1.ToString = "00:00:00" Then
                                average = "00:00"
                            Else
                                Dim result As Integer = totalTime1.ToString.Split(":")(0) * 60 + totalTime1.ToString.Split(":")(1) + (totalTime1.ToString.Split(":")(2) \ 60)
                                Dim days As Integer = result / nonholiday

                                Dim hours As Integer = days \ 60
                                Dim minutes As Integer = days - (hours * 60)
                                Dim timeElapsed As String = CType(hours.ToString("d2"), String) & ":" & CType(minutes.ToString("d2"), String)
                                average = timeElapsed
                            End If
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " " & "(" & average & ")"
                            xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                        End If


                        x = Nothing


                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Interior.ColorIndex = 50
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                        OffS2 = OffS2 + 1
                        OffS1 = 2
                    Next
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Value = "Name"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Borders.LineStyle = XlLineStyle.xlContinuous

                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(1, 2).Value = "Month"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Range(xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0), xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 1)).Merge
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).Value = "Total hours + Average hours"
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).HorizontalAlignment = 3
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 1).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Range("A1").Offset(nr, 0).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Save()
                    xlBook.Close(True)
                    xlApp.Quit()
                    xlSheet = Nothing
                    xlBook = Nothing
                    Dim proc As Process
                    For Each proc In Process.GetProcessesByName("EXCEL")
                        proc.Kill()
                    Next
                    MsgBox("Export Done!", MsgBoxStyle.Information)
                Else

                    Dim nn(ListBox5.Items.Count) As String
                    ListBox5.Items.CopyTo(nn, 0)
                    Dim oItem As String
                    Dim OffS1 As Integer = 2
                    Dim OffS2 As Integer = 0
                    xlApp = GetObject("", "Excel.Application")
                    xlBook = xlApp.Workbooks.Open(Label25.Text)
                    xlSheet = xlBook.Worksheets(ComboBox1.SelectedItem)
                    xlApp.Visible = True
                    Dim totalHours1 As Integer
                    Dim totalMinutes1 As Integer
                    Dim totalseconds1 As Integer
                    Dim minutesx, mm As Integer
                    For index As Integer = 0 To ListBox5.Items.Count - 1
                        For Each oItem In ListBox4.Items
                            Dim emer As String
                            emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                            If emer.Contains(nn(index)) And emer.Length = nn(index).Length Then
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).ColumnWidth = 23
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(0, OffS2).Value = "Hours ↓"
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Value = nn(index)
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Interior.ColorIndex = 15
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = oItem.Split(" "c)(3)
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                                xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).HorizontalAlignment = 2
                                If emer.Contains(nn(index)) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
                                    totalHours1 += oItem.Split(" "c)(3).Split(":")(0)
                                    totalMinutes1 += oItem.Split(" "c)(3).Split(":")(1)
                                    totalseconds1 += oItem.Split(" "c)(3).Split(":")(2)
                                    'x += TimeSpan.Parse(oItem.Split(" "c)(3).ToString)
                                End If
                                OffS1 = OffS1 + 1
                            End If
                        Next oItem
                        'xlBook.Sheets(ComboBox1.SelectedItem).Rows("34").delete()
                        Dim sec = totalseconds1 Mod 60
                        Dim remainder1 = totalMinutes1 Mod 60
                        minutesx = totalMinutes1 \ 60 + (totalseconds1 \ 60)
                        mm = (remainder1 + (totalseconds1 \ 60)) Mod 60
                        totalseconds1 = 0
                        totalHours1 += totalMinutes1 \ 60 + minutesx \ 60
                        totalTime1 = totalHours1.ToString("d2") & ":" & mm.ToString("d2") & ":" & sec.ToString("d2")
                        totalHours1 = 0
                        totalMinutes1 = 0



                        Dim countsun As Integer = 0
                        Dim countsat As Integer = 0
                        Dim nonholiday As Integer = 0
                        Dim totalDays = (DateTimePicker2.Value - DateTimePicker1.Value).Days
                        For i = 0 To totalDays
                            Dim Weekday As DayOfWeek = DateTimePicker1.Value.Date.AddDays(i).DayOfWeek
                            If Weekday = DayOfWeek.Saturday Then
                                countsat += 1
                            End If
                            If Weekday = DayOfWeek.Sunday Then
                                countsun += 1
                            End If
                            If Weekday <> DayOfWeek.Saturday AndAlso Weekday <> DayOfWeek.Sunday Then
                                nonholiday += 1
                            End If
                        Next
                        Dim average
                        If Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd")) = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd")) Then
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString

                            xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                        Else
                            If totalTime1.ToString = "00:00:00" Then
                                average = "00:00"
                            Else
                                Dim result As Integer = totalTime1.ToString.Split(":")(0) * 60 + totalTime1.ToString.Split(":")(1) + (totalTime1.ToString.Split(":")(2) \ 60)
                                Dim days As Integer = result / nonholiday

                                Dim hours As Integer = days \ 60
                                Dim minutes As Integer = days - (hours * 60)
                                Dim timeElapsed As String = CType(hours.ToString("d2"), String) & ":" & CType(minutes.ToString("d2"), String)
                                average = timeElapsed
                            End If
                            xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " " & "(" & average & ")"
                        End If



                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Interior.ColorIndex = 50
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Borders.LineStyle = XlLineStyle.xlContinuous
                        OffS2 = OffS2 + 1
                        OffS1 = 2
                    Next
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Value = "Name"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 1).Borders.LineStyle = XlLineStyle.xlContinuous

                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Range(xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1), xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 2)).Merge
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).Value = "Total hours + Average hours"
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).Font.Bold = True
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).HorizontalAlignment = 3
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 2).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Sheets(ComboBox1.SelectedItem).Cells(34, 1).Borders.LineStyle = XlLineStyle.xlContinuous
                    xlBook.Save()
                    xlBook.Close(True)
                    xlApp.Quit()
                    xlSheet = Nothing
                    xlBook = Nothing
                    Dim proc As Process
                    For Each proc In Process.GetProcessesByName("EXCEL")
                        proc.Kill()
                    Next
                    MsgBox("Export Done!", MsgBoxStyle.Information)
                End If
            End If

            '

        End If
    End Sub
End Class