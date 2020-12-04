Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows
Imports Microsoft.Office.Interop.Excel
Public Class Form1
    Private ReadOnly Required As String = "Time Out"
    Dim OrigLst As New List(Of String)
    Dim xlApp As Application
    Dim xlBook As Workbook
    Dim xlSheet As Worksheet
    Private Sub ListBox5_DoubleClick(sender As Object, e As EventArgs) Handles ListBox5.DoubleClick
        If ListBox5.SelectedIndex >= 0 Then
            ListBox5.SetSelected(ListBox5.SelectedIndex, False)
        Else
        End If
    End Sub
    Private Sub ListBox5_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox5.MouseDown
        If e.Button = Forms.MouseButtons.Right And Not ListBox5.SelectedIndex = -1 Then
            ContextMenuStrip1.Items(1).Visible = True
            ContextMenuStrip1.Show(MousePosition)
        ElseIf e.Button = Forms.MouseButtons.Right And ListBox5.SelectedIndex = -1 Then
            ContextMenuStrip1.Show(MousePosition)
            ContextMenuStrip1.Items(1).Visible = False
        End If
    End Sub
    Private Sub OriginButton4_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles OriginButton4.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Dim pth As String
    Private Sub OriginButton4_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles OriginButton4.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
            For Each path In files
                pth = path
            Next
        End If
        Dim shkurt As String = pth.Substring(pth.LastIndexOf("\")).Replace("\", "")
        Label8.Text = pth
        Label13.Text = "File loaded!"
        Label19.Text = shkurt
        Label19.ForeColor = Color.ForestGreen
        Label13.ForeColor = Color.ForestGreen
        Dim lines() As String = IO.File.ReadAllLines(Label8.Text)
        Dim x, y, z As String
        Dim largestday As Integer = Integer.MinValue
        Dim smallestday As Integer = Integer.MaxValue
        Dim largestmonth As Integer = Integer.MinValue
        Dim smallestmonth As Integer = Integer.MaxValue
        Dim largestyear As Integer = Integer.MinValue
        Dim smallestyear As Integer = Integer.MaxValue
        For Each element As String In lines
            x = element.ToString.Split(" ")(2).Split("-")(2)
            y = element.ToString.Split(" ")(2).Split("-")(1)
            z = element.ToString.Split(" ")(2).Split("-")(0)
            largestday = Math.Max(largestday, CInt(x))
            smallestday = Math.Min(smallestday, CInt(x))
            largestmonth = Math.Max(largestmonth, CInt(y))
            smallestmonth = Math.Min(smallestmonth, CInt(y))
            largestyear = Math.Max(largestyear, CInt(z))
            smallestyear = Math.Min(smallestyear, CInt(z))
        Next
        Dim startdt, enddt As String
        startdt = smallestyear & "-" & smallestmonth & "-" & smallestday
        enddt = largestyear & "-" & largestmonth & "-" & largestday
        DateTimePicker1.Value = startdt
        DateTimePicker2.Value = enddt
        DateTimePicker2.Enabled = True
        DateTimePicker1.Enabled = True
    End Sub
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
    Private Sub ListBox5_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox5.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
    Dim path As String
    Private Sub ListBox5_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles ListBox5.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer
            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                path = (MyFiles(i))
            Next
            Dim sr As StreamReader = New StreamReader(path)
            Dim strLine As String
            Do While sr.Peek() >= 0
                strLine = sr.ReadLine
                If ListBox5.Items.Contains(strLine) Then
                Else
                    ListBox5.Items.Add(strLine)
                End If
            Loop
            sr.Close()
        End If
        Label17.Text = "(" & ListBox5.Items.Count.ToString & ")"
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
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
        ListBox4.Sorted = True
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
    End Sub
    Function StripDups(s As String, lst As List(Of String)) As List(Of String)
        Dim inx As New List(Of Integer)
        For i As Integer = 0 To lst.Count - 2
            If lst(i).EndsWith(s) AndAlso lst(i + 1).EndsWith(s) Then
                inx.Add(i + 1)
            End If
        Next
        inx.Reverse()
        For Each i As Integer In inx
            lst.RemoveAt(i)
        Next
        Return lst
    End Function
    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim strings(ListBox5.Items.Count - 1) As String
        ListBox5.Items.CopyTo(strings, 0)
        My.Settings.list = New Specialized.StringCollection
        My.Settings.list.AddRange(strings)
        My.Settings.Save()



    End Sub
    Public Sub ReleaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Dim st, sm As String
    Private ReadOnly x As String
    Dim totalTime1
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        If ListBox5.SelectedIndex = -1 Then
            MsgBox("Select the user to remove first!", MsgBoxStyle.Information)
        Else
            ListBox5.Items.RemoveAt(ListBox5.SelectedIndex)
            Merged_List.ListBox1.Items.Remove(ListBox5.SelectedIndex)
            Label17.Text = "(" & ListBox5.Items.Count.ToString & ")"
        End If
    End Sub
    Dim iix
    Private Sub AddEmployeeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddEmployeeToolStripMenuItem.Click
        Dim StatusDate As String
        StatusDate = InputBox("Enter the employee name here", "Employee Name", "")
        If StatusDate = "" Then
        Else
            ListBox5.Items.Add(StatusDate)
            Label17.Text = "(" & ListBox5.Items.Count.ToString & ")"
        End If



        Dim StatusDate1 As String
        StatusDate1 = InputBox("Enter the employee group here", "Employee Group", "")
        If StatusDate1 = "" Then
        Else
            'Merged_List.ListBox1.Items.Add(StatusDate1)
            Merged_List.ListBox1.Items.Add(StatusDate & " " & StatusDate1)

        End If
    End Sub
    Private Sub ClearListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearListToolStripMenuItem.Click
        ListBox5.Items.Clear()
        Merged_List.ListBox1.Items.Clear()
        Label17.Text = "(" & ListBox5.Items.Count.ToString & ")"
        Dim strings1(Merged_List.ListBox1.Items.Count - 1) As String
        Merged_List.ListBox1.Items.CopyTo(strings1, 0)
        My.Settings.list2 = New Specialized.StringCollection
        My.Settings.list2.AddRange(strings1)
        My.Settings.Save() '
    End Sub
    Private Sub ExportListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportListToolStripMenuItem.Click
        SaveFileDialog1.Filter = "TXT Files (*.txt*)|*.txt"
        If SaveFileDialog1.ShowDialog = Forms.DialogResult.OK _
       Then
            Dim SW As IO.StreamWriter = IO.File.CreateText(SaveFileDialog1.FileName)
            For Each S As String In ListBox5.Items
                SW.WriteLine(S)
            Next
            SW.Close()
            MsgBox("List Exported!", MsgBoxStyle.Information)
        End If
    End Sub
    Dim fdate As String
    Dim y As String
    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        MsgBox("Enter the date format that is shown in the unformatted list!", MsgBoxStyle.Information)
    End Sub
    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        MsgBox("Working days based on the date range you choose!", MsgBoxStyle.Information)
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
    Private Sub OriginCheckBox2_CheckedChanged(sender As Object) Handles OriginCheckBox2.CheckedChanged
        If OriginCheckBox2.Checked = True Then
            TextBox2.Enabled = True
        Else
            TextBox2.Enabled = False
            TextBox2.Text = ""
        End If
    End Sub
    Private Sub OriginButton5_Click(sender As Object, e As EventArgs) Handles OriginButton5.Click
        Dim thread As New Thread(
 Sub()
     If Label19.Text = "-" Then
         MsgBox("Load the list first!", MsgBoxStyle.Information)
     Else
         Dim readingFile As System.IO.StreamReader = New System.IO.StreamReader(Label8.Text)
         Dim readingLine As String = readingFile.ReadLine()
         readingFile.Close()
         If OriginCheckBox2.Checked = True Then
             ListBox1.Items.Clear()
             ListBox2.Items.Clear()
             ListBox3.Items.Clear()
             ListBox4.Items.Clear()
             ListBox6.Items.Clear()
             Dim items() As String = ListBox5.Items.Cast(Of Object).Select(Function(o) ListBox5.GetItemText(o)).ToArray
             Dim list1 As New ArrayList
             Dim starting As DateTime = Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd"))
             Dim ending As DateTime = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd"))
             Dim dates As String() = Enumerable.Range(0, 1 + ending.Subtract(starting).Days).[Select](Function(i) starting.AddDays(i).ToString("yyyy-MM-dd")).ToArray()
             list1.AddRange(dates)
             For Each num1 In list1
                 Dim sr As StreamReader = New StreamReader(Label8.Text)
                 Dim strLine As String
                 Do While sr.Peek() >= 0
                     strLine = String.Empty
                     strLine = sr.ReadLine
                     If strLine.Contains(TextBox2.Text) And strLine.Contains(num1) Then
                         ListBox1.Items.Add(strLine)
                         OrigLst.Add(strLine)
                     End If
                 Loop
                 sr.Close()
                 For i As Integer = 0 To OrigLst.Count - 1
                     OrigLst(i) = Trim((OrigLst(i).Replace(TextBox2.Text & " ", Nothing)).ToLower)
                 Next
                 OrigLst = StripDups("c/in", OrigLst)
                 OrigLst = StripDups("c/out", OrigLst)
                 Dim dts As New List(Of DateTime)
                 Dim dt As DateTime = Nothing
                 For i As Integer = 0 To OrigLst.Count - 1
                     OrigLst(i) = OrigLst(i).Replace(" c/out", Nothing).Replace(" c/in", Nothing)
                     If DateTime.TryParse(OrigLst(i), dt) Then
                         dts.Add(dt)
                     End If
                 Next
                 For Each obj As Object In OrigLst
                     ListBox2.Items.Add(TextBox2.Text & " " & num1 & " " & obj)
                 Next
                 Dim d2 As TimeSpan
                 For i As Integer = 0 To OrigLst.Count - 2 Step 2
                     Dim d1 As TimeSpan = dts(i + 1) - dts(i)
                     ListBox3.Items.Add(TextBox2.Text + " " + num1 + " " + dts(i + 1).ToString("HH:mm:ss") & " - " & dts(i).ToString("HH:mm:ss") & " = " & d1.ToString)
                     d2 += d1
                 Next
                 If d2.ToString() = "00:00:00" Then
                     ListBox4.Items.Add(TextBox2.Text & " " + num1 & " " + "Absent")
                 Else
                     ListBox4.Items.Add(TextBox2.Text & " " + num1 & " " + d2.ToString())
                 End If
                 d2 = Nothing
                 OrigLst.Clear()
             Next
             Dim nn(ListBox5.Items.Count) As String
             ListBox5.Items.CopyTo(nn, 0)
             Dim totalHours1 As Integer
             Dim totalMinutes1 As Integer
             Dim totalseconds1 As Integer
             Dim minutesx, mm As Integer
             For Each oItem In ListBox4.Items
                 Dim emer As String
                 emer = oItem.Split(" "c)(0).ToString + " " + oItem.Split(" "c)(1).ToString
                 If emer.Contains(TextBox2.Text) And emer.Length = TextBox2.Text.Length Then
                     If emer.Contains(TextBox2.Text) And Not oItem.Split(" "c)(3).ToString.Contains("Absent") Then
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
                 ListBox6.Items.Add(TextBox2.Text & " " & totalTime1.ToString)
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
                 ListBox6.Items.Add(TextBox2.Text & " " & totalTime1.ToString & " " & "(" & average & ")")
             End If
         Else
             ListBox1.Items.Clear()
             ListBox2.Items.Clear()
             ListBox3.Items.Clear()
             ListBox4.Items.Clear()
             ListBox6.Items.Clear()
             Dim items() As String = ListBox5.Items.Cast(Of Object).Select(Function(o) ListBox5.GetItemText(o)).ToArray
             Dim list1 As New ArrayList
             Dim starting As DateTime = Format(DateTimePicker1.Value.Date.ToString("yyyy/MM/dd"))
             Dim ending As DateTime = Format(DateTimePicker2.Value.Date.ToString("yyyy/MM/dd"))
             Dim dates As String() = Enumerable.Range(0, 1 + ending.Subtract(starting).Days).[Select](Function(i) starting.AddDays(i).ToString("yyyy-MM-dd")).ToArray()
             list1.AddRange(dates)
             For Each num1 In list1
                 Dim list As New ArrayList
                 list.AddRange(items.ToArray)
                 For Each num In list
                     Dim sr As StreamReader = New StreamReader(Label8.Text)
                     Dim strLine As String
                     Do While sr.Peek() >= 0
                         strLine = String.Empty
                         strLine = sr.ReadLine
                         If strLine.Contains(num) And strLine.Contains(num1) Then
                             ListBox1.Items.Add(strLine)
                             OrigLst.Add(strLine)
                         End If
                     Loop
                     sr.Close()
                     For i As Integer = 0 To OrigLst.Count - 1
                         OrigLst(i) = Trim((OrigLst(i).Replace(num & " ", Nothing)).ToLower)
                     Next
                     OrigLst = StripDups("c/in", OrigLst)
                     OrigLst = StripDups("c/out", OrigLst)
                     Dim dts As New List(Of DateTime)
                     Dim dt As DateTime = Nothing
                     For i As Integer = 0 To OrigLst.Count - 1
                         OrigLst(i) = OrigLst(i).Replace(" c/out", Nothing).Replace(" c/in", Nothing)
                         If DateTime.TryParse(OrigLst(i), dt) Then
                             dts.Add(dt)
                         End If
                     Next
                     For Each obj As Object In OrigLst
                         ListBox2.Items.Add(num & " " & num1 & " " & obj)
                     Next
                     Dim d2 As TimeSpan
                     For i As Integer = 0 To OrigLst.Count - 2 Step 2
                         Dim d1 As TimeSpan = dts(i + 1) - dts(i)
                         ListBox3.Items.Add(num + " " + num1 + " " + dts(i + 1).ToString("HH:mm:ss") & " - " & dts(i).ToString("HH:mm:ss") & " = " & d1.ToString)
                         d2 += d1
                     Next
                     If d2.ToString() = "00:00:00" Then
                         ListBox4.Items.Add(num & " " + num1 & " " + "Absent")
                     Else
                         ListBox4.Items.Add(num & " " + num1 & " " + d2.ToString())
                     End If
                     d2 = Nothing
                     OrigLst.Clear()
                 Next
             Next
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
         OriginCheckBox4.Enabled = True
     End If
 End Sub
)
        thread.Start()
    End Sub
    Private Sub OriginButton6_Click(sender As Object, e As EventArgs) Handles OriginButton6.Click
        If Label9.Text = "No excel file loaded..." Then
            MsgBox("Load excel file first!", MsgBoxStyle.Information)
        Else
            If OriginCheckBox2.Checked = True Then
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
                    For Each oItem In ListBox4.Items
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
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " (" & average & ")"
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
                        xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " (" & average & ")"
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
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " (" & average & ")"
                            xlBook.Sheets(ComboBox1.SelectedItem).Cells(2, 2).Value = ComboBox1.SelectedItem & "(" & nonholiday & " weekdays" & ")"
                        End If
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
                            xlBook.Sheets(ComboBox1.SelectedItem).Range("C1").Offset(OffS1, OffS2).Value = totalTime1.ToString & " (" & average & ")"
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
        End If
    End Sub
    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        Me.Close()
    End Sub
    Private Sub OriginButton4_Click_1(sender As Object, e As EventArgs) Handles OriginButton4.Click
        Dim open As New OpenFileDialog With {
        .Filter = "Text File(*.txt*)|*.txt*"
    }
        open.ShowDialog()
        Label8.Text = (open.FileName)
        Dim strFilename As String
        strFilename = Trim(Label8.Text)
        OriginButton5.Enabled = True
        If strFilename.Length = 0 Then
        Else
            Dim shkurt As String = strFilename.Substring(strFilename.LastIndexOf("\")).Replace("\", "")
            Label13.Text = "File loaded!"
            Label19.Text = shkurt
            Label19.ForeColor = Color.ForestGreen
        End If
        Label13.ForeColor = Color.ForestGreen
        If Label8.Text.Length = 0 Then
        Else

        End If
    End Sub
    Private Sub OriginButton1_Click_1(sender As Object, e As EventArgs) Handles OriginButton1.Click
        MsgBox("What format should have the list?" & vbCrLf & vbCrLf & "*Name Surname Checktime(Date Time) Checktype(IN,OUT)" & vbCrLf & "*No header text" & vbCrLf & "*Run button 'Format List' only one time for each list!" & vbCrLf & "*Sort list from first check to last check!" & vbCrLf & "*Excel needs to be installed" & vbCrLf & "*Date format should be(yyyy-MM-dd)" & vbCrLf & vbCrLf & "List not formatted example: Gabriel Lami 2020-07-15 09:41:10 IN", MsgBoxStyle.Information)
    End Sub
    Dim reve As String
    Private Sub OriginButton2_Click_1(sender As Object, e As EventArgs) Handles OriginButton2.Click
        Dim open As New OpenFileDialog With {
            .Filter = "Text File(*.txt*)|*.txt*"
        }
        If open.ShowDialog() = DialogResult.OK Then
            Label22.Text = (open.FileName)
            Dim strFilename As String
            strFilename = Trim(Label22.Text)
            Dim readingFile As System.IO.StreamReader = New System.IO.StreamReader(Label22.Text)
            Dim readingLine As String = readingFile.ReadLine()
            readingFile.Close()
            ListBox1.Items.Clear()
            Dim sr As StreamReader = New StreamReader(Label22.Text)
            Dim strLine As String
            Do While sr.Peek() >= 0
                strLine = sr.ReadLine
                Dim regex As Regex = New Regex("[ ]{2,}", RegexOptions.None)
                strLine = regex.Replace(strLine, " ")
                strLine = System.Text.RegularExpressions.Regex.Replace(strLine, "\s+", " ")
                If strLine.Split(" ").Count < 5 Then
                Else
                    y = strLine.Split(" ")(0) & " " & strLine.Split(" ")(1) & " " & strLine.Split(" ")(2).Replace("/", "-") & " " & strLine.Split(" ")(3) & " " & strLine.Split(" ")(4)
                End If
                If OriginCheckBox1.Checked = True Then
                    sm = y.Replace("IN", "C/In").Replace("OUT", "C/Out").Replace(TextBox1.Text, Label12.Text).Replace(TextBox3.Text, Label15.Text).Replace("Ã§", "ç").Replace("Ã‡", "Ç")
                Else
                    sm = y.Replace("IN", "C/In").Replace("OUT", "C/Out").Replace("Ã§", "ç").Replace("Ã‡", "Ç")
                End If
                If TextBox4.Text = "dd/MM/yyyy" Or TextBox4.Text = "dd-MM-yyyy" Then
                    Dim date2 As Date = Convert.ToDateTime(sm.Split(" ")(2))
                    fdate = sm.Split(" ")(0) & " " & sm.Split(" ")(1) & " " & date2.ToString("yyyy-dd-MM", CultureInfo.InvariantCulture) & " " & sm.Split(" ")(3) & " " & sm.Split(" ")(4)
                ElseIf TextBox4.Text = "MM/dd/yyyy" Or TextBox4.Text = "MM-dd-yyyy" Then
                    Dim date2 As Date = Convert.ToDateTime(sm.Split(" ")(2))
                    fdate = sm.Split(" ")(0) & " " & sm.Split(" ")(1) & " " & date2.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) & " " & sm.Split(" ")(3) & " " & sm.Split(" ")(4)
                Else
                    fdate = sm.Split(" ")(0) & " " & sm.Split(" ")(1) & " " & sm.Split(" ")(2) & " " & sm.Split(" ")(3) & " " & sm.Split(" ")(4)
                End If
                If OriginCheckBox3.Checked = True Then
                    reve = fdate.Replace(TextBox5.Text, "C/In1").Replace(TextBox6.Text, TextBox5.Text).Replace("C/In1", TextBox6.Text).Replace("�", "ç").Replace("Ã§", "ç").Replace("Ã‡", "Ç")
                Else
                    reve = fdate
                End If
                ListBox1.Items.Add(reve)
            Loop
            sr.Close()
            Dim SW As IO.StreamWriter = IO.File.CreateText(Label22.Text)
            For Each S As String In ListBox1.Items
                SW.WriteLine(S)
            Next
            SW.Close()
            ListBox1.Items.Clear()
            MsgBox("Data formatted successfully!", MsgBoxStyle.Information)
            OriginCheckBox1.Checked = False
            OriginCheckBox3.Checked = False
            TextBox1.Text = ""
            TextBox3.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
            TextBox1.Enabled = False
            TextBox3.Enabled = False
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox4.Text = "yyyy-MM-dd"
        End If
    End Sub
    Private Sub OriginTheme1_Click(sender As Object, e As EventArgs) Handles OriginTheme1.Click
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
    Private Sub DateTimePicker2_ValueChanged_1(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
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
    Private Sub DateTimePicker1_ValueChanged_1(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
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
    Private Sub OriginCheckBox3_CheckedChanged(sender As Object) Handles OriginCheckBox3.CheckedChanged
        If OriginCheckBox3.Checked = True Then
            TextBox5.Enabled = True
            TextBox6.Enabled = True
        Else
            TextBox5.Enabled = False
            TextBox6.Enabled = False
            TextBox5.Text = ""
            TextBox6.Text = ""
        End If
    End Sub
    Private Sub Label5_Click_1(sender As Object, e As EventArgs) Handles Label5.Click
        MsgBox("You can drag and drop list of users on a .txt file or right click for more options!", MsgBoxStyle.Information)
    End Sub
    Private Sub OriginButton7_Click(sender As Object, e As EventArgs) Handles OriginButton7.Click
        If OriginRadioButton3.Checked = True Then
            SaveFileDialog1.Filter = "TXT Files (*.txt*)|*.txt"
            If SaveFileDialog1.ShowDialog = Forms.DialogResult.OK _
           Then
                Dim SW As IO.StreamWriter = IO.File.CreateText(SaveFileDialog1.FileName)
                For Each S As String In ListBox2.Items
                    SW.WriteLine(S)
                Next
                SW.Close()
                MsgBox("List Exported!", MsgBoxStyle.Information)
            End If
        ElseIf OriginRadioButton4.Checked = True Then
            SaveFileDialog1.Filter = "TXT Files (*.txt*)|*.txt"
            If SaveFileDialog1.ShowDialog = Forms.DialogResult.OK _
           Then
                Dim SW As IO.StreamWriter = IO.File.CreateText(SaveFileDialog1.FileName)
                For Each S As String In ListBox3.Items
                    SW.WriteLine(S)
                Next
                SW.Close()
                MsgBox("List Exported!", MsgBoxStyle.Information)
            End If
        ElseIf OriginRadioButton5.Checked = True Then
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
        ElseIf OriginRadioButton6.Checked = True Then
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
        Else
            MsgBox("Select one list to export first!", MsgBoxStyle.Information)
        End If
    End Sub
    Private Sub OriginCheckBox4_CheckedChanged(sender As Object) Handles OriginCheckBox4.CheckedChanged
        If OriginCheckBox4.Checked = True Then
            Second_List.Show()
            Me.WindowState = FormWindowState.Minimized
        End If
    End Sub
    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub OriginCheckBox1_CheckedChanged_1(sender As Object) Handles OriginCheckBox1.CheckedChanged
        If OriginCheckBox1.Checked = True Then
            TextBox1.Enabled = True
            TextBox3.Enabled = True
        Else
            TextBox1.Enabled = False
            TextBox3.Enabled = False
            TextBox1.Text = ""
            TextBox3.Text = ""
        End If
    End Sub
End Class