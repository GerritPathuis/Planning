Imports System.Data.SqlClient
Imports System.Data
Imports System.Text
Imports System
Imports System.IO
Imports System.Configuration
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Security.Permissions
Imports System.Security
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Threading

Public Class Form1
    Dim connectionString As String
    'Change connection string in App.config to connect new database.................
    'Opgelet is SQL server "SQLlocalDB.msi" al geinstalleerd ??


    ' Dim connectionString1 As String = "Data Source=(LocalDB)\v11.0;AttachDbFilename=N:\Engineering\DB_vtk_planning\VTK-Planning.mdf;Integrated Security=True;Connect Timeout=60"
    Dim connectionString2 As String = "Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\KenPlan_database_directory\E-Planning.mdf;Integrated Security=True;Connect Timeout=60"

    Private streamToPrint As StreamReader
    Private documentContents As String  'Contains complete document for printing
    Private stringToPrint As String 'Contains part document not yet printing
    Dim FILE_NAME As String = "MyjobsFile2.txt"
    Dim DATABASE_NAME As String = "E-Planning.mdf"
    Dim DATABASE_LOG_NAME As String = "E-Planning_log.ldf"

    Dim PATH As String
    'Dim PATH1 As String = "N:\Engineering\DB_vtk_planning\"
    Dim PATH2 As String = "c:\KenPlan_database_directory\"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '----------------------Set the default date format-------------------- 
        Dim newCulture As CultureInfo = DirectCast(System.Threading.Thread.CurrentThread.CurrentCulture.Clone(), CultureInfo)
        newCulture.DateTimeFormat.ShortDatePattern = "dd-MM-yyyy"
        newCulture.DateTimeFormat.DateSeparator = "-"
        Thread.CurrentThread.CurrentCulture = newCulture

        find_database() 'Make a Database file copy 
        copy_database() 'Make a Database file copy 

        Try
            Me.StaffTableAdapter.Fill(Me._VTK_PlanningDataSet1.Staff)
        Catch ex As Exception
            MessageBox.Show("Staff table login failed" & ex.Message)
        End Try

        Try
            Me.JobsTableAdapter.Fill(Me._VTK_PlanningDataSet1.Jobs)
        Catch ex As Exception
            MessageBox.Show("Jobs table login failed" & ex.Message)
        End Try

        ' Resize the master DataGridView columns to fit the newly loaded data.
        JobsDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        StaffDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Dim column4 As DataGridViewColumn = JobsDataGridView.Columns(4)
        column4.Width = 200

        Label1.Text = Now
        Label2.Text = "Weeknummer " + WeekNummer(Now).ToString

        Try
            Update_0()  'Check database

            Fill_ComboBox1()
            Fill_ComboBox2()
            Fill_ComboBox3()
        Catch ex As Exception
            MessageBox.Show("First update failled " & ex.Message)
        End Try
    End Sub

    Public Function WeekNummer(ByVal datum As Date) As Double
        Dim week As Integer
        Dim wtest As Integer
        Dim dag_nr As Integer
        Dim jaar_nr As Integer
        Dim wkn_str As String
        Dim wkn As Double

        week = DatePart(DateInterval.WeekOfYear, datum, Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        wtest = DatePart(DateInterval.WeekOfYear, datum.AddDays(7), Microsoft.VisualBasic.FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)
        If week = 53 And wtest = 2 Then 'weeknummer corrigeren
            week = 1
        End If
        dag_nr = DatePart("w", datum) - 1           'Monday = 1
        jaar_nr = DatePart("yyyy", datum) Mod 10
        wkn_str = jaar_nr.ToString("D1") + week.ToString("D2") + "." + dag_nr.ToString
        Double.TryParse(wkn_str, wkn)
        Return wkn
    End Function

    'Populate the Job numbers in combobox
    Private Sub Fill_ComboBox1()
        Dim staff_table As DataTable = Me._VTK_PlanningDataSet1.Staff

        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("ALL")
        ComboBox1.Items.Add("ALL FINISHED")
        ComboBox1.Items.Add("ALL PENDING -x wks")
        ComboBox1.Items.Add("BOOK DATE")
        For hh = 0 To staff_table.Rows.Count - 1
            ComboBox1.Items.Add(staff_table.Rows(hh).Item("TC"))
        Next hh
        ComboBox1.Text = "ALL PENDING -x wks"
    End Sub
    'Populate the Job_nr in combo box
    Private Sub Fill_ComboBox2()
        Dim jobs_table As DataTable = Me._VTK_PlanningDataSet1.Jobs
        Dim lstOfStrings As New List(Of String)
        Dim str As String

        'Add only unique job_nr to the string list
        For hj = 0 To jobs_table.Rows.Count - 1
            str = jobs_table.Rows(hj).Item("Job_nr").ToString
            If (Not lstOfStrings.Contains(str)) Then
                lstOfStrings.Add(jobs_table.Rows(hj).Item("Job_nr"))
            End If
        Next hj
        lstOfStrings.Sort()  'Now sort the list

        'Now move the string list to the Combobox
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("Job_nr")
        For hk = 0 To lstOfStrings.Count - 1
            ComboBox2.Items.Add(lstOfStrings.Item(hk))
        Next hk
        ComboBox2.Text = "Job_nr"
    End Sub
    'Department numbers
    Private Sub Fill_ComboBox3()
        Dim jobs_table As DataTable = Me._VTK_PlanningDataSet1.Jobs
        Dim lstOfStrings As New List(Of String)
        Dim str As String

        'Add only unique Department to the string list
        For hj = 0 To jobs_table.Rows.Count - 1
            str = jobs_table.Rows(hj).Item("Department").ToString
            If (Not lstOfStrings.Contains(str)) Then
                lstOfStrings.Add(jobs_table.Rows(hj).Item("Department"))
            End If
        Next hj
        lstOfStrings.Sort()  'Now sort the list

        ComboBox3.Items.Clear()
        ComboBox3.Items.Add("ALL")
        For hk = 0 To lstOfStrings.Count - 1
            ComboBox3.Items.Add(lstOfStrings.Item(hk))
        Next hk

        ComboBox3.Text = "ALL"
    End Sub
    'Enter TC names in the combobox1
    Private Sub ComboBox1_Enter(sender As Object, e As EventArgs) Handles ComboBox1.Enter
        Fill_ComboBox1()    'Staff
    End Sub
    'Enter Jobs in the combobox2
    Private Sub ComboBox2_Enter(sender As Object, e As EventArgs) Handles ComboBox1.Enter, ComboBox2.Enter
        Fill_ComboBox2()    'Jobs
    End Sub
    Private Sub ComboBox3_Enter(sender As Object, e As EventArgs) Handles ComboBox3.Enter
        Fill_ComboBox3()
    End Sub

    'Update button
    Private Sub Update_ToolStripButton_Click(sender As Object, e As EventArgs) Handles Update_ToolStripButton.Click, TabPage2.Enter
        Label5.Text = "Update_0"
        MyBase.Update()
        Update_0()
        Label5.Text = "Update_1"
        MyBase.Update()
        Update_1()
        Label5.Text = "Update_2"
        MyBase.Update()
        Update_2()
        Label5.Text = "Update_3"
        MyBase.Update()
        Update_3()
        Label5.Text = "Update done"
        JobsDataGridView.Refresh()
    End Sub
    'Save button 
    Private Sub Save_ToolStripButton2_Click(sender As Object, e As EventArgs) Handles Save_ToolStripButton2.Click

        Dim cnn As SqlConnection
        cnn = New SqlConnection(connectionString)
        cnn.Open()
        Try
            TableAdapterManager1.UpdateAll(_VTK_PlanningDataSet1)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        cnn.Close() 'connection close here , that is disconnected from data source
        JobsDataGridView.Refresh()
    End Sub
    'Save button
    Private Sub Save_ToolStripButton_Click(sender As Object, e As EventArgs) Handles Save_ToolStripButton.Click
        Save_ToolStripButton2.PerformClick()
    End Sub
    'Remove row?
    Private Sub ToolStripButton1_Click_1(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim cnn As SqlConnection
        Dim rowindex As Integer
        Dim result As Integer

        rowindex = Me.JobsDataGridView.CurrentCell.RowIndex()
        JobsDataGridView.Rows(rowindex).DefaultCellStyle.BackColor = Color.Red

        result = MessageBox.Show("     Remove row ?", "DELETE LINE", MessageBoxButtons.YesNo)
        JobsDataGridView.Rows(rowindex).DefaultCellStyle.BackColor = Color.White

        If result = DialogResult.Yes Then
            Try
                cnn = New SqlConnection(connectionString)
                cnn.Open()
                Me.JobsDataGridView.Rows.RemoveAt(rowindex)

                TableAdapterManager1.UpdateAll(_VTK_PlanningDataSet1)
                cnn.Close() 'connection close here , that is disconnected from data source
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Private Sub JobsDataGridView_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles JobsDataGridView.CellValidating

        Me.JobsDataGridView.Rows(e.RowIndex).ErrorText = ""

        ' Don't try to validate the 'new row' until finished  
        ' editing since there 
        ' is not any point in validating its initial value. 
        'If JobsDataGridView.Rows(e.RowIndex).IsNewRow Then

        ' MsgBox("Bingo line 220")
        Try
            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(0).Value) Then 'Booking
                JobsDataGridView.Rows(e.RowIndex).Cells(0).Value = Now()
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(1).Value) Then 'Department
                JobsDataGridView.Rows(e.RowIndex).Cells(1).Value = 11
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(2).Value) Then 'Jobnr
                JobsDataGridView.Rows(e.RowIndex).Cells(2).Value = 1000
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(3).Value) Then  'TC
                JobsDataGridView.Rows(e.RowIndex).Cells(3).Value = "GPA"
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(4).Value) Then  'Description
                '  JobsDataGridView.Rows(e.RowIndex).Cells(4).Value = "Descrip"
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(5).Value) Then 'Length
                JobsDataGridView.Rows(e.RowIndex).Cells(5).Value = 1
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(6).Value) Then     'Spent_hrs
                JobsDataGridView.Rows(e.RowIndex).Cells(6).Value = 0
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(7).Value) Then    'Finished (%)
                ' JobsDataGridView.Rows(e.RowIndex).Cells(7).Value = 0
            End If

            'Remove trailing spaces
            'str = JobsDataGridView.Rows(e.RowIndex).Cells(2).Value.ToString.Trim
            ' JobsDataGridView.Rows(e.RowIndex).Cells(2).Value = str

            'JobsDataGridView.Rows(e.RowIndex).Cells(5).Value = Now          'Start_Date
            'JobsDataGridView.Rows(e.RowIndex).Cells(6).Value = Now          'BB_date
            'JobsDataGridView.Rows(e.RowIndex).Cells(7).Value = "1-1-2000"   'Finished_Date

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(11).Value) Then     'Remarks
                JobsDataGridView.Rows(e.RowIndex).Cells(11).Value = "Client"
            End If

            If IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(12).Value) Then     'Priority
                JobsDataGridView.Rows(e.RowIndex).Cells(12).Value = 0
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Function RandomString(r As Random)
        Dim s As String = "ABCDEFGH"
        Dim sb As New StringBuilder
        Dim cnt As Integer = r.Next(8, 10)
        For i As Integer = 1 To cnt
            Dim idx As Integer = r.Next(0, s.Length)
            sb.Append(s.Substring(idx, 1))
        Next
        Return sb.ToString()
    End Function

    'Check for DBnull in database and cleanup
    Private Sub Update_0()
        Dim jobs_table As DataTable = _VTK_PlanningDataSet1.Jobs
        Dim staff_table As DataTable = _VTK_PlanningDataSet1.Staff
        Dim r As New Random
        Dim str As String

        'Get All jobs    
        Save_ToolStripButton2.PerformClick()    'Make sure that there is NO pending stuff


        For hc = 0 To jobs_table.Rows.Count - 1
            If IsDBNull(jobs_table.Rows(hc).Item("TC")) Then jobs_table.Rows(hc).Item("TC") = "GPA"
            'Covert to Upper case
            str = jobs_table.Rows(hc).Item("TC").ToString
            jobs_table.Rows(hc).Item("TC") = str.Trim().ToUpper

            'Check for DBNull prevent problems------------- 
            If IsDBNull(jobs_table.Rows(hc).Item("Job_nr")) Then jobs_table.Rows(hc).Item("Job_nr") = 1000
            If IsDBNull(jobs_table.Rows(hc).Item("Spent_hrs")) Then jobs_table.Rows(hc).Item("Spent_hrs") = 0
            If IsDBNull(jobs_table.Rows(hc).Item("Priority")) Then jobs_table.Rows(hc).Item("Priority") = 0
            If IsDBNull(jobs_table.Rows(hc).Item("Department")) Then jobs_table.Rows(hc).Item("Department") = 11
            If IsDBNull(jobs_table.Rows(hc).Item("Finished")) Then jobs_table.Rows(hc).Item("Finished") = 0
            If IsDBNull(jobs_table.Rows(hc).Item("Finish_date")) Then jobs_table.Rows(hc).Item("Finish_date") = Now
            If IsDBNull(jobs_table.Rows(hc).Item("Start_date")) Then jobs_table.Rows(hc).Item("Start_date") = Now
            If IsDBNull(jobs_table.Rows(hc).Item("BB_date")) Then jobs_table.Rows(hc).Item("BB_date") = Now
            'If IsDBNull(jobs_table.Rows(hc).Item("Finished_date")) Then jobs_table.Rows(hc).Item("Finished_date") = "1-1-2000"
            If IsDBNull(jobs_table.Rows(hc).Item("Ship_date")) Then jobs_table.Rows(hc).Item("Ship_date") = "1-1-2000"
            If IsDBNull(jobs_table.Rows(hc).Item("Prod_time")) Then jobs_table.Rows(hc).Item("Prod_time") = "1-1-2000"
            If IsDBNull(jobs_table.Rows(hc).Item("Remarks")) Then jobs_table.Rows(hc).Item("Remarks") = "-"
            If IsDBNull(jobs_table.Rows(hc).Item("Descrip_job")) Then jobs_table.Rows(hc).Item("Descrip_job") = "-"
            'MessageBox.Show(jobs_table.Rows(hc).Item("Booking"))

        Next hc
        Save_ToolStripButton2.PerformClick()

        '-----------------------------------------------------------------
        'Get All Staff

        For hd = 0 To staff_table.Rows.Count - 1
            'Check for DBNull prevent problems-------------
            If IsDBNull(staff_table.Rows(hd).Item("TC")) Then staff_table.Rows(hd).Item("TC") = RandomString(r)

            'Covert TC to Upper case
            str = staff_table.Rows(hd).Item("TC").ToString
            staff_table.Rows(hd).Item("TC") = str.Trim().ToUpper
            If IsDBNull(staff_table.Rows(hd).Item("Id")) Then staff_table.Rows(hd).Item("Id") = 99
            If IsDBNull(staff_table.Rows(hd).Item("Autoriteit")) Then staff_table.Rows(hd).Item("Autoriteit") = 5
            If IsDBNull(staff_table.Rows(hd).Item("Dprtmnt")) Then staff_table.Rows(hd).Item("Dprtmnt") = 11
            If IsDBNull(staff_table.Rows(hd).Item("Wk_hrs")) Then staff_table.Rows(hd).Item("Wk_hrs") = 40
            If IsDBNull(staff_table.Rows(hd).Item("Bussy")) Then staff_table.Rows(hd).Item("Bussy") = Now
            If IsDBNull(staff_table.Rows(hd).Item("Bussy_wk")) Then staff_table.Rows(hd).Item("Bussy_wk") = "0"

            'If IsDBNull(staff_table.Rows(hd).Item("TC_load")) Then staff_table.Rows(hd).Item("TC_load") = 0
            If IsDBNull(staff_table.Rows(hd).Item("Vak_wk1")) Then staff_table.Rows(hd).Item("Vak_wk1") = 0
            If IsDBNull(staff_table.Rows(hd).Item("Vak_wk2")) Then staff_table.Rows(hd).Item("Vak_wk2") = 0
            If IsDBNull(staff_table.Rows(hd).Item("Vak_wk3")) Then staff_table.Rows(hd).Item("Vak_wk3") = 0
            If IsDBNull(staff_table.Rows(hd).Item("Vak_wk4")) Then staff_table.Rows(hd).Item("Vak_wk4") = 0

            If IsDBNull(staff_table.Rows(hd).Item("Descrip")) Then staff_table.Rows(hd).Item("Descrip") = "Descrip"

        Next hd
        Save_ToolStripButton2.PerformClick()
    End Sub

    'Calculate the the job start and stop time of the jobs
    Private Sub Update_1()
        Dim jobs_table As DataTable = _VTK_PlanningDataSet1.Jobs
        Dim staff_table As DataTable = _VTK_PlanningDataSet1.Staff

        Dim found_rows_job() As DataRow 'Job selection
        Dim tc_name As String
        Dim filter As String
        Dim job_length As Integer
        Dim temp As DateTime
        Dim vak1, vak2, vak3, vak4 As Integer

        'Do for every TC of the staff table
        For hh = 0 To staff_table.Rows.Count - 1
            ' Get the first TC name
            tc_name = staff_table.Rows(hh).Item(0)

            '-------------------- Get All unfinished jobs from first TC -------------
            filter = "TC = '" + tc_name.ToString + "'" + " and Finished < 100 "
            found_rows_job = jobs_table.Select(filter, "Priority ASC")
            ' MessageBox.Show("Found Jobs are " + filter)

            For ii = 0 To found_rows_job.GetUpperBound(0)
                'MessageBox.Show("Job number " + i.ToString)
                If (ii = 0) Then
                    temp = New DateTime(Now.Year, Now.Month, Now.Day, 8, 0, 0, 0)  'start time first job
                    temp = DateAdd("d", 1, temp)                                    'start tomorrow 8:00
                    If (temp.DayOfWeek() = DayOfWeek.Saturday) Then temp = DateAdd("d", 2, temp)
                    If (temp.DayOfWeek() = DayOfWeek.Sunday) Then temp = DateAdd("d", 1, temp)
                    found_rows_job(ii).Item("Start_date") = temp

                    ' MessageBox.Show(temp.ToLongDateString)
                End If
                '--------------- Just making sure----------------
                If found_rows_job(ii).Item("Finished") < 0 Then found_rows_job(ii).Item("Finished") = 0
                If found_rows_job(ii).Item("Finished") > 100 Then found_rows_job(ii).Item("Finished") = 100
                If found_rows_job(ii).Item("Length") > 2000 Then found_rows_job(ii).Item("Length") = 2000
                If found_rows_job(ii).Item("Length") < 0 Then found_rows_job(ii).Item("Length") = -1

                '--------------- Done making sure----------------
                job_length = found_rows_job(ii).Item("Length") * (100 - found_rows_job(ii).Item("Finished")) / 100

                '--------------- add the load factor----------------
                If CheckBox1.Checked = True Then
                    job_length = job_length / 0.8
                End If

                found_rows_job(ii).Item("Start_date") = temp
                vak1 = staff_table.Rows(hh).Item("vak_wk1")
                vak2 = staff_table.Rows(hh).Item("vak_wk2")
                vak3 = staff_table.Rows(hh).Item("vak_wk3")
                vak4 = staff_table.Rows(hh).Item("vak_wk4")

                vak1 = Int(vak1 / 10) * 10 + 1      'Set to the first day of the week eg 5333 becomes 5331
                vak2 = Int(vak2 / 10) * 10 + 1      'Set to the first day of the week eg 5333 becomes 5331
                vak3 = Int(vak3 / 10) * 10 + 1      'Set to the first day of the week eg 5333 becomes 5331
                vak4 = Int(vak4 / 10) * 10 + 1      'Set to the first day of the week eg 5333 becomes 5331

                While (job_length > 0)

                    '-------Vacation1 then add 7 days ---------------
                    'MessageBox.Show("vak1=" + vak1.ToString + " C_DayOfYear(vak1)=" + C_DayOfYear(vak1).ToString)

                    If (vak1 > 0 And temp >= C_DayOfYear(vak1) And temp <= C_DayOfYear(vak1 + 5)) Then
                        temp = DateAdd("d", 7, temp)
                    End If
                    '-------Vacation1 then add 7 days ---------------
                    If (vak2 > 0 And temp >= C_DayOfYear(vak2) And temp <= C_DayOfYear(vak2 + 5)) Then
                        temp = DateAdd("d", 7, temp)
                    End If
                    '-------Vacation1 then add 7 days ---------------
                    If (vak3 > 0 And temp >= C_DayOfYear(vak3) And temp <= C_DayOfYear(vak3 + 5)) Then
                        temp = DateAdd("d", 7, temp)
                    End If
                    '-------Vacation1 then add 7 days ---------------
                    If (vak4 > 0 And temp >= C_DayOfYear(vak4) And temp <= C_DayOfYear(vak4 + 5)) Then
                        temp = DateAdd("d", 7, temp)
                    End If


                    '-------is temp Saterday ? then add 2 days ---------------
                    If (temp.DayOfWeek() = DayOfWeek.Saturday) Then temp = DateAdd("d", 2, temp)

                    '-------is temp Sunday ? then add 1 days ---------------
                    If (temp.DayOfWeek() = DayOfWeek.Sunday) Then temp = DateAdd("d", 1, temp)


                    '-------is it after work hours -----------------------
                    '-------- (calculate from 8.00 to 16.00 hrs) -------
                    If (temp.Hour >= 16) Then
                        temp = DateAdd("h", 16, temp)
                    End If
                    '-------during work hours, add job time-----------------
                    If (temp.Hour >= 8 And temp.Hour < 16) Then
                        temp = DateAdd("h", 1, temp) ' ADD ONE HOUR
                        job_length -= 1
                    End If
                End While
                found_rows_job(ii).Item("BB_date") = temp                'Store the Bedrijfsburo datum
                staff_table.Rows(hh).Item("Bussy") = temp                'test
                staff_table.Rows(hh).Item("Bussy_wk") = WeekNummer(temp) 'test
            Next ii
        Next hh
    End Sub

    Public Function SubttractWorkingDays(ByVal Date_start As DateTime, ByVal Date_end As DateTime) As Integer
        Dim DayCount As Integer = 0

        If Date_start > Date_end Then Return 0 'Impossible Date_end must be bigger

        ' Loop around until we get the need non-weekend day
        While Date_start < Date_end
            Date_start = Date_start.AddDays(1)
            If Weekday(Date_start) <> 7 And Weekday(Date_start) <> 1 Then DayCount += 1
        End While
        Return DayCount
    End Function
    'Calculate the department load in hours

    Private Sub Update_2()
        Dim total As Double
        Dim table As DataTable = _VTK_PlanningDataSet1.Jobs
        Dim done As Integer

        'Get All unfinished jobs
        Dim result() As DataRow = table.Select("Finished < 100")

        For i = 0 To result.GetUpperBound(0)
            'Calculate the department load in hours
            total += result(i).Item("Length") * (100 - result(i).Item("Finished")) / 100

            'Calculate the finished percentage
            If (result(i).Item("Finished") <> 100 And Not IsDBNull(result(i).Item("Finished")) And Not IsDBNull(result(i).Item("Spent_hrs")) And result(i).Item("Length") > 0) Then

                'Make sure finished >= spent
                If result(i).Item("Spent_hrs") > result(i).Item("Finished") Then
                    result(i).Item("Finished") = result(i).Item("Spent_hrs")
                End If

                done = Int(100 * result(i).Item("Spent_hrs") / result(i).Item("Length"))

                'This calculation is not allowed to closed the job at 100%
                If done < 100 Then
                    result(i).Item("Finished") = done
                Else
                    result(i).Item("Finished") = 99
                End If
            End If

            'Calculate workingdays available for the production
            result(i).Item("Prod_time") = SubttractWorkingDays(result(i).Item("BB_date"), result(i).Item("Ship_date"))
        Next i
    End Sub
    ' Calc total load of individual TC's

    Private Sub Update_3()
        Dim jobs_table As DataTable = _VTK_PlanningDataSet1.Jobs
        Dim staff_table As DataTable = _VTK_PlanningDataSet1.Staff
        Dim TC_total As Double
        Dim department_total As Double = 0
        Dim tc_name As String
        Dim filter As String
        Dim found_rows_job() As DataRow 'Job selection

        '--------------- Do for every TC of the staff table -------------------------
        For j = 0 To staff_table.Rows.Count - 1
            ' Get the first TC name
            tc_name = staff_table.Rows(j).Item(0)

            '-------------------- Get All unfinished jobs from first TC -------------
            TC_total = 0
            filter = "TC = '" + tc_name.ToString + "'"
            found_rows_job = jobs_table.Select(filter)
            'MessageBox.Show(filter)
            For i = 0 To found_rows_job.GetUpperBound(0)
                TC_total += found_rows_job(i).Item("Length") * (100 - found_rows_job(i).Item("Finished")) / 100
            Next i
            'Store the result
            staff_table.Rows(j).Item("TC_Load") = TC_total
            department_total += TC_total
        Next j
        department_total = Math.Round(department_total)
        Label3.Text = "Department total load is " + department_total.ToString + " hours"
    End Sub
    'Convert DayOfYer to Date
    '5011 Translate t0 2015, week 01, day 1
    Public Function C_DayOfYear(ByVal dwkd As Integer) As Date
        Dim c_year, c_week, c_day As Integer
        Dim dayofyear As Integer
        Dim StartOfYear As Date = "#1/1/2015#"
        Dim ReturnDate As Date

        c_year = Int(dwkd / 1000)
        c_week = Int((dwkd - (c_year * 1000)) / 10)
        c_day = dwkd - (c_year * 1000) - (c_week * 10)
        c_year += 2010

        dayofyear = c_day + c_week * 7 - 11
        ReturnDate = StartOfYear.AddDays(dayofyear)

        Return ReturnDate
    End Function

    Private Sub JobsDataGridView_DataError(sender As Object, e As DataGridViewDataErrorEventArgs)
        MessageBox.Show("Error happened " & e.Context.ToString())
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Dim ps As New PaperSize("A4", 850, 1100)
        MyBase.Update()
        Update_ToolStripButton.PerformClick()
        Create_jobs_file()
        ReadDocument()

        Label5.Text = "Printer setting"
        MyBase.Update()
        For ix As Integer = 0 To PrintDocument1.PrinterSettings.PaperSizes.Count - 1
            If PrintDocument1.PrinterSettings.PaperSizes(ix).Kind = PaperKind.A4 Then
                ps = PrintDocument1.PrinterSettings.PaperSizes(ix)
                PrintDocument1.DefaultPageSettings.PaperSize = ps
            End If
        Next

        Label5.Text = "Choose Printer"
        MyBase.Update()
        Dim result As DialogResult = PrintDialog1.ShowDialog()
        If (result = DialogResult.OK) Then
            Label5.Text = "Page Setup"
            MyBase.Update()
            'Page setup------------------
            PageSetupDialog1.Document = PrintDocument1
            PageSetupDialog1.PageSettings.Landscape = True
            PageSetupDialog1.PageSettings.Margins.Left = 50
            PageSetupDialog1.PageSettings.Margins.Right = 30
            PageSetupDialog1.PageSettings.Margins.Top = 60
            PageSetupDialog1.PageSettings.Margins.Bottom = 70
            Dim result2 As DialogResult = PageSetupDialog1.ShowDialog()
            If (result2 = DialogResult.OK) Then
                Label5.Text = "Preview Setup"
                MyBase.Update()
                'Preview setup------------------
                PrintPreviewDialog1.PrintPreviewControl.Zoom = 1.0
                PrintPreviewDialog1.Document = PrintDocument1
                PrintPreviewDialog1.ShowDialog()
                Label5.Text = "Print done"
            End If
        End If
    End Sub

    'Read the document from disk and store it into documentcontents
    Private Sub ReadDocument()
        Dim docName As String = PATH + FILE_NAME

        PrintDocument1.DocumentName = docName
        Dim stream As New FileStream(docName, FileMode.Open)
        Try
            Dim reader As New StreamReader(stream)
            Try
                documentContents = reader.ReadToEnd()
            Finally
                reader.Dispose()
            End Try
        Finally
            stream.Dispose()
        End Try
        stringToPrint = documentContents
    End Sub

    Sub printDocument1_PrintPage(ByVal sender As Object,
    ByVal e As PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim charactersOnPage As Integer = 0
        Dim linesPerPage As Integer = 0
        Dim drawFont As New Font("Courier New", 9) 'Courier New has fixed width

        ' Sets the value of charactersOnPage to the number of characters  
        ' of stringToPrint that will fit within the bounds of the page. 
        e.Graphics.MeasureString(stringToPrint, drawFont, e.MarginBounds.Size, StringFormat.GenericTypographic, charactersOnPage, linesPerPage)

        ' Draws the string within the bounds of the page.
        e.Graphics.DrawString(stringToPrint, drawFont, Brushes.Black, e.MarginBounds, StringFormat.GenericTypographic)

        ' Remove the portion of the string that has been printed.
        stringToPrint = stringToPrint.Substring(charactersOnPage)

        ' Check to see if more pages are to be printed.
        e.HasMorePages = stringToPrint.Length > 0

        ' If there are no more pages, reset the string to be printed. 
        If Not e.HasMorePages Then
            stringToPrint = documentContents
        End If
    End Sub

    '------ Create file containing the jobs info-------------
    Private Sub Create_jobs_file()
        Dim str, desc_short, rem_short, job_nr_short, pt_short, TC_short, prio_short, BB_wk, PP_wk, ship_wk As String
        Dim TC_name As String
        Dim format As String = "{0,-5} {1,-8} {2,-60} {3,4} {4,3} {5,3} {6,3} {7,-6} {8,-6} {9,-6} {10,-6} {11,4} {12,-10}"
        Dim now_min_2wks As DateTime

        Try
            If System.IO.File.Exists(PATH + FILE_NAME) = False Then
                Using sw As StreamWriter = File.CreateText(PATH + FILE_NAME) ' Create a file to write to. 
                    sw.WriteLine("File for Eplanning")
                End Using
                MsgBox("File Created for Eplanning")
            End If
            Dim objWriter As New System.IO.StreamWriter(PATH + FILE_NAME)

            '---------------Get All unfinished jobs and finished jobs less 2 weeks----------------
            now_min_2wks = New DateTime(Now.Year, Now.Month, Now.Day, 8, 0, 0, 0)  'start time 
            now_min_2wks = DateAdd("d", (NumericUpDown1.Value * -7 + 1), now_min_2wks)                          'Subtract 2 weeks


            '---------------What to print---------------------------------------------------------
            Dim table As DataTable = _VTK_PlanningDataSet1.Jobs
            Dim result() As DataRow

            Dim message, title, defaultValue As String
            Dim myValue As Object

            '---------------------What to print-----------------------
            ' Set prompt.
            message = "1) Op TC gesorteerd, Jobs finished + last 2 wks" & Chr(13) & Chr(10) & "2) Op TC gesorteed, Jobs finished" &
                Chr(13) & Chr(10) & "3) Op Jobs gesorteed"
            ' Set title.
            title = "Print Selection"
            defaultValue = "1"   ' Set default value.
            myValue = InputBox(message, title, defaultValue)
            ' If user has clicked Cancel, set myValue to defaultValue 
            If myValue Is "" Then myValue = defaultValue


            '-----------------Select the proper dataset and sort------------------------
            str = "[Finished] < 100"
            result = table.Select(str, "[Job_nr] ASC, [Start_date] ASC, [Finish_date] ASC")

            If myValue = 1 Then
                str = String.Format("[Finished] < 100 OR [Finished_date] > #{0:MM/dd/yyyy hh:mm:ss}#", now_min_2wks) & " OR [Finished] < 100"
                result = table.Select(str, "[TC] ASC, [Start_date] ASC, [Finish_date] ASC")
                MessageBox.Show(str)
            ElseIf myValue = 2 Then
                str = "[Finished] < 100"
                result = table.Select(str, "[TC] ASC, [Start_date] ASC, [Finish_date] ASC")
            End If

            '---------------Print Header to file-----------
            str = String.Format("Engineering planning wk " + WeekNummer(Now).ToString + "; " + Now.ToShortTimeString.ToString) + "; " + Label3.Text
            objWriter.Write(str + Environment.NewLine + Environment.NewLine)

            str = String.Format(format, "Book", "Job_nr", "Description", "Hrs", "Hrs", "TC", "Pri", "Start", "BBuro", "PPlan", "Ship", "Prod", "Remarks")
            objWriter.Write(str + Environment.NewLine)

            str = String.Format(format, "date", "      ", "          ", "tot", "act", "   ", "   ", "ywkd ", "ywkd ", "ywkd ", "ywkd", "days", " ")
            objWriter.Write(str + Environment.NewLine + Environment.NewLine)

            TC_name = result(0).Item("TC")                                      'For adding separation line betwee different TC

            '---------------Print to file-----------
            For i = 0 To result.GetUpperBound(0)
                '--- limit length of job_nr
                job_nr_short = result(i).Item("job_nr") & "      "   'Append blanks 
                job_nr_short = job_nr_short.Substring(0, 8)          'Now reduce the length 

                '--- limit length of the description
                desc_short = result(i).Item("Descrip_job").trim() & "-----------------------------------------------------------"   'Append blanks 
                desc_short = desc_short.Substring(0, 60)             'Now reduce the length 

                TC_short = result(i).Item("TC")
                TC_short = TC_short.Substring(0, 3)

                prio_short = result(i).Item("Priority").ToString
                prio_short = prio_short.Trim

                '--- limit length of the Product time
                pt_short = result(i).Item("Prod_time") & "        "   'Append blanks 
                pt_short = pt_short.Substring(0, 3)                   'Now reduce the length 

                BB_wk = WeekNummer(result(i).Item("BB_date"))

                If Not IsDBNull(result(i).Item("Production_Planning")) Then
                    PP_wk = WeekNummer(result(i).Item("Production_Planning"))
                Else
                    PP_wk = "-----"
                End If

                ship_wk = WeekNummer(result(i).Item("Ship_Date"))

                '--- limit length of the Remarks
                rem_short = result(i).Item("Remarks").trim() & "-----------------"   'Append blanks 
                rem_short = rem_short.Substring(0, 17)                'Now reduce the length 

                '---------------------- Make the string ---------------------------------------
                str = String.Format(format, WeekNummer(result(i).Item("Booking")), job_nr_short,
               desc_short, result(i).Item("Length"), result(i).Item("Spent_hrs"),
               TC_short, prio_short, WeekNummer(result(i).Item("Start_Date")),
                BB_wk, PP_wk, ship_wk, pt_short, rem_short)

                '---------------For adding separation line betwee different TC----------------
                If String.Compare(TC_name, result(i).Item("TC")) And myValue <> 3 Then
                    objWriter.Write(Environment.NewLine)
                    TC_name = result(i).Item("TC")
                End If

                objWriter.Write(str + Environment.NewLine)
            Next i
            objWriter.Close()
            Label5.Text = "Txt to " + PATH + FILE_NAME
            MyBase.Update()
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Create file containing the jobs info ")
        End Try
    End Sub

    'Make a copy of the database file
    Private Sub copy_database()
        Dim sss, ssr As String
        sss = PATH + "E-Planning(" + WeekNummer(Now).ToString + ").mdf#"
        ssr = PATH + "E-Planning_log(" + WeekNummer(Now).ToString + ").ldf#"
        Try
            If System.IO.File.Exists(sss) = False Then My.Computer.FileSystem.CopyFile(PATH + DATABASE_NAME, sss)
            If System.IO.File.Exists(ssr) = False Then My.Computer.FileSystem.CopyFile(PATH + DATABASE_LOG_NAME, ssr)
            Label5.Text = DATABASE_NAME + " copied"
            MyBase.Update()
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Problem making copy of the Database files ")
        End Try
    End Sub
    'Find the database file (different for home and VTK)
    Private Sub find_database()
        Try
            'If System.IO.File.Exists(PATH1 + DATABASE_NAME) Then
            '    connectionString = connectionString1
            '    PATH = PATH1
            'End If

            If System.IO.File.Exists(PATH2 + DATABASE_NAME) Then
                connectionString = connectionString2
                PATH = PATH2
            End If

            Label4.Text = connectionString
            Label5.Text = DATABASE_NAME & " found"
            MyBase.Update()
        Catch ex As Exception
            MessageBox.Show("Oops did not find the database " + ex.Message)
        End Try

    End Sub
    'Select the jobs you want to see
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        change_job_selection()
    End Sub
    Private Sub change_job_selection()
        Dim now_min_2wks As DateTime
        Dim sql_str, filter_str, sort_str As String

        now_min_2wks = New DateTime(Now.Year, Now.Month, Now.Day, 8, 0, 0, 0)  'start time 
        now_min_2wks = DateAdd("d", (NumericUpDown1.Value * -7 + 1), now_min_2wks)                          'Subtract x weeks
        sql_str = String.Format("[Finished_date] > #{0:MM/dd/yyyy hh:mm:ss}# ", now_min_2wks)

        Try
            '---------------Making sure that comboboxes are not empty-----------------
            If ComboBox1.Items.Count > 0 And ComboBox2.Items.Count > 0 And ComboBox3.Items.Count > 0 Then

                If ComboBox1.Text = "ALL" Then
                    filter_str = " "
                    sort_str = "Job_nr, Start_date, Finish_date ASC"
                ElseIf ComboBox1.Text = "ALL FINISHED" Then
                    filter_str = "[Finished] = 100"
                    sort_str = "Job_nr, Start_date, Finish_date ASC"
                ElseIf ComboBox1.Text = "ALL PENDING -x wks" Then
                    filter_str = sql_str & " OR [Finished] < 100"
                    sort_str = "Job_nr, Start_date, Finish_date ASC"
                ElseIf ComboBox1.Text = "BOOK DATE" Then
                    filter_str = "[Booking] > #01/01/2014#"
                    sort_str = "Booking, Job_nr, Start_date, Finish_date ASC"
                Else
                    filter_str = "[TC]='" & ComboBox1.Text & "' AND ([Finished] < 100 OR " & sql_str & ")"
                    sort_str = "Start_date, Finish_date ASC"
                End If

                '-------------------- add the department selection --------------------
                If ComboBox3.Text <> "ALL" Then
                    filter_str = filter_str & " AND [Department]= " & ComboBox3.Text
                End If
                JobsBindingSource.Filter = filter_str
                JobsBindingSource.Sort = sort_str
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim filter_str, sort_str As String
        Try
            '---------------Making sure that comboboxes are not empty-----------------
            If ComboBox1.Items.Count > 0 And ComboBox2.Items.Count > 0 And ComboBox3.Items.Count > 0 Then

                If Not ComboBox2.Text = "Job_nr" Then
                    filter_str = "[Job_nr]= '" & ComboBox2.Text & "'"
                    sort_str = "Job_nr, Start_date, Finish_date ASC"
                Else
                    filter_str = " "
                    sort_str = "Job_nr, Start_date, Finish_date ASC"
                End If

                '-------------------- add the department selection --------------------
                If ComboBox3.Text <> "ALL" And Not ComboBox2.Text = "Job_nr" Then
                    filter_str = filter_str & " AND [Department]= " & ComboBox3.Text
                End If
                JobsBindingSource.Filter = filter_str
                JobsBindingSource.Sort = sort_str
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    'Job done ?
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        Dim cnn As SqlConnection
        Dim result As Integer
        Dim rowindex As Integer

        'Indicate the row with Green background
        rowindex = Me.JobsDataGridView.CurrentCell.RowIndex()
        JobsDataGridView.Rows(rowindex).DefaultCellStyle.BackColor = Color.GreenYellow
        result = MessageBox.Show("       Job Done ?", "JOB FINISHED", MessageBoxButtons.YesNoCancel)
        JobsDataGridView.Rows(rowindex).DefaultCellStyle.BackColor = Color.White

        If result = DialogResult.Yes Then
            Try
                cnn = New SqlConnection(connectionString)
                cnn.Open()
                JobsDataGridView.Rows(rowindex).Cells(10).Value = Now         'Finish_date/Done_date
                JobsDataGridView.Rows(rowindex).Cells(7).Value = 100          'Finished (%)
                JobsDataGridView.Refresh()
                TableAdapterManager1.UpdateAll(_VTK_PlanningDataSet1)
                cnn.Close() 'connection close here , that is disconnected from data source
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
    'Row finished job change color
    Private Sub JobsDataGridView_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs) Handles JobsDataGridView.RowPrePaint

        If Not IsDBNull(JobsDataGridView.Rows(e.RowIndex).Cells(7).Value) Then
            If (JobsDataGridView.Rows(e.RowIndex).Cells(7).Value = 100) And (JobsDataGridView.Rows(e.RowIndex).DefaultCellStyle.BackColor <> Color.Red) Then
                JobsDataGridView.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Beige
            End If
            If (JobsDataGridView.Rows(e.RowIndex).Cells(7).Value < 100) And (JobsDataGridView.Rows(e.RowIndex).DefaultCellStyle.BackColor <> Color.Red) And
                (JobsDataGridView.Rows(e.RowIndex).DefaultCellStyle.BackColor <> Color.GreenYellow) Then
                JobsDataGridView.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
            End If
        End If
    End Sub
    'Week selection changed
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        change_job_selection()
    End Sub
    'Department
    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        change_job_selection()
    End Sub

End Class
