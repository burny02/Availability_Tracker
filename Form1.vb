Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUp()

        Me.Text = SolutionName

        If Bank = False Then
            Label8.Text = "Unavailability"
        End If

        RefreshList()

    End Sub

    Public Sub RefreshList()

        Dim FirstOfMonth As Date = DateSerial(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, 1)
        Dim LastOfMonth As Date = DateSerial(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month + 1, 0)
        Dim FirstWeekDay As Integer
        Dim FirstGridNumber As Integer
        Dim Table As String

        If Bank = True Then
            Table = "Bank"
        Else
            Table = "Other"
        End If

        FirstWeekDay = Weekday(FirstOfMonth, FirstDayOfWeek.Monday)

        Dim Ctl As Control() = Controls.Find("DataGridView" & FirstWeekDay, True)
        FirstGridNumber = CInt(Replace(Ctl(0).Name, "DataGridView", ""))

        Dim NumberOfMonthDays As Integer = DateDiff(DateInterval.Day, FirstOfMonth, LastOfMonth) + 1

        Dim i As Integer = 0
        Dim SQLArray() As String
        Do While i < NumberOfMonthDays
            ReDim Preserve SQLArray(i)
            Dim CurDate As Date = DateAdd(DateInterval.Day, i, FirstOfMonth)
            SQLArray(i) = "SELECT format(StartDate,'HH:mm') & ' > ' & format(StopDate,'HH:mm') FROM " & Table & " WHERE StaffID=" & StaffID & " AND format(StartDate,'dd-MMM-yyyy')='" & Format(CurDate, "dd-MMM-yyyy") & "'"
            i += 1
        Loop

        i -= 1
        ReDim Preserve SQLArray(i)
        Dim DtArray() As DataTable = OverClass.MultiTempDataTable(SQLArray)

        i = 1
        Do While i <= 42
            Dim TempCtl As Control() = Controls.Find("DataGridView" & i, True)
            Dim TempDGV As DataGridView = TempCtl(0)
            TempDGV.Columns.Clear()
            TempDGV.Visible = False
            i += 1
        Loop

        i = 0
        Do While i < NumberOfMonthDays
            Dim TempCtl As Control() = Controls.Find("DataGridView" & FirstGridNumber + i, True)
            Dim TempDGV As DataGridView = TempCtl(0)
            TempDGV.DataSource = DtArray(i)
            Dim CurDate As Date = DateAdd(DateInterval.Day, i, FirstOfMonth)
            TempDGV.Columns(0).HeaderText = Format(CurDate, "dd-MMM")
            TempDGV.Visible = True
            i += 1
        Loop




    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        RefreshList()
    End Sub

    Private Sub DataGridView7_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView9.DoubleClick, DataGridView8.DoubleClick, DataGridView7.DoubleClick, DataGridView6.DoubleClick, DataGridView5.DoubleClick, DataGridView42.DoubleClick, DataGridView41.DoubleClick, DataGridView40.DoubleClick, DataGridView4.DoubleClick, DataGridView39.DoubleClick, DataGridView38.DoubleClick, DataGridView37.DoubleClick, DataGridView36.DoubleClick, DataGridView35.DoubleClick, DataGridView34.DoubleClick, DataGridView33.DoubleClick, DataGridView32.DoubleClick, DataGridView31.DoubleClick, DataGridView30.DoubleClick, DataGridView3.DoubleClick, DataGridView29.DoubleClick, DataGridView28.DoubleClick, DataGridView27.DoubleClick, DataGridView26.DoubleClick, DataGridView25.DoubleClick, DataGridView24.DoubleClick, DataGridView23.DoubleClick, DataGridView22.DoubleClick, DataGridView21.DoubleClick, DataGridView20.DoubleClick, DataGridView2.DoubleClick, DataGridView19.DoubleClick, DataGridView18.DoubleClick, DataGridView17.DoubleClick, DataGridView16.DoubleClick, DataGridView15.DoubleClick, DataGridView14.DoubleClick, DataGridView13.DoubleClick, DataGridView12.DoubleClick, DataGridView11.DoubleClick, DataGridView10.DoubleClick, DataGridView1.DoubleClick
        MsgBox("ok")
    End Sub
End Class
