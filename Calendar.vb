Imports System
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.VisualBasic

Public Class Calendar
    Private Sub Calendar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Cursor = Cursors.WaitCursor
        Me.CenterToScreen()
        lblDate.Hide()
        lblTitre.Hide()
        initBDD()
        Dim cmd As OleDbCommand
        Dim jeuEnr As OleDbDataReader

        cnn.Open()
        cmd = New OleDbCommand
        cmd.Connection = cnn
        cmd.CommandText = "SELECT DateEvent, Titre FROM EVENEMENTS "

        jeuEnr = cmd.ExecuteReader

        While jeuEnr.Read()
            Dim dtArrSpecialDates As Date
            Dim titre As String
            titre = jeuEnr.GetValue(1)
            dtArrSpecialDates = jeuEnr.GetValue(0)
            MonthCalendar1.AddBoldedDate(dtArrSpecialDates)
        End While
        MonthCalendar1.UpdateBoldedDates()
        MonthCalendar1.ShowTodayCircle = True
        cnn.Close()

        cnn.Open()
        cmd = New OleDbCommand
        cmd.Connection = cnn
        cmd.CommandText = "SELECT DateEvent, Titre FROM EVENEMENTS "

        jeuEnr = cmd.ExecuteReader

        While jeuEnr.Read()
            Dim dtArrSpecialDates As Date
            Dim titre As String
            titre = jeuEnr.GetValue(1)
            dtArrSpecialDates = jeuEnr.GetValue(0)
            MonthCalendar1.AddBoldedDate(dtArrSpecialDates)
        End While
        MonthCalendar1.UpdateBoldedDates()
        MonthCalendar1.ShowTodayCircle = True
        cnn.Close()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TousLesEvents.Show()
        Me.Close()
    End Sub

    Private Sub MonthCalendar1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        lblDate.Show()
        lblTitre.Show()

        Dim cmd As OleDbCommand
        Dim jeuEnr As OleDbDataReader

        cnn.Open()
        cmd = New OleDbCommand
        cmd.Connection = cnn
        cmd.CommandText = "SELECT Titre FROM EVENEMENTS WHERE DateEvent LIKE '" & MonthCalendar1.SelectionStart & "'"

        jeuEnr = cmd.ExecuteReader
        Dim titre As String
        If IsDate(MonthCalendar1.SelectionStart) Then
            If jeuEnr.Read() Then


                titre = jeuEnr.GetValue(0)
                lblTitre.Text = titre
            Else
                titre = ""
                lblTitre.Text = "Pas d'évènement prévu pour cette date"
            End If
            lblDate.Text = MonthCalendar1.SelectionStart
            cnn.Close()
        End If

    End Sub

End Class