Imports System.Data
Imports System.Data.OleDb

Public Class Form6
    Dim firma As String = Form1.ComboBox1.Text
    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Dim newbagla As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
                             "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;"
    Dim myConnection As New OleDbConnection(newbagla)
    Dim myDataReader As OleDbDataReader
    Dim objDataSet As DataSet
    Dim objDataView As DataView
    Dim objDataView1 As DataView
    Dim sqlString As String
    Dim gıcık As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
    "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb; Persist Security Info=False "
    Dim objConnection As New OleDbConnection(gıcık)

    Private Sub FillDataSetAndView()

        Dim objDataAdapter As New OleDbDataAdapter("Select * FROM Riskler order by bölüm asc ", newbagla)
        objDataSet = New DataSet()
        objDataAdapter.Fill(objDataSet, "Riskler")
        objDataView = New DataView(objDataSet.Tables("Riskler"))

    End Sub
    Private Sub DataGridBind()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView
        DataGridView2.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView1
    End Sub
    Private Sub Combox01()
        Dim adapter As New OleDbDataAdapter("SELECT distinct bölüm FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim bölüm As String

        For i As Integer = 0 To table.Rows.Count - 1
            bölüm = table(i)(0)
            ComboBox1.Items.Add(bölüm)
        Next
    End Sub
    Private Sub Combox02()
        Dim adapter As New OleDbDataAdapter("SELECT distinct prosesekipmanadı FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim prosesekipmanadı As String

        For i As Integer = 0 To table.Rows.Count - 1
            prosesekipmanadı = table(i)(0)
            ComboBox2.Items.Add(prosesekipmanadı)
        Next
    End Sub
    Private Sub Combox04()
        Dim adapter As New OleDbDataAdapter("SELECT distinct kategori FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim kategori As String

        For i As Integer = 0 To table.Rows.Count - 1
            kategori = table(i)(0)
            ComboBox3.Items.Add(kategori)
        Next
    End Sub
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillDataSetAndView()
        DataGridBind()
        Combox01()
        Combox02()
        Combox04()
    End Sub
    Private Sub bölümara()
        Dim adtr As New OleDbDataAdapter("selecet * from Riskler  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Riskler" & " where(bölüm like '%") + ComboBox1.Text & "%' )"
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Riskler")

    End Sub
    Private Sub makineara()
        Dim adtr As New OleDbDataAdapter("selecet * from Riskler ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Riskler" & " where(prosesekipmanadı like '%") + ComboBox2.Text & "%' )"
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Riskler")
    End Sub
    Private Sub kategoriara()
        Dim adtr As New OleDbDataAdapter("selecet * from Riskler  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Riskler" & " where(kategori like '%") + ComboBox3.Text & "%' )"
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Riskler")
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ComboBox1.SelectedItem IsNot Nothing Then bölümara()
        If ComboBox2.SelectedItem IsNot Nothing Then makineara()
        If ComboBox3.SelectedItem IsNot Nothing Then kategoriara()

        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
        ComboBox3.SelectedItem = Nothing
        ComboBox3.Text = ""
        ComboBox1.Text = ""
        ComboBox2.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form10.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataSetAndView()
        DataGridBind()
    End Sub

    Private Sub Form6_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form10.Show()
        Me.Hide()
    End Sub
End Class