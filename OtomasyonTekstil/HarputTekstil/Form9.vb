Imports System.Data.OleDb
Imports System.Data
Public Class Form9
    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Dim firma As String = Form1.ComboBox1.Text

    Dim newbagla As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
                            "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;"
    Dim myConnection As New OleDbConnection(newbagla)
    Dim myDataReader As OleDbDataReader
    Dim objDataSet As DataSet
    Dim objDataView As DataView
    Dim gıcık As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
   "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb; Persist Security Info=False "
    Dim objConnection As New OleDbConnection(gıcık)
    Dim objcurrencyManager As CurrencyManager
    Dim lst As String
    Dim int As String

    Private Sub FillDataSetAndView()

        Dim objDataAdapter As New OleDbDataAdapter("Select * FROM Sorgu WHERE HedefTarih is null " & "order by say1", objConnection)
        objDataSet = New DataSet()
        objDataAdapter.Fill(objDataSet, "Sorgu")
        objDataView = New DataView(objDataSet.Tables("Sorgu"))

        objcurrencyManager = CType(Me.BindingContext(objDataView), CurrencyManager)
    End Sub
    Private Sub DataGridBind()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView
        DataGridView1.Columns("HedefTarih").Visible = False
        DataGridView1.Columns("GerçekleşmeTarihi").Visible = False
        DataGridView1.Columns("Tamamlandı").Visible = False
    End Sub
    Private Sub BindFields()
        TextBox1.DataBindings.Clear()
        TextBox1.DataBindings.Add("Text", objDataView, "say1")
    End Sub
    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillDataSetAndView()
        DataGridBind()
        Combox01()
        Combox02()
        Combox3()
        BindFields()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form10.Show()
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FillDataSetAndView()
        DataGridBind()
        BindFields()
    End Sub
    Private Sub Combox01()
        Dim adapter As New OleDbDataAdapter("SELECT distinct bölüm FROM sorgu", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim bölüm As String

        For i As Integer = 0 To table.Rows.Count - 1
            bölüm = table(i)(0)
            ComboBox1.Items.Add(bölüm)
        Next
    End Sub
    Private Sub Combox02()
        Dim adapter As New OleDbDataAdapter("SELECT distinct prosesekipmanadı FROM sorgu", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim prosesekipmanadı As String

        For i As Integer = 0 To table.Rows.Count - 1
            prosesekipmanadı = table(i)(0)
            ComboBox2.Items.Add(prosesekipmanadı)
        Next
    End Sub
    Private Sub Combox3()
        Dim adapter As New OleDbDataAdapter("SELECT distinct kategori FROM sorgu", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim kategori As String

        For i As Integer = 0 To table.Rows.Count - 1
            kategori = table(i)(0)
            ComboBox3.Items.Add(kategori)
        Next
    End Sub
    Private Sub bölümara()
        Dim adtr As New OleDbDataAdapter("select * from sorgu  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From sorgu where HedefTarih is null AND bölüm like '%" & ComboBox1.Text & "%' ")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "sorgu")

    End Sub
    Private Sub makineara()
        Dim adtr As New OleDbDataAdapter("select * from sorgu ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From sorgu where HedefTarih is null AND prosesekipmanadı like '%" & ComboBox2.Text & "%' ")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "sorgu")
    End Sub
    Private Sub kategoriara()
        Dim adtr As New OleDbDataAdapter("select * from sorgu  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From sorgu where HedefTarih is null AND kategori like '%" & ComboBox3.Text & "%' ")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "sorgu")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Int = objcurrencyManager.Position
        Dim objCommand As OleDbCommand = New OleDbCommand()

        objCommand.Connection = objConnection
        objCommand.Parameters.AddWithValue("@Say1", CType(TextBox1.Text, Integer))

        objCommand.CommandText = "DELETE FROM AksiyonPlanı " & "WHERE [Say1] = @Say1 "

        objConnection.Open()
        ' Execute the SqlCommand object to insert the new data...
        Try
            objCommand.ExecuteNonQuery()
        Catch oledbExceptionErr As OleDbException
            MessageBox.Show(oledbExceptionErr.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ' Close the connection...
            objConnection.Close()
        End Try
        FillDataSetAndView()
        BindFields()
        DataGridBind()
        objcurrencyManager.Position = int
    End Sub

    Private Sub Form9_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form10.Show()
        Me.Hide()
    End Sub
End Class
