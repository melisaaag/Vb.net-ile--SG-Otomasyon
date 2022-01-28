Imports System.Data.OleDb
Imports System.Data
Public Class Form8
    Dim firma As String = Form1.ComboBox1.Text
    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
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

        Dim objDataAdapter As New OleDbDataAdapter("Select * FROM Sorgu WHERE HedefTarih is not null " & "order by say1", objConnection)
        objDataSet = New DataSet()
        objDataAdapter.Fill(objDataSet, "Sorgu")
        objDataView = New DataView(objDataSet.Tables("Sorgu"))
        objcurrencyManager = CType(Me.BindingContext(objDataView), CurrencyManager)

    End Sub
    Private Sub DataGridBind()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView
        DataGridView1.Columns("MevcutÖnlemler").Visible = False

    End Sub
    Private Sub BindFields()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        TextBox7.DataBindings.Clear()


        TextBox1.DataBindings.Add("Text", objDataView, "HedefTarih")
        TextBox2.DataBindings.Add("Text", objDataView, "AlınmasıGerekenÖnlemler")
        TextBox3.DataBindings.Add("Text", objDataView, "GerçekleşmeTarihi")
        TextBox5.DataBindings.Add("Text", objDataView, "Sorumlu")
        CheckBox1.DataBindings.Add("Checked", objDataView, "Tamamlandı")
        TextBox7.DataBindings.Add("Text", objDataView, "say1")


    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillDataSetAndView()
        DataGridBind()
        BindFields()
        tarıh()
        DataGridView()
        Combox01()
        Combox02()
        Combox3()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form10.Show()
        Me.Close()
    End Sub
    Private Sub tarıh()

        Dim tarih2 As Date = Today
        TextBox4.Text = tarih2

    End Sub
    ''kalan gün hesapla butonu
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles BtnGo.Click

        Dim tarih1 As Date = TextBox1.Text
        Dim tarih2 As Date = TextBox4.Text
        Dim zaman As Integer
        Dim gun As Integer = -1 * DateDiff(DateInterval.Day, tarih2, tarih1)

        zaman = -1 * (gun)

        If gun > 0 Then
            zaman *= -1
            MessageBox.Show("Aksiyon kapatma " & zaman & " gün gecikti", "UYARI")
            TextBox6.BackColor = Color.Red

            TextBox6.Text = zaman
        End If
        If gun < 0 Then

            MessageBox.Show("Aksiyonu kapatmak için " & -1 * gun & " gün kaldı", "UYARI")
            TextBox6.BackColor = Color.LightGreen
            TextBox6.Text = zaman
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        int = objcurrencyManager.Position
        Dim objCommand As OleDbCommand = New OleDbCommand()

        objCommand.Connection = objConnection
        objCommand.Parameters.AddWithValue("@Say1", CType(TextBox7.Text, Integer))

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
        DataGridView()
        objcurrencyManager.Position = int
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        int = objcurrencyManager.Position
        Dim objCommand As OleDbCommand = New OleDbCommand()
        objCommand.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox2.Text, String))
        objCommand.Parameters.AddWithValue("@Sorumlu", CType(TextBox5.Text, String))
        If TextBox3.Text IsNot "" Then objCommand.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox3.Text, Date))
        objCommand.Parameters.AddWithValue("@HedefTarih", CType(TextBox1.Text, Date))
        objCommand.Parameters.AddWithValue("@Tamamlandı", CType(CheckBox1.Checked, Boolean))
        objCommand.Parameters.AddWithValue("@Say1", CType(TextBox7.Text, Integer))
        objCommand.Connection = objConnection

        If TextBox3.Text IsNot "" Then objCommand.CommandText = "UPDATE [AksiyonPlanı] SET [AlınmasıGerekenÖnlemler] = @AlınmasıGerekenÖnlemler, [Sorumlu] = @Sorumlu, [GerçekleşmeTarihi]=@GerçekleşmeTarihi, [HedefTarih] = @HedefTarih, [Tamamlandı] = @Tamamlandı WHERE [Say1] = @Say1 ;"
        If TextBox3.Text Is "" Then objCommand.CommandText = "UPDATE [AksiyonPlanı] SET [AlınmasıGerekenÖnlemler] = @AlınmasıGerekenÖnlemler, [Sorumlu] = @Sorumlu, [HedefTarih] = @HedefTarih, [Tamamlandı] = @Tamamlandı WHERE [Say1] = @Say1 ;"
        Try

            objConnection.Open()
            'Execute the SqlCommand object to insert the New data...
            objCommand.ExecuteNonQuery()

        Catch oledbExceptionErr As OleDbException
            MessageBox.Show(oledbExceptionErr.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            objConnection.Close()
        End Try

        'Fill the dataset And bind the fields...
        FillDataSetAndView()
        BindFields()
        DataGridBind()
        DataGridView()
        objcurrencyManager.Position = int
    End Sub
    Private Sub DataGridView()


        For i As Integer = 0 To DataGridView1.Rows.Count - 1



            Dim tarih2 As Date = Today
            Dim tarih As Date = DataGridView1.Rows(i).Cells(7).Value
            Dim gun = DateDiff(DateInterval.Day, tarih2, tarih)
            If gun <= 30 And gun >= 0 Then
                DataGridView1.Rows(i).Cells(7).Style.BackColor = Color.Yellow

            End If
            If gun <= 0 Then
                DataGridView1.Rows(i).Cells(7).Style.BackColor = Color.OrangeRed

            End If
            If DataGridView1(9, i).Value() = True Then
                DataGridView1.Rows(i).Cells(7).Style.BackColor = Color.LightGreen
            End If
        Next

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        DataGridView()
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
        Dim adtr As New OleDbDataAdapter("select * from Sorgu  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Sorgu where HedefTarih is not null AND Bölüm like '%" & ComboBox1.Text & "%'")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Sorgu")

    End Sub
    Private Sub makineara()
        Dim adtr As New OleDbDataAdapter("select * from Sorgu ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Sorgu where HedefTarih is not null AND ProsesEkipmanAdı like '%" & ComboBox2.Text & "%'")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "sorgu")
    End Sub
    Private Sub kategoriara()
        Dim adtr As New OleDbDataAdapter("select * from Sorgu  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Sorgu where HedefTarih is not null AND Kategori like '%" & ComboBox3.Text & "%'")
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
        BindFields()
        DataGridView()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FillDataSetAndView()
        DataGridBind()
        DataGridView()
    End Sub

    Private Sub Form8_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form10.Show()
        Me.Hide()
    End Sub
End Class
