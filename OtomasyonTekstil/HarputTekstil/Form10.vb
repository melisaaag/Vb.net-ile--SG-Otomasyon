Imports System.Data
Imports System.Data.OleDb
Imports System.IO


Public Class Form10

    Dim firma As String = Form1.ComboBox1.Text
    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

    Dim newbagla As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
                              "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;"
    Dim myConnection As New OleDbConnection(newbagla)
    Dim strSQL As String = "Select * FROM Riskler"
    Dim myCommand As New OleDbCommand(strSQL, myConnection)
    Dim myDataReader As OleDbDataReader
    Dim objDataSet As DataSet
    Dim objDataView As DataView
    Dim gıcık As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
    "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;Persist Security Info=False "
    Dim objConnection As New OleDbConnection(gıcık)
    Dim objcurrencyManager As CurrencyManager
    Dim lst As String
    Dim int As String
    Private Sub FillDataSetAndView()

        Dim objDataAdapter As New OleDbDataAdapter("Select * FROM Riskler order by Sıra", newbagla)
        objDataSet = New DataSet()
        objDataAdapter.Fill(objDataSet, "Riskler")
        objDataView = New DataView(objDataSet.Tables("Riskler"))
        objcurrencyManager = CType(Me.BindingContext(objDataView), CurrencyManager)

    End Sub
    Private Sub DataGridBind()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView
        DataGridView1.Columns("Sıra").Visible = False
    End Sub
    Private Sub BindFields()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        TextBox10.DataBindings.Clear()
        TextBox11.DataBindings.Clear()
        TextBox12.DataBindings.Clear()
        TextBox13.DataBindings.Clear()
        TextBox14.DataBindings.Clear()
        TextBox15.DataBindings.Clear()
        TextBox16.DataBindings.Clear()
        TextBox17.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox01.DataBindings.Clear()
        ComboBox02.DataBindings.Clear()
        ComboBox04.DataBindings.Clear()
        ComboBox05.DataBindings.Clear()
        ComboBox06.DataBindings.Clear()
        ComboBox07.DataBindings.Clear()
        TextBox18.DataBindings.Clear()
        TextBox19.DataBindings.Clear()
        TextBox20.DataBindings.Clear()
        TextBox21.DataBindings.Clear()
        TextBox22.DataBindings.Clear()
        TextBox23.DataBindings.Clear()
        TextBox24.DataBindings.Clear()
        TextBox25.DataBindings.Clear()
        TextBox26.DataBindings.Clear()
        TextBox27.DataBindings.Clear()
        TextBox28.DataBindings.Clear()
        TextBox29.DataBindings.Clear()
        TextBox30.DataBindings.Clear()
        TextBox31.DataBindings.Clear()
        TextBox32.DataBindings.Clear()
        TextBox33.DataBindings.Clear()
        TextBox34.DataBindings.Clear()
        TextBox35.DataBindings.Clear()
        TextBox36.DataBindings.Clear()
        TextBox37.DataBindings.Clear()
        TextBox38.DataBindings.Clear()
        TextBox39.DataBindings.Clear()
        TextBox40.DataBindings.Clear()
        TextBox41.DataBindings.Clear()
        TextBox42.DataBindings.Clear()
        TextBox43.DataBindings.Clear()
        TextBox44.DataBindings.Clear()
        TextBox45.DataBindings.Clear()
        TextBox49.DataBindings.Clear()
        TextBox46.DataBindings.Clear()
        TextBox47.DataBindings.Clear()
        TextBox048.DataBindings.Clear()
        TextBox50.DataBindings.Clear()
        TextBox51.DataBindings.Clear()
        TextBox52.DataBindings.Clear()

        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox15.Clear()
        TextBox16.Clear()
        TextBox17.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox01.Text = ""
        ComboBox02.Text = ""
        ComboBox04.Text = ""
        ComboBox05.Text = ""
        ComboBox06.Text = ""
        ComboBox07.Text = ""
        TextBox18.Clear()
        TextBox19.Clear()
        TextBox20.Clear()
        TextBox21.Clear()
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox24.Clear()
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox30.Clear()
        TextBox31.Clear()
        TextBox32.Clear()
        TextBox33.Clear()
        TextBox34.Clear()
        TextBox35.Clear()
        TextBox36.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        TextBox39.Clear()
        TextBox40.Clear()
        TextBox41.Clear()
        TextBox42.Clear()
        TextBox43.Clear()
        TextBox44.Clear()
        TextBox45.Clear()
        TextBox51.Clear()
        TextBox52.Clear()
        TextBox50.Clear()
        TextBox048.Clear()
        TextBox46.Clear()
        TextBox47.Clear()
        TextBox49.Clear()

        TextBox1.DataBindings.Add("Text", objDataView, "No")
        TextBox2.DataBindings.Add("Text", objDataView, "Yapılanİş")
        TextBox5.DataBindings.Add("Text", objDataView, "Risk")
        TextBox3.DataBindings.Add("Text", objDataView, "RiskDeğeri")
        TextBox4.DataBindings.Add("Text", objDataView, "ÖncelikDurumu")
        TextBox6.DataBindings.Add("Text", objDataView, "MevcutÖnlemler")
        ComboBox01.DataBindings.Add("Text", objDataView, "Bölüm")
        ComboBox02.DataBindings.Add("Text", objDataView, "ProsesEkipmanAdı")
        TextBox16.DataBindings.Add("Text", objDataView, "Tehlike")
        ComboBox04.DataBindings.Add("Text", objDataView, "Kategori")
        ComboBox05.DataBindings.Add("Text", objDataView, "Etkilenenler")
        ComboBox06.DataBindings.Add("Text", objDataView, "İhtimal")
        ComboBox07.DataBindings.Add("Text", objDataView, "Şiddet")
        TextBox17.DataBindings.Add("Text", objDataView, "YasalZorunluluk")
        TextBox18.DataBindings.Add("Text", objDataView, "AlınmasıGerekenÖnlemler")
        TextBox51.DataBindings.Add("Text", objDataView, "Sorumlu")
        TextBox52.DataBindings.Add("Text", objDataView, "GerçekleşmeTarihi")
        TextBox50.DataBindings.Add("Text", objDataView, "HedefTarih")
        TextBox49.DataBindings.Add("Text", objDataView, "İlgiliFotoğrafNo")
        ComboBox1.DataBindings.Add("Text", objDataView, "Güncelİhtimal")
        ComboBox2.DataBindings.Add("Text", objDataView, "GüncelŞiddet")
        TextBox46.DataBindings.Add("Text", objDataView, "GüncelRiskDeğeri")
        TextBox47.DataBindings.Add("Text", objDataView, "GüncelÖncelikDurumu")
        TextBox048.DataBindings.Add("Text", objDataView, "GüncelDurum")
        Me.Text = firma

    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form1.Show()
        Me.Close()
    End Sub

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillDataSetAndView()
        DataGridBind()
        BindFields()
        Combox01()
        Combox02()
        Combox04()
        Combox05()

    End Sub
    Private Sub Combox01()
        Dim adapter As New OleDbDataAdapter("Select distinct bölüm FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim bölüm As String

        For i As Integer = 0 To table.Rows.Count - 1
            bölüm = table(i)(0)
            ComboBox01.Items.Add(bölüm)
        Next
    End Sub
    Private Sub Combox02()
        Dim adapter As New OleDbDataAdapter("Select distinct prosesekipmanadı FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim prosesekipmanadı As String

        For i As Integer = 0 To table.Rows.Count - 1
            prosesekipmanadı = table(i)(0)
            ComboBox02.Items.Add(prosesekipmanadı)
        Next
    End Sub
    Private Sub Combox04()
        Dim adapter As New OleDbDataAdapter("Select distinct kategori FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim kategori As String

        For i As Integer = 0 To table.Rows.Count - 1
            kategori = table(i)(0)
            ComboBox04.Items.Add(kategori)
        Next
    End Sub
    Private Sub Combox05()
        Dim adapter As New OleDbDataAdapter("Select distinct etkilenenler FROM Riskler", newbagla)

        Dim table As New DataTable()
        adapter.Fill(table)

        Dim etkilenenler As String

        For i As Integer = 0 To table.Rows.Count - 1
            etkilenenler = table(i)(0)
            ComboBox05.Items.Add(etkilenenler)
        Next
    End Sub

    Private Sub ComboBox06_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox06.SelectedIndexChanged
        TextBox3.Text = ComboBox06.SelectedItem * ComboBox07.SelectedItem
        If Val(TextBox3.Text) <= 1 Then
            TextBox4.Text = "Anlamsız"
            TextBox4.BackColor = Color.MistyRose
        ElseIf Val(TextBox3.Text) >= 1 And Val(TextBox3.Text) <= 6 Then
            TextBox4.Text = "Düşük"
            TextBox4.BackColor = Color.LightGreen
        ElseIf Val(TextBox3.Text) >= 6 And Val(TextBox3.Text) <= 12 Then
            TextBox4.Text = "Orta"
            TextBox4.BackColor = Color.Yellow
        ElseIf Val(TextBox3.Text) >= 12 And Val(TextBox3.Text) <= 20 Then
            TextBox4.Text = "Yüksek"
            TextBox4.BackColor = Color.Red
        ElseIf Val(TextBox3.Text) = 25 Then
            TextBox4.Text = "Kabul Edilemez"
            TextBox4.BackColor = Color.MistyRose
        ElseIf Val(TextBox3.Text) > 25 Then
            TextBox4.Text = "Değerlerinizi kontrol ediniz!"
        End If
    End Sub

    Private Sub ComboBox07_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox07.SelectedIndexChanged
        TextBox3.Text = ComboBox06.SelectedItem * ComboBox07.SelectedItem
        If Val(TextBox3.Text) <= 1 Then
            TextBox4.Text = "Anlamsız"
            TextBox4.BackColor = Color.MistyRose
        ElseIf Val(TextBox3.Text) >= 1 And Val(TextBox3.Text) <= 6 Then
            TextBox4.Text = "Düşük"
            TextBox4.BackColor = Color.LightGreen
        ElseIf Val(TextBox3.Text) >= 6 And Val(TextBox3.Text) <= 12 Then
            TextBox4.Text = "Orta"
            TextBox4.BackColor = Color.Yellow
        ElseIf Val(TextBox3.Text) >= 12 And Val(TextBox3.Text) <= 20 Then
            TextBox4.Text = "Yüksek"
            TextBox4.BackColor = Color.Red
        ElseIf Val(TextBox3.Text) = 25 Then
            TextBox4.Text = "Kabul Edilemez"
        ElseIf Val(TextBox3.Text) > 25 Then
            TextBox4.Text = "Değerlerinizi kontrol ediniz!"
            TextBox4.BackColor = Color.MistyRose
        End If

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox46.Text = ComboBox1.SelectedItem * ComboBox2.SelectedItem
        If Val(TextBox46.Text) <= 1 Then
            TextBox47.Text = "Anlamsız"
            TextBox47.BackColor = Color.MistyRose
        ElseIf Val(TextBox46.Text) >= 1 And Val(TextBox46.Text) <= 6 Then
            TextBox47.Text = "Düşük"
            TextBox47.BackColor = Color.LightGreen
        ElseIf Val(TextBox46.Text) >= 6 And Val(TextBox46.Text) <= 12 Then
            TextBox47.Text = "Orta"
            TextBox47.BackColor = Color.Yellow
        ElseIf Val(TextBox46.Text) >= 12 And Val(TextBox46.Text) <= 20 Then
            TextBox47.Text = "Yüksek"
            TextBox47.BackColor = Color.Red
        ElseIf Val(TextBox46.Text) = 25 Then
            TextBox47.Text = "Kabul Edilemez"
            TextBox47.BackColor = Color.MistyRose
        ElseIf Val(TextBox46.Text) > 25 Then
            TextBox47.Text = "Değerlerinizi kontrol ediniz!"
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox46.Text = ComboBox1.SelectedItem * ComboBox2.SelectedItem
        If Val(TextBox46.Text) <= 1 Then
            TextBox47.Text = "Anlamsız"
            TextBox47.BackColor = Color.MistyRose
        ElseIf Val(TextBox46.Text) >= 1 And Val(TextBox46.Text) <= 6 Then
            TextBox47.Text = "Düşük"
            TextBox47.BackColor = Color.LightGreen
        ElseIf Val(TextBox46.Text) >= 6 And Val(TextBox46.Text) <= 12 Then
            TextBox47.Text = "Orta"
            TextBox47.BackColor = Color.Yellow
        ElseIf Val(TextBox46.Text) >= 12 And Val(TextBox46.Text) <= 20 Then
            TextBox47.Text = "Yüksek"
            TextBox47.BackColor = Color.Red
        ElseIf Val(TextBox46.Text) = 25 Then
            TextBox4.Text = "Kabul Edilemez"
        ElseIf Val(TextBox46.Text) > 25 Then
            TextBox47.Text = "Değerlerinizi kontrol ediniz!"
            TextBox47.BackColor = Color.MistyRose
        End If

    End Sub
    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        If TextBox7.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox7.Text
        If TextBox8.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox8.Text
        If TextBox9.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox9.Text
        If TextBox10.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox10.Text
        If TextBox11.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox11.Text
        If TextBox12.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox12.Text
        If TextBox13.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox13.Text
        If TextBox14.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox14.Text
        If TextBox15.Text IsNot "" Then TextBox18.Text = TextBox18.Text & vbCrLf & TextBox15.Text

        If TextBox28.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox28.Text
        If TextBox29.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox29.Text
        If TextBox30.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox30.Text
        If TextBox31.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox31.Text
        If TextBox32.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox32.Text
        If TextBox33.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox33.Text
        If TextBox34.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox34.Text
        If TextBox35.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox35.Text
        If TextBox36.Text IsNot "" Then TextBox51.Text = TextBox51.Text & vbCrLf & TextBox36.Text

        If TextBox19.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox19.Text
        If TextBox20.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox20.Text
        If TextBox21.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox21.Text
        If TextBox22.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox22.Text
        If TextBox23.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox23.Text
        If TextBox24.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox24.Text
        If TextBox25.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox25.Text
        If TextBox26.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox26.Text
        If TextBox27.Text IsNot "" Then TextBox50.Text = TextBox50.Text & vbCrLf & TextBox27.Text

        If TextBox45.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox45.Text
        If TextBox44.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox44.Text
        If TextBox43.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox43.Text
        If TextBox42.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox42.Text
        If TextBox41.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox41.Text
        If TextBox40.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox40.Text
        If TextBox39.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox39.Text
        If TextBox38.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox38.Text
        If TextBox37.Text IsNot "" Then TextBox52.Text = TextBox52.Text & vbCrLf & TextBox37.Text
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        int = objcurrencyManager.Position
        Dim objCommand As OleDbCommand = New OleDbCommand()

        objCommand.Parameters.AddWithValue("@Bölüm", CType(ComboBox01.Text, String))
        objCommand.Parameters.AddWithValue("@ProsesEkipmanAdı", CType(ComboBox02.Text, String))
        objCommand.Parameters.AddWithValue("@Yapılanİş", CType(TextBox2.Text, String))
        objCommand.Parameters.AddWithValue("@Tehlike", CType(TextBox16.Text, String))
        objCommand.Parameters.AddWithValue("@Kategori", CType(ComboBox04.Text, String))
        objCommand.Parameters.AddWithValue("@Etkilenenler", CType(ComboBox05.Text, String))
        objCommand.Parameters.AddWithValue("@Risk", CType(TextBox5.Text, String))
        objCommand.Parameters.AddWithValue("@İhtimal", CType(ComboBox06.Text, Integer))
        objCommand.Parameters.AddWithValue("@Şiddet", CType(ComboBox07.Text, Integer))
        objCommand.Parameters.AddWithValue("@RiskDeğeri", CType(TextBox3.Text, Integer))
        objCommand.Parameters.AddWithValue("@ÖncelikDurumu", CType(TextBox4.Text, String))
        objCommand.Parameters.AddWithValue("@YasalZorunluluk", CType(TextBox17.Text, String))
        objCommand.Parameters.AddWithValue("@MevcutÖnlemler", CType(TextBox6.Text, String))
        objCommand.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
        objCommand.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox18.Text, String))
        objCommand.Parameters.AddWithValue("@Sorumlu", CType(TextBox51.Text, String))
        objCommand.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox52.Text, String))
        objCommand.Parameters.AddWithValue("@HedefTarih", CType(TextBox50.Text, String))
        objCommand.Parameters.AddWithValue("@İlgiliFotoğrafNo", CType(TextBox49.Text, String))
        objCommand.Parameters.AddWithValue("@Güncelİhtimal", CType(ComboBox1.Text, Integer))
        objCommand.Parameters.AddWithValue("@GüncelŞiddet", CType(ComboBox2.Text, Integer))
        objCommand.Parameters.AddWithValue("@GüncelRiskDeğeri", CType(TextBox46.Text, Integer))
        objCommand.Parameters.AddWithValue("@GüncelÖncelikDurumu", CType(TextBox47.Text, String))
        objCommand.Parameters.AddWithValue("@GüncelDurum", CType(TextBox048.Text, String))

        objCommand.Connection = objConnection
        aks()

        objCommand.CommandText = "UPDATE [Riskler] Set [Bölüm] = @Bölüm, [ProsesEkipmanAdı] = @ProsesEkipmanAdı, [Yapılanİş] = @Yapılanİş, [Tehlike] = @Tehlike, [Kategori] = @Kategori, [Etkilenenler] = @Etkilenenler, [Risk] = @Risk, [İhtimal] = @İhtimal, [Şiddet] = @Şiddet, [RiskDeğeri] = @RiskDeğeri, [ÖncelikDurumu] = @ÖncelikDurumu , [YasalZorunluluk] =@YasalZorunluluk, [MevcutÖnlemler] =@MevcutÖnlemler, [No] =@No, [AlınmasıGerekenÖnlemler] =@AlınmasıGerekenÖnlemler , [Sorumlu] =@Sorumlu, [GerçekleşmeTarihi] =@GerçekleşmeTarihi, [HedefTarih] =@HedefTarih, [İlgiliFotoğrafNo] =@ilgiliFotoğrafNo, [Güncelİhtimal] =@Güncelİhtimal , [GüncelŞiddet] =@GüncelŞiddet , [GüncelRiskDeğeri] =@GüncelRiskDeğeri, [GüncelÖncelikDurumu] =@GüncelÖncelikDurumu , [GüncelDurum] = @GüncelDurum  WHERE [No] = @No;"

        Try

            objConnection.Open()
            'Execute the SqlCommand object to insert the New data...
            objCommand.ExecuteNonQuery()

        Catch oledbExceptionErr As OleDbException
            MessageBox.Show(oledbExceptionErr.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Button4.Enabled = False
            objConnection.Close()
        End Try

        'Fill the dataset And bind the fields...
        FillDataSetAndView()
        BindFields()
        DataGridBind()
        objcurrencyManager.Position = int
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()
        TextBox9.DataBindings.Clear()
        TextBox10.DataBindings.Clear()
        TextBox11.DataBindings.Clear()
        TextBox12.DataBindings.Clear()
        TextBox13.DataBindings.Clear()
        TextBox14.DataBindings.Clear()
        TextBox15.DataBindings.Clear()
        TextBox16.DataBindings.Clear()
        TextBox17.DataBindings.Clear()
        TextBox49.DataBindings.Clear()
        ComboBox1.DataBindings.Clear()
        ComboBox2.DataBindings.Clear()
        ComboBox01.DataBindings.Clear()
        ComboBox02.DataBindings.Clear()
        ComboBox04.DataBindings.Clear()
        ComboBox05.DataBindings.Clear()
        ComboBox06.DataBindings.Clear()
        ComboBox07.DataBindings.Clear()
        TextBox18.DataBindings.Clear()
        TextBox19.DataBindings.Clear()
        TextBox20.DataBindings.Clear()
        TextBox21.DataBindings.Clear()
        TextBox22.DataBindings.Clear()
        TextBox23.DataBindings.Clear()
        TextBox24.DataBindings.Clear()
        TextBox25.DataBindings.Clear()
        TextBox26.DataBindings.Clear()
        TextBox27.DataBindings.Clear()
        TextBox28.DataBindings.Clear()
        TextBox29.DataBindings.Clear()
        TextBox30.DataBindings.Clear()
        TextBox31.DataBindings.Clear()
        TextBox32.DataBindings.Clear()
        TextBox33.DataBindings.Clear()
        TextBox34.DataBindings.Clear()
        TextBox35.DataBindings.Clear()
        TextBox36.DataBindings.Clear()
        TextBox37.DataBindings.Clear()
        TextBox38.DataBindings.Clear()
        TextBox39.DataBindings.Clear()
        TextBox40.DataBindings.Clear()
        TextBox41.DataBindings.Clear()
        TextBox42.DataBindings.Clear()
        TextBox43.DataBindings.Clear()
        TextBox44.DataBindings.Clear()
        TextBox45.DataBindings.Clear()
        TextBox46.DataBindings.Clear()
        TextBox47.DataBindings.Clear()
        TextBox048.DataBindings.Clear()
        TextBox49.DataBindings.Clear()
        TextBox50.DataBindings.Clear()
        TextBox51.DataBindings.Clear()
        TextBox52.DataBindings.Clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        TextBox12.Clear()
        TextBox13.Clear()
        TextBox14.Clear()
        TextBox15.Clear()
        TextBox16.Clear()
        TextBox17.Clear()
        ComboBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox01.Text = ""
        ComboBox02.Text = ""
        ComboBox04.Text = ""
        ComboBox05.Text = ""
        ComboBox06.Text = ""
        ComboBox07.Text = ""
        TextBox18.Clear()
        TextBox19.Clear()
        TextBox20.Clear()
        TextBox21.Clear()
        TextBox22.Clear()
        TextBox23.Clear()
        TextBox24.Clear()
        TextBox25.Clear()
        TextBox26.Clear()
        TextBox27.Clear()
        TextBox28.Clear()
        TextBox29.Clear()
        TextBox30.Clear()
        TextBox31.Clear()
        TextBox32.Clear()
        TextBox33.Clear()
        TextBox34.Clear()
        TextBox35.Clear()
        TextBox36.Clear()
        TextBox37.Clear()
        TextBox38.Clear()
        TextBox39.Clear()
        TextBox40.Clear()
        TextBox41.Clear()
        TextBox42.Clear()
        TextBox43.Clear()
        TextBox44.Clear()
        TextBox45.Clear()
        TextBox46.Clear()
        TextBox47.Clear()
        TextBox048.Clear()
        TextBox49.Clear()
        TextBox50.Clear()
        TextBox51.Clear()
        TextBox52.Clear()
        ComboBox01.SelectedItem = Nothing
        ComboBox02.SelectedItem = Nothing
        ComboBox04.SelectedItem = Nothing
        ComboBox06.SelectedItem = Nothing
        ComboBox07.SelectedItem = Nothing
        ComboBox1.SelectedItem = Nothing
        ComboBox2.SelectedItem = Nothing
        Button4.Enabled = True
        Button8.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim objCommand As OleDbCommand = New OleDbCommand()
        Dim objCommand1 As OleDbCommand = New OleDbCommand()
        Dim objCommand2 As OleDbCommand = New OleDbCommand()
        Dim objCommand3 As OleDbCommand = New OleDbCommand()
        Dim objCommand4 As OleDbCommand = New OleDbCommand()
        Dim objCommand5 As OleDbCommand = New OleDbCommand()
        Dim objCommand6 As OleDbCommand = New OleDbCommand()
        Dim objCommand7 As OleDbCommand = New OleDbCommand()
        Dim objCommand8 As OleDbCommand = New OleDbCommand()
        Dim objCommand9 As OleDbCommand = New OleDbCommand()
        Dim str = Guid.NewGuid.ToString()
        objCommand.Parameters.AddWithValue("@Bölüm", CType(ComboBox01.Text, String))
        objCommand.Parameters.AddWithValue("@ProsesEkipmanAdı", CType(ComboBox02.Text, String))
        objCommand.Parameters.AddWithValue("@No", str)
        objCommand.Parameters.AddWithValue("@Yapılanİş", CType(TextBox2.Text, String))
        objCommand.Parameters.AddWithValue("@Tehlike", CType(TextBox16.Text, String))
        objCommand.Parameters.AddWithValue("@Kategori", CType(ComboBox04.Text, String))
        objCommand.Parameters.AddWithValue("@Etkilenenler", CType(ComboBox05.Text, String))
        objCommand.Parameters.AddWithValue("@Risk", CType(TextBox5.Text, String))
        If ComboBox06.Text IsNot "" Then
            objCommand.Parameters.AddWithValue("@İhtimal", CType(ComboBox06.Text, Integer))
        Else objCommand.Parameters.AddWithValue("@İhtimal", 0)
        End If
        If ComboBox07.Text IsNot "" Then
            objCommand.Parameters.AddWithValue("@Şiddet", CType(ComboBox07.Text, Integer))
        Else objCommand.Parameters.AddWithValue("@Şiddet", 0)
        End If
        If TextBox3.Text IsNot "" Then
            objCommand.Parameters.AddWithValue("@RiskDeğeri", CType(TextBox3.Text, Integer))
        Else objCommand.Parameters.AddWithValue("@RiskDeğeri", 0)
        End If
        objCommand.Parameters.AddWithValue("@ÖncelikDurumu", CType(TextBox4.Text, String))
        objCommand.Parameters.AddWithValue("@YasalZorunluluk", CType(TextBox17.Text, String))
        objCommand.Parameters.AddWithValue("@MevcutÖnlemler", CType(TextBox6.Text, String))
        objCommand.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox18.Text, String))
        objCommand.Parameters.AddWithValue("@Sorumlu", CType(TextBox51.Text, String))
        objCommand.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox52.Text, String))
        objCommand.Parameters.AddWithValue("@HedefTarih", CType(TextBox50.Text, String))
        objCommand.Parameters.AddWithValue("@İlgiliFotoğrafNo", CType(TextBox49.Text, String))
        If ComboBox1.Text IsNot "" Then
            objCommand.Parameters.AddWithValue("@Güncelİhtimal", CType(ComboBox1.Text, Integer))
        Else objCommand.Parameters.AddWithValue("@Güncelİhtimal", 0)
        End If

        If ComboBox2.Text IsNot "" Then
            objCommand.Parameters.AddWithValue("@GüncelŞiddet", CType(ComboBox2.Text, Integer))
        Else objCommand.Parameters.AddWithValue("@GüncelŞiddet", 0)
        End If

        If TextBox46.Text IsNot "" Then
            objCommand.Parameters.AddWithValue("@GüncelRiskDeğeri", CType(TextBox46.Text, Integer))
        Else objCommand.Parameters.AddWithValue("@GüncelRiskDeğeri", 0)
        End If

        objCommand.Parameters.AddWithValue("@GüncelÖncelikDurumu", CType(TextBox47.Text, String))
        objCommand.Parameters.AddWithValue("@GüncelDurum", CType(TextBox048.Text, String))
        objCommand.Connection = objConnection

        objCommand.CommandText = "INSERT INTO Riskler " & "([Bölüm], [ProsesEkipmanAdı], [No], [Yapılanİş], [Tehlike], [Kategori], [Etkilenenler], [Risk], [İhtimal], [Şiddet], [RiskDeğeri], [ÖncelikDurumu], [YasalZorunluluk], [MevcutÖnlemler], [AlınmasıGerekenÖnlemler], [Sorumlu], [GerçekleşmeTarihi], [HedefTarih], [İlgiliFotoğrafNo], [Güncelİhtimal], [GüncelŞiddet], [GüncelRiskDeğeri], [GüncelÖncelikDurumu], [GüncelDurum])" & "VALUES(@Bölüm,@ProsesEkipmanAdı,@No,@Yapılanİş,@Tehlike,@Kategori,@Etkilenenler,@Risk,@İhtimal,@Şiddet,@RiskDeğeri,@ÖncelikDurumu,@YasalZorunluluk,@MevcutÖnlemmler,@AlınmasıGerekenÖnlemler,@Sorumlu,@GerçekleşmeTarihi,@HedefTarih,@İlgiliFotoğrafNo,@Güncelİhtimal,@GüncelŞiddet,@GüncelRiskDeğeri,@GüncelÖncelikDurumu, @GüncelDurum);"

        If TextBox7.Text IsNot "" Then
            objCommand1.Parameters.AddWithValue("@No", str)
            objCommand1.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox7.Text, String))
            objCommand1.Parameters.AddWithValue("@Sorumlu", CType(TextBox28.Text, String))
            If TextBox19.Text IsNot "" Then objCommand1.Parameters.AddWithValue("@HedefTarih", CType(TextBox19.Text, Date))
            If TextBox45.Text IsNot "" Then objCommand1.Parameters.AddWithValue("@GerçekleşmeTarihi", CType((TextBox45.Text), Date))

            objCommand1.Connection = objConnection
            If TextBox45.Text Is "" Then objCommand1.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox45.Text Is "" And TextBox19.Text Is "" Then objCommand1.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox45.Text IsNot "" And TextBox19.Text IsNot "" Then objCommand1.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox8.Text IsNot "" Then
            objCommand2.Parameters.AddWithValue("@No", str)
            objCommand2.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox8.Text, String))
            objCommand2.Parameters.AddWithValue("@Sorumlu", CType(TextBox29.Text, String))
            If TextBox20.Text IsNot "" Then objCommand2.Parameters.AddWithValue("@HedefTarih", CType(TextBox20.Text, Date))
            If TextBox44.Text IsNot "" Then objCommand2.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox44.Text, Date))

            objCommand2.Connection = objConnection
            If TextBox44.Text Is "" Then objCommand2.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox44.Text Is "" And TextBox20.Text Is "" Then objCommand2.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox44.Text IsNot "" And TextBox20.Text IsNot "" Then objCommand2.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If

        If TextBox9.Text IsNot "" Then
            objCommand3.Parameters.AddWithValue("@No", str)
            objCommand3.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox9.Text, String))
            objCommand3.Parameters.AddWithValue("@Sorumlu", CType(TextBox30.Text, String))
            If TextBox21.Text IsNot "" Then objCommand3.Parameters.AddWithValue("@HedefTarih", CType(TextBox21.Text, Date))
            If TextBox43.Text IsNot "" Then objCommand3.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox43.Text, Date))

            objCommand3.Connection = objConnection
            If TextBox43.Text Is "" Then objCommand3.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox43.Text Is "" And TextBox21.Text Is "" Then objCommand3.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox43.Text IsNot "" And TextBox21.Text IsNot "" Then objCommand3.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If

        If TextBox10.Text IsNot "" Then
            objCommand4.Parameters.AddWithValue("@No", str)
            objCommand4.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox10.Text, String))
            objCommand4.Parameters.AddWithValue("@Sorumlu", CType(TextBox31.Text, String))
            If TextBox22.Text IsNot "" Then objCommand4.Parameters.AddWithValue("@HedefTarih", CType(TextBox22.Text, Date))
            If TextBox42.Text IsNot "" Then objCommand4.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox42.Text, Date))

            objCommand4.Connection = objConnection
            If TextBox42.Text Is "" Then objCommand4.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox42.Text Is "" And TextBox22.Text Is "" Then objCommand4.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox42.Text IsNot "" And TextBox22.Text IsNot "" Then objCommand4.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox11.Text IsNot "" Then
            objCommand5.Parameters.AddWithValue("@No", str)
            objCommand5.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox11.Text, String))
            objCommand5.Parameters.AddWithValue("@Sorumlu", CType(TextBox32.Text, String))
            If TextBox23.Text IsNot "" Then objCommand5.Parameters.AddWithValue("@HedefTarih", CType(TextBox23.Text, Date))
            If TextBox41.Text IsNot "" Then objCommand5.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox41.Text, Date))

            objCommand5.Connection = objConnection
            If TextBox41.Text Is "" Then objCommand5.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox41.Text Is "" And TextBox23.Text Is "" Then objCommand5.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox41.Text IsNot "" And TextBox23.Text IsNot "" Then objCommand5.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox12.Text IsNot "" Then
            objCommand6.Parameters.AddWithValue("@No", str)
            objCommand6.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox12.Text, String))
            objCommand6.Parameters.AddWithValue("@Sorumlu", CType(TextBox33.Text, String))
            If TextBox24.Text IsNot "" Then objCommand6.Parameters.AddWithValue("@HedefTarih", CType(TextBox24.Text, Date))
            If TextBox40.Text IsNot "" Then objCommand6.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox40.Text, Date))

            objCommand6.Connection = objConnection
            If TextBox40.Text Is "" Then objCommand6.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox40.Text Is "" And TextBox24.Text Is "" Then objCommand6.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox40.Text IsNot "" And TextBox24.Text IsNot "" Then objCommand6.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox13.Text IsNot "" Then
            objCommand7.Parameters.AddWithValue("@No", str)
            objCommand7.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox13.Text, String))
            objCommand7.Parameters.AddWithValue("@Sorumlu", CType(TextBox34.Text, String))
            If TextBox25.Text IsNot "" Then objCommand7.Parameters.AddWithValue("@HedefTarih", CType(TextBox25.Text, Date))
            If TextBox39.Text IsNot "" Then objCommand7.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox39.Text, Date))

            objCommand7.Connection = objConnection
            If TextBox39.Text Is "" Then objCommand7.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox39.Text Is "" And TextBox25.Text Is "" Then objCommand7.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox39.Text IsNot "" And TextBox25.Text IsNot "" Then objCommand7.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox14.Text IsNot "" Then
            objCommand8.Parameters.AddWithValue("@No", str)
            objCommand8.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox14.Text, String))
            objCommand8.Parameters.AddWithValue("@Sorumlu", CType(TextBox35.Text, String))
            If TextBox26.Text IsNot "" Then objCommand8.Parameters.AddWithValue("@HedefTarih", CType(TextBox26.Text, Date))
            If TextBox38.Text IsNot "" Then objCommand8.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox38.Text, Date))

            objCommand8.Connection = objConnection
            If TextBox38.Text Is "" Then objCommand8.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox38.Text Is "" And TextBox26.Text Is "" Then objCommand8.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox38.Text IsNot "" And TextBox26.Text IsNot "" Then objCommand8.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox15.Text IsNot "" Then
            objCommand9.Parameters.AddWithValue("@No", str)
            objCommand9.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox15.Text, String))
            objCommand9.Parameters.AddWithValue("@Sorumlu", CType(TextBox36.Text, String))
            If TextBox27.Text IsNot "" Then objCommand9.Parameters.AddWithValue("@HedefTarih", CType(TextBox27.Text, Date))
            If TextBox37.Text IsNot "" Then objCommand9.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox37.Text, Date))

            objCommand9.Connection = objConnection
            If TextBox37.Text Is "" Then objCommand9.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox37.Text Is "" And TextBox27.Text Is "" Then objCommand9.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox37.Text IsNot "" And TextBox27.Text IsNot "" Then objCommand9.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        Try

            objConnection.Open()
            objCommand.ExecuteNonQuery()
            If TextBox7.Text IsNot "" Then objCommand1.ExecuteNonQuery()
            If TextBox8.Text IsNot "" Then objCommand2.ExecuteNonQuery()
            If TextBox9.Text IsNot "" Then objCommand3.ExecuteNonQuery()
            If TextBox10.Text IsNot "" Then objCommand4.ExecuteNonQuery()
            If TextBox11.Text IsNot "" Then objCommand5.ExecuteNonQuery()
            If TextBox12.Text IsNot "" Then objCommand6.ExecuteNonQuery()
            If TextBox13.Text IsNot "" Then objCommand7.ExecuteNonQuery()
            If TextBox14.Text IsNot "" Then objCommand8.ExecuteNonQuery()
            If TextBox15.Text IsNot "" Then objCommand9.ExecuteNonQuery()
        Catch oledbExceptionErr As OleDbException
            MessageBox.Show(oledbExceptionErr.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            Button4.Enabled = False
            objConnection.Close()
        End Try
        lst = objcurrencyManager.Count + 1
        objcurrencyManager.Position = lst
        'Fill the dataset And bind the fields...
        FillDataSetAndView()
        BindFields()
        DataGridBind()
        objcurrencyManager.Position = lst
    End Sub
    Sub aks()
        Dim objCommand1 As OleDbCommand = New OleDbCommand()
        Dim objCommand2 As OleDbCommand = New OleDbCommand()
        Dim objCommand3 As OleDbCommand = New OleDbCommand()
        Dim objCommand4 As OleDbCommand = New OleDbCommand()
        Dim objCommand5 As OleDbCommand = New OleDbCommand()
        Dim objCommand6 As OleDbCommand = New OleDbCommand()
        Dim objCommand7 As OleDbCommand = New OleDbCommand()
        Dim objCommand8 As OleDbCommand = New OleDbCommand()
        Dim objCommand9 As OleDbCommand = New OleDbCommand()



        If TextBox7.Text IsNot "" Then
            objCommand1.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand1.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox7.Text, String))
            objCommand1.Parameters.AddWithValue("@Sorumlu", CType(TextBox28.Text, String))
            If TextBox19.Text IsNot "" Then objCommand1.Parameters.AddWithValue("@HedefTarih", CType(TextBox19.Text, Date))
            If TextBox45.Text IsNot "" Then objCommand1.Parameters.AddWithValue("@GerçekleşmeTarihi", CType((TextBox45.Text), Date))

            objCommand1.Connection = objConnection
            If TextBox45.Text Is "" Then objCommand1.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox45.Text Is "" And TextBox19.Text Is "" Then objCommand1.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox45.Text IsNot "" And TextBox19.Text IsNot "" Then objCommand1.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If

        If TextBox8.Text IsNot "" Then
            objCommand2.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand2.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox8.Text, String))
            objCommand2.Parameters.AddWithValue("@Sorumlu", CType(TextBox29.Text, String))
            If TextBox20.Text IsNot "" Then objCommand2.Parameters.AddWithValue("@HedefTarih", CType(TextBox20.Text, Date))
            If TextBox44.Text IsNot "" Then objCommand2.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox44.Text, Date))

            objCommand2.Connection = objConnection
            If TextBox44.Text Is "" Then objCommand2.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox44.Text Is "" And TextBox20.Text Is "" Then objCommand2.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox44.Text IsNot "" And TextBox20.Text IsNot "" Then objCommand2.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If

        If TextBox9.Text IsNot "" Then
            objCommand3.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand3.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox9.Text, String))
            objCommand3.Parameters.AddWithValue("@Sorumlu", CType(TextBox30.Text, String))
            If TextBox21.Text IsNot "" Then objCommand3.Parameters.AddWithValue("@HedefTarih", CType(TextBox21.Text, Date))
            If TextBox43.Text IsNot "" Then objCommand3.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox43.Text, Date))

            objCommand3.Connection = objConnection
            If TextBox43.Text Is "" Then objCommand3.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox43.Text Is "" And TextBox21.Text Is "" Then objCommand3.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox43.Text IsNot "" And TextBox21.Text IsNot "" Then objCommand3.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If

        If TextBox10.Text IsNot "" Then
            objCommand4.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand4.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox10.Text, String))
            objCommand4.Parameters.AddWithValue("@Sorumlu", CType(TextBox31.Text, String))
            If TextBox22.Text IsNot "" Then objCommand4.Parameters.AddWithValue("@HedefTarih", CType(TextBox22.Text, Date))
            If TextBox42.Text IsNot "" Then objCommand4.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox42.Text, Date))

            objCommand4.Connection = objConnection
            If TextBox42.Text Is "" Then objCommand4.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox42.Text Is "" And TextBox22.Text Is "" Then objCommand4.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox42.Text IsNot "" And TextBox22.Text IsNot "" Then objCommand4.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox11.Text IsNot "" Then
            objCommand5.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand5.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox11.Text, String))
            objCommand5.Parameters.AddWithValue("@Sorumlu", CType(TextBox32.Text, String))
            If TextBox23.Text IsNot "" Then objCommand5.Parameters.AddWithValue("@HedefTarih", CType(TextBox23.Text, Date))
            If TextBox41.Text IsNot "" Then objCommand5.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox41.Text, Date))

            objCommand5.Connection = objConnection
            If TextBox41.Text Is "" Then objCommand5.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox41.Text Is "" And TextBox23.Text Is "" Then objCommand5.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox41.Text IsNot "" And TextBox23.Text IsNot "" Then objCommand5.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox12.Text IsNot "" Then
            objCommand6.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand6.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox12.Text, String))
            objCommand6.Parameters.AddWithValue("@Sorumlu", CType(TextBox33.Text, String))
            If TextBox24.Text IsNot "" Then objCommand6.Parameters.AddWithValue("@HedefTarih", CType(TextBox24.Text, Date))
            If TextBox40.Text IsNot "" Then objCommand6.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox40.Text, Date))

            objCommand6.Connection = objConnection
            If TextBox40.Text Is "" Then objCommand6.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox40.Text Is "" And TextBox24.Text Is "" Then objCommand6.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox40.Text IsNot "" And TextBox24.Text IsNot "" Then objCommand6.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox13.Text IsNot "" Then
            objCommand7.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand7.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox13.Text, String))
            objCommand7.Parameters.AddWithValue("@Sorumlu", CType(TextBox34.Text, String))
            If TextBox25.Text IsNot "" Then objCommand7.Parameters.AddWithValue("@HedefTarih", CType(TextBox25.Text, Date))
            If TextBox39.Text IsNot "" Then objCommand7.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox39.Text, Date))

            objCommand7.Connection = objConnection
            If TextBox39.Text Is "" Then objCommand7.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox39.Text Is "" And TextBox25.Text Is "" Then objCommand7.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox39.Text IsNot "" And TextBox25.Text IsNot "" Then objCommand7.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox14.Text IsNot "" Then
            objCommand8.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand8.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox14.Text, String))
            objCommand8.Parameters.AddWithValue("@Sorumlu", CType(TextBox35.Text, String))
            If TextBox26.Text IsNot "" Then objCommand8.Parameters.AddWithValue("@HedefTarih", CType(TextBox26.Text, Date))
            If TextBox38.Text IsNot "" Then objCommand8.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox38.Text, Date))

            objCommand8.Connection = objConnection
            If TextBox38.Text Is "" Then objCommand8.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox38.Text Is "" And TextBox26.Text Is "" Then objCommand8.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox38.Text IsNot "" And TextBox26.Text IsNot "" Then objCommand8.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        If TextBox15.Text IsNot "" Then
            objCommand9.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))
            objCommand9.Parameters.AddWithValue("@AlınmasıGerekenÖnlemler", CType(TextBox15.Text, String))
            objCommand9.Parameters.AddWithValue("@Sorumlu", CType(TextBox36.Text, String))
            If TextBox27.Text IsNot "" Then objCommand9.Parameters.AddWithValue("@HedefTarih", CType(TextBox27.Text, Date))
            If TextBox37.Text IsNot "" Then objCommand9.Parameters.AddWithValue("@GerçekleşmeTarihi", CType(TextBox37.Text, Date))

            objCommand9.Connection = objConnection
            If TextBox37.Text Is "" Then objCommand9.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih);"
            If TextBox37.Text Is "" And TextBox27.Text Is "" Then objCommand9.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu);"
            If TextBox37.Text IsNot "" And TextBox27.Text IsNot "" Then objCommand9.CommandText = "INSERT INTO AksiyonPlanı" & "([No],[AlınmasıGerekenÖnlemler],[Sorumlu],[HedefTarih],[GerçekleşmeTarihi])" & "VALUES(@No,@AlınmasıGerekenÖnlemler,@Sorumlu,@HedefTarih,@GerçekleşmeTarihi);"
        End If
        Try

            objConnection.Open()

            If TextBox7.Text IsNot "" Then objCommand1.ExecuteNonQuery()
            If TextBox8.Text IsNot "" Then objCommand2.ExecuteNonQuery()
            If TextBox9.Text IsNot "" Then objCommand3.ExecuteNonQuery()
            If TextBox10.Text IsNot "" Then objCommand4.ExecuteNonQuery()
            If TextBox11.Text IsNot "" Then objCommand5.ExecuteNonQuery()
            If TextBox12.Text IsNot "" Then objCommand6.ExecuteNonQuery()
            If TextBox13.Text IsNot "" Then objCommand7.ExecuteNonQuery()
            If TextBox14.Text IsNot "" Then objCommand8.ExecuteNonQuery()
            If TextBox15.Text IsNot "" Then objCommand9.ExecuteNonQuery()
        Catch oledbExceptionErr As OleDbException
            MessageBox.Show(oledbExceptionErr.Message, "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally

            objConnection.Close()
        End Try

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        int = objcurrencyManager.Position
        Dim objCommand As OleDbCommand = New OleDbCommand()

        objCommand.Connection = objConnection
        objCommand.Parameters.AddWithValue("@No", CType(TextBox1.Text, String))

        objCommand.CommandText = "DELETE FROM Riskler " & "WHERE [No] = @No"

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

    Private Sub BölümSorgulamaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BölümSorgulamaToolStripMenuItem.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub KategoriSorgulamaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KategoriSorgulamaToolStripMenuItem.Click
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub AksiyonSorgulamaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AksiyonSorgulamaToolStripMenuItem.Click
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub RaporlamaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RaporlamaToolStripMenuItem.Click
        Form7.Show()
        Me.Hide()
    End Sub

    Private Sub YetkiliKişilerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YetkiliKişilerToolStripMenuItem.Click
        Form6.Show()
        Me.Hide()
    End Sub


    Private Sub bölümara()
        Dim adtr As New OleDbDataAdapter("selecet * from Riskler  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Riskler" & " where bölüm Like '%" & ComboBox01.Text & "%' ")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Riskler")

    End Sub
    Private Sub makineara()
        Dim adtr As New OleDbDataAdapter("selecet * from Riskler ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Riskler" & " where prosesekipmanadı like '%" & ComboBox02.Text & "%' ")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Riskler")
    End Sub
    Private Sub kategoriara()
        Dim adtr As New OleDbDataAdapter("selecet * from Riskler  ", newbagla)
        adtr.SelectCommand.CommandText = (" Select * From Riskler" & " where kategori like '%" & ComboBox04.Text & "%' ")
        objDataSet.Clear()
        adtr.Fill(objDataSet, "Riskler")
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        If ComboBox01.SelectedItem IsNot Nothing Then bölümara()
        If ComboBox02.SelectedItem IsNot Nothing Then makineara()
        If ComboBox04.SelectedItem IsNot Nothing Then kategoriara()


        Button4.Enabled = False
        Button8.Enabled = False
        BindFields()


    End Sub


    Private Sub TerminiGecikmişİşlerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TerminiGecikmişİşlerToolStripMenuItem.Click
        Form8.Show()
        Me.Hide()
    End Sub

    Private Sub SürekliİşlerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SürekliİşlerToolStripMenuItem.Click
        Form9.Show()
        Me.Hide()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FillDataSetAndView()
        DataGridBind()
        ComboBox01.Items.Clear()
        ComboBox02.Items.Clear()
        ComboBox04.Items.Clear()
        ComboBox05.Items.Clear()
        Combox01()
        Combox02()
        Combox04()
        Combox05()
        BindFields()

        Button8.Enabled = False
        Button4.Enabled = False

    End Sub


    Private Sub Form10_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form1.Show()
        Me.Hide()
    End Sub

End Class