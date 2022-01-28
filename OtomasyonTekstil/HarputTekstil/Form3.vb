Imports System.Data.OleDb
Public Class Form3
    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Dim firma As String = Form1.ComboBox1.Text

    Dim newbagla As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
                           "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;"
    Dim objDataSet As DataSet
    Dim objDataView As DataView
    Private Sub FillDataSetAndView()

        Dim objDataAdapter As New OleDbDataAdapter("Select * FROM Bölüm_Çapraz", newbagla)
        objDataSet = New DataSet()
        objDataAdapter.Fill(objDataSet, "Bölüm_Çapraz")
        objDataView = New DataView(objDataSet.Tables("Bölüm_Çapraz"))

    End Sub
    Private Sub DataGridBind()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView
    End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillDataSetAndView()
        DataGridBind()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form10.Show()
        Me.Close()
    End Sub

    Private Sub Form3_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Me.Hide()
        Form10.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataSetAndView()
        DataGridBind()
    End Sub
End Class