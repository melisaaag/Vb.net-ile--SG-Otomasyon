Imports System.Data.OleDb
Public Class Form5
    Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Dim firma As String = Form1.ComboBox1.Text

    Dim newbagla As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
                             "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;"
    Dim objDataSet As DataSet
    Dim objDataView As DataView
    Private Sub FillDataSetAndView()

        Dim objDataAdapter As New OleDbDataAdapter("Select * FROM Aks_Çapraz", newbagla)
        objDataSet = New DataSet()
        objDataAdapter.Fill(objDataSet, "Aks_Çapraz")
        objDataView = New DataView(objDataSet.Tables("Aks_Çapraz"))

    End Sub
    Private Sub DataGridBind()
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = objDataView
    End Sub
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FillDataSetAndView()
        DataGridBind()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Form10.Show()
        Me.Close()
    End Sub

    Private Sub Form5_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form10.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        FillDataSetAndView()
        DataGridBind()
    End Sub
End Class