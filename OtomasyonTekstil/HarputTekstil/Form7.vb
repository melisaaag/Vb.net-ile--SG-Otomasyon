Imports System.Data
Imports System.Data.OleDb
Imports mexcel = Microsoft.Office.Interop.Excel

Public Class Form7
    Dim TabloAl As DataSet
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

        Dim firma As String = Form1.ComboBox1.Text

        Dim newbagla As String = "provider=Microsoft.ACE.OLEDB.12.0;" &
                          "Data Source=" + path + "\HarputTekstilBST\" + firma + ".accdb;"
        Dim bag As New OleDbConnection(newbagla)

        bag.Open()
        Dim yenibaglanti As New OleDbConnection(bag.ConnectionString)
        Dim sorgu As String = "Select * from Riskler order by prosesekipmanadı"
        Dim verial As New OleDbDataAdapter(sorgu, yenibaglanti)
        TabloAl = New DataSet()
        verial.Fill(TabloAl, "Riskler")

        DataGridView1.DataSource = TabloAl.Tables(0)
        bag.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try

            Dim exceluygulama As New mexcel.Application()
            Dim excelkitap As mexcel.Workbook
            excelkitap = exceluygulama.Workbooks.Add(True)
            Dim excelsayfa As mexcel.Worksheet
            excelsayfa = CType(exceluygulama.ActiveSheet, mexcel.Worksheet)
            exceluygulama.Visible = True

            Dim sutunnumarasi As Integer = 0
            Dim satirnumarasi As Integer = 1

            For Each sutun As DataColumn In TabloAl.Tables(0).Columns

                sutunnumarasi += 1
                excelsayfa.Cells(1, sutunnumarasi).Font.Bold = True
                excelsayfa.Cells(1, sutunnumarasi) = sutun.ColumnName

            Next

            For Each satir As DataRow In TabloAl.Tables(0).Rows
Melisa:
                satirnumarasi += 1
                sutunnumarasi = 0
                For Each sutun As DataColumn In TabloAl.Tables(0).Columns

                    sutunnumarasi += 1
                    excelsayfa.Cells(satirnumarasi, sutunnumarasi) = satir(sutun.ColumnName).ToString()
                Next
                If satirnumarasi >= 3 Then
                    Dim isim As String = excelsayfa.Cells(satirnumarasi - 1, 2).value
                    If isim.Length > 31 Then isim = Microsoft.VisualBasic.Left(isim, 29) & Microsoft.VisualBasic.Right(isim, 2)

                    excelsayfa.Name = isim

                    If excelsayfa.Cells(satirnumarasi, 2).value <> excelsayfa.Cells(satirnumarasi - 1, 2).value Then
                        excelsayfa.Rows(satirnumarasi).Delete
                        excelsayfa = excelkitap.Worksheets.Add()
                        satirnumarasi = 1
                        sutunnumarasi = 0
                        For Each sutun As DataColumn In TabloAl.Tables(0).Columns
                            sutunnumarasi += 1
                            excelsayfa.Cells(1, sutunnumarasi).Font.Bold = True
                            excelsayfa.Cells(1, sutunnumarasi) = sutun.ColumnName
                        Next
                        GoTo Melisa
                    End If
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form10.Show()
        Me.Close()
    End Sub

    Private Sub Form7_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form10.Show()
        Me.Hide()
    End Sub
End Class