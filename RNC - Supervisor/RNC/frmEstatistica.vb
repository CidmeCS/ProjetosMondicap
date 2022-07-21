Imports System.Data.OleDb
Public Class frmEstatistica
    Dim conRNC As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb;Jet OLEDB:Database Password= projetornc;")
    Private Sub frmEstatistica_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim da1 As New OleDbDataAdapter
            Dim ds1 As New DataSet
            Dim da2 As New OleDbDataAdapter
            Dim ds2 As New DataSet
            Dim da3 As New OleDbDataAdapter
            Dim ds3 As New DataSet

            conRNC.Open()
            Dim sel As String = "select RNC, count (RNC) from tblRNC group by RNC " 'conta quantas RNCs exitem
            da1 = New OleDbDataAdapter(sel, conRNC)
            ds1.Clear()
            da1.Fill(ds1, "tblRNC")
            conRNC.Close()
            Me.DataGridView1.DataSource = ds1
            Me.DataGridView1.DataMember = "tblRNC"
            DataGridView1.Columns(0).HeaderText = "RNC"
            DataGridView1.Columns(1).HeaderText = "Qt por RNC´s"
            Label1.Text = DataGridView1.RowCount 'conta quantas RNCs exitem

            conRNC.Open()
            Dim sel2 As String = "select Status, count (Status) from tblRNC Where Status = 'Pendente' group by Status " 'conta quantas RNCs exitem
            da2 = New OleDbDataAdapter(sel2, conRNC)
            ds2.Clear()
            da2.Fill(ds2, "tblRNC")
            conRNC.Close()
            Me.DataGridView2.DataSource = ds2
            Me.DataGridView2.DataMember = "tblRNC"
            DataGridView2.Columns(0).HeaderText = "Status"
            DataGridView2.Columns(1).HeaderText = "Qt de Pendentes"


            conRNC.Open()
            Dim sel3 As String = "select Status, count (Status) from tblRNC Where Status = 'Fechada' group by Status " 'conta quantas RNCs exitem
            da3 = New OleDbDataAdapter(sel3, conRNC)
            ds3.Clear()
            da3.Fill(ds3, "tblRNC")
            conRNC.Close()
            Me.DataGridView3.DataSource = ds3
            Me.DataGridView3.DataMember = "tblRNC"
            DataGridView3.Columns(0).HeaderText = "Status"
            DataGridView3.Columns(1).HeaderText = "Qt de Fechada"

        Catch ex As Exception
            Beep()
            MsgBox("Erro 1 " & ex.Message)
        End Try
    End Sub
End Class