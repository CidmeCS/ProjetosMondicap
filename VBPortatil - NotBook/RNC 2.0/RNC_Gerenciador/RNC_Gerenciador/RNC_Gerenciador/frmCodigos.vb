Imports System.Data.OleDb
Public Class frmCodigos
    Private Sub frmCodigos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conDefeito As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Receb.Mat.Prima\Banco_Dados\RNC_Defeito.accdb;Jet OLEDB:Database Password= projetornc;")
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conDefeito.Open()
            Dim sel As String = "Select * from tblDefeitos order by Nao_conformidade asc"
            da = New OleDbDataAdapter(sel, conDefeito)
            ds.Clear()
            da.Fill(ds, "tblDefeitos")
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblDefeitos"
            conDefeito.Close()
        Catch ex As Exception
            Beep()
            MsgBox("Erro 1 frmDefeitos " & ex.Message)
        End Try
    End Sub
End Class