Imports System.Data.OleDb
Imports RNC.Module1
Public Class frmMaquina
    Private Sub frmCodigos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conMaquina As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb;Jet OLEDB:Database Password= projetornc;")
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conMaquina.Open()
            Dim sel As String = "Select * from tblMaquina order by Maquina asc"
            da = New OleDbDataAdapter(sel, conMaquina)
            ds.Clear()
            da.Fill(ds, "tblMaquina")
            conMaquina.Close()
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblMaquina"
        Catch ex As Exception
            conMaquina.Close()
            Beep()
            MsgBox("Erro 1 frmMaquina " & ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try

            Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

            Dim Maquina = row.Cells(0)
            Dim Celula = row.Cells(1)

            Me.Label1.Text = Maquina.Value
            Me.Label2.Text = Celula.Value

        Catch ex As Exception
            MsgBox("Erro 70 " & ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        txtMMaquina = Label1.Text
        Close()
    End Sub

    
End Class