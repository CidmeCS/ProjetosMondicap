Imports System.Data.OleDb
Imports RNC_Supervisor.Module1
Public Class frmCodigoConsulta
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

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Try

            Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

            Dim Cod = row.Cells(0)
            Dim RNC = row.Cells(1)

            Me.Label1.Text = Cod.Value
            Me.Label2.Text = RNC.Value

        Catch ex As Exception
            MsgBox("Erro 70 " & ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        MCod1 = Label1.Text
        MRNC1 = Label2.Text

        MCod2 = Label1.Text
        MRNC2 = Label2.Text

        MCod3 = Label1.Text
        MRNC3 = Label2.Text

        MCod4 = Label1.Text
        MRNC4 = Label2.Text

        MCod5 = Label1.Text
        MRNC5 = Label2.Text

        MCod6 = Label1.Text
        MRNC6 = Label2.Text

        MCod7 = Label1.Text
        MRNC7 = Label2.Text

        MCod8 = Label1.Text
        MRNC8 = Label2.Text

        MCod9 = Label1.Text
        MRNC9 = Label2.Text

        MCod10 = Label1.Text
        MRNC10 = Label2.Text

        Close()
    End Sub
End Class