Imports System.Data.OleDb
Public Class frmDefeito
    Dim conDefeito As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Defeito.accdb;Jet OLEDB:Database Password= projetornc;")
    Private Sub frmDefeito_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        TesteAbertoDefeito()
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conDefeito.Open()
            Dim sel As String = "Select * from tblDefeitos order by Nao_Conformidade asc"
            da = New OleDbDataAdapter(sel, conDefeito)
            ds.Clear()
            da.Fill(ds, "tblDefeitos")
            conDefeito.Close()
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblDefeitos"


            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(0).HeaderText = "Código"
            DataGridView1.Columns(1).HeaderText = "Não Conformidade"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 50
            'DataGridView1.Columns(1).Width = 350
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro D1D " & ex.Message)
            conDefeito.Close()
        End Try

    End Sub
    Sub Atualizar()
        TesteAbertoDefeito()
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conDefeito.Open()
            Dim sel As String = "Select * from tblDefeitos order by Nao_Conformidade asc"
            da = New OleDbDataAdapter(sel, conDefeito)
            ds.Clear()
            da.Fill(ds, "tblDefeitos")
            conDefeito.Close()
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblDefeitos"


            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(1).HeaderText = "Não Conformidade"
            DataGridView1.Columns(0).HeaderText = "Código"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(1).Width = 50
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro D1f " & ex.Message)
            conDefeito.Close()
        End Try

    End Sub
    Sub TesteAbertoDefeito()
        Dim RNC_Defeito As Boolean
        RNC_Defeito = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Defeito.accdb")
        If RNC_Defeito = True Then
            Dim RNCDefeito As Integer = 0
            For RNCDefeito = 5 To 20
                RNC_Defeito = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Defeito.accdb")
                If RNC_Defeito = True Then
                    RNCDefeito = 5
                    If (MsgBox("O Arquivo 'RNC_Defeito.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_Defeito.accdb")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNC_Defeito = False Then
                    RNCDefeito = 20
                End If
            Next
        End If
    End Sub
    Function Test(ByVal pathfile As String) As Boolean
        Dim ff As Integer
        If System.IO.File.Exists(pathfile) Then
            Try
                ff = FreeFile()
                Microsoft.VisualBasic.FileOpen(ff, pathfile, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
                Return False
            Catch
                Return True
            Finally
                FileClose(ff)
            End Try
            Return True
        Else
        End If
        Return True
    End Function

    Private Sub btInserir_Click(sender As System.Object, e As System.EventArgs) Handles btInserir.Click


        TesteAbertoDefeito()
        If btInserir.Text = "Inserir" Then
            If MsgBox("Deseja Incluir um novo Defeito?", vbYesNo, "Novo Defeito") = vbYes Then
                btInserir.Text = "Aplicar"
                btAlterar.Enabled = False
                btExcluir.Enabled = False
                DataGridView1.Enabled = False
                txtCodigoRNC1.Clear()
                txtCodigoRNC1.Enabled = False
                txtDescricaoRNC1.Focus()
            Else

            End If

        Else
            If txtDescricaoRNC1.TextLength = 0 Then
                MsgBox("Insira uma Não Conformidade", , "Não Conformidade")
            Else

                Try
                    TesteAbertoDefeito()
                    conDefeito.Open()
                    Dim da4 As New OleDbDataAdapter
                    Dim ds4 As New DataSet
                    ds4 = New DataSet
                    da4 = New OleDbDataAdapter("INSERT INTO tblDefeitos (Nao_Conformidade) Values ('" & txtDescricaoRNC1.Text & "') ", conDefeito)
                    ds4.Clear()
                    da4.Fill(ds4, "tblDefeito")
                    conDefeito.Close()
                    MsgBox("Registro Inserido com sucesso!")
                    Atualizar()
                    btInserir.Text = "Inserir"
                    btAlterar.Enabled = True
                    btExcluir.Enabled = True
                    DataGridView1.Enabled = True
                    txtDescricaoRNC1.Clear()
                    txtCodigoRNC1.Enabled = True
                    btInserir.Focus()

                Catch ex As Exception
                    MsgBox("Erro D10 " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub btExcluir_Click(sender As System.Object, e As System.EventArgs) Handles btExcluir.Click
        Try
            TesteAbertoDefeito()
            Dim da21 As New OleDbDataAdapter
            Dim ds21 As New DataSet
            If btExcluir.Text = "Excluir" Then
                If MsgBox("Deseja Excluir um Código?", vbYesNo, "Excluir Código") = vbYes Then
                    txtCodigoRNC1.Clear()
                    txtCodigoRNC1.Focus()
                    btExcluir.Text = "Aplicar"
                    btInserir.Enabled = False
                    btAlterar.Enabled = False
                Else
                End If
            Else
                If txtCodigoRNC1.TextLength = 0 Then
                    MsgBox("Selecione um Código na tabela abaixo", , "Selecione um Código")
                Else
                    TesteAbertoDefeito()
                    conDefeito.Open()
                    ds21 = New DataSet
                    da21 = New OleDbDataAdapter("delete from tblDefeitos where Codigo = " & txtCodigoRNC1.Text & " ", conDefeito)
                    ds21.Clear()
                    da21.Fill(ds21, "tblDefeitos")
                    conDefeito.Close()
                    MsgBox("Registro Excluido com sucesso!")
                    Atualizar()
                    txtCodigoRNC1.Clear()
                    btExcluir.Focus()
                    btExcluir.Text = "Excluir"
                    btInserir.Enabled = True
                    btAlterar.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro D84 " & ex.Message)
        End Try
    End Sub

    Private Sub btAlterar_Click(sender As System.Object, e As System.EventArgs) Handles btAlterar.Click
        Try
            TesteAbertoDefeito()

            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet


            If btAlterar.Text = "Alterar" Then
                If MsgBox("Deseja Alterar um Defeito?", vbYesNo, "Alterar Defeito") = vbYes Then
                    txtCodigoRNC1.Clear()
                    txtDescricaoRNC1.Clear()
                    txtCodigoRNC1.Focus()
                    btAlterar.Text = "Aplicar"
                    btInserir.Enabled = False
                    btExcluir.Enabled = False
                Else
                End If
            Else
                If txtCodigoRNC1.TextLength = 0 Then
                    MsgBox("Selecione um Código na tabela abaixo", , "Selecione um Código")
                Else
                    TesteAbertoDefeito()
                    conDefeito.Open()
                    ds20 = New DataSet
                    da20 = New OleDbDataAdapter("UPDATE tblDefeitos SET  Nao_Conformidade = '" & txtDescricaoRNC1.Text & "' WHERE Codigo = " & txtCodigoRNC1.Text & "", conDefeito)
                    ds20.Clear()
                    da20.Fill(ds20, "tblDefeitos")
                    MsgBox("Registro Alterado com sucesso!")
                    conDefeito.Close()
                    Atualizar()
                    txtCodigoRNC1.Clear()
                    txtDescricaoRNC1.Clear()
                    btAlterar.Focus()
                    btAlterar.Text = "Alterar"
                    btInserir.Enabled = True
                    btExcluir.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro D73 " & ex.Message)
            conDefeito.Close()
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

        Dim Codigo = row.Cells(0)
        Dim Descricao = row.Cells(1)
        Me.txtCodigoRNC1.Text = Codigo.Value
        Me.txtDescricaoRNC1.Text = Descricao.Value

    End Sub

    Private Sub btCancelar_Click(sender As System.Object, e As System.EventArgs) Handles btCancelar.Click

        txtCodigoRNC1.Clear()
        txtDescricaoRNC1.Clear()

        btInserir.Enabled = True
        btExcluir.Enabled = True
        btAlterar.Enabled = True
       
        btExcluir.Text = "Excluir"
        btAlterar.Text = "Alterar"
        btInserir.Text = "Inserir"

        DataGridView1.Enabled = True
        txtCodigoRNC1.Enabled = True
        btInserir.Focus()
        Atualizar()
        conDefeito.Close()

    End Sub

    Private Sub btPesquisa_Click(sender As System.Object, e As System.EventArgs) Handles btPesquisa.Click
        Try
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            Dim seleccion As String
            If txtDescricaoRNC1.TextLength = 0 Then
                MsgBox("O Campo de Descrição está Vazio")
            Else
                seleccion = txtDescricaoRNC1.Text
                seleccion = "%" & seleccion & "%"
                DataGridView1.DataSource.clear()
                conDefeito.Open()
                Dim sel_ As String = "SELECT * FROM tblDefeitos WHERE Nao_COnformidade LIKE '" & seleccion & "' ORDER BY Nao_Conformidade ASC "
                da19 = New OleDbDataAdapter(sel_, conDefeito)
                ds19.Clear()
                da19.Fill(ds19, "tblDefeitos")
                conDefeito.Close()
                Me.DataGridView1.DataSource = ds19
                Me.DataGridView1.DataMember = "tblDefeitos"

                '1 - Coloca o Cabeçalho na coluna 
                DataGridView1.Columns(0).HeaderText = "Código"
                DataGridView1.Columns(1).HeaderText = "Não Conformidade"
                '2 - Acerta a largura da coluna em pixels
                DataGridView1.Columns(0).Width = 50
                'DataGridView1.Columns(1).Width = 350
                '3 - faz a coluna ajustar no resto do grid
                DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

            End If
        Catch ex As Exception
            MsgBox("Erro D71 " & ex.Message)
        End Try
    End Sub
    Private Sub frm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Try

            If e.KeyChar = Convert.ToChar(13) Then
                e.Handled = True

                SendKeys.Send("{TAB}")
            End If
        Catch ex As Exception
            MsgBox("Erro 53 " & ex.Message)
        End Try
    End Sub


    Private Sub txtDescricaoRNC1_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDescricaoRNC1.LostFocus
        Try
            If btInserir.Text = "Aplicar" Then
                btInserir.Focus()
            ElseIf btAlterar.Text = "Aplicar" Then
                btAlterar.Focus()
            ElseIf btExcluir.Text = "Aplicar" Then
                btExcluir.Focus()
            Else
                txtCodigoRNC1.Focus()
            End If
        Catch ex As Exception
            MsgBox("Erro D89 " & ex.Message)
        End Try

    End Sub

    Private Sub Quantidades2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoRNC1.KeyPress
        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(Numero(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro D55 " & ex.Message)
        End Try
    End Sub
    Function Numero(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            Numero = 0
        Else
            Numero = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                Numero = Keyascii
            Case 13
                Numero = Keyascii
                'Case 32 'permite espaço
                '   SoNumeros = Keyascii
        End Select
    End Function
End Class