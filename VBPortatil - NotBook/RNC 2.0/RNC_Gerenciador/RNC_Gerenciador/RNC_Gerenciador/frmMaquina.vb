Imports System.Data.OleDb
Public Class frmMaquina
    Dim conMaquina As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb;Jet OLEDB:Database Password= projetornc;")
    Private Sub frmMaquina_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        testeAbertoMaquina()
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


            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(0).HeaderText = "Máquina"
            DataGridView1.Columns(1).HeaderText = "Célula"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 190
            DataGridView1.Columns(2).Width = 40
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro M1 " & ex.Message)
            conMaquina.Close()
        End Try
    End Sub
    Sub testeAbertoMaquina()
        Dim RNC_Maquina As Boolean
        RNC_Maquina = Test("F:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb")
        If RNC_Maquina = True Then
            Dim RNCMaquina As Integer = 0
            For RNCMaquina = 5 To 20
                RNC_Maquina = Test("F:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb")
                If RNC_Maquina = True Then
                    RNCMaquina = 5
                    If (MsgBox("O Arquivo 'RNC_Maquina.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_Maquina.accdb")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNC_Maquina = False Then
                    RNCMaquina = 20
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

    Private Sub btInserir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btInserir.Click

        testeAbertoMaquina()
        If btInserir.Text = "Inserir" Then
            If MsgBox("Deseja Incluir um novo Registro?", vbYesNo, "Novo Registro") = vbYes Then
                btInserir.Text = "Aplicar"
                btAlterar.Enabled = False
                btExcluir.Enabled = False
                DataGridView1.Enabled = False
                lblID.Text = 0
                txtMaquina.Clear()
                txtCelula.Clear()
                txtMaquina.Focus()
            Else

            End If

        Else


            Try
                testeAbertoMaquina()
                conMaquina.Open()
                Dim da4 As New OleDbDataAdapter
                Dim ds4 As New DataSet
                ds4 = New DataSet
                da4 = New OleDbDataAdapter("INSERT INTO tblMaquina (Maquina, Celula) Values ('" & txtMaquina.Text & "', '" & txtCelula.Text & "') ", conMaquina)
                ds4.Clear()
                da4.Fill(ds4, "tblMaquina")
                conMaquina.Close()
                MsgBox("Registro Inserido com sucesso!")
                Atualizar()
                btInserir.Text = "Inserir"
                btAlterar.Enabled = True
                btExcluir.Enabled = True
                DataGridView1.Enabled = True
                txtMaquina.Clear()
                txtCelula.Clear()
                btInserir.Focus()

            Catch ex As Exception
                MsgBox("Erro M10 " & ex.Message)
            End Try

        End If

    End Sub
    Sub Atualizar()
        testeAbertoMaquina()
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


            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(0).HeaderText = "Máquina"
            DataGridView1.Columns(1).HeaderText = "Célula"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 190
            DataGridView1.Columns(2).Width = 40
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro M1X " & ex.Message)
            conMaquina.Close()
        End Try
    End Sub
    Private Sub frm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Try

            If e.KeyChar = Convert.ToChar(13) Then
                e.Handled = True

                SendKeys.Send("{TAB}")
            End If
        Catch ex As Exception
            MsgBox("Erro M53 " & ex.Message)
        End Try
    End Sub

    Private Sub txtCelula_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCelula.LostFocus
        Try
            If btInserir.Text = "Aplicar" Then
                btInserir.Focus()
            ElseIf btAlterar.Text = "Aplicar" Then
                btAlterar.Focus()
            ElseIf btExcluir.Text = "Aplicar" Then
                btExcluir.Focus()
            Else
                txtMaquina.Focus()
            End If
        Catch ex As Exception
            MsgBox("Erro M89 " & ex.Message)
        End Try
    End Sub

    Private Sub btAlterar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterar.Click
        Try
            testeAbertoMaquina()

            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet


            If btAlterar.Text = "Alterar" Then
                If MsgBox("Deseja Alterar um Registro?", vbYesNo, "Alterar Registro") = vbYes Then
                    txtMaquina.Clear()
                    lblID.Text = 0
                    txtCelula.Clear()
                    txtMaquina.Focus()
                    btAlterar.Text = "Aplicar"
                    btInserir.Enabled = False
                    btExcluir.Enabled = False
                Else
                End If
            Else
                If txtMaquina.TextLength = 0 Or txtCelula.TextLength = 0 Or lblID.Text = 0 Then
                    MsgBox("Selecione um Registro na tabela abaixo", , "Selecione um Registro")
                Else
                    testeAbertoMaquina()
                    conMaquina.Open()
                    ds20 = New DataSet
                    da20 = New OleDbDataAdapter("UPDATE tblMaquina SET  Maquina = '" & txtMaquina.Text & "', Celula = '" & txtCelula.Text & "' WHERE ID = " & lblID.Text & "", conMaquina)
                    ds20.Clear()
                    da20.Fill(ds20, "tblMaquina")
                    MsgBox("Registro Alterado com sucesso!")
                    conMaquina.Close()
                    Atualizar()
                    txtMaquina.Clear()
                    txtCelula.Clear()
                    lblID.Text = 0
                    btAlterar.Focus()
                    btAlterar.Text = "Alterar"
                    btInserir.Enabled = True
                    btExcluir.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro D73 " & ex.Message)
            conMaquina.Close()
        End Try
    End Sub
    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

        Dim Maquina = row.Cells(0)
        Dim Celula = row.Cells(1)
        Dim ID = row.Cells(2)

        Me.txtMaquina.Text = Maquina.Value
        Me.txtCelula.Text = Celula.Value
        Me.lblID.Text = ID.Value

    End Sub

    Private Sub btExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExcluir.Click


        Try
            testeAbertoMaquina()
            Dim da21 As New OleDbDataAdapter
            Dim ds21 As New DataSet
            If btExcluir.Text = "Excluir" Then
                If MsgBox("Deseja Excluir um Registro?", vbYesNo, "Excluir Registro") = vbYes Then
                    txtMaquina.Clear()
                    txtCelula.Clear()
                    lblID.Text = 0
                    btExcluir.Text = "Aplicar"
                    btInserir.Enabled = False
                    btAlterar.Enabled = False
                Else
                End If
            Else
                If txtMaquina.TextLength = 0 Or txtCelula.TextLength = 0 Or lblID.Text = 0 Then
                    MsgBox("Selecione um Registro na tabela abaixo", , "Selecione um Registro")
                Else
                    testeAbertoMaquina()
                    conMaquina.Open()
                    ds21 = New DataSet
                    da21 = New OleDbDataAdapter("delete from tblMaquina where ID = " & lblID.Text & " ", conMaquina)
                    ds21.Clear()
                    da21.Fill(ds21, "tblMaquina")
                    conMaquina.Close()
                    MsgBox("Registro Excluido com sucesso!")
                    Atualizar()
                    txtMaquina.Clear()
                    txtCelula.Clear()
                    lblID.Text = 0
                    btExcluir.Focus()
                    btExcluir.Text = "Excluir"
                    btInserir.Enabled = True
                    btAlterar.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro M84 " & ex.Message)
        End Try


    End Sub

    Private Sub btCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelar.Click

        rbCelula.Checked = False
        rbMaquina.Checked = False

        txtMaquina.Clear()
        txtCelula.Clear()
        lblID.Text = 0

        btInserir.Enabled = True
        btExcluir.Enabled = True
        btAlterar.Enabled = True

        btExcluir.Text = "Excluir"
        btAlterar.Text = "Alterar"
        btInserir.Text = "Inserir"

        DataGridView1.Enabled = True
        btInserir.Focus()
        Atualizar()
        conMaquina.Close()
    End Sub

    Private Sub btPesquisa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPesquisa.Click
        Try
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            Dim Maquina As String
            Dim Celula As String
            If rbMaquina.Checked = True Then
                If txtMaquina.TextLength = 0 Then
                    MsgBox("O Campo Máquina está Vázio")
                Else
                    Maquina = txtMaquina.Text
                    Maquina = "%" & Maquina & "%"
                    DataGridView1.DataSource.clear()
                    conMaquina.Open()
                    Dim sel_ As String = "SELECT * FROM tblMaquina WHERE Maquina LIKE '" & Maquina & "' ORDER BY Maquina ASC "
                    da19 = New OleDbDataAdapter(sel_, conMaquina)
                    ds19.Clear()
                    da19.Fill(ds19, "tblMaquina")
                    conMaquina.Close()
                    Me.DataGridView1.DataSource = ds19
                    Me.DataGridView1.DataMember = "tblMaquina"

                    '1 - Coloca o Cabeçalho na coluna 
                    DataGridView1.Columns(0).HeaderText = "Máquina"
                    DataGridView1.Columns(1).HeaderText = "Célula"
                    '2 - Acerta a largura da coluna em pixels
                    DataGridView1.Columns(0).Width = 190
                    DataGridView1.Columns(2).Width = 40
                    '3 - faz a coluna ajustar no resto do grid
                    DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice
                End If
            ElseIf rbCelula.Checked = True Then
                If txtCelula.TextLength = 0 Then
                    MsgBox("O Campo Célula está Vázio")
                Else
                    Celula = txtCelula.Text
                    Celula = "%" & Celula & "%"
                    DataGridView1.DataSource.clear()
                    conMaquina.Open()
                    Dim sel_ As String = "SELECT * FROM tblMaquina WHERE Celula LIKE '" & Celula & "' ORDER BY Maquina ASC "
                    da19 = New OleDbDataAdapter(sel_, conMaquina)
                    ds19.Clear()
                    da19.Fill(ds19, "tblMaquina")
                    conMaquina.Close()
                    Me.DataGridView1.DataSource = ds19
                    Me.DataGridView1.DataMember = "tblMaquina"

                    '1 - Coloca o Cabeçalho na coluna 
                    DataGridView1.Columns(0).HeaderText = "Máquina"
                    DataGridView1.Columns(1).HeaderText = "Célula"
                    '2 - Acerta a largura da coluna em pixels
                    DataGridView1.Columns(0).Width = 190
                    DataGridView1.Columns(2).Width = 40
                    '3 - faz a coluna ajustar no resto do grid
                    DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice
                End If
            Else
                MsgBox("Selecione um Tipo de Pesquisa", , "Tipo de Pesquisa")
            End If

        Catch ex As Exception
            MsgBox("Erro M71 " & ex.Message)
        End Try
    End Sub
End Class