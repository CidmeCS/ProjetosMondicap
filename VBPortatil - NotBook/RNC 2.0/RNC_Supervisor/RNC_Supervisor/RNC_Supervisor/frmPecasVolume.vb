﻿Imports System.Data.OleDb
Public Class frmPecasVolume
    Dim conPecasVolume As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb;Jet OLEDB:Database Password= projetornc;")
    Private Sub frmPecasVolume_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TesteAbertoPecasVolume()
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conPecasVolume.Open()
            Dim sel As String = "Select * from tblPecasVolume order by Cliente ASC"
            da = New OleDbDataAdapter(sel, conPecasVolume)
            ds.Clear()
            da.Fill(ds, "tblPecasVolume")
            conPecasVolume.Close()
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblPecasVolume"

            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(0).HeaderText = "Código"
            DataGridView1.Columns(3).HeaderText = "Peças por Volume"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(3).Width = 100
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro PV1 " & ex.Message)
            conPecasVolume.Close()
        End Try

    End Sub
    Sub Atualizar()
        TesteAbertoPecasVolume()
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conPecasVolume.Open()
            Dim sel As String = "Select * from tblPecasVolume order by Cliente ASC"
            da = New OleDbDataAdapter(sel, conPecasVolume)
            ds.Clear()
            da.Fill(ds, "tblPecasVolume")
            conPecasVolume.Close()
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblPecasVolume"

            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(0).HeaderText = "Código"
            DataGridView1.Columns(3).HeaderText = "Peças por Volume"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 100
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(3).Width = 100
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro PV2 " & ex.Message)
            conPecasVolume.Close()
        End Try

    End Sub
    Sub TesteAbertoPecasVolume()
        Dim RNC_RE As Boolean
        RNC_RE = Test("F:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb")
        If RNC_RE = True Then
            Dim RNCRE As Integer = 0
            For RNCRE = 5 To 20
                RNC_RE = Test("F:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb")
                If RNC_RE = True Then
                    RNCRE = 5
                    If (MsgBox("O Arquivo 'RNC_PecasVolume.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_PecasVolume.accdb")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNC_RE = False Then
                    RNCRE = 20
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


        TesteAbertoPecasVolume()
        If btInserir.Text = "Inserir" Then
            If MsgBox("Deseja Incluir um novo Registro?", vbYesNo, "Novo Registro") = vbYes Then
                btInserir.Text = "Aplicar"
                btAlterar.Enabled = False
                btExcluir.Enabled = False
                DataGridView1.Enabled = False
                txCodProduto.Clear()
                txtCliente.Clear()
                txtProduto.Clear()
                txtPecasVolume.Clear()
                txCodProduto.Focus()
            Else

            End If

        Else
            If txCodProduto.TextLength = 0 Then
                MsgBox("Insira um Código do Produto", , "Código do Produto")
                txCodProduto.Focus()
            ElseIf txtCliente.TextLength = 0 Then
                MsgBox("Insira um Cliente", , "Cliente")
                txtCliente.Focus()
            ElseIf txtProduto.TextLength = 0 Then
                MsgBox("Insira um Produto", , "Produto")
                txtProduto.Focus()
            ElseIf txtPecasVolume.TextLength = 0 Then
                MsgBox("Insira a quantidade de Peças por Volume", , "Peças por Volume")
                txtPecasVolume.Focus()
            Else

                Try
                    TesteAbertoPecasVolume()
                    conPecasVolume.Open()
                    Dim da4 As New OleDbDataAdapter
                    Dim ds4 As New DataSet
                    ds4 = New DataSet
                    da4 = New OleDbDataAdapter("INSERT INTO tblPecasVolume (Cod_Produto, Cliente, Produto, PecasVolume) Values (" & txCodProduto.Text & ", '" & txtCliente.Text & "', '" & txtProduto.Text & "', '" & txtPecasVolume.Text & "') ", conPecasVolume)
                    ds4.Clear()
                    da4.Fill(ds4, "tblPecasVolume")
                    conPecasVolume.Close()
                    MsgBox("Registro Inserido com sucesso!")
                    Atualizar()
                    btInserir.Text = "Inserir"
                    btAlterar.Enabled = True
                    btExcluir.Enabled = True
                    DataGridView1.Enabled = True
                    txCodProduto.Clear()
                    txtCliente.Clear()
                    txtProduto.Clear()
                    txtPecasVolume.Clear()
                    btInserir.Focus()

                Catch ex As Exception
                    MsgBox("Erro PV10 " & ex.Message)
                    conPecasVolume.Close()
                End Try
            End If
        End If
    End Sub

    Private Sub btExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExcluir.Click
        Try
            TesteAbertoPecasVolume()
            Dim da21 As New OleDbDataAdapter
            Dim ds21 As New DataSet
            If btExcluir.Text = "Excluir" Then
                If MsgBox("Deseja Excluir um Registro?", vbYesNo, "Excluir Registro") = vbYes Then
                    txCodProduto.Clear()
                    txtCliente.Clear()
                    txtProduto.Clear()
                    txtPecasVolume.Clear()
                    btExcluir.Text = "Aplicar"
                    btInserir.Enabled = False
                    btAlterar.Enabled = False
                Else
                End If
            Else
                If txCodProduto.Text = 0 Then
                    MsgBox("Selecione um Registro na tabela abaixo", , "Selecione um Registro")
                Else
                    TesteAbertoPecasVolume()
                    conPecasVolume.Open()
                    ds21 = New DataSet
                    da21 = New OleDbDataAdapter("delete from tblPecasVolume where Cod_Produto = " & txCodProduto.Text & " ", conPecasVolume)
                    ds21.Clear()
                    da21.Fill(ds21, "tblPecasVolume")
                    conPecasVolume.Close()
                    MsgBox("Registro Excluido com sucesso!")
                    Atualizar()
                    txCodProduto.Clear()
                    txtCliente.Clear()
                    txtProduto.Clear()
                    txtPecasVolume.Clear()
                    btExcluir.Focus()
                    btExcluir.Text = "Excluir"
                    btInserir.Enabled = True
                    btAlterar.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro PV33 " & ex.Message)
        End Try
    End Sub

    Private Sub btAlterar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterar.Click
        Try
            TesteAbertoPecasVolume()

            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet


            If btAlterar.Text = "Alterar" Then
                If MsgBox("Deseja Alterar um Registro?", vbYesNo, "Alterar Registro") = vbYes Then
                    txCodProduto.Clear()
                    txtCliente.Clear()
                    txtProduto.Clear()
                    txtPecasVolume.Clear()
                    txCodProduto.Focus()
                    btAlterar.Text = "Aplicar"
                    btInserir.Enabled = False
                    btExcluir.Enabled = False
                Else
                End If
            Else
                If txCodProduto.TextLength = 0 Then
                    MsgBox("Insira um Código do Produto", , "Código do Produto")
                    txCodProduto.Focus()
                ElseIf txtCliente.TextLength = 0 Then
                    MsgBox("Insira um Cliente", , "Cliente")
                    txtCliente.Focus()
                ElseIf txtProduto.TextLength = 0 Then
                    MsgBox("Insira um Produto", , "Produto")
                    txtProduto.Focus()
                ElseIf txtPecasVolume.TextLength = 0 Then
                    MsgBox("Insira a quantidade de Peças por Volume", , "Peças por Volume")
                    txtPecasVolume.Focus()
                Else
                    TesteAbertoPecasVolume()
                    conPecasVolume.Open()
                    ds20 = New DataSet
                    da20 = New OleDbDataAdapter("UPDATE tblPecasVolume SET Cliente = '" & txtCliente.Text & "', Produto = '" & txtProduto.Text & "', PecasVolume = " & txtPecasVolume.Text & " WHERE Cod_Produto = " & txCodProduto.Text & "", conPecasVolume)
                    ds20.Clear()
                    da20.Fill(ds20, "tblPecasVolume")
                    MsgBox("Registro Alterado com sucesso!")
                    conPecasVolume.Close()
                    Atualizar()
                    txCodProduto.Clear()
                    txtCliente.Clear()
                    txtProduto.Clear()
                    txtPecasVolume.Clear()
                    btAlterar.Focus()
                    btAlterar.Text = "Alterar"
                    btInserir.Enabled = True
                    btExcluir.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro PV73 " & ex.Message)
            conPecasVolume.Close()
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

        Dim Codigo = row.Cells(0)
        Dim Cliente = row.Cells(1)
        Dim Produto = row.Cells(2)
        Dim PecasVolume = row.Cells(3)

        Me.txCodProduto.Text = Codigo.Value
        Me.txtCliente.Text = Cliente.Value
        Me.txtProduto.Text = Produto.Value
        Me.txtPecasVolume.Text = PecasVolume.Value

    End Sub

    Private Sub btCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelar.Click

        txCodProduto.Clear()
        txtCliente.Clear()
        txtProduto.Clear()
        txtPecasVolume.Clear()

        btInserir.Enabled = True
        btExcluir.Enabled = True
        btAlterar.Enabled = True

        btExcluir.Text = "Excluir"
        btAlterar.Text = "Alterar"
        btInserir.Text = "Inserir"

        DataGridView1.Enabled = True
        btInserir.Focus()
        Atualizar()
        conPecasVolume.Close()

    End Sub

    Private Sub btPesquisa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPesquisa.Click
        Try
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            Dim seleccion As String
            If txtProduto.TextLength = 0 Then
                MsgBox("O Campo de Produto está Vázio")
            Else
                seleccion = txtProduto.Text
                seleccion = "%" & seleccion & "%"
                DataGridView1.DataSource.clear()
                conPecasVolume.Open()
                Dim sel_ As String = "SELECT * FROM tblPecasVolume WHERE Produto LIKE '" & seleccion & "' ORDER BY Produto ASC "
                da19 = New OleDbDataAdapter(sel_, conPecasVolume)
                ds19.Clear()
                da19.Fill(ds19, "tblPecasVolume")
                conPecasVolume.Close()
                Me.DataGridView1.DataSource = ds19
                Me.DataGridView1.DataMember = "tblPecasVolume"


                '1 - Coloca o Cabeçalho na coluna 
                DataGridView1.Columns(0).HeaderText = "Código"
                DataGridView1.Columns(3).HeaderText = "Peças por Volume"
                '2 - Acerta a largura da coluna em pixels
                DataGridView1.Columns(0).Width = 100
                DataGridView1.Columns(1).Width = 80
                DataGridView1.Columns(3).Width = 100
                '3 - faz a coluna ajustar no resto do grid
                DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

            End If
        Catch ex As Exception
            MsgBox("Erro PV71 " & ex.Message)
            conPecasVolume.Close()
        End Try
    End Sub
    Private Sub frm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Try

            If e.KeyChar = Convert.ToChar(13) Then
                e.Handled = True

                SendKeys.Send("{TAB}")
            End If
        Catch ex As Exception
            MsgBox("Erro PV53 " & ex.Message)
        End Try
    End Sub


    Private Sub SetorChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPecasVolume.LostFocus
        Try
            If btInserir.Text = "Aplicar" Then
                btInserir.Focus()
            ElseIf btAlterar.Text = "Aplicar" Then
                btAlterar.Focus()
            ElseIf btExcluir.Text = "Aplicar" Then
                btExcluir.Focus()
            Else
                txCodProduto.Focus()
            End If
        Catch ex As Exception
            MsgBox("Erro PV89 " & ex.Message)
        End Try

    End Sub

    Private Sub Quantidades2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txCodProduto.KeyPress, txtPecasVolume.KeyPress
        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(Numero(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro RE55 " & ex.Message)
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
