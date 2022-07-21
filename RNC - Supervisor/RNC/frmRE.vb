Imports System.Data.OleDb
Public Class frmRE
    Dim conRE As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RE.accdb;Jet OLEDB:Database Password= projetornc;")

    Private Sub frmRE_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        btInserir.Focus()
        TesteAbertoRE()
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conRE.Open()
            Dim sel As String = "Select * from tblRE order by Inspetor asc"
            da = New OleDbDataAdapter(sel, conRE)
            ds.Clear()
            da.Fill(ds, "tblRE")

            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblRE"

            conRE.Close()
            '1 - Coloca o Cabeçalho na coluna 
            'DataGridView1.Columns(0).HeaderText = "Código"
            'DataGridView1.Columns(1).HeaderText = "Não Conformidade"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 50
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(3).Width = 50
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice


        Catch ex As Exception
            Beep()
            MsgBox("Erro RE1 " & ex.Message)
            conRE.Close()
        End Try

    End Sub
    Sub Atualizar()
        TesteAbertoRE()
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet
            conRE.Open()
            Dim sel As String = "Select * from tblRE order by Inspetor asc"
            da = New OleDbDataAdapter(sel, conRE)
            ds.Clear()
            da.Fill(ds, "tblRE")
            conRE.Close()
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblRE"


            '1 - Coloca o Cabeçalho na coluna 
            'DataGridView1.Columns(0).HeaderText = "Código"
            'DataGridView1.Columns(1).HeaderText = "Não Conformidade"
            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 50
            DataGridView1.Columns(2).Width = 150
            DataGridView1.Columns(3).Width = 50
            '3 - faz a coluna ajustar no resto do grid
            DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

        Catch ex As Exception
            Beep()
            MsgBox("Erro RE2 " & ex.Message)
            conRE.Close()
        End Try

    End Sub
    Sub TesteAbertoRE()
        Dim RNC_RE As Boolean
        RNC_RE = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RE.accdb")
        If RNC_RE = True Then
            Dim RNCRE As Integer = 0
            For RNCRE = 5 To 20
                RNC_RE = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RE.accdb")
                If RNC_RE = True Then
                    RNCRE = 5
                    If (MsgBox("O Arquivo 'RNC_RE.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_RE.accdb")) = vbRetry Then
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

    Private Sub btInserir_Click(sender As System.Object, e As System.EventArgs) Handles btInserir.Click


        TesteAbertoRE()
        If btInserir.Text = "Inserir" Then
            If MsgBox("Deseja Incluir um novo Inspetor?", vbYesNo, "Novo Inspetor") = vbYes Then
                btInserir.Text = "Aplicar"
                btAlterar.Enabled = False
                btExcluir.Enabled = False
                DataGridView1.Enabled = False
                txtRE.Clear()
                txtInspetor.Clear()
                txtSetor.Clear()
                lblID.Text = 0
                txtRE.Focus()
            Else

            End If

        Else

            If txtRE.TextLength = 0 Then
                MsgBox("Insira um RE", , "RE")
                txtRE.Focus()
            ElseIf txtInspetor.TextLength = 0 Then
                MsgBox("Insira um Inspetor", , "Inspetor")
                txtInspetor.Focus()
            ElseIf txtSetor.TextLength = 0 Then
                MsgBox("Insira um Setor", , "Setor")
                txtSetor.Focus()
            Else
                Try
                    TesteAbertoRE()
                    conRE.Open()
                    Dim da4 As New OleDbDataAdapter
                    Dim ds4 As New DataSet
                    ds4 = New DataSet
                    da4 = New OleDbDataAdapter("INSERT INTO tblRE (RE, Inspetor, Setor) Values (" & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "') ", conRE)
                    ds4.Clear()
                    da4.Fill(ds4, "tblRE")
                    conRE.Close()
                    MsgBox("Registro Inserido com sucesso!")
                    Atualizar()
                    btInserir.Text = "Inserir"
                    btAlterar.Enabled = True
                    btExcluir.Enabled = True
                    DataGridView1.Enabled = True
                    txtRE.Clear()
                    txtInspetor.Clear()
                    txtSetor.Clear()
                    lblID.Text = 0
                    btInserir.Focus()

                Catch ex As Exception
                    MsgBox("Erro RE10 " & ex.Message)
                    conRE.Close()
                End Try
            End If
        End If
    End Sub

    Private Sub btExcluir_Click(sender As System.Object, e As System.EventArgs) Handles btExcluir.Click
        Try
            TesteAbertoRE()
            Dim da21 As New OleDbDataAdapter
            Dim ds21 As New DataSet
            If btExcluir.Text = "Excluir" Then
                If MsgBox("Deseja Excluir um Registro?", vbYesNo, "Excluir Registro") = vbYes Then
                    txtRE.Clear()
                    txtInspetor.Clear()
                    txtSetor.Clear()
                    lblID.Text = 0
                    btExcluir.Text = "Aplicar"
                    btInserir.Enabled = False
                    btAlterar.Enabled = False
                Else
                End If
            Else
                If lblID.Text = 0 Then
                    MsgBox("Selecione um Registro na tabela abaixo", , "Selecione um Registro")
                Else
                    TesteAbertoRE()
                    conRE.Open()
                    ds21 = New DataSet
                    da21 = New OleDbDataAdapter("delete from tblRE where ID = " & lblID.Text & " ", conRE)
                    ds21.Clear()
                    da21.Fill(ds21, "tblRE")
                    conRE.Close()
                    MsgBox("Registro Excluido com sucesso!")
                    Atualizar()
                    txtRE.Clear()
                    txtInspetor.Clear()
                    txtSetor.Clear()
                    lblID.Text = 0
                    btExcluir.Focus()
                    btExcluir.Text = "Excluir"
                    btInserir.Enabled = True
                    btAlterar.Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro RE33 " & ex.Message)
        End Try
    End Sub

    Private Sub btAlterar_Click(sender As System.Object, e As System.EventArgs) Handles btAlterar.Click
        Try
            TesteAbertoRE()

            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet


            If btAlterar.Text = "Alterar" Then
                If MsgBox("Deseja Alterar um Registro?", vbYesNo, "Alterar Registro") = vbYes Then
                    txtRE.Clear()
                    txtInspetor.Clear()
                    txtSetor.Clear()
                    lblID.Text = 0
                    txtRE.Focus()
                    btAlterar.Text = "Aplicar"
                    btInserir.Enabled = False
                    btExcluir.Enabled = False
                Else
                End If
            Else
                If txtRE.TextLength = 0 Then
                    MsgBox("Insira um RE", , "RE")
                    txtRE.Focus()
                ElseIf txtInspetor.TextLength = 0 Then
                    MsgBox("Insira um Inspetor", , "Inspetor")
                    txtInspetor.Focus()
                ElseIf txtSetor.TextLength = 0 Then
                    MsgBox("Insira um Setor", , "Setor")
                    txtInspetor.Focus()
                ElseIf lblID.Text = 0 Then
                    MsgBox("Selecione um Registro na tabela abaixo!", , "ID")
                Else
                    TesteAbertoRE()
                    conRE.Open()
                    ds20 = New DataSet
                    da20 = New OleDbDataAdapter("UPDATE tblRE SET  RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "' WHERE ID = " & lblID.Text & "", conRE)
                    ds20.Clear()
                    da20.Fill(ds20, "tblRE")
                    MsgBox("Registro Alterado com sucesso!")
                    conRE.Close()
                    Atualizar()
                    txtRE.Clear()
                    txtInspetor.Clear()
                    txtSetor.Clear()
                    lblID.Text = 0
                    btAlterar.Focus()
                    btAlterar.Text = "Alterar"
                    btInserir.Enabled = True
                    btExcluir.Enabled = True
                End If
                End If
        Catch ex As Exception
            MsgBox("Erro RE73 " & ex.Message)
            conRE.Close()
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

        Dim RE = row.Cells(0)
        Dim Inspetor = row.Cells(1)
        Dim Setor = row.Cells(2)
        Dim ID = row.Cells(3)

        Me.txtRE.Text = RE.Value
        Me.txtInspetor.Text = Inspetor.Value
        Me.txtSetor.Text = Setor.Value
        Me.lblID.Text = ID.Value

    End Sub

    Private Sub btCancelar_Click(sender As System.Object, e As System.EventArgs) Handles btCancelar.Click

        txtRE.Clear()
        txtInspetor.Clear()
        txtSetor.Clear()
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
        conRE.Close()

    End Sub

    Private Sub btPesquisa_Click(sender As System.Object, e As System.EventArgs) Handles btPesquisa.Click
        Try
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            Dim seleccion As String
            If txtInspetor.TextLength = 0 Then
                MsgBox("O Campo de Inspetor está Vazio")
            Else
                seleccion = txtInspetor.Text
                seleccion = "%" & seleccion & "%"
                DataGridView1.DataSource.clear()
                conRE.Open()
                Dim sel_ As String = "SELECT * FROM tblRE WHERE Inspetor LIKE '" & seleccion & "' ORDER BY Inspetor ASC "
                da19 = New OleDbDataAdapter(sel_, conRE)
                ds19.Clear()
                da19.Fill(ds19, "tblRE")
                conRE.Close()
                Me.DataGridView1.DataSource = ds19
                Me.DataGridView1.DataMember = "tblRE"

                '1 - Coloca o Cabeçalho na coluna 
                'DataGridView1.Columns(0).HeaderText = "Código"
                'DataGridView1.Columns(1).HeaderText = "Não Conformidade"
                '2 - Acerta a largura da coluna em pixels
                DataGridView1.Columns(0).Width = 50
                DataGridView1.Columns(2).Width = 150
                DataGridView1.Columns(3).Width = 50
                '3 - faz a coluna ajustar no resto do grid
                DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice
            End If
        Catch ex As Exception
            MsgBox("Erro RE71 " & ex.Message)
            conRE.Close()
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


    Private Sub SetorChanged(sender As System.Object, e As System.EventArgs) Handles txtSetor.LostFocus
        Try
            If btInserir.Text = "Aplicar" Then
                btInserir.Focus()
            ElseIf btAlterar.Text = "Aplicar" Then
                btAlterar.Focus()
            ElseIf btExcluir.Text = "Aplicar" Then
                btExcluir.Focus()
            Else
                txtRE.Focus()
            End If
        Catch ex As Exception
            MsgBox("Erro RE89 " & ex.Message)
        End Try

    End Sub

    Private Sub Quantidades2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtRE.KeyPress
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