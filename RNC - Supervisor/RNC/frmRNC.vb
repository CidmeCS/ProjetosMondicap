Imports System.Data.OleDb
Imports System.DBNull
Imports System.Diagnostics
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports RNC.Module1
Public Class frmRNC
    Dim conConsulta_OP As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conDefeito As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Defeito.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conMaquina As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Maquina.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conPecasVolume As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_PecasVolume.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conRE As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RE.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conRNC As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim cs As ConnectionState
    Dim Mes_ As String
    Dim Cliente As String
    Dim Celula As String
    Dim Defeito1, Defeito2, Defeito3, Defeito4, Defeito5, Defeito6, Defeito7, Defeito8, Defeito9, Defeito10 As String
    Dim Alteradu As String
    Dim seleccion3 As String
    Dim ID1, ID2, ID3, ID4, ID5, ID6, ID7, ID8, ID9, ID10 As Integer
    Dim compara As Int64
    Dim L1, L2, L3, L4, L5, L6, L7, L8, L9, L10 As String
    Dim Status1, Status2, Status3, Status4, Status5, Status6, Status7, Status8, Status9, Status10 As String
    Dim StatusAll As String
    Dim OPRetrabalho1, OPRetrabalho2, OPRetrabalho3, OPRetrabalho4, OPRetrabalho5, OPRetrabalho6, OPRetrabalho7, OPRetrabalho8, OPRetrabalho9, OPRetrabalho10 As String
    Dim Valor1, Valor2, Valor3, Valor4, Valor5, Valor6, Valor7, Valor8, Valor9, Valor10 As Int16
    Dim ValorX1, ValorX2, ValorX3, ValorX4, ValorX5, ValorX6, ValorX7, ValorX8, ValorX9, ValorX10 As Int16
    Dim ValorT, ValorXT As Int16
    Dim Limpo, Posicao As String
    Dim SMC As Int64 = 0
    'variaveis do FromTo
    Dim ds1FT As New DataSet
    Dim ds2FT As New DataSet
    Dim ds3FT As New DataSet
    Dim daFT As OleDbDataAdapter
    Dim connFT As OleDbConnection
    Dim cbFT As OleDbCommandBuilder
    Dim _connFT As String
    Dim da2FT As OleDbDataAdapter
    Dim conn2FT As OleDbConnection
    Dim cb2FT As OleDbCommandBuilder
    Dim AccessFT As Boolean
    Dim ExcelFT As Boolean

    Private Sub frmRNC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Today > "01/11/2099" Then
            MsgBox("Contate o Programador: Cid (15) 81797980 - cidevangelista@hotmail.com")
            Close()
        Else
            Call Teste_AbertoFT()
            Call PriMeiro_Passo()
            TesteAbertoRNC()
            Try
                Dim da As New OleDbDataAdapter
                Dim ds As New DataSet

                conRNC.Open()
                Dim sel As String = "Select top 100 * from tblRNC where Status = 'Pendente' and Disposicao <> 'Sem Disposição' order by ID desc"
                'Dim sel As String = "select Contador, count (*) from tblRNC group by Contador order by contador desc" 'conta quantas RNCs exitem
                da = New OleDbDataAdapter(sel, conRNC)
                ds.Clear()
                da.Fill(ds, "tblRNC")
                conRNC.Close()
                Me.DataGridView1.DataSource = ds
                Me.DataGridView1.DataMember = "tblRNC"
                FormatacaoGrid()
                'lblCodProduto.Text = DataGridView1.RowCount 'conta quantas RNCs exitem
                lblData.Text = Today
                lblHora.Text = TimeOfDay.ToShortTimeString

            Catch ex As Exception
                Beep()
                MsgBox("Erro 1 " & ex.Message)
            End Try
        End If
    End Sub
    'Inicio do FromTo
    Sub PriMeiro_Passo() 'Handles MyBase.Shown
        Try

            _connFT = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.xlsx;Extended Properties=Excel 8.0")
            Dim _connectionFT As OleDbConnection = New OleDbConnection(_connFT)
            Dim da As OleDbDataAdapter = New OleDbDataAdapter()
            Dim _commandFT As OleDbCommand = New OleDbCommand()
            _commandFT.Connection = _connectionFT
            _commandFT.CommandText = "SELECT top 100 * FROM [tblOP$] where OP > 0 order by OP desc "
            da.SelectCommand = _commandFT
            da.Fill(ds1FT, "tblOP")
            _connectionFT.Close()
            'Me.DataGridView1.DataSource = ds1
            'Me.DataGridView1.DataMember = "tblOP"
        Catch e1 As Exception
            ' lblMensagem.Text = "Falha"
            MessageBox.Show(" Erro 1TRE!", e1.Message)
        End Try
        Call segundo_Passo()
    End Sub

    Sub segundo_Passo() ' Handles MyBase.Shown
        Try
            conConsulta_OP.Open()
            Dim selFT As String = "SELECT top 100 * FROM tblOP where OP > 0  order by OP desc "
            daFT = New OleDbDataAdapter(selFT, conConsulta_OP)
            cbFT = New OleDbCommandBuilder(daFT)
            daFT.MissingSchemaAction = MissingSchemaAction.AddWithKey
            daFT.Fill(ds2FT, "tblOP")
            conConsulta_OP.Close()
            ' Me.DataGridView1.DataSource = ds2
            'Me.DataGridView1.DataMember = "tblOP"
        Catch e1 As Exception
            conConsulta_OP.Close()
            ' lblMensagem.Text = "Falha"
            MessageBox.Show("Erro 2TRE", e1.Message)
        End Try
        Call Terceiro_Passo()

    End Sub
    Sub Terceiro_Passo() 'Handles MyBase.Shown
        For Each dr As DataRow In ds1FT.Tables(0).Rows
            Dim expressionFT As String
            expressionFT = "OP = " + CType(dr.Item(0), Integer).ToString
            Dim drsFT() As DataRow = ds2FT.Tables(0).Select(expressionFT)
            If (drsFT.Length = 1) Then
                For i As Integer = 1 To ds2FT.Tables(0).Columns.Count - 1
                    drsFT(0).Item(i) = dr.Item(i)
                Next
            Else
                Dim drnewFT As DataRow = ds2FT.Tables(0).NewRow
                For i As Integer = 0 To ds2FT.Tables(0).Columns.Count - 1
                    drnewFT.Item(i) = dr.Item(i)
                Next
                ds2FT.Tables(0).Rows.Add(drnewFT)
            End If
        Next
        'Me.DataGridView1.DataSource = ds2
        ' Me.DataGridView1.DataMember = "tblOP"
        daFT.Update(ds2FT.Tables(0))
        Call Quarto_Passo()
    End Sub

    Sub Quarto_Passo() 'Handles MyBase.Shown
        Try
            conConsulta_OP.Open()
            Dim sel2FT As String = "SELECT top 100 * FROM tblOP where OP > 0 order by OP Desc"
            da2FT = New OleDbDataAdapter(sel2FT, conConsulta_OP)
            cb2FT = New OleDbCommandBuilder(da2FT)
            da2FT.MissingSchemaAction = MissingSchemaAction.AddWithKey
            da2FT.Fill(ds3FT, "tblOP")
            conConsulta_OP.Close()
            'DataGridView1.DataSource = ds3FT
            'DataGridView1.DataMember = "tblOP"
            ' lblMensagem.Text = "Mensagem: A Importação dos dados está Completa!"
            ' btFechar.Focus()
        Catch e1 As Exception
            conConsulta_OP.Close()
            'lblMensagem.Text = "Falha"
            MessageBox.Show("Erro 3!", e1.Message)
        End Try
    End Sub
    ' Private Sub btFechar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btFechar.Click
    ' Close()
    ' End Sub
    Sub Teste_AbertoFT()
        ExcelFT = TestFT("C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.xlsx")
        AccessFT = TestFT("C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.accdb")
        If ExcelFT = True Then
            MsgBox("O Arquivo Excel de importação está aberto, Feche-o para para continuar")
            Close()
        ElseIf AccessFT = True Then
            MsgBox("O Arquivo Access de importação está aberto, Feche-o para para continuar")
            Close()
        Else
        End If
    End Sub
    Function TestFT(ByVal pathfile As String) As Boolean
        Dim ffFT As Integer
        If System.IO.File.Exists(pathfile) Then
            Try
                ffFT = FreeFile()
                Microsoft.VisualBasic.FileOpen(ffFT, pathfile, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
                Return False
            Catch
                Return True
            Finally
                FileClose(ffFT)
            End Try
            Return True
        Else
        End If
        Return True
    End Function
    'Fim do FromTo


    Sub FormatacaoGrid()
        Try
            '1 - Coloca o Cabeçalho na coluna 
            DataGridView1.Columns(0).HeaderText = "ID"
            DataGridView1.Columns(1).HeaderText = "RNC"
            DataGridView1.Columns(2).HeaderText = "Státus"
            DataGridView1.Columns(3).HeaderText = "Origem"
            DataGridView1.Columns(4).HeaderText = "Data Abertura"
            DataGridView1.Columns(5).HeaderText = "Hora"
            DataGridView1.Columns(6).HeaderText = "Mês"
            DataGridView1.Columns(7).HeaderText = "Cód Produto"
            DataGridView1.Columns(8).HeaderText = "Cliente"
            DataGridView1.Columns(9).HeaderText = "Produto"
            DataGridView1.Columns(10).HeaderText = "OP Reprovada"
            DataGridView1.Columns(11).HeaderText = "Turno"
            DataGridView1.Columns(12).HeaderText = "Nº das Caixas"
            DataGridView1.Columns(13).HeaderText = "QT de Caixas"
            DataGridView1.Columns(14).HeaderText = "QT Reclamadada"
            DataGridView1.Columns(15).HeaderText = "Cód Defeito"
            DataGridView1.Columns(16).HeaderText = "Ñ Conformidade"
            DataGridView1.Columns(17).HeaderText = "Máquina"
            DataGridView1.Columns(18).HeaderText = "Célula"
            DataGridView1.Columns(19).HeaderText = "Disposição"
            DataGridView1.Columns(20).HeaderText = "OP Retrabalho"
            DataGridView1.Columns(21).HeaderText = "QT Reprovada"
            DataGridView1.Columns(22).HeaderText = "QT Aprovada"
            DataGridView1.Columns(23).HeaderText = "Data Encerramento"
            DataGridView1.Columns(24).HeaderText = " "
            DataGridView1.Columns(25).HeaderText = "Observação"
            DataGridView1.Columns(26).HeaderText = "RE"
            DataGridView1.Columns(27).HeaderText = "Inspetor"
            DataGridView1.Columns(28).HeaderText = "Setor"
            DataGridView1.Columns(29).HeaderText = "Turno Detector"
            DataGridView1.Columns(30).HeaderText = "Data Hora Alteração"
            DataGridView1.Columns(31).HeaderText = "Data Hora Fechamento"

            '2 - Acerta a largura da coluna em pixels
            DataGridView1.Columns(0).Width = 80
            DataGridView1.Columns(1).Width = 80
            DataGridView1.Columns(2).Width = 64
            DataGridView1.Columns(3).Width = 100
            DataGridView1.Columns(4).Width = 80
            DataGridView1.Columns(5).Width = 40
            DataGridView1.Columns(6).Width = 60
            DataGridView1.Columns(7).Width = 75
            DataGridView1.Columns(8).Width = 100
            DataGridView1.Columns(9).Width = 200
            DataGridView1.Columns(10).Width = 85
            DataGridView1.Columns(11).Width = 36
            DataGridView1.Columns(12).Width = 80
            DataGridView1.Columns(13).Width = 78
            DataGridView1.Columns(14).Width = 100
            DataGridView1.Columns(15).Width = 70
            DataGridView1.Columns(16).Width = 110
            DataGridView1.Columns(17).Width = 56
            DataGridView1.Columns(18).Width = 42
            DataGridView1.Columns(19).Width = 100
            DataGridView1.Columns(20).Width = 85
            DataGridView1.Columns(21).Width = 85
            DataGridView1.Columns(22).Width = 80
            DataGridView1.Columns(23).Width = 105
            DataGridView1.Columns(24).Width = 5
            DataGridView1.Columns(25).Width = 100
            DataGridView1.Columns(26).Width = 30
            DataGridView1.Columns(27).Width = 80
            DataGridView1.Columns(28).Width = 72
            DataGridView1.Columns(29).Width = 88
            DataGridView1.Columns(30).Width = 110
            DataGridView1.Columns(31).Width = 125
        Catch e1 As Exception
            MsgBox(e1.Message)
            MessageBox.Show("Erro 3y!", e1.Message)
        End Try

        '3 - faz a coluna ajustar no resto do grid
        'DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

        'lblCodProduto.Text = DataGridView1.RowCount 'conta quantas RNCs exitem
    End Sub
    Private Sub btInserir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btInserir.Click
        Try
            TesteAbertoRNC()
            If btInserir.Text = "Inserir" Then
                If MsgBox("Deseja Incluir uma Nova RNC?", vbYesNo, "Nova RNC") = vbYes Then
                    Call Limpar()
                    txtOP.Focus()
                    btInserir.Text = "Aplicar"
                    btAlterar.Enabled = False
                    btExcluir.Enabled = False
                    btImprimir.Enabled = False
                    btImprimirEtiqueta.Enabled = False
                    btEmail.Enabled = False
                    lblData.Text = Today
                    lblHora.Text = TimeOfDay.ToShortTimeString
                    DataGridView1.Enabled = False


                    conRNC.Open()
                    Dim sel2 As String = "SELECT top 1 ID, RNC FROM tblRNC order by ID desc"
                    Dim da2 As New OleDbDataAdapter
                    Dim ds2 As New DataSet
                    da2 = New OleDbDataAdapter(sel2, conRNC)
                    ds2.Clear()
                    da2.Fill(ds2, "tblRNC")
                    lblRNC.Text = ds2.Tables("tblRNC").Rows(0)("RNC") + 1
                    lblID.Text = ds2.Tables("tblRNC").Rows(0)("ID") + 1
                    conRNC.Close()


                Else

                End If

                'se botão Inserir for  = Aplicar
            Else
                ContarCaixas()

                If (MsgBox("O Total de caixas que você está reprovando é " & SMC & " ?", vbYesNo, "Confirmação de Quantidade de Caixas!!") = vbYes) Then

                    'radiobuton1
                    If rb1T.Checked = True Then
                        Call Verificacao1()

                        'radiobutton 2
                    ElseIf rb2T.Checked = True Then
                        Call Verificacao2()

                        'radiobutto 3
                    ElseIf rb3T.Checked = True Then
                        Call Verificacao3()
                        'radiobutton 4
                    ElseIf rb4T.Checked = True Then
                        Call Verificacao4()
                    ElseIf rb5T.Checked = True Then
                        Call Verificacao5()
                    ElseIf rb6T.Checked = True Then
                        Call Verificacao6()
                    ElseIf rb7T.Checked = True Then
                        Call Verificacao7()
                    ElseIf rb8T.Checked = True Then
                        Call Verificacao8()
                    ElseIf rb9T.Checked = True Then
                        Call Verificacao9()
                    ElseIf rb10T.Checked = True Then
                        Call Verificacao10()
                    Else
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 2 " & ex.Message)
        End Try
    End Sub

    Sub Verificacao10()
        Try
            If cb10Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb10Turno.Focus()
            ElseIf txtCaixas10Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas10Turno.Focus()
            ElseIf txtQtCaixasReprovada10.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada10.Focus()
            ElseIf txtCodigoRNC10.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC10.Focus()
            ElseIf txtDescricaoRNC10.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC10.Focus()
            Else
                Call Verificacao9()
            End If
        Catch ex As Exception
            MsgBox("Erro 4 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao9()
        Try
            If cb9Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb9Turno.Focus()
            ElseIf txtCaixas9Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas9Turno.Focus()
            ElseIf txtQtCaixasReprovada9.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada9.Focus()
            ElseIf txtCodigoRNC9.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC9.Focus()
            ElseIf txtDescricaoRNC9.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC9.Focus()
            Else
                Call Verificacao8()
            End If
        Catch ex As Exception
            MsgBox("Erro 5 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao8()
        Try
            If cb8Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb8Turno.Focus()
            ElseIf txtCaixas8Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas8Turno.Focus()
            ElseIf txtQtCaixasReprovada8.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada8.Focus()
            ElseIf txtCodigoRNC8.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC8.Focus()
            ElseIf txtDescricaoRNC8.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC8.Focus()
            Else
                Call Verificacao7()
            End If
        Catch ex As Exception
            MsgBox("Erro 6 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao7()
        Try
            If cb7Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb7Turno.Focus()
            ElseIf txtCaixas7Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas7Turno.Focus()
            ElseIf txtQtCaixasReprovada7.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada7.Focus()
            ElseIf txtCodigoRNC7.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC7.Focus()
            ElseIf txtDescricaoRNC7.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC7.Focus()
            Else
                Call Verificacao6()
            End If
        Catch ex As Exception
            MsgBox("Erro 7 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao6()
        Try
            If cb6Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb6Turno.Focus()
            ElseIf txtCaixas6Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas6Turno.Focus()
            ElseIf txtQtCaixasReprovada6.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada6.Focus()
            ElseIf txtCodigoRNC6.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC6.Focus()
            ElseIf txtDescricaoRNC6.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC6.Focus()
            Else
                Call Verificacao5()
            End If
        Catch ex As Exception
            MsgBox("Erro 8 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao5()
        Try
            If cb5Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb5Turno.Focus()
            ElseIf txtCaixas5Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas5Turno.Focus()
            ElseIf txtQtCaixasReprovada5.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada5.Focus()
            ElseIf txtCodigoRNC5.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC5.Focus()
            ElseIf txtDescricaoRNC5.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC5.Focus()
            Else
                Call Verificacao4()
            End If
        Catch ex As Exception
            MsgBox("Erro 9 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao4()
        Try
            If cb4Turnos.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb4Turnos.Focus()
            ElseIf txtCaixas4Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas4Turno.Focus()
            ElseIf txtQtCaixasReprovada4.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada4.Focus()
            ElseIf txtCodigoRNC4.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC4.Focus()
            ElseIf txtDescricaoRNC4.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC4.Focus()
            Else
                Call Verificacao3()
            End If
        Catch ex As Exception
            MsgBox("Erro 10 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao3()
        Try
            If cb3Turnos.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb3Turnos.Focus()
            ElseIf txtCaixas3Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas3Turno.Focus()
            ElseIf txtQtCaixasReprovada3.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada3.Focus()
            ElseIf txtCodigoRNC3.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC3.Focus()
            ElseIf txtDescricaoRNC3.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC3.Focus()
            Else
                Call Verificacao2()
            End If
        Catch ex As Exception
            MsgBox("Erro 11 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao2()
        Try
            If cb2Turnos.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb2Turnos.Focus()
            ElseIf txtCaixas2Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas2Turno.Focus()
            ElseIf txtQtCaixasReprovada2.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada2.Focus()
            ElseIf txtCodigoRNC2.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC2.Focus()
            ElseIf txtDescricaoRNC2.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC2.Focus()
            Else
                Call Verificacao1()
            End If
        Catch ex As Exception
            MsgBox("Erro 12 " & ex.Message)
        End Try
    End Sub
    Sub Verificacao1()
        Try
            If txtOP.TextLength = 0 Then
                MsgBox("O campo 'OP Reprovada' está vazio", , "OP Reprovada")
                txtOP.Focus()
            ElseIf cbDetectado.Text = "" Then
                MsgBox("O campo 'Detectado' está vazio", , "Detectado")
                cbDetectado.Focus()
            ElseIf txtMaquina.TextLength = 0 Then
                MsgBox("O campo 'Máquina' está vazio", , "Máquina")
                txtMaquina.Focus()
            ElseIf rb1T.Checked = False And rb2T.Checked = False And rb3T.Checked = False And rb4T.Checked = False And rb5T.Checked = False And rb6T.Checked = False And rb7T.Checked = False And rb8T.Checked = False And rb9T.Checked = False And rb10T.Checked = False Then
                MsgBox("Selecione a 'Quantidade' de Turnos que geraram a RNC", , "Quantidade de Turnos")
            ElseIf txtPecasPorVolume.Text = "" Then
                MsgBox("O campo 'Peças por Volume' está vazio", , "Peças por Volume")
                txtPecasPorVolume.Focus()
            ElseIf txtRE.Text = "" Then
                MsgBox("O campo 'RE' está Vazio", , "RE")
                txtRE.Focus()
            ElseIf txtInspetor.Text = "" Then
                MsgBox("O campo 'Inspetor' está Vazio", , "Inspetor")
                txtRE.Focus()
            ElseIf txtSetor.Text = "" Then
                MsgBox("O campo 'Setor' está Vazio", , "Setor")
                txtRE.Focus()
            ElseIf cbTurno.Text = "" Then
                MsgBox("O campo 'Turno Detector' está Vazio", , "Turno Detector")
                cbTurno.Focus()
            ElseIf cb1Turno.Text = "" Then
                MsgBox("O campo 'Turno' está vazio", , "Turno")
                cb1Turno.Focus()
            ElseIf txtCaixas1Turno.Text = "" Then
                MsgBox("O campo 'Nº das Caixas Reprovadas' está vazio", , "Caixas Reprovadas")
                txtCaixas1Turno.Focus()
            ElseIf txtQtCaixasReprovada1.Text = "" Then
                MsgBox("O campo 'Quantidade de Caixas Reprovadas' está vazio", , "Quantidade de Caixas Reprovadas")
                txtQtCaixasReprovada1.Focus()
            ElseIf txtCodigoRNC1.Text = "" Then
                MsgBox("O campo 'Codigo da RNC' está vazio", , "Codigo da RNC")
                txtCodigoRNC1.Focus()
            ElseIf txtDescricaoRNC1.Text = "" Then
                MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                txtDescricaoRNC1.Focus()
            Else
                If btInserir.Text = "Aplicar" Then
                    If rb1T.Checked = True Then
                        Call Inserir1()
                    ElseIf rb2T.Checked = True Then
                        Call Inserir2()
                    ElseIf rb3T.Checked = True Then
                        Call Inserir3()
                    ElseIf rb4T.Checked = True Then
                        Call Inserir4()
                    ElseIf rb5T.Checked = True Then
                        Call Inserir5()
                    ElseIf rb6T.Checked = True Then
                        Call Inserir6()
                    ElseIf rb7T.Checked = True Then
                        Call Inserir7()
                    ElseIf rb8T.Checked = True Then
                        Call Inserir8()
                    ElseIf rb9T.Checked = True Then
                        Call Inserir9()
                    ElseIf rb10T.Checked = True Then
                        Call Inserir10()
                    End If
                    conRNC.Close()
                    MsgBox("Dados Incluidos com Sucesso!")
                    If MsgBox("Você deseja 'Imprimir a RNC?'", vbYesNo, "Imprimir a RNC") = vbYes Then
                        Call ImprimirRNC()
                    End If
                    If MsgBox("Você deseja 'Imprimir as Etiquetas?'", vbYesNo, "Imprimir Etiquetas") = vbYes Then
                        ImprimirEtiqueta()
                    End If
                    If MsgBox("Você deseja enviar a RNC por 'E-mail?'", vbYesNo, "Enviar Email") = vbYes Then
                        Call email()
                    End If
                    Call Atualizar()
                End If

                If btAlterar.Text = "Aplicar" Then
                    If rb1T.Checked = True Then
                        Alterar1()
                    ElseIf rb2T.Checked = True Then
                        Alterar2()
                    ElseIf rb3T.Checked = True Then
                        Alterar3()
                    ElseIf rb4T.Checked = True Then
                        Alterar4()
                    ElseIf rb5T.Checked = True Then
                        Alterar5()
                    ElseIf rb6T.Checked = True Then
                        Alterar6()
                    ElseIf rb7T.Checked = True Then
                        Alterar7()
                    ElseIf rb8T.Checked = True Then
                        Alterar8()
                    ElseIf rb9T.Checked = True Then
                        Alterar9()
                    ElseIf rb10T.Checked = True Then
                        Alterar10()
                    End If
                    conRNC.Close()
                    MsgBox("Dados Alterados com Sucesso!")
                    If MsgBox("Você deseja 'Imprimir a RNC?'", vbYesNo, "Imprimir a RNC") = vbYes Then
                        Call ImprimirRNC()
                    End If
                    If MsgBox("Você deseja 'Imprimir as Etiquetas?'", vbYesNo, "Imprimir Etiquetas") = vbYes Then
                        ImprimirEtiqueta()
                    End If
                    If MsgBox("Você deseja enviar a RNC por 'E-mail?'", vbYesNo, "Enviar Email") = vbYes Then
                        Call email()
                    End If
                    Call Atualizar()
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 13 " & ex.Message)
        End Try
    End Sub
    Sub Atualizar()
        Try
            Dim da3 As New OleDbDataAdapter
            Dim ds3 As New DataSet
            Dim x As Integer
            x = lblRNC.Text
            Dim sel3 As String = "SELECT * FROM tblRNC where RNC = " & x & ""
            da3 = New OleDbDataAdapter(sel3, conRNC)
            ds3.Clear()
            da3.Fill(ds3, "tblRNC")
            Me.DataGridView1.DataSource = ds3
            Me.DataGridView1.DataMember = "tblRNC"
            FormatacaoGrid()
            conRNC.Close()
            Call Limpar()
        Catch ex As Exception
            MsgBox("Erro 14 " & ex.Message)
        End Try
    End Sub
    Sub Inserir1()
        Try
            '            Dim command As OleDbCommand = New OleDbCommand("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, 
            'Produto, OP_Reprovado, Turno, NúmerosCaixas, 
            'QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) values (@RNC, @Status, @Origem, @Data_Abertura, @Hora, @Mes, @Cod_Produto, @Cliente, @Produto, @OP_Reprovado, @Turno, @NúmerosCaixas, @QT_Caixas, @QT_Reprovado, @Cod_Defeito, @Nao_Conformidade, @Maquina, @Celula, @Observacao, @RE, @Inspetor, @Setor, @TurnoDetector")
            '            'Values (" & .Text & ", '', '" & .Text & "', '" & .Text & "', '" & .Text & "', '" &  & "', " & .Text & ", '" &  & "', '" & .Text & "', " & .Text & ", '" & .Text & "', '" & .Text & "', " & .Text & ", " & .Text & ", " & .Text & ", '" &  & "', '" & .Text & "', 
            '            '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            '            command.Parameters.Add("@RNC", OleDbType.Integer).Value = Convert.ToInt32(lblRNC.Text)
            '            command.Parameters.Add("@Status", OleDbType.VarChar).Value = "Pendente"
            '            command.Parameters.Add("@Origem", OleDbType.VarChar).Value = (cbDetectado.Text)
            '            command.Parameters.Add("@Data_Abertura", OleDbType.VarChar).Value = (lblData.Text)
            '            command.Parameters.Add("@Hora", OleDbType.VarChar).Value = (lblHora.Text)
            '            command.Parameters.Add("@Mes", OleDbType.VarChar).Value = (Mes_)
            '            command.Parameters.Add("@Cod_Produto", OleDbType.VarChar).Value = (lblCodProduto.Text)
            '            command.Parameters.Add("@Cliente", OleDbType.VarChar).Value = (Cliente)
            '            command.Parameters.Add("@Produto", OleDbType.VarChar).Value = (lblProduto.Text)
            '            command.Parameters.Add("@OP_Reprovado", OleDbType.VarChar).Value = (txtOP.Text)
            '            command.Parameters.Add("@Turno", OleDbType.VarChar).Value = (cb1Turno.Text)

            '            command.Parameters.Add("@NúmerosCaixas", OleDbType.VarChar).Value = (txtCaixas1Turno.Text)
            '            command.Parameters.Add("@QT_Caixas", OleDbType.VarChar).Value = (txtQtCaixasReprovada1.Text)
            '            command.Parameters.Add("@QT_Reprovado", OleDbType.VarChar).Value = (lblQtPorTurno1.Text)
            '            command.Parameters.Add("@Cod_Defeito", OleDbType.VarChar).Value = (txtCodigoRNC1.Text)
            '            command.Parameters.Add("@Nao_Conformidade", OleDbType.VarChar).Value = (Defeito1 & txtDescricaoRNC1.Text)
            '            command.Parameters.Add("@Maquina", OleDbType.VarChar).Value = (txtMaquina.Text)
            '            command.Parameters.Add("@Celula", OleDbType.VarChar).Value = (Celula)
            '            command.Parameters.Add("@Observacao", OleDbType.VarChar).Value = (txtOBS.Text)
            '            command.Parameters.Add("@RE", OleDbType.VarChar).Value = (txtRE.Text)

            '            command.Parameters.Add("@Inspetor", OleDbType.VarChar).Value = (txtInspetor.Text)
            '            command.Parameters.Add("@Setor", OleDbType.VarChar).Value = (cbTurno.Text)
            '            command.Parameters.Add("@TurnoDetector", OleDbType.VarChar).Value = (cbTurno.Text)

            '            conRNC.Open()
            '            command.ExecuteNonQuery()
            '            conRNC.Close()





            Dim da4 As New OleDbDataAdapter
            Dim ds4 As New DataSet
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds4 = New DataSet
            da4 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb1Turno.Text & "', '" & txtCaixas1Turno.Text & "', " & txtQtCaixasReprovada1.Text & ", " & lblQtPorTurno1.Text & ", " & txtCodigoRNC1.Text & ", '" & Defeito1 & txtDescricaoRNC1.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds4.Clear()
            da4.Fill(ds4, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 15 " & ex.Message)
        End Try
    End Sub
    Sub Inserir2()
        Try
            Dim da5 As New OleDbDataAdapter
            Dim ds5 As New DataSet
            Call Inserir1()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds5 = New DataSet
            da5 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb2Turnos.Text & "', '" & txtCaixas2Turno.Text & "', " & txtQtCaixasReprovada2.Text & ", '" & lblQtPorTurno2.Text & "', " & txtCodigoRNC2.Text & ", '" & Defeito2 & txtDescricaoRNC2.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds5.Clear()
            da5.Fill(ds5, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 16 " & ex.Message)
        End Try
    End Sub
    Sub Inserir3()
        Try
            Dim da6 As New OleDbDataAdapter
            Dim ds6 As New DataSet
            Call Inserir2()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds6 = New DataSet
            da6 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb3Turnos.Text & "', '" & txtCaixas3Turno.Text & "', " & txtQtCaixasReprovada3.Text & ", '" & lblQtPorTurno3.Text & "', " & txtCodigoRNC3.Text & ", '" & Defeito3 & txtDescricaoRNC3.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "','" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds6.Clear()
            da6.Fill(ds6, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 17 " & ex.Message)
        End Try
    End Sub
    Sub Inserir4()
        Try
            Dim da7 As New OleDbDataAdapter
            Dim ds7 As New DataSet
            Call Inserir3()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds7 = New DataSet
            da7 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb4Turnos.Text & "', '" & txtCaixas4Turno.Text & "', " & txtQtCaixasReprovada4.Text & ", '" & lblQtPorTurno4.Text & "', " & txtCodigoRNC4.Text & ", '" & Defeito4 & txtDescricaoRNC4.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds7.Clear()
            da7.Fill(ds7, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 18 " & ex.Message)
        End Try
    End Sub
    Sub Inserir5()
        Try
            Dim da8 As New OleDbDataAdapter
            Dim ds8 As New DataSet
            Call Inserir4()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds8 = New DataSet
            da8 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb5Turno.Text & "', '" & txtCaixas5Turno.Text & "', " & txtQtCaixasReprovada5.Text & ", '" & lblQtPorTurno5.Text & "', " & txtCodigoRNC5.Text & ", '" & Defeito5 & txtDescricaoRNC5.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds8.Clear()
            da8.Fill(ds8, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 19 " & ex.Message)
        End Try
    End Sub
    Sub Inserir6()
        Try
            Dim da8 As New OleDbDataAdapter
            Dim ds8 As New DataSet
            Call Inserir5()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds8 = New DataSet
            da8 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb6Turno.Text & "', '" & txtCaixas6Turno.Text & "', " & txtQtCaixasReprovada6.Text & ", '" & lblQtPorTurno6.Text & "', " & txtCodigoRNC6.Text & ", '" & Defeito6 & txtDescricaoRNC6.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds8.Clear()
            da8.Fill(ds8, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 20 " & ex.Message)
        End Try
    End Sub
    Sub Inserir7()
        Try
            Dim da8 As New OleDbDataAdapter
            Dim ds8 As New DataSet
            Call Inserir6()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds8 = New DataSet
            da8 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb7Turno.Text & "', '" & txtCaixas7Turno.Text & "', " & txtQtCaixasReprovada7.Text & ", '" & lblQtPorTurno7.Text & "', " & txtCodigoRNC7.Text & ", '" & Defeito7 & txtDescricaoRNC7.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds8.Clear()
            da8.Fill(ds8, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 21 " & ex.Message)
        End Try
    End Sub
    Sub Inserir8()
        Try
            Dim da9 As New OleDbDataAdapter
            Dim ds9 As New DataSet
            Call Inserir7()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds9 = New DataSet
            da9 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb8Turno.Text & "', '" & txtCaixas8Turno.Text & "', " & txtQtCaixasReprovada8.Text & ", '" & lblQtPorTurno8.Text & "', " & txtCodigoRNC8.Text & ", '" & Defeito8 & txtDescricaoRNC8.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds9.Clear()
            da9.Fill(ds9, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 22 " & ex.Message)
        End Try
    End Sub
    Sub Inserir9()
        Try
            Dim da10 As New OleDbDataAdapter
            Dim ds10 As New DataSet
            Call Inserir8()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds10 = New DataSet
            da10 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb9Turno.Text & "', '" & txtCaixas9Turno.Text & "', " & txtQtCaixasReprovada9.Text & ", '" & lblQtPorTurno9.Text & "', " & txtCodigoRNC9.Text & ", '" & Defeito9 & txtDescricaoRNC9.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds10.Clear()
            da10.Fill(ds10, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 23 " & ex.Message)
        End Try
    End Sub
    Sub Inserir10()
        Try
            Dim da11 As New OleDbDataAdapter
            Dim ds11 As New DataSet
            Call Inserir9()
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds11 = New DataSet
            da11 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb10Turno.Text & "', '" & txtCaixas10Turno.Text & "', " & txtQtCaixasReprovada10.Text & ", '" & lblQtPorTurno10.Text & "', " & txtCodigoRNC10.Text & ", '" & Defeito10 & txtDescricaoRNC10.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
            ds11.Clear()
            da11.Fill(ds11, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 24 " & ex.Message)
        End Try
    End Sub
    Sub Mes()
        Try
            Dim Mes As Int16
            Mes = Today.Month
            Select Case Mes
                Case 1
                    Mes_ = "Janeiro"
                Case 2
                    Mes_ = "Fevereiro"
                Case 3
                    Mes_ = "Março"
                Case 4
                    Mes_ = "Abril"
                Case 5
                    Mes_ = "Maio"
                Case 6
                    Mes_ = "Junho"
                Case 7
                    Mes_ = "Julho"
                Case 8
                    Mes_ = "Agosto"
                Case 9
                    Mes_ = "Setembro"
                Case 10
                    Mes_ = "Outubro"
                Case 11
                    Mes_ = "Novembro"
                Case 12
                    Mes_ = "Dezembro"
            End Select
        Catch ex As Exception
            MsgBox("Erro 25 " & ex.Message)
        End Try
    End Sub
    Sub Clientex()
        Try
            TesteAbertoPecasVolume()
            Dim da8 As New OleDbDataAdapter
            Dim ds8 As New DataSet
            Dim dt8 As New DataTable

            Dim sel7 As String = "SELECT Cliente FROM tblPecasVolume where Cod_Produto = " & lblCodProduto.Text & " "
            da8 = New OleDbDataAdapter(sel7, conPecasVolume)
            ds8.Clear()
            dt8.Clear()
            da8.Fill(dt8)
            da8.Fill(ds8, "tblPecasVolume")
            If dt8.Rows.Count = 0 Then
            Else
                Cliente = ds8.Tables("tblPecasVolume").Rows(0)("Cliente")
            End If
        Catch ex As Exception
            MsgBox("Erro 26 " & ex.Message)
        End Try
    End Sub
    Sub Celulax()
        Try
            testeAbertoMaquina()
            Dim da9 As New OleDbDataAdapter
            Dim ds9 As New DataSet
            Dim dt9 As New DataTable
            conMaquina.Open()
            Dim sel8 As String = "SELECT Celula FROM tblMaquina where Maquina = '" & txtMaquina.Text & "' "
            da9 = New OleDbDataAdapter(sel8, conMaquina)
            dt9.Clear()
            da9.Fill(dt9)
            If dt9.Rows.Count = 0 Then
                conMaquina.Close()
                If (MsgBox("A Maquina não está cadastrada, deseja inserir um valor padrão e solicitar o cadastro?", vbYesNo, "Maquina")) = vbYes Then
                    txtMaquina.Text = "0"
                    rb1T.Focus()
                Else
                    txtMaquina.Clear()
                    txtMaquina.Focus()
                End If

            Else
                ds9.Clear()
                da9.Fill(ds9, "tblMaquina")
                Celula = ds9.Tables("tblMaquina").Rows(0)("Celula")
                lblcelula.Text = ds9.Tables("tblMaquina").Rows(0)("Celula")
                conMaquina.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 27 " & ex.Message)
        End Try
    End Sub
    Sub Limpar()
        Try
            DataGridView1.Enabled = True
            lblID.Text = "*"
            lblRNC.Text = "*"
            lblData.Text = "*"
            lblHora.Text = "*"
            Limpo = "limpo"
            txtOP.Clear()
            Limpo = ""
            lblCodProduto.Text = "*"
            lblProduto.Text = "*"
            txtMaquina.Clear()
            cbDetectado.ResetText()
            rb1T.Checked = True

            cb1Turno.ResetText()
            cb2Turnos.ResetText()
            cb3Turnos.ResetText()
            cb4Turnos.ResetText()
            cb5Turno.ResetText()
            cb6Turno.ResetText()
            cb7Turno.ResetText()
            cb8Turno.ResetText()
            cb9Turno.ResetText()
            cb10Turno.ResetText()

            txtCaixas1Turno.Clear()
            txtCaixas2Turno.Clear()
            txtCaixas3Turno.Clear()
            txtCaixas4Turno.Clear()
            txtCaixas5Turno.Clear()
            txtCaixas6Turno.Clear()
            txtCaixas7Turno.Clear()
            txtCaixas8Turno.Clear()
            txtCaixas9Turno.Clear()
            txtCaixas10Turno.Clear()

            txtQtCaixasReprovada1.Clear()
            txtQtCaixasReprovada2.Clear()
            txtQtCaixasReprovada3.Clear()
            txtQtCaixasReprovada4.Clear()
            txtQtCaixasReprovada5.Clear()
            txtQtCaixasReprovada6.Clear()
            txtQtCaixasReprovada7.Clear()
            txtQtCaixasReprovada8.Clear()
            txtQtCaixasReprovada9.Clear()
            txtQtCaixasReprovada10.Clear()

            lblQtPorTurno1.Text = 0
            lblQtPorTurno2.Text = 0
            lblQtPorTurno3.Text = 0
            lblQtPorTurno4.Text = 0
            lblQtPorTurno5.Text = 0
            lblQtPorTurno6.Text = 0
            lblQtPorTurno7.Text = 0
            lblQtPorTurno8.Text = 0
            lblQtPorTurno9.Text = 0
            lblQtPorTurno10.Text = 0

            txtPecasPorVolume.Clear()
            lblTotalPecas.Text = 0
            txtCodigoRNC1.Clear()
            txtDescricaoRNC1.Clear()
            txtCodigoRNC2.Clear()
            txtDescricaoRNC2.Clear()
            txtCodigoRNC3.Clear()
            txtDescricaoRNC3.Clear()
            txtCodigoRNC4.Clear()
            txtDescricaoRNC4.Clear()
            txtCodigoRNC5.Clear()
            txtDescricaoRNC5.Clear()
            txtCodigoRNC6.Clear()
            txtDescricaoRNC6.Clear()
            txtCodigoRNC7.Clear()
            txtDescricaoRNC7.Clear()
            txtCodigoRNC8.Clear()
            txtDescricaoRNC8.Clear()
            txtCodigoRNC9.Clear()
            txtDescricaoRNC9.Clear()
            txtCodigoRNC10.Clear()
            txtDescricaoRNC10.Clear()


            txtOBS.Clear()
            txtRE.Clear()
            txtInspetor.Clear()
            btInserir.Text = "Inserir"
            btInserir.Enabled = True
            btAlterar.Text = "Alterar"
            btAlterar.Enabled = True
            btExcluir.Text = "Excluir"
            btExcluir.Enabled = True
            btImprimir.Enabled = True
            btImprimirEtiqueta.Enabled = True
            txtInspetor.Enabled = True
            btEmail.Enabled = True

            rb1T.Enabled = True
            rb2T.Enabled = True
            rb3T.Enabled = True
            rb4T.Enabled = True
            rb5T.Enabled = True
            rb6T.Enabled = True
            rb7T.Enabled = True
            rb8T.Enabled = True
            rb9T.Enabled = True
            rb10T.Enabled = True


            txtSetor.Enabled = True
            txtSetor.Text = ""
            cbTurno.Text = ""

            cb1Turno.Enabled = True
            cb2Turnos.Enabled = True
            cb3Turnos.Enabled = True
            cb4Turnos.Enabled = True
            cb5Turno.Enabled = True
            cb6Turno.Enabled = True
            cb7Turno.Enabled = True
            cb8Turno.Enabled = True
            cb9Turno.Enabled = True
            cb10Turno.Enabled = True

            txtCodigoRNC1.Enabled = True
            txtCodigoRNC2.Enabled = True
            txtCodigoRNC3.Enabled = True
            txtCodigoRNC4.Enabled = True
            txtCodigoRNC5.Enabled = True
            txtCodigoRNC6.Enabled = True
            txtCodigoRNC7.Enabled = True
            txtCodigoRNC8.Enabled = True
            txtCodigoRNC9.Enabled = True
            txtCodigoRNC10.Enabled = True

            lblCarregada.Text = "*"
            compara = 0

            conConsulta_OP.Close()
            conDefeito.Close()
            conMaquina.Close()
            conPecasVolume.Close()
            conRE.Close()
            conRNC.Close()
        Catch ex As Exception
            MsgBox("Erro 28 " & ex.Message)
        End Try
    End Sub
    Private Sub txtMaquina_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMaquina.LostFocus
        Try
            If txtMaquina.TextLength = 0 Then

            Else
                Call Celulax()
            End If
        Catch ex As Exception
            MsgBox("Erro 29 " & ex.Message)
        End Try
    End Sub
    'os dois metodos abaixo é para carregar os dados da OP nas labels codigo do produto e produto
    Private Sub txtOP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP.LostFocus
        Try
            If txtOP.Text = "" Or txtOP.Text = "0" Or txtOP.Text = "00" Or txtOP.Text = "000" Or txtOP.Text = "0000" Or txtOP.Text = "00000" Or txtOP.Text = "000000" Then
                If Limpo = "" Then
                    MsgBox("Insira uma 'OP' válida", , "OP")
                End If
            Else
                TesteAbertoConsultaOP()
                TesteAbertoPecasVolume()
                seleccion3 = txtOP.Text
                seleccion3 = "" & seleccion3 & ""
                Dim da10 As New OleDbDataAdapter
                Dim ds10 As New DataSet
                Dim dt10 As New DataTable
                Dim cb10 As New OleDbCommandBuilder
                conConsulta_OP.Open()
                Dim sel12 As String = "SELECT top 1 Cod_Mondicap, Descricao FROM tblOP where OP = " & seleccion3 & "  "
                da10 = New OleDbDataAdapter(sel12, conConsulta_OP)
                dt10.Clear()
                da10.Fill(dt10)
                If dt10.Rows.Count = 0 Then
                    conConsulta_OP.Close()
                    MsgBox("A OP não existe")
                    lblCodProduto.Text = "*"
                    lblProduto.Text = "*"
                    txtOP.Focus()
                Else
                    lblCodProduto.Text = dt10.Rows(0)("Cod_Mondicap")
                    lblProduto.Text = dt10.Rows(0)("Descricao")
                    conConsulta_OP.Close()
                    Dim da12 As New OleDbDataAdapter
                    Dim dt12 As New DataTable
                    Dim ds12 As New DataSet
                    conPecasVolume.Open()
                    Dim sel5 As String = "SELECT top 1 PecasVolume FROM tblPecasVolume where Cod_Produto = " & lblCodProduto.Text & ""
                    da12 = New OleDbDataAdapter(sel5, conPecasVolume)
                    dt12.Clear()
                    da12.Fill(dt12)
                    If dt12.Rows.Count = 0 Then
                        conPecasVolume.Close()
                        conConsulta_OP.Close()
                        MsgBox("'Peçcas Por Volume' Não Cadastrado, insira a quantidade manualmente e solicite o cadastro para este item", , "Peças por Volume")
                        txtPecasPorVolume.Clear()
                        'txtPecasPorVolume.Focus()
                    Else
                        da12.Fill(ds12, "tblPecasVolume")
                        txtPecasPorVolume.Text = ds12.Tables("tblPecasVolume").Rows(0)("PecasVolume")
                        conPecasVolume.Close()
                    End If
                    conConsulta_OP.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 30 " & ex.Message)
        End Try
    End Sub
    Sub QuantidadeTurnos()
        Try
            If rb1T.Checked = True Then
                rb1v()

                rb2f()
                rb3f()
                rb4f()
                rb5f()
                rb6f()
                rb7f()
                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb2T.Checked = True Then
                rb1v()
                rb2v()

                rb3f()
                rb4f()
                rb5f()
                rb6f()
                rb7f()
                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb3T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()

                rb4f()
                rb5f()
                rb6f()
                rb7f()
                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb4T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()

                rb5f()
                rb6f()
                rb7f()
                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb5T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()
                rb5v()

                rb6f()
                rb7f()
                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb6T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()
                rb5v()
                rb6v()

                rb7f()
                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb7T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()
                rb5v()
                rb6v()
                rb7v()

                rb8f()
                rb9f()
                rb10f()
                Calcular()
            ElseIf rb8T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()
                rb5v()
                rb6v()
                rb7v()
                rb8v()

                rb9f()
                rb10f()
                Calcular()
            ElseIf rb9T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()
                rb5v()
                rb6v()
                rb7v()
                rb8v()
                rb9v()

                rb10f()
                Calcular()
            ElseIf rb10T.Checked = True Then
                rb1v()
                rb2v()
                rb3v()
                rb4v()
                rb5v()
                rb6v()
                rb7v()
                rb8v()
                rb9v()
                rb10v()
                Calcular()
            End If
        Catch ex As Exception
            MsgBox("Erro 31 " & ex.Message)
        End Try
    End Sub
    Sub rb1v()
        Try
            cb1Turno.Visible = True
            txtCaixas1Turno.Visible = True
            txtQtCaixasReprovada1.Visible = True
            lblQtPorTurno1.Visible = True
            txtCodigoRNC1.Visible = True
            txtDescricaoRNC1.Visible = True
            txtQTReprovado1.Visible = True
            txtQTAprovado1.Visible = True
            gb1.Visible = True
        Catch ex As Exception
            MsgBox("Erro 32 " & ex.Message)
        End Try
    End Sub
    Sub rb2v()
        Try
            cb2Turnos.Visible = True
            txtCaixas2Turno.Visible = True
            txtQtCaixasReprovada2.Visible = True
            lblQtPorTurno2.Visible = True
            txtCodigoRNC2.Visible = True
            txtDescricaoRNC2.Visible = True
            txtQTReprovado2.Visible = True
            txtQTAprovado2.Visible = True
            gb2.Visible = True
        Catch ex As Exception
            MsgBox("Erro 33 " & ex.Message)
        End Try
    End Sub
    Sub rb3v()
        Try
            cb3Turnos.Visible = True
            txtCaixas3Turno.Visible = True
            txtQtCaixasReprovada3.Visible = True
            lblQtPorTurno3.Visible = True
            txtCodigoRNC3.Visible = True
            txtDescricaoRNC3.Visible = True
            txtQTReprovado3.Visible = True
            txtQTAprovado3.Visible = True
            gb3.Visible = True
        Catch ex As Exception
            MsgBox("Erro 34 " & ex.Message)
        End Try
    End Sub
    Sub rb4v()
        Try
            cb4Turnos.Visible = True
            txtCaixas4Turno.Visible = True
            txtQtCaixasReprovada4.Visible = True
            lblQtPorTurno4.Visible = True
            txtCodigoRNC4.Visible = True
            txtDescricaoRNC4.Visible = True
            txtQTReprovado4.Visible = True
            txtQTAprovado4.Visible = True
            gb4.Visible = True
        Catch ex As Exception
            MsgBox("Erro 35 " & ex.Message)
        End Try
    End Sub
    Sub rb5v()
        Try
            cb5Turno.Visible = True
            txtCaixas5Turno.Visible = True
            txtQtCaixasReprovada5.Visible = True
            lblQtPorTurno5.Visible = True
            txtCodigoRNC5.Visible = True
            txtDescricaoRNC5.Visible = True
            txtQTReprovado5.Visible = True
            txtQTAprovado5.Visible = True
            gb5.Visible = True
        Catch ex As Exception
            MsgBox("Erro 36 " & ex.Message)
        End Try
    End Sub
    Sub rb6v()
        Try
            cb6Turno.Visible = True
            txtCaixas6Turno.Visible = True
            txtQtCaixasReprovada6.Visible = True
            lblQtPorTurno6.Visible = True
            txtCodigoRNC6.Visible = True
            txtDescricaoRNC6.Visible = True
            txtQTReprovado6.Visible = True
            txtQTAprovado6.Visible = True
            gb6.Visible = True
        Catch ex As Exception
            MsgBox("Erro 37 " & ex.Message)
        End Try
    End Sub
    Sub rb7v()
        Try
            cb7Turno.Visible = True
            txtCaixas7Turno.Visible = True
            txtQtCaixasReprovada7.Visible = True
            lblQtPorTurno7.Visible = True
            txtCodigoRNC7.Visible = True
            txtDescricaoRNC7.Visible = True
            txtQTReprovado7.Visible = True
            txtQTAprovado7.Visible = True
            gb7.Visible = True
        Catch ex As Exception
            MsgBox("Erro 38 " & ex.Message)
        End Try
    End Sub
    Sub rb8v()
        Try
            cb8Turno.Visible = True
            txtCaixas8Turno.Visible = True
            txtQtCaixasReprovada8.Visible = True
            lblQtPorTurno8.Visible = True
            txtCodigoRNC8.Visible = True
            txtDescricaoRNC8.Visible = True
            txtQTReprovado8.Visible = True
            txtQTAprovado8.Visible = True
            gb8.Visible = True
        Catch ex As Exception
            MsgBox("Erro 39 " & ex.Message)
        End Try
    End Sub
    Sub rb9v()
        Try
            cb9Turno.Visible = True
            txtCaixas9Turno.Visible = True
            txtQtCaixasReprovada9.Visible = True
            lblQtPorTurno9.Visible = True
            txtCodigoRNC9.Visible = True
            txtDescricaoRNC9.Visible = True
            txtQTReprovado9.Visible = True
            txtQTAprovado9.Visible = True
            gb9.Visible = True
        Catch ex As Exception
            MsgBox("Erro 40 " & ex.Message)
        End Try
    End Sub
    Sub rb10v()
        Try
            cb10Turno.Visible = True
            txtCaixas10Turno.Visible = True
            txtQtCaixasReprovada10.Visible = True
            lblQtPorTurno10.Visible = True
            txtCodigoRNC10.Visible = True
            txtDescricaoRNC10.Visible = True
            txtQTReprovado10.Visible = True
            txtQTAprovado10.Visible = True
            gb10.Visible = True
        Catch ex As Exception
            MsgBox("Erro 41 " & ex.Message)
        End Try
    End Sub
    Sub rb1f()
        Try
            cb1Turno.Visible = False
            txtCaixas1Turno.Visible = False
            txtQtCaixasReprovada1.Visible = False
            lblQtPorTurno1.Visible = False
            txtCodigoRNC1.Visible = False
            txtDescricaoRNC1.Visible = False
            txtQTReprovado1.Visible = False
            txtQTAprovado1.Visible = False
            gb1.Visible = False
        Catch ex As Exception
            MsgBox("Erro 42 " & ex.Message)
        End Try
    End Sub
    Sub rb2f()
        Try
            cb2Turnos.Visible = False
            txtCaixas2Turno.Visible = False
            txtQtCaixasReprovada2.Visible = False
            lblQtPorTurno2.Visible = False
            txtCodigoRNC2.Visible = False
            txtDescricaoRNC2.Visible = False
            txtQTReprovado2.Visible = False
            txtQTAprovado2.Visible = False
            gb2.Visible = False
        Catch ex As Exception
            MsgBox("Erro 43 " & ex.Message)
        End Try
    End Sub
    Sub rb3f()
        Try
            cb3Turnos.Visible = False
            txtCaixas3Turno.Visible = False
            txtQtCaixasReprovada3.Visible = False
            lblQtPorTurno3.Visible = False
            txtCodigoRNC3.Visible = False
            txtDescricaoRNC3.Visible = False
            txtQTReprovado3.Visible = False
            txtQTAprovado3.Visible = False
            gb3.Visible = False
        Catch ex As Exception
            MsgBox("Erro 44 " & ex.Message)
        End Try
    End Sub
    Sub rb4f()
        Try
            cb4Turnos.Visible = False
            txtCaixas4Turno.Visible = False
            txtQtCaixasReprovada4.Visible = False
            lblQtPorTurno4.Visible = False
            txtCodigoRNC4.Visible = False
            txtDescricaoRNC4.Visible = False
            txtQTReprovado4.Visible = False
            txtQTAprovado4.Visible = False
            gb4.Visible = False
        Catch ex As Exception
            MsgBox("Erro 45 " & ex.Message)
        End Try
    End Sub
    Sub rb5f()
        Try
            cb5Turno.Visible = False
            txtCaixas5Turno.Visible = False
            txtQtCaixasReprovada5.Visible = False
            lblQtPorTurno5.Visible = False
            txtCodigoRNC5.Visible = False
            txtDescricaoRNC5.Visible = False
            txtQTReprovado5.Visible = False
            txtQTAprovado5.Visible = False
            gb5.Visible = False
        Catch ex As Exception
            MsgBox("Erro 46 " & ex.Message)
        End Try
    End Sub
    Sub rb6f()
        Try
            cb6Turno.Visible = False
            txtCaixas6Turno.Visible = False
            txtQtCaixasReprovada6.Visible = False
            lblQtPorTurno6.Visible = False
            txtCodigoRNC6.Visible = False
            txtDescricaoRNC6.Visible = False
            txtQTReprovado6.Visible = False
            txtQTAprovado6.Visible = False
            gb6.Visible = False
        Catch ex As Exception
            MsgBox("Erro 47 " & ex.Message)
        End Try
    End Sub
    Sub rb7f()
        Try
            cb7Turno.Visible = False
            txtCaixas7Turno.Visible = False
            txtQtCaixasReprovada7.Visible = False
            lblQtPorTurno7.Visible = False
            txtCodigoRNC7.Visible = False
            txtDescricaoRNC7.Visible = False
            txtQTReprovado7.Visible = False
            txtQTAprovado7.Visible = False
            gb7.Visible = False
        Catch ex As Exception
            MsgBox("Erro 48 " & ex.Message)
        End Try
    End Sub
    Sub rb8f()
        Try
            cb8Turno.Visible = False
            txtCaixas8Turno.Visible = False
            txtQtCaixasReprovada8.Visible = False
            lblQtPorTurno8.Visible = False
            txtCodigoRNC8.Visible = False
            txtDescricaoRNC8.Visible = False
            txtQTReprovado8.Visible = False
            txtQTAprovado8.Visible = False
            gb8.Visible = False
        Catch ex As Exception
            MsgBox("Erro 49 " & ex.Message)
        End Try
    End Sub
    Sub rb9f()
        Try
            cb9Turno.Visible = False
            txtCaixas9Turno.Visible = False
            txtQtCaixasReprovada9.Visible = False
            lblQtPorTurno9.Visible = False
            txtCodigoRNC9.Visible = False
            txtDescricaoRNC9.Visible = False
            txtQTReprovado9.Visible = False
            txtQTAprovado9.Visible = False
            gb9.Visible = False
        Catch ex As Exception
            MsgBox("Erro 50 " & ex.Message)
        End Try
    End Sub
    Sub rb10f()
        Try
            cb10Turno.Visible = False
            txtCaixas10Turno.Visible = False
            txtQtCaixasReprovada10.Visible = False
            lblQtPorTurno10.Visible = False
            txtCodigoRNC10.Visible = False
            txtDescricaoRNC10.Visible = False
            txtQTReprovado10.Visible = False
            txtQTAprovado10.Visible = False
            gb10.Visible = False
        Catch ex As Exception
            MsgBox("Erro 51 " & ex.Message)
        End Try
    End Sub
    Private Sub rb1T_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb1T.CheckedChanged, rb2T.CheckedChanged, rb3T.CheckedChanged, rb4T.CheckedChanged, rb5T.CheckedChanged, rb6T.CheckedChanged, rb7T.CheckedChanged, rb8T.CheckedChanged, rb9T.CheckedChanged, rb10T.CheckedChanged
        Try
            Call QuantidadeTurnos()
            Call Calcular()
        Catch ex As Exception
            MsgBox("Erro 52 " & ex.Message)
        End Try
    End Sub
    'Enter com a função de TAB, Altere a propriedade KeyPreview do formulário para True;
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
    Private Sub Quantidades(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCaixas1Turno.KeyPress, txtCaixas2Turno.KeyPress, txtCaixas3Turno.KeyPress, txtCaixas4Turno.KeyPress, txtCaixas5Turno.KeyPress, txtCaixas6Turno.KeyPress, txtCaixas7Turno.KeyPress, txtCaixas8Turno.KeyPress, txtCaixas9Turno.KeyPress, txtCaixas10Turno.KeyPress
        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(NumeroVirgSpace(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 54 " & ex.Message)
        End Try
    End Sub
    Function NumeroVirgSpace(ByVal Keyascii As Short) As Short

        If InStr("1234567890-atéecixsplt=,", Chr(Keyascii)) = 0 Then
            NumeroVirgSpace = 0
        Else
            NumeroVirgSpace = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                NumeroVirgSpace = Keyascii
            Case 13
                NumeroVirgSpace = Keyascii
            Case 32 'permite espaço
                NumeroVirgSpace = Keyascii
        End Select
    End Function
    Private Sub Quantidades2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoRNC1.KeyPress, txtCodigoRNC2.KeyPress, txtCodigoRNC3.KeyPress, txtCodigoRNC4.KeyPress, txtCodigoRNC5.KeyPress, txtCodigoRNC6.KeyPress, txtCodigoRNC7.KeyPress, txtCodigoRNC8.KeyPress, txtCodigoRNC9.KeyPress, txtCodigoRNC10.KeyPress, txtRE.KeyPress, txtOP.KeyPress, txtMaquina.KeyPress, txtQtCaixasReprovada1.KeyPress, txtQtCaixasReprovada2.KeyPress, txtQtCaixasReprovada3.KeyPress, txtQtCaixasReprovada4.KeyPress, txtQtCaixasReprovada5.KeyPress, txtQtCaixasReprovada6.KeyPress, txtQtCaixasReprovada7.KeyPress, txtQtCaixasReprovada8.KeyPress, txtQtCaixasReprovada9.KeyPress, txtQtCaixasReprovada10.KeyPress
        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(Numero(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 55 " & ex.Message)
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
    Private Sub Quantidades3(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPecasPorVolume.KeyPress, txtQTReprovado1.KeyPress, txtQTAprovado1.KeyPress, txtQTReprovado2.KeyPress, txtQTAprovado2.KeyPress, txtQTReprovado3.KeyPress, txtQTAprovado3.KeyPress, txtQTReprovado4.KeyPress, txtQTAprovado4.KeyPress, txtQTReprovado5.KeyPress, txtQTAprovado5.KeyPress, txtQTReprovado6.KeyPress, txtQTAprovado6.KeyPress, txtQTReprovado7.KeyPress, txtQTAprovado7.KeyPress, txtQTReprovado8.KeyPress, txtQTAprovado8.KeyPress, txtQTReprovado9.KeyPress, txtQTAprovado9.KeyPress, txtQTReprovado10.KeyPress, txtQTAprovado10.KeyPress
        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(NumeroVir(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 56 " & ex.Message)
        End Try
    End Sub
    Function NumeroVir(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            NumeroVir = 0
        Else
            NumeroVir = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                NumeroVir = Keyascii
            Case 13
                NumeroVir = Keyascii
                'Case 32 'permite espaço
                '   SoNumeros = Keyascii
        End Select
    End Function

    Private Sub txtQtCaixasReprovada1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQtCaixasReprovada1.TextChanged, txtQtCaixasReprovada2.TextChanged, txtQtCaixasReprovada3.TextChanged, txtQtCaixasReprovada4.TextChanged, txtQtCaixasReprovada5.TextChanged, txtQtCaixasReprovada6.TextChanged, txtQtCaixasReprovada7.TextChanged, txtQtCaixasReprovada8.TextChanged, txtQtCaixasReprovada9.TextChanged, txtQtCaixasReprovada10.TextChanged, txtPecasPorVolume.TextChanged
        Call Calcular()
    End Sub


    Sub Calcular()
        Try
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada1.TextLength > 0 Then
                lblQtPorTurno1.Text = Double.Parse(txtQtCaixasReprovada1.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno1.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada2.TextLength > 0 Then
                lblQtPorTurno2.Text = Double.Parse(txtQtCaixasReprovada2.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno2.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada3.TextLength > 0 Then
                lblQtPorTurno3.Text = Double.Parse(txtQtCaixasReprovada3.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno3.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada4.TextLength > 0 Then
                lblQtPorTurno4.Text = Double.Parse(txtQtCaixasReprovada4.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno4.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada5.TextLength > 0 Then
                lblQtPorTurno5.Text = Double.Parse(txtQtCaixasReprovada5.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno5.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada6.TextLength > 0 Then
                lblQtPorTurno6.Text = Double.Parse(txtQtCaixasReprovada6.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno6.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada7.TextLength > 0 Then
                lblQtPorTurno7.Text = Double.Parse(txtQtCaixasReprovada7.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno7.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada8.TextLength > 0 Then
                lblQtPorTurno8.Text = Double.Parse(txtQtCaixasReprovada8.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno8.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada9.TextLength > 0 Then
                lblQtPorTurno9.Text = Double.Parse(txtQtCaixasReprovada9.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno9.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada10.TextLength > 0 Then
                lblQtPorTurno10.Text = Double.Parse(txtQtCaixasReprovada10.Text * txtPecasPorVolume.Text)
            Else
                lblQtPorTurno10.Text = 0
            End If


            Dim x1 As Double
            Dim x2 As Double
            Dim x3 As Double
            Dim x4 As Double
            Dim x5 As Double
            Dim x6 As Double
            Dim x7 As Double
            Dim x8 As Double
            Dim x9 As Double
            Dim x10 As Double

            x1 = lblQtPorTurno1.Text
            x2 = lblQtPorTurno2.Text
            x3 = lblQtPorTurno3.Text
            x4 = lblQtPorTurno4.Text
            x5 = lblQtPorTurno5.Text
            x6 = lblQtPorTurno6.Text
            x7 = lblQtPorTurno7.Text
            x8 = lblQtPorTurno8.Text
            x9 = lblQtPorTurno9.Text
            x10 = lblQtPorTurno10.Text



            If rb1T.Checked = True Then
                lblTotalPecas.Text = x1
            ElseIf rb2T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2)
            ElseIf rb3T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3)
            ElseIf rb4T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4)
            ElseIf rb5T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4 + x5)
            ElseIf rb6T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4 + x5 + x6)
            ElseIf rb7T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4 + x5 + x6 + x7)
            ElseIf rb8T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4 + x5 + x6 + x7 + x8)
            ElseIf rb9T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4 + x5 + x6 + x7 + x8 + x9)
            ElseIf rb10T.Checked = True Then
                lblTotalPecas.Text = (x1 + x2 + x3 + x4 + x5 + x6 + x7 + x8 + x9 + x10)
            End If
        Catch ex As Exception
            MsgBox("Erro 57 " & ex.Message)
        End Try
    End Sub
    '27160
    Private Sub txtRE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRE.LostFocus
        Try
            TesteAbertoRE()
            Dim da13 As New OleDbDataAdapter
            Dim ds13 As New DataSet
            Dim re As Integer
            conRE.Open()
            If txtRE.TextLength = 0 Then
                conRE.Close()
            End If
            Dim sel6 As String = "SELECT Inspetor, Setor FROM tblRE where RE = " & txtRE.Text & " "
            da13 = New OleDbDataAdapter(sel6, conRE)
            ds13.Clear()
            da13.Fill(ds13, "tblRE")
            txtInspetor.Clear()
            re = ds13.Tables("tblRE").Rows.Count
            If re <= 0 Then
                conRE.Close()
                MsgBox("'RE' inexitente! Insira um RE válido ou se não possui digite 0 no campo RE e o campo Inspetor preencha com seu nome", , "RE")
                txtRE.Text = "0"
                txtInspetor.Enabled = True
                txtInspetor.Focus()
            Else
                txtInspetor.Text = ds13.Tables("tblRE").Rows(0)("Inspetor")
                txtSetor.Text = ds13.Tables("tblRE").Rows(0)("Setor")
                conRE.Close()
                txtInspetor.Enabled = False
                txtSetor.Enabled = False
            End If
        Catch ex As Exception
            MsgBox("Erro 58 " & ex.Message)
        End Try
    End Sub
    Sub codRNC1()
        Try
            TesteAbertoDefeito()
            Dim da15 As New OleDbDataAdapter
            Dim ds15 As New DataSet
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC1.Text & " "
            da15 = New OleDbDataAdapter(sel9, conDefeito)
            ds15.Clear()
            da15.Fill(ds15, "tblDefeitos")
            Defeito1 = ds15.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
            conDefeito.Close()
        Catch ex As Exception
            MsgBox("Erro 60 " & ex.Message)
        End Try
    End Sub
    Private Sub txtCodigoRNC1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC1.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC1.Text = MCod1
        txtDescricaoRNC1.Text = " - "
        Defeito1 = ""
        Defeito1 = MRNC1
        txtDescricaoRNC1.Focus()

    End Sub

    Private Sub txtCodigoRNC2_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC2.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC2.Text = MCod2
        txtDescricaoRNC2.Text = " - "
        Defeito2 = ""
        Defeito2 = MRNC2
        txtDescricaoRNC2.Focus()

    End Sub
    Private Sub txtCodigoRNC3_TextChanged_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC3.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC3.Text = MCod3
        txtDescricaoRNC3.Text = " - "
        Defeito3 = ""
        Defeito3 = MRNC3
        txtDescricaoRNC3.Focus()

    End Sub
    Private Sub txtCodigoRNC4_TextChanged_4(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC4.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC4.Text = MCod4
        txtDescricaoRNC4.Text = " - "
        Defeito4 = ""
        Defeito4 = MRNC4
        txtDescricaoRNC4.Focus()

    End Sub
    Private Sub txtCodigoRNC5_TextChanged_5(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC5.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC5.Text = MCod5
        txtDescricaoRNC5.Text = " - "
        Defeito5 = ""
        Defeito5 = MRNC5
        txtDescricaoRNC5.Focus()

    End Sub
    Private Sub txtCodigoRNC6_TextChanged_6(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC6.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC6.Text = MCod6
        txtDescricaoRNC6.Text = " - "
        Defeito6 = ""
        Defeito6 = MRNC6
        txtDescricaoRNC6.Focus()

    End Sub
    Private Sub txtCodigoRNC7_TextChanged_7(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC7.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC7.Text = MCod7
        txtDescricaoRNC7.Text = " - "
        Defeito7 = ""
        Defeito7 = MRNC7
        txtDescricaoRNC7.Focus()

    End Sub
    Private Sub txtCodigoRNC8_TextChanged_8(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC8.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC8.Text = MCod8
        txtDescricaoRNC8.Text = " - "
        Defeito8 = ""
        Defeito8 = MRNC8
        txtDescricaoRNC8.Focus()

    End Sub
    Private Sub txtCodigoRNC9_TextChanged_9(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC9.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC9.Text = MCod9
        txtDescricaoRNC9.Text = " - "
        Defeito9 = ""
        Defeito9 = MRNC9
        txtDescricaoRNC9.Focus()

    End Sub
    Private Sub txtCodigoRNC10_TextChanged_10(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC10.MouseClick
        frmCodigoConsulta.ShowDialog()
        txtCodigoRNC10.Text = MCod10
        txtDescricaoRNC10.Text = " - "
        Defeito10 = ""
        Defeito10 = MRNC10
        txtDescricaoRNC10.Focus()

    End Sub



    Private Sub btCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelar.Click
        Call Limpar()
        LimparDisposicao()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Try
            TesteAbertoRNC()

            Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

            Dim ID = row.Cells(0)

            Dim RNC = row.Cells(1)
            If lblRNC.Text = "*" Then
                LimparDisposicao()
                Me.lblRNC.Text = RNC.Value
                AlterarCarregar2()
                AlterarCarregar()
            ElseIf RNC.Value = lblRNC.Text Then
            Else
                Me.lblRNC.Text = RNC.Value
                LimparDisposicao()
                AlterarCarregar2()
                AlterarCarregar()
            End If
            Dim Status = row.Cells(2)
            Dim Origem = row.Cells(3)
            Dim Data_Abertura = row.Cells(4)
            Dim Hora = row.Cells(5)
            Dim Cliente = row.Cells(8)
            Dim OP_Reprovado = row.Cells(10)
            Dim Maquina = row.Cells(17)
            Dim L1x = row.Cells(19)
            Dim OPRetrabalho = row.Cells(20)
            Dim Observacao = row.Cells(25)
            Dim RE = row.Cells(26)
            Dim Inspetor = row.Cells(27)
            Dim Setor = row.Cells(28)
            Dim TurnoDetector = row.Cells(29)
            Dim Alterado = row.Cells(30)

            Me.L1 = L1x.Value.ToString()
            Me.lblStatus.Text = Status.Value
            If OPRetrabalho.Value Is DBNull.Value Then
                txtOPRetrabalho.Text = ""
            Else
                Me.txtOPRetrabalho.Text = OPRetrabalho.Value
            End If
            Me.lblID.Text = ID.Value
            Me.txtOP.Text = OP_Reprovado.Value
            Me.cbDetectado.Text = Origem.Value
            Me.lblData.Text = Data_Abertura.Value
            Me.lblHora.Text = Hora.Value

            Me.txtMaquina.Text = Maquina.Value
            Me.txtOBS.Text = Observacao.Value
            Me.txtRE.Text = RE.Value
            Me.txtInspetor.Text = Inspetor.Value
            Me.txtSetor.Text = Setor.Value
            Me.cbTurno.Text = TurnoDetector.Value
            If Alterado.Value Is DBNull.Value Then
                Me.Alteradu = ""
            Else
                Me.Alteradu = Alterado.Value
            End If

            If Status.Value = "Pendente" Then

            End If
            ContarCaixas()

            If rbRT1.Checked = True Or rbRT2.Checked = True Or rbRT3.Checked = True Or rbRT4.Checked = True Or rbRT5.Checked = True Or rbRT6.Checked = True Or rbRT7.Checked = True Or rbRT8.Checked = True Or rbRT9.Checked = True Or rbRT10.Checked = True Then
                txtOPRetrabalho.Clear()
                txtOPRetrabalho.Enabled = True
            Else
                txtOPRetrabalho.Clear()
                txtOPRetrabalho.Enabled = False
                txtOPRetrabalho.Text = txtOP.Text
            End If

            txtOP.Focus()
            txtOBS.Focus()

        Catch e1 As Exception
            MessageBox.Show("Erro 3x!", e1.Message)
        End Try
    End Sub

    Private Sub btPesquisa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPesquisa.Click
        Try
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            Dim Linhas As Integer
            Dim seleccion As String
            If cbColuna.Text = "" Or txtDadoColuna.TextLength = 0 Or txtLinhas.TextLength = 0 Or cbOrdenadoPor.Text = "" Then
                MsgBox("Há Campos de Pesquisa Vazio")
            Else 'If cbColuna.Text = "RNC" Or cbColuna.Text = "Origem" Or cbColuna.Text = "Data_Abertura" Or cbColuna.Text = "Cod_Produto" Or cbColuna.Text = "Produto" Or cbColuna.Text = "OP_Reprovado" Or cbColuna.Text = "Turno" Or cbColuna.Text = "NúmerosCaixas" Or cbColuna.Text = "QT_Caixas" Or cbColuna.Text = "QT_P_Caixa" Or cbColuna.Text = "QT_Reprovado" Or cbColuna.Text = "Cod_Defeito" Or cbColuna.Text = "Nao_Conformidade" Or cbColuna.Text = "Maquina" Or cbColuna.Text = "Observacao" Or cbColuna.Text = "RE" Or cbColuna.Text = "Inspetor" Then
                seleccion = txtDadoColuna.Text
                seleccion = "%" & seleccion & "%"
                Linhas = txtLinhas.Text
                'Linhas = "" & Linhas & ""
                DataGridView1.DataSource.clear()
                conRNC.Open()
                Dim sel_ As String = "SELECT TOP " & Linhas & " * FROM tblRNC WHERE " & cbColuna.Text & " LIKE '" & seleccion & "' ORDER BY " & cbOrdenadoPor.Text & " DESC "
                da19 = New OleDbDataAdapter(sel_, conRNC)
                ds19.Clear()
                da19.Fill(ds19, "tblRNC")
                Me.DataGridView1.DataSource = ds19
                Me.DataGridView1.DataMember = "tblRNC"
                FormatacaoGrid()
                conRNC.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 71 " & ex.Message)
        End Try
    End Sub

    Private Sub btAlterar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterar.Click
        Try
            TesteAbertoRNC()
            If lblRNC.Text = "*" Or lblRNC.Text = "" Then
                MsgBox("Selecione um RNC na tabela abaixo", , "Selecione uma RNC")
            Else
                If btAlterar.Text = "Alterar" Then
                    AlterarCarregar()

                    If MsgBox("Deseja realmente 'Alterar' uma RNC?", vbYesNo, "Alterar RNC") = vbYes Then
                        'Call Limpar()
                        btAlterar.Text = "Aplicar"
                        btInserir.Enabled = False
                        btExcluir.Enabled = False
                        btImprimir.Enabled = False
                        btImprimirEtiqueta.Enabled = False
                        rb1T.Enabled = False
                        rb2T.Enabled = False
                        rb3T.Enabled = False
                        rb4T.Enabled = False
                        rb5T.Enabled = False
                        rb6T.Enabled = False
                        rb7T.Enabled = False
                        rb8T.Enabled = False
                        rb9T.Enabled = False
                        rb10T.Enabled = False
                        DataGridView1.Enabled = False
                        btEmail.Enabled = False
                    Else
                    End If

                Else
                    ContarCaixas()

                    If (MsgBox("O Total de caixas que você está reprovando é " & SMC & " ?", vbYesNo, "Confirmação de Quantidade de Caixas!!") = vbYes) Then

                        'radiobuton1

                        If rb1T.Checked = True Then
                            Call Verificacao1()

                            'radiobutton 2
                        ElseIf rb2T.Checked = True Then
                            Call Verificacao2()

                            'radiobutto 3
                        ElseIf rb3T.Checked = True Then
                            Call Verificacao3()
                            'radiobutton 4
                        ElseIf rb4T.Checked = True Then
                            Call Verificacao4()
                        ElseIf rb5T.Checked = True Then
                            Call Verificacao5()

                            'radiobutton 2
                        ElseIf rb6T.Checked = True Then
                            Call Verificacao6()

                            'radiobutto 3
                        ElseIf rb7T.Checked = True Then
                            Call Verificacao7()
                            'radiobutton 4
                        ElseIf rb8T.Checked = True Then
                            Call Verificacao8()

                        ElseIf rb9T.Checked = True Then
                            Call Verificacao9()

                            'radiobutto 3
                        ElseIf rb10T.Checked = True Then
                            Call Verificacao10()
                            'radiobutton 4
                        Else
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 72 " & ex.Message)
        End Try
    End Sub
    Sub ContarCaixas()


        Dim CAIXA1 As Int32 = 0
        Dim CAIXA2 As Int32 = 0
        Dim CAIXA3 As Int32 = 0
        Dim CAIXA4 As Int32 = 0
        Dim CAIXA5 As Int32 = 0
        Dim CAIXA6 As Int32 = 0
        Dim CAIXA7 As Int32 = 0
        Dim CAIXA8 As Int32 = 0
        Dim CAIXA9 As Int32 = 0
        Dim CAIXA10 As Int32 = 0



        CAIXA1 = txtQtCaixasReprovada1.Text
        If rb2T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
        End If
        If rb3T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
        End If
        If rb4T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
        End If
        If rb5T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
            CAIXA5 = txtQtCaixasReprovada5.Text
        End If
        If rb6T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
            CAIXA5 = txtQtCaixasReprovada5.Text
            CAIXA6 = txtQtCaixasReprovada6.Text
        End If
        If rb7T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
            CAIXA5 = txtQtCaixasReprovada5.Text
            CAIXA6 = txtQtCaixasReprovada6.Text
            CAIXA7 = txtQtCaixasReprovada7.Text
        End If
        If rb8T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
            CAIXA5 = txtQtCaixasReprovada5.Text
            CAIXA6 = txtQtCaixasReprovada6.Text
            CAIXA7 = txtQtCaixasReprovada7.Text
            CAIXA8 = txtQtCaixasReprovada8.Text
        End If
        If rb9T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
            CAIXA5 = txtQtCaixasReprovada5.Text
            CAIXA6 = txtQtCaixasReprovada6.Text
            CAIXA7 = txtQtCaixasReprovada7.Text
            CAIXA8 = txtQtCaixasReprovada8.Text
            CAIXA9 = txtQtCaixasReprovada9.Text
        End If
        If rb10T.Checked = True Then
            CAIXA2 = txtQtCaixasReprovada2.Text
            CAIXA3 = txtQtCaixasReprovada3.Text
            CAIXA4 = txtQtCaixasReprovada4.Text
            CAIXA5 = txtQtCaixasReprovada5.Text
            CAIXA6 = txtQtCaixasReprovada6.Text
            CAIXA7 = txtQtCaixasReprovada7.Text
            CAIXA8 = txtQtCaixasReprovada8.Text
            CAIXA9 = txtQtCaixasReprovada9.Text
            CAIXA10 = txtQtCaixasReprovada10.Text
        End If
        SMC = 0
        SMC = CAIXA1 + CAIXA2 + CAIXA3 + CAIXA4 + CAIXA5 + CAIXA6 + CAIXA7 + CAIXA8 + CAIXA9 + CAIXA10
    End Sub
    Sub Alterar1()
        Try
            Call codRNC1()
            Call Celulax()
            conRNC.Open()
            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet
            ds20 = New DataSet
            da20 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb1Turno.Text & "', NúmerosCaixas = '" & txtCaixas1Turno.Text & "', 
QT_Caixas = " & txtQtCaixasReprovada1.Text & ", QT_Reprovado = '" & lblQtPorTurno1.Text & "', Cod_Defeito = " & txtCodigoRNC1.Text & ", Nao_Conformidade = '" & txtDescricaoRNC1.Text & "', 
Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', 
TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID1 & "", conRNC)
            ds20.Clear()
            da20.Fill(ds20, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 73 " & ex.Message)
        End Try
    End Sub
    Sub Alterar2()
        Try
            Alterar1()
            Dim da20_2 As New OleDbDataAdapter
            Dim ds20_2 As New DataSet
            ds20_2 = New DataSet
            da20_2 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb2Turnos.Text & "', NúmerosCaixas = '" & txtCaixas2Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada2.Text & ", QT_Reprovado = '" & lblQtPorTurno2.Text & "', Cod_Defeito = " & txtCodigoRNC2.Text & ", Nao_Conformidade = '" & txtDescricaoRNC2.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID2 & "", conRNC)
            ds20_2.Clear()
            da20_2.Fill(ds20_2, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 74 " & ex.Message)
        End Try
    End Sub
    Sub Alterar3()
        Try
            Alterar2()
            Dim da20_3 As New OleDbDataAdapter
            Dim ds20_3 As New DataSet
            ds20_3 = New DataSet
            da20_3 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb3Turnos.Text & "', NúmerosCaixas = '" & txtCaixas3Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada3.Text & ", QT_Reprovado = '" & lblQtPorTurno3.Text & "', Cod_Defeito = " & txtCodigoRNC3.Text & ", Nao_Conformidade = '" & txtDescricaoRNC3.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID3 & "", conRNC)
            ds20_3.Clear()
            da20_3.Fill(ds20_3, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 75 " & ex.Message)
        End Try
    End Sub
    Sub Alterar4()
        Try
            Alterar3()
            Dim da20_4 As New OleDbDataAdapter
            Dim ds20_4 As New DataSet
            ds20_4 = New DataSet
            da20_4 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb4Turnos.Text & "', NúmerosCaixas = '" & txtCaixas4Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada4.Text & ", QT_Reprovado = '" & lblQtPorTurno4.Text & "', Cod_Defeito = " & txtCodigoRNC4.Text & ", Nao_Conformidade = '" & txtDescricaoRNC4.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID4 & "", conRNC)
            ds20_4.Clear()
            da20_4.Fill(ds20_4, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 76 " & ex.Message)
        End Try
    End Sub
    Sub Alterar5()
        Try
            Alterar4()
            Dim da20_5 As New OleDbDataAdapter
            Dim ds20_5 As New DataSet
            ds20_5 = New DataSet
            da20_5 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb5Turno.Text & "', NúmerosCaixas = '" & txtCaixas5Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada5.Text & ", QT_Reprovado = '" & lblQtPorTurno5.Text & "', Cod_Defeito = " & txtCodigoRNC5.Text & ", Nao_Conformidade = '" & txtDescricaoRNC5.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID5 & "", conRNC)
            ds20_5.Clear()
            da20_5.Fill(ds20_5, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 77 " & ex.Message)
        End Try
    End Sub
    Sub Alterar6()
        Try
            Alterar5()
            Dim da20_6 As New OleDbDataAdapter
            Dim ds20_6 As New DataSet
            ds20_6 = New DataSet
            da20_6 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb6Turno.Text & "', NúmerosCaixas = '" & txtCaixas6Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada6.Text & ", QT_Reprovado = '" & lblQtPorTurno6.Text & "', Cod_Defeito = " & txtCodigoRNC6.Text & ", Nao_Conformidade = '" & txtDescricaoRNC6.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID6 & "", conRNC)
            ds20_6.Clear()
            da20_6.Fill(ds20_6, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 78 " & ex.Message)
        End Try
    End Sub
    Sub Alterar7()
        Try
            Alterar6()
            Dim da20_7 As New OleDbDataAdapter
            Dim ds20_7 As New DataSet
            ds20_7 = New DataSet
            da20_7 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb7Turno.Text & "', NúmerosCaixas = '" & txtCaixas7Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada7.Text & ", QT_Reprovado = '" & lblQtPorTurno7.Text & "', Cod_Defeito = " & txtCodigoRNC7.Text & ", Nao_Conformidade = '" & txtDescricaoRNC7.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID7 & "", conRNC)
            ds20_7.Clear()
            da20_7.Fill(ds20_7, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 79 " & ex.Message)
        End Try
    End Sub
    Sub Alterar8()
        Try
            Alterar7()
            Dim da20_8 As New OleDbDataAdapter
            Dim ds20_8 As New DataSet
            ds20_8 = New DataSet
            da20_8 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb8Turno.Text & "', NúmerosCaixas = '" & txtCaixas8Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada8.Text & ", QT_Reprovado = '" & lblQtPorTurno8.Text & "', Cod_Defeito = " & txtCodigoRNC8.Text & ", Nao_Conformidade = '" & txtDescricaoRNC8.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID8 & "", conRNC)
            ds20_8.Clear()
            da20_8.Fill(ds20_8, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 80 " & ex.Message)
        End Try
    End Sub
    Sub Alterar9()
        Try
            Alterar8()
            Dim da20_9 As New OleDbDataAdapter
            Dim ds20_9 As New DataSet
            ds20_9 = New DataSet
            da20_9 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb9Turno.Text & "', NúmerosCaixas = '" & txtCaixas9Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada9.Text & ", QT_Reprovado = '" & lblQtPorTurno9.Text & "', Cod_Defeito = " & txtCodigoRNC9.Text & ", Nao_Conformidade = '" & txtDescricaoRNC9.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID9 & "", conRNC)
            ds20_9.Clear()
            da20_9.Fill(ds20_9, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 81 " & ex.Message)
        End Try
    End Sub
    Sub Alterar10()
        Try
            Alterar9()
            Dim da20_10 As New OleDbDataAdapter
            Dim ds20_10 As New DataSet
            ds20_10 = New DataSet
            da20_10 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb10Turno.Text & "', NúmerosCaixas = '" & txtCaixas10Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada10.Text & ", QT_Reprovado = '" & lblQtPorTurno10.Text & "', Cod_Defeito = " & txtCodigoRNC10.Text & ", Nao_Conformidade = '" & txtDescricaoRNC10.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID10 & "", conRNC)
            ds20_10.Clear()
            da20_10.Fill(ds20_10, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 82 " & ex.Message)
        End Try
    End Sub

    Sub AlterarCarregar2()
        Try
            'apos o status e disposições


            conRNC.Open()
            Dim selPRINT As String = "SELECT top 10 RNC, OP_Retrabalho FROM tblRNC where RNC = " & lblRNC.Text & " order by OP_Retrabalho desc"
            Dim daPRINT As New OleDbDataAdapter
            Dim dsPRINT As New DataSet
            Dim dtPrint As New DataTable
            daPRINT = New OleDbDataAdapter(selPRINT, conRNC)
            dsPRINT.Clear()
            daPRINT.Fill(dsPRINT, "tblRNC")

            If dsPRINT.Tables("tblRNC").Rows(0)("OP_Retrabalho") Is DBNull.Value Then
                txtOPRetrabalho.Clear()
            Else
                txtOPRetrabalho.Text = dsPRINT.Tables("tblRNC").Rows(0)("OP_Retrabalho").ToString()

            End If

            Dim selPRINT2 As String = "SELECT top 10 RNC, OP_Retrabalho, Status, Data_Encerramento FROM tblRNC where RNC = " & lblRNC.Text & " order by Status desc"
            daPRINT = New OleDbDataAdapter(selPRINT2, conRNC)
            dsPRINT.Clear()
            daPRINT.Fill(dsPRINT, "tblRNC")
            conRNC.Close()
            lblStatus.Text = dsPRINT.Tables("tblRNC").Rows(0)("Status")
            If dsPRINT.Tables("tblRNC").Rows(0)("Status") = "Pendente" Then
                lblDataEncerramento.Text = "*"
            Else
                lblDataEncerramento.Text = dsPRINT.Tables("tblRNC").Rows(0)("Data_Encerramento")
            End If

        Catch ex As Exception
            MsgBox("Erro 70 " & ex.Message)
            conRNC.Close()
        End Try
    End Sub
    Sub limpar2()

        cb2Turnos.Text = ""
        txtCaixas2Turno.Clear()
        txtQtCaixasReprovada2.Clear()
        lblQtPorTurno2.Text = ""
        txtCodigoRNC2.Clear()
        txtDescricaoRNC2.Clear()

        cb3Turnos.Text = ""
        txtCaixas3Turno.Clear()
        txtQtCaixasReprovada3.Clear()
        lblQtPorTurno3.Text = ""
        txtCodigoRNC3.Clear()
        txtDescricaoRNC3.Clear()

        cb4Turnos.Text = ""
        txtCaixas4Turno.Clear()
        txtQtCaixasReprovada4.Clear()
        lblQtPorTurno4.Text = ""
        txtCodigoRNC4.Clear()
        txtDescricaoRNC4.Clear()

        cb5Turno.Text = ""
        txtCaixas5Turno.Clear()
        txtQtCaixasReprovada5.Clear()
        lblQtPorTurno5.Text = ""
        txtCodigoRNC5.Clear()
        txtDescricaoRNC5.Clear()

        cb6Turno.Text = ""
        txtCaixas6Turno.Clear()
        txtQtCaixasReprovada6.Clear()
        lblQtPorTurno6.Text = ""
        txtCodigoRNC6.Clear()
        txtDescricaoRNC6.Clear()

        cb7Turno.Text = ""
        txtCaixas7Turno.Clear()
        txtQtCaixasReprovada7.Clear()
        lblQtPorTurno7.Text = ""
        txtCodigoRNC7.Clear()
        txtDescricaoRNC7.Clear()

        cb8Turno.Text = ""
        txtCaixas8Turno.Clear()
        txtQtCaixasReprovada8.Clear()
        lblQtPorTurno8.Text = ""
        txtCodigoRNC8.Clear()
        txtDescricaoRNC8.Clear()

        cb9Turno.Text = ""
        txtCaixas9Turno.Clear()
        txtQtCaixasReprovada9.Clear()
        lblQtPorTurno9.Text = ""
        txtCodigoRNC9.Clear()
        txtDescricaoRNC9.Clear()

        cb10Turno.Text = ""
        txtCaixas10Turno.Clear()
        txtQtCaixasReprovada10.Clear()
        lblQtPorTurno10.Text = ""
        txtCodigoRNC10.Clear()
        txtDescricaoRNC10.Clear()


    End Sub

    Private Sub AlterarCarregar()
        Try
            conRNC.Open()
            Dim selPRINT As String = "SELECT top 10 * FROM tblRNC where RNC = " & lblRNC.Text & " order by ID asc"
            Dim daPRINT As New OleDbDataAdapter
            Dim dsPRINT As New DataSet
            Dim dtPrint As New DataTable
            daPRINT = New OleDbDataAdapter(selPRINT, conRNC)
            dsPRINT.Clear()
            dtPrint.Clear()
            daPRINT.Fill(dsPRINT, "tblRNC")
            daPRINT.Fill(dtPrint)
            conRNC.Close()

            rb1T.Checked = True

            ID1 = dsPRINT.Tables("tblRNC").Rows(0)("ID")
            cb1Turno.Text = dsPRINT.Tables("tblRNC").Rows(0)("Turno")
            txtCaixas1Turno.Text = dsPRINT.Tables("tblRNC").Rows(0)("NúmerosCaixas")
            txtQtCaixasReprovada1.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_Caixas")
            lblQtPorTurno1.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_Reprovado")
            txtCodigoRNC1.Text = dsPRINT.Tables("tblRNC").Rows(0)("Cod_Defeito")
            txtDescricaoRNC1.Text = dsPRINT.Tables("tblRNC").Rows(0)("Nao_Conformidade")

            Try
                If dsPRINT.Tables("tblRNC").Rows(0)("Disposicao").ToString() = "Sem Disposição" Then
                    txtQTReprovado1.Text = ""
                    txtQTAprovado1.Text = ""
                ElseIf dsPRINT.Tables("tblRNC").Rows(0)("Disposicao").ToString() = "Retrabalhar" Then
                    rbRT1.Checked = True
                    If dsPRINT.Tables("tblRNC").Rows(0)("QT_ReprovadoR") Is DBNull.Value Then
                        txtQTReprovado1.Text = ""
                        txtQTAprovado1.Text = ""
                    Else
                        txtQTReprovado1.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_ReprovadoR").ToString()
                        txtQTAprovado1.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_AprovadoR").ToString()
                    End If
                ElseIf dsPRINT.Tables("tblRNC").Rows(0)("Disposicao").ToString() = "Refugar" Then
                    rbRF1.Checked = True
                Else
                    rbLC1.Checked = True
                End If
            Catch ex As Exception
                MsgBox("Erro oo " & ex.Message)
            End Try

            limpar2()
            If dtPrint.Rows.Count >= 2 Then
                rb2T.Checked = True
                ID2 = dsPRINT.Tables("tblRNC").Rows(1)("ID")
                cb2Turnos.Text = dsPRINT.Tables("tblRNC").Rows(1)("Turno")
                txtCaixas2Turno.Text = dsPRINT.Tables("tblRNC").Rows(1)("NúmerosCaixas")
                txtQtCaixasReprovada2.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_Caixas")
                lblQtPorTurno2.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_Reprovado")
                txtCodigoRNC2.Text = dsPRINT.Tables("tblRNC").Rows(1)("Cod_Defeito")
                txtDescricaoRNC2.Text = dsPRINT.Tables("tblRNC").Rows(1)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(1)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado2.Text = ""
                        txtQTAprovado2.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(1)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT2.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(1)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado2.Text = ""
                            txtQTAprovado2.Text = ""
                        Else
                            txtQTReprovado2.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_ReprovadoR").ToString()
                            txtQTAprovado2.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(1)("Disposicao").ToString() = "Refugar" Then
                        rbRF2.Checked = True
                    Else
                        rbLC2.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro pp " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 3 Then
                ID3 = dsPRINT.Tables("tblRNC").Rows(2)("ID")
                rb3T.Checked = True
                cb3Turnos.Text = dsPRINT.Tables("tblRNC").Rows(2)("Turno")
                txtCaixas3Turno.Text = dsPRINT.Tables("tblRNC").Rows(2)("NúmerosCaixas")
                txtQtCaixasReprovada3.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_Caixas")
                lblQtPorTurno3.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_Reprovado")
                txtCodigoRNC3.Text = dsPRINT.Tables("tblRNC").Rows(2)("Cod_Defeito")
                txtDescricaoRNC3.Text = dsPRINT.Tables("tblRNC").Rows(2)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(2)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado3.Text = ""
                        txtQTAprovado3.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(2)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT3.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(2)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado3.Text = ""
                            txtQTAprovado3.Text = ""
                        Else
                            txtQTReprovado3.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_ReprovadoR").ToString()
                            txtQTAprovado3.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(2)("Disposicao").ToString() = "Refugar" Then
                        rbRF3.Checked = True
                    Else
                        rbLC3.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro hh " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 4 Then
                ID4 = dsPRINT.Tables("tblRNC").Rows(3)("ID")
                rb4T.Checked = True
                cb4Turnos.Text = dsPRINT.Tables("tblRNC").Rows(3)("Turno")
                txtCaixas4Turno.Text = dsPRINT.Tables("tblRNC").Rows(3)("NúmerosCaixas")
                txtQtCaixasReprovada4.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_Caixas")
                lblQtPorTurno4.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_Reprovado")
                txtCodigoRNC4.Text = dsPRINT.Tables("tblRNC").Rows(3)("Cod_Defeito")
                txtDescricaoRNC4.Text = dsPRINT.Tables("tblRNC").Rows(3)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(3)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado4.Text = ""
                        txtQTAprovado4.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(3)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT4.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(3)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado4.Text = ""
                            txtQTAprovado4.Text = ""
                        Else
                            txtQTReprovado4.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_ReprovadoR").ToString()
                            txtQTAprovado4.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(3)("Disposicao").ToString() = "Refugar" Then
                        rbRF4.Checked = True
                    Else
                        rbLC4.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro ff " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 5 Then
                ID5 = dsPRINT.Tables("tblRNC").Rows(4)("ID")
                rb5T.Checked = True
                cb5Turno.Text = dsPRINT.Tables("tblRNC").Rows(4)("Turno")
                txtCaixas5Turno.Text = dsPRINT.Tables("tblRNC").Rows(4)("NúmerosCaixas")
                txtQtCaixasReprovada5.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_Caixas")
                lblQtPorTurno5.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_Reprovado")
                txtCodigoRNC5.Text = dsPRINT.Tables("tblRNC").Rows(4)("Cod_Defeito")
                txtDescricaoRNC5.Text = dsPRINT.Tables("tblRNC").Rows(4)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(4)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado5.Text = ""
                        txtQTAprovado5.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(4)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT5.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(4)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado5.Text = ""
                            txtQTAprovado5.Text = ""
                        Else
                            txtQTReprovado5.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_ReprovadoR").ToString()
                            txtQTAprovado5.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(4)("Disposicao").ToString() = "Refugar" Then
                        rbRF5.Checked = True
                    Else
                        rbLC5.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro ee " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 6 Then
                ID6 = dsPRINT.Tables("tblRNC").Rows(5)("ID")
                rb6T.Checked = True
                cb6Turno.Text = dsPRINT.Tables("tblRNC").Rows(5)("Turno")
                txtCaixas6Turno.Text = dsPRINT.Tables("tblRNC").Rows(5)("NúmerosCaixas")
                txtQtCaixasReprovada6.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_Caixas")
                lblQtPorTurno6.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_Reprovado")
                txtCodigoRNC6.Text = dsPRINT.Tables("tblRNC").Rows(5)("Cod_Defeito")
                txtDescricaoRNC6.Text = dsPRINT.Tables("tblRNC").Rows(5)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(5)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado6.Text = ""
                        txtQTAprovado6.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(5)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT6.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(5)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado6.Text = ""
                            txtQTAprovado6.Text = ""
                        Else
                            txtQTReprovado6.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_ReprovadoR").ToString()
                            txtQTAprovado6.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(5)("Disposicao").ToString() = "Refugar" Then
                        rbRF6.Checked = True
                    Else
                        rbLC6.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro dd " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 7 Then
                ID7 = dsPRINT.Tables("tblRNC").Rows(6)("ID")
                rb7T.Checked = True
                cb7Turno.Text = dsPRINT.Tables("tblRNC").Rows(6)("Turno")
                txtCaixas7Turno.Text = dsPRINT.Tables("tblRNC").Rows(6)("NúmerosCaixas")
                txtQtCaixasReprovada7.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_Caixas")
                lblQtPorTurno7.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_Reprovado")
                txtCodigoRNC7.Text = dsPRINT.Tables("tblRNC").Rows(6)("Cod_Defeito")
                txtDescricaoRNC7.Text = dsPRINT.Tables("tblRNC").Rows(6)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(6)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado7.Text = ""
                        txtQTAprovado7.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(6)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT7.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(6)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado7.Text = ""
                            txtQTAprovado7.Text = ""
                        Else
                            txtQTReprovado7.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_ReprovadoR").ToString()
                            txtQTAprovado7.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(6)("Disposicao").ToString() = "Refugar" Then
                        rbRF7.Checked = True
                    Else
                        rbLC7.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro cc " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 8 Then
                ID8 = dsPRINT.Tables("tblRNC").Rows(7)("ID")
                rb8T.Checked = True
                cb8Turno.Text = dsPRINT.Tables("tblRNC").Rows(7)("Turno")
                txtCaixas8Turno.Text = dsPRINT.Tables("tblRNC").Rows(7)("NúmerosCaixas")
                txtQtCaixasReprovada8.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_Caixas")
                lblQtPorTurno8.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_Reprovado")
                txtCodigoRNC8.Text = dsPRINT.Tables("tblRNC").Rows(7)("Cod_Defeito")
                txtDescricaoRNC8.Text = dsPRINT.Tables("tblRNC").Rows(7)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(7)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado8.Text = ""
                        txtQTAprovado8.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(7)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT8.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(7)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado8.Text = ""
                            txtQTAprovado8.Text = ""
                        Else
                            txtQTReprovado8.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_ReprovadoR").ToString()
                            txtQTAprovado8.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(7)("Disposicao").ToString() = "Refugar" Then
                        rbRF8.Checked = True
                    Else
                        rbLC8.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro bb " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count >= 9 Then
                ID9 = dsPRINT.Tables("tblRNC").Rows(8)("ID")
                rb9T.Checked = True
                cb9Turno.Text = dsPRINT.Tables("tblRNC").Rows(8)("Turno")
                txtCaixas9Turno.Text = dsPRINT.Tables("tblRNC").Rows(8)("NúmerosCaixas")
                txtQtCaixasReprovada9.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_Caixas")
                lblQtPorTurno9.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_Reprovado")
                txtCodigoRNC9.Text = dsPRINT.Tables("tblRNC").Rows(8)("Cod_Defeito")
                txtDescricaoRNC9.Text = dsPRINT.Tables("tblRNC").Rows(8)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(8)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado9.Text = ""
                        txtQTAprovado9.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(8)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT9.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(8)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado9.Text = ""
                            txtQTAprovado9.Text = ""
                        Else
                            txtQTReprovado9.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_ReprovadoR").ToString()
                            txtQTAprovado9.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(8)("Disposicao").ToString() = "Refugar" Then
                        rbRF9.Checked = True
                    Else
                        rbLC9.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox("Erro aa " & ex.Message)
                End Try
            End If
            If dtPrint.Rows.Count = 10 Then
                ID10 = dsPRINT.Tables("tblRNC").Rows(9)("ID")
                rb10T.Checked = True
                cb10Turno.Text = dsPRINT.Tables("tblRNC").Rows(9)("Turno")
                txtCaixas10Turno.Text = dsPRINT.Tables("tblRNC").Rows(9)("NúmerosCaixas")
                txtQtCaixasReprovada10.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_Caixas")
                lblQtPorTurno10.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_Reprovado")
                txtCodigoRNC10.Text = dsPRINT.Tables("tblRNC").Rows(9)("Cod_Defeito")
                txtDescricaoRNC10.Text = dsPRINT.Tables("tblRNC").Rows(9)("Nao_Conformidade")
                Try
                    If dsPRINT.Tables("tblRNC").Rows(9)("Disposicao").ToString() = "Sem Disposição" Then
                        txtQTReprovado10.Text = ""
                        txtQTAprovado10.Text = ""
                    ElseIf dsPRINT.Tables("tblRNC").Rows(9)("Disposicao").ToString() = "Retrabalhar" Then
                        rbRT10.Checked = True
                        If dsPRINT.Tables("tblRNC").Rows(9)("QT_ReprovadoR") Is DBNull.Value Then
                            txtQTReprovado10.Text = ""
                            txtQTAprovado10.Text = ""
                        Else
                            txtQTReprovado10.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_ReprovadoR").ToString()
                            txtQTAprovado10.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_AprovadoR").ToString()
                        End If
                    ElseIf dsPRINT.Tables("tblRNC").Rows(9)("Disposicao").ToString() = "Refugar" Then
                        rbRF10.Checked = True
                    Else
                        rbLC10.Checked = True
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        Catch ex As Exception
            MsgBox("Erro 83 " & ex.Message)
        End Try
    End Sub

    Private Sub btExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExcluir.Click
        Try
            TesteAbertoRNC()
            If lblRNC.Text = "*" Or lblRNC.Text = "" Then
                MsgBox("Selecione um RNC na tabela abaixo", , "Selecione uma RNC")
            Else
                Dim da21 As New OleDbDataAdapter
                Dim ds21 As New DataSet
                If btExcluir.Text = "Excluir" Then
                    If MsgBox("Deseja Excluir uma RNC?", vbYesNo, "Excluir RNC") = vbYes Then
                        'Call Limpar()
                        btExcluir.Text = "Aplicar"
                        btInserir.Enabled = False
                        btAlterar.Enabled = False
                        btImprimir.Enabled = False
                        btImprimirEtiqueta.Enabled = False
                        rb1T.Enabled = False
                        rb2T.Enabled = False
                        rb3T.Enabled = False
                        rb4T.Enabled = False
                        rb5T.Enabled = False
                        rb6T.Enabled = False
                        rb7T.Enabled = False
                        rb8T.Enabled = False
                        rb9T.Enabled = False
                        rb10T.Enabled = False
                        btEmail.Enabled = False
                        DataGridView1.Enabled = False
                    Else
                    End If
                Else
                    conRNC.Open()
                    ds21 = New DataSet
                    da21 = New OleDbDataAdapter("delete from tblRNC where ID = " & lblID.Text & " ", conRNC)
                    ds21.Clear()
                    da21.Fill(ds21, "tblRNC")
                    conRNC.Close()
                    Call Atualizar()
                    MsgBox("Registro deletado com sucesso!")
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 84 " & ex.Message)
        End Try
    End Sub

    Sub ImprimirEtiqueta()
        TesteAbertoEtiqueta()

        Dim Excell_ETQ As New Microsoft.Office.Interop.Excel.Application
        Dim Documento_xlsx_ETQ As Microsoft.Office.Interop.Excel.Workbook
        Dim Planilha_do_Documento_xlsx_ETQ As Microsoft.Office.Interop.Excel.Worksheet

        Dim RNC As Microsoft.Office.Interop.Excel.Range
        Dim OP As Microsoft.Office.Interop.Excel.Range
        Dim Produto As Microsoft.Office.Interop.Excel.Range
        Dim Data As Microsoft.Office.Interop.Excel.Range
        Dim TurnoProduzido As Microsoft.Office.Interop.Excel.Range
        Dim Descricao As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade1 As Microsoft.Office.Interop.Excel.Range
        Dim Maquina As Microsoft.Office.Interop.Excel.Range
        'Dim RE As Microsoft.Office.Interop.Excel.Range
        Dim Inspetor As Microsoft.Office.Interop.Excel.Range
        Dim TurnoDetector As Microsoft.Office.Interop.Excel.Range


        On Error GoTo ErrHandler

        '3º Abrir o arquivo Excel
        Documento_xlsx_ETQ = Excell_ETQ.Workbooks.Open("C:\Users\Cid\Documents\Projetos\BancoDados\RNCEtiqueta.xlsx")

        '4º Abrir a planilha para inserir texto
        Planilha_do_Documento_xlsx_ETQ = Documento_xlsx_ETQ.Sheets.Item("RNCEtiqueta")

        '5º Atribuir uma célula na planilha

        RNC = Planilha_do_Documento_xlsx_ETQ.Cells(3, 6)
        OP = Planilha_do_Documento_xlsx_ETQ.Cells(3, 1)
        Produto = Planilha_do_Documento_xlsx_ETQ.Cells(2, 1)
        Data = Planilha_do_Documento_xlsx_ETQ.Cells(7, 5)
        TurnoProduzido = Planilha_do_Documento_xlsx_ETQ.Cells(6, 1)
        Descricao = Planilha_do_Documento_xlsx_ETQ.Cells(4, 1)
        Quantidade1 = Planilha_do_Documento_xlsx_ETQ.Cells(3, 3)
        Maquina = Planilha_do_Documento_xlsx_ETQ.Cells(7, 4)
        'RE = Planilha_do_Documento_xlsx_ETQ.Cells(7, 4)
        Inspetor = Planilha_do_Documento_xlsx_ETQ.Cells(7, 1)
        TurnoDetector = Planilha_do_Documento_xlsx_ETQ.Cells(6, 5)

        'conectar com nº da rnc e transferir as rows/ cells para as variaveis abaixo

        RNC.Value = "RNC: " & Today.Year - 2000 & "/" & lblRNC.Text
        OP.Value = "OP: " & txtOP.Text
        Produto.Value = "Produto: " & lblProduto.Text
        Data.Value = "Data: " & lblData.Text
        TurnoProduzido.Value = "T Produzido: " & cb1Turno.Text & " " & cb2Turnos.Text & " " & cb3Turnos.Text & " " & cb4Turnos.Text & " " & cb5Turno.Text & " " & cb6Turno.Text & " " & cb7Turno.Text & " " & cb8Turno.Text & " " & cb9Turno.Text & " " & cb10Turno.Text
        Quantidade1.Value = "Quantidade: " & lblTotalPecas.Text
        If btInserir.Text = "Aplicar" Then
            Descricao.Value = "Descrição: " & Defeito1 & " " & txtDescricaoRNC1.Text & ", " & Defeito2 & " " & txtDescricaoRNC2.Text & ", " & Defeito3 & " " & txtDescricaoRNC3.Text & ", " & Defeito4 & " " & txtDescricaoRNC4.Text & ", " & Defeito5 & " " & txtDescricaoRNC5.Text & ", " & Defeito6 & " " & txtDescricaoRNC6.Text & ", " & Defeito7 & " " & txtDescricaoRNC7.Text & ", " & Defeito8 & " " & txtDescricaoRNC8.Text & ", " & Defeito9 & " " & txtDescricaoRNC9.Text & ", " & Defeito10 & " " & txtDescricaoRNC10.Text
        Else
            Descricao.Value = "Descrição: " & txtDescricaoRNC1.Text & ", " & txtDescricaoRNC2.Text & ", " & txtDescricaoRNC3.Text & ", " & txtDescricaoRNC4.Text & ", " & txtDescricaoRNC5.Text & ", " & txtDescricaoRNC6.Text & ", " & txtDescricaoRNC7.Text & ", " & txtDescricaoRNC8.Text & ", " & txtDescricaoRNC9.Text & ", " & txtDescricaoRNC10.Text
        End If

        Maquina.Value = "Maquina: " & txtMaquina.Text
        'RE.Value = "RE: " & txtRE.Text
        Inspetor.Value = "Inspetor: " & txtInspetor.Text
        TurnoDetector.Value = "T Detector: " & cbTurno.Text

        'Dim QT_Informada As Int16 = InputBox("Informe a Quantidade de Etiquetas!!")
        Dim QT_Final As Int16 = SMC / 8
        If QT_Final = 0 Then
            QT_Final = 1
        End If
        MsgBox("Prepare: " & QT_Final & " Folhas na Impressora," _
               & Chr(13) _
               & "E informe este valor no dialogo de Impressão", , "Imprimir Etiquetas")
        'For i = 1 To QT_Final Step 1

        ' Next

        '7º Abrindo o excel
        Excell_ETQ.Visible = True

        '8º Salvando a Planilha
        Documento_xlsx_ETQ.Save()

        'imprimir
        'Documento_xlsx_ETQ.PrintOutEx() ' imprime direto

        'imprime com dialogo
        Documento_xlsx_ETQ.PrintPreview()


        '9º encerra os processos EXCEL.EXE no gerenciador de tarefas do windows 
ExitHere:
        Excell_ETQ.Quit()
        Exit Sub
ErrHandler:
        MsgBox(Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source, vbCritical, "ERRO 85 ")
        Resume ExitHere

    End Sub
    Dim Documento_xlsx As Microsoft.Office.Interop.Excel.Workbook
    Sub ImprimirRNC()

        TesteAbertoDoc()
        Dim Excell As New Microsoft.Office.Interop.Excel.Application

        Dim Planilha_do_Documento_xlsx As Microsoft.Office.Interop.Excel.Worksheet

        Dim RNC As Microsoft.Office.Interop.Excel.Range
        Dim OP As Microsoft.Office.Interop.Excel.Range
        Dim CodProduto As Microsoft.Office.Interop.Excel.Range
        Dim Produto As Microsoft.Office.Interop.Excel.Range
        Dim Data As Microsoft.Office.Interop.Excel.Range
        Dim Hora As Microsoft.Office.Interop.Excel.Range
        Dim Deteccao As Microsoft.Office.Interop.Excel.Range
        Dim Maquina As Microsoft.Office.Interop.Excel.Range

        Dim Turno1 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas1 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa1 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade1 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC1 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC1 As Microsoft.Office.Interop.Excel.Range

        Dim Turno2 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas2 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa2 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade2 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC2 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC2 As Microsoft.Office.Interop.Excel.Range

        Dim Turno3 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas3 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa3 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade3 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC3 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC3 As Microsoft.Office.Interop.Excel.Range

        Dim Turno4 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas4 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa4 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade4 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC4 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC4 As Microsoft.Office.Interop.Excel.Range

        Dim Turno5 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas5 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa5 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade5 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC5 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC5 As Microsoft.Office.Interop.Excel.Range

        Dim Turno6 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas6 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa6 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade6 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC6 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC6 As Microsoft.Office.Interop.Excel.Range

        Dim Turno7 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas7 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa7 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade7 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC7 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC7 As Microsoft.Office.Interop.Excel.Range

        Dim Turno8 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas8 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa8 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade8 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC8 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC8 As Microsoft.Office.Interop.Excel.Range

        Dim Turno9 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas9 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa9 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade9 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC9 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC9 As Microsoft.Office.Interop.Excel.Range

        Dim Turno10 As Microsoft.Office.Interop.Excel.Range
        Dim Caixas10 As Microsoft.Office.Interop.Excel.Range
        Dim QT_Caixa10 As Microsoft.Office.Interop.Excel.Range
        Dim Quantidade10 As Microsoft.Office.Interop.Excel.Range
        Dim CodRNC10 As Microsoft.Office.Interop.Excel.Range
        Dim DescricaoRNC10 As Microsoft.Office.Interop.Excel.Range

        Dim RE As Microsoft.Office.Interop.Excel.Range
        Dim Inspetor As Microsoft.Office.Interop.Excel.Range
        Dim Setor As Microsoft.Office.Interop.Excel.Range
        Dim TurnoDetector As Microsoft.Office.Interop.Excel.Range

        Dim Obs As Microsoft.Office.Interop.Excel.Range
        On Error GoTo ErrHandler

        '3º Abrir o arquivo Excel
        Documento_xlsx = Excell.Workbooks.Open("C:\Users\Cid\Documents\Projetos\BancoDados\RNCDoc.xlsx")

        '4º Abrir a planilha para inserir texto
        Planilha_do_Documento_xlsx = Documento_xlsx.Sheets.Item("RNCForm")

        '5º Atribuir uma célula na planilha

        RNC = Planilha_do_Documento_xlsx.Cells(1, 10)
        OP = Planilha_do_Documento_xlsx.Cells(5, 10)
        CodProduto = Planilha_do_Documento_xlsx.Cells(5, 6)
        Produto = Planilha_do_Documento_xlsx.Cells(5, 1)
        Data = Planilha_do_Documento_xlsx.Cells(2, 10)
        Hora = Planilha_do_Documento_xlsx.Cells(2, 11)
        Deteccao = Planilha_do_Documento_xlsx.Cells(20, 3)
        Maquina = Planilha_do_Documento_xlsx.Cells(5, 9)

        Turno1 = Planilha_do_Documento_xlsx.Cells(8, 1)
        Caixas1 = Planilha_do_Documento_xlsx.Cells(8, 2)
        QT_Caixa1 = Planilha_do_Documento_xlsx.Cells(8, 5)
        Quantidade1 = Planilha_do_Documento_xlsx.Cells(8, 6)
        CodRNC1 = Planilha_do_Documento_xlsx.Cells(8, 7)
        DescricaoRNC1 = Planilha_do_Documento_xlsx.Cells(8, 8)

        Turno2 = Planilha_do_Documento_xlsx.Cells(9, 1)
        Caixas2 = Planilha_do_Documento_xlsx.Cells(9, 2)
        QT_Caixa2 = Planilha_do_Documento_xlsx.Cells(9, 5)
        Quantidade2 = Planilha_do_Documento_xlsx.Cells(9, 6)
        CodRNC2 = Planilha_do_Documento_xlsx.Cells(9, 7)
        DescricaoRNC2 = Planilha_do_Documento_xlsx.Cells(9, 8)

        Turno3 = Planilha_do_Documento_xlsx.Cells(10, 1)
        Caixas3 = Planilha_do_Documento_xlsx.Cells(10, 2)
        QT_Caixa3 = Planilha_do_Documento_xlsx.Cells(10, 5)
        Quantidade3 = Planilha_do_Documento_xlsx.Cells(10, 6)
        CodRNC3 = Planilha_do_Documento_xlsx.Cells(10, 7)
        DescricaoRNC3 = Planilha_do_Documento_xlsx.Cells(10, 8)

        Turno4 = Planilha_do_Documento_xlsx.Cells(11, 1)
        Caixas4 = Planilha_do_Documento_xlsx.Cells(11, 2)
        QT_Caixa4 = Planilha_do_Documento_xlsx.Cells(11, 5)
        Quantidade4 = Planilha_do_Documento_xlsx.Cells(11, 6)
        CodRNC4 = Planilha_do_Documento_xlsx.Cells(11, 7)
        DescricaoRNC4 = Planilha_do_Documento_xlsx.Cells(11, 8)

        Turno5 = Planilha_do_Documento_xlsx.Cells(12, 1)
        Caixas5 = Planilha_do_Documento_xlsx.Cells(12, 2)
        QT_Caixa5 = Planilha_do_Documento_xlsx.Cells(12, 5)
        Quantidade5 = Planilha_do_Documento_xlsx.Cells(12, 6)
        CodRNC5 = Planilha_do_Documento_xlsx.Cells(12, 7)
        DescricaoRNC5 = Planilha_do_Documento_xlsx.Cells(12, 8)

        Turno6 = Planilha_do_Documento_xlsx.Cells(13, 1)
        Caixas6 = Planilha_do_Documento_xlsx.Cells(13, 2)
        QT_Caixa6 = Planilha_do_Documento_xlsx.Cells(13, 5)
        Quantidade6 = Planilha_do_Documento_xlsx.Cells(13, 6)
        CodRNC6 = Planilha_do_Documento_xlsx.Cells(13, 7)
        DescricaoRNC6 = Planilha_do_Documento_xlsx.Cells(13, 8)

        Turno7 = Planilha_do_Documento_xlsx.Cells(14, 1)
        Caixas7 = Planilha_do_Documento_xlsx.Cells(14, 2)
        QT_Caixa7 = Planilha_do_Documento_xlsx.Cells(14, 5)
        Quantidade7 = Planilha_do_Documento_xlsx.Cells(14, 6)
        CodRNC7 = Planilha_do_Documento_xlsx.Cells(14, 7)
        DescricaoRNC7 = Planilha_do_Documento_xlsx.Cells(14, 8)

        Turno8 = Planilha_do_Documento_xlsx.Cells(15, 1)
        Caixas8 = Planilha_do_Documento_xlsx.Cells(15, 2)
        QT_Caixa8 = Planilha_do_Documento_xlsx.Cells(15, 5)
        Quantidade8 = Planilha_do_Documento_xlsx.Cells(15, 6)
        CodRNC8 = Planilha_do_Documento_xlsx.Cells(15, 7)
        DescricaoRNC8 = Planilha_do_Documento_xlsx.Cells(15, 8)

        Turno9 = Planilha_do_Documento_xlsx.Cells(16, 1)
        Caixas9 = Planilha_do_Documento_xlsx.Cells(16, 2)
        QT_Caixa9 = Planilha_do_Documento_xlsx.Cells(16, 5)
        Quantidade9 = Planilha_do_Documento_xlsx.Cells(16, 6)
        CodRNC9 = Planilha_do_Documento_xlsx.Cells(16, 7)
        DescricaoRNC9 = Planilha_do_Documento_xlsx.Cells(16, 8)

        Turno10 = Planilha_do_Documento_xlsx.Cells(17, 1)
        Caixas10 = Planilha_do_Documento_xlsx.Cells(17, 2)
        QT_Caixa10 = Planilha_do_Documento_xlsx.Cells(17, 5)
        Quantidade10 = Planilha_do_Documento_xlsx.Cells(17, 6)
        CodRNC10 = Planilha_do_Documento_xlsx.Cells(17, 7)
        DescricaoRNC10 = Planilha_do_Documento_xlsx.Cells(17, 8)

        RE = Planilha_do_Documento_xlsx.Cells(26, 4)
        Inspetor = Planilha_do_Documento_xlsx.Cells(26, 5)
        TurnoDetector = Planilha_do_Documento_xlsx.Cells(26, 8)
        Setor = Planilha_do_Documento_xlsx.Cells(26, 9)

        Obs = Planilha_do_Documento_xlsx.Cells(22, 2)

        'conectar com nº da rnc e transferir as rows/ cells para as variaveis abaixo

        RNC.Value = Today.Year - 2000 & "/" & lblRNC.Text
        OP.Value = Integer.Parse(txtOP.Text)
        CodProduto.Value = lblCodProduto.Text
        Produto.Value = lblProduto.Text


        If btInserir.Text = "Aplicar" Then
            Data.Value = "Gerado em: " & Today
            Hora.Value = TimeOfDay
        ElseIf btAlterar.Text = "Aplicar" Then
            Data.Value = "Alterado em: " & Today
            Hora.Value = TimeOfDay
        Else
            If Alteradu = "" Then
                Data.Value = "Gerado em: " & lblData.Text
                Hora.Value = lblHora.Text

            Else
                'remove caracteres de uma string
                Data.Value = "Alterado em: " & Alteradu.Remove(10, 6)
                Hora.Value = Alteradu.Remove(0, 10)
            End If
        End If

        Deteccao.Value = cbDetectado.Text
        Maquina.Value = txtMaquina.Text

        Turno1.Value = cb1Turno.Text
        Caixas1.Value = txtCaixas1Turno.Text
        QT_Caixa1.Value = txtQtCaixasReprovada1.Text
        Quantidade1.Value = Double.Parse(lblQtPorTurno1.Text)
        CodRNC1.Value = txtCodigoRNC1.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC1.Value = Defeito1 & txtDescricaoRNC1.Text
        Else
            DescricaoRNC1.Value = txtDescricaoRNC1.Text
        End If
        Turno2.Value = cb2Turnos.Text
        Caixas2.Value = txtCaixas2Turno.Text
        QT_Caixa2.Value = txtQtCaixasReprovada2.Text
        If lblQtPorTurno2.Text = "" Then
            Quantidade2.Value = 0
        Else
            Quantidade2.Value = Double.Parse(lblQtPorTurno2.Text)
        End If
        CodRNC2.Value = txtCodigoRNC2.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC2.Value = Defeito2 & " - " & txtDescricaoRNC2.Text
        Else
            DescricaoRNC2.Value = txtDescricaoRNC2.Text
        End If

        Turno3.Value = cb3Turnos.Text
        Caixas3.Value = txtCaixas3Turno.Text
        QT_Caixa3.Value = txtQtCaixasReprovada3.Text
        If lblQtPorTurno3.Text = "" Then
            Quantidade3.Value = 0
        Else
            Quantidade3.Value = Double.Parse(lblQtPorTurno3.Text)
        End If
        CodRNC3.Value = txtCodigoRNC3.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC3.Value = Defeito3 & " - " & txtDescricaoRNC3.Text
        Else
            DescricaoRNC3.Value = txtDescricaoRNC3.Text
        End If

        Turno4.Value = cb4Turnos.Text
        Caixas4.Value = txtCaixas4Turno.Text
        QT_Caixa4.Value = txtQtCaixasReprovada4.Text
        If lblQtPorTurno4.Text = "" Then
            Quantidade4.Value = 0
        Else
            Quantidade4.Value = Double.Parse(lblQtPorTurno4.Text)
        End If
        CodRNC4.Value = txtCodigoRNC4.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC4.Value = Defeito4 & " - " & txtDescricaoRNC4.Text
        Else
            DescricaoRNC4.Value = txtDescricaoRNC4.Text
        End If

        Turno5.Value = cb5Turno.Text
        Caixas5.Value = txtCaixas5Turno.Text
        QT_Caixa5.Value = txtQtCaixasReprovada5.Text
        If lblQtPorTurno5.Text = "" Then
            Quantidade5.Value = 0
        Else
            Quantidade5.Value = Double.Parse(lblQtPorTurno5.Text)
        End If
        CodRNC5.Value = txtCodigoRNC5.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC5.Value = Defeito5 & " - " & txtDescricaoRNC5.Text
        Else
            DescricaoRNC5.Value = txtDescricaoRNC5.Text
        End If

        Turno6.Value = cb6Turno.Text
        Caixas6.Value = txtCaixas6Turno.Text
        QT_Caixa6.Value = txtQtCaixasReprovada6.Text
        If lblQtPorTurno6.Text = "" Then
            Quantidade6.Value = 0
        Else
            Quantidade6.Value = Double.Parse(lblQtPorTurno6.Text)
        End If
        CodRNC6.Value = txtCodigoRNC6.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC6.Value = Defeito6 & " - " & txtDescricaoRNC6.Text
        Else
            DescricaoRNC6.Value = txtDescricaoRNC6.Text
        End If

        Turno7.Value = cb7Turno.Text
        Caixas7.Value = txtCaixas7Turno.Text
        QT_Caixa7.Value = txtQtCaixasReprovada7.Text
        If lblQtPorTurno7.Text = "" Then
            Quantidade7.Value = 0
        Else
            Quantidade7.Value = Double.Parse(lblQtPorTurno7.Text)
        End If
        CodRNC7.Value = txtCodigoRNC7.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC7.Value = Defeito7 & " - " & txtDescricaoRNC7.Text
        Else
            DescricaoRNC7.Value = txtDescricaoRNC7.Text
        End If

        Turno8.Value = cb8Turno.Text
        Caixas8.Value = txtCaixas8Turno.Text
        QT_Caixa8.Value = txtQtCaixasReprovada8.Text
        If lblQtPorTurno8.Text = "" Then
            Quantidade8.Value = 0
        Else
            Quantidade8.Value = Double.Parse(lblQtPorTurno8.Text)
        End If
        CodRNC8.Value = txtCodigoRNC8.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC8.Value = Defeito8 & " - " & txtDescricaoRNC8.Text
        Else
            DescricaoRNC8.Value = txtDescricaoRNC8.Text
        End If

        Turno9.Value = cb9Turno.Text
        Caixas9.Value = txtCaixas9Turno.Text
        QT_Caixa9.Value = txtQtCaixasReprovada9.Text
        If lblQtPorTurno9.Text = "" Then
            Quantidade9.Value = 0
        Else
            Quantidade9.Value = Double.Parse(lblQtPorTurno9.Text)
        End If
        CodRNC9.Value = txtCodigoRNC9.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC9.Value = Defeito9 & " - " & txtDescricaoRNC9.Text
        Else
            DescricaoRNC9.Value = txtDescricaoRNC9.Text
        End If

        Turno10.Value = cb10Turno.Text
        Caixas10.Value = txtCaixas10Turno.Text
        QT_Caixa10.Value = txtQtCaixasReprovada10.Text
        If lblQtPorTurno10.Text = "" Then
            Quantidade10.Value = 0
        Else
            Quantidade10.Value = Double.Parse(lblQtPorTurno10.Text)
        End If
        CodRNC10.Value = txtCodigoRNC10.Text
        If btInserir.Text = "Aplicar" Then
            DescricaoRNC10.Value = Defeito10 & " - " & txtDescricaoRNC10.Text
        Else
            DescricaoRNC10.Value = txtDescricaoRNC10.Text
        End If

        RE.Value = "RE: " & txtRE.Text
        Inspetor.Value = "Nome: " & txtInspetor.Text
        Setor.Value = "Setor: " & txtSetor.Text
        TurnoDetector.Value = "Turno: " & cbTurno.Text
        Obs.Value = txtOBS.Text



        'Documento_xlsx.PrintOutEx() ' imprime direto

        If btInserir.Text = "Aplicar" Then
            '7º Abrindo o excel
            Excell.Visible = True
            '8º Salvando a Planilha
            Documento_xlsx.Save()
            'imprimircom dialogo
            printview()
        ElseIf btAlterar.Text = "Aplicar" Then
            '7º Abrindo o excel
            Excell.Visible = True
            '8º Salvando a Planilha
            Documento_xlsx.Save()
            'imprimircom dialogo
            printview()
        ElseIf btImprimir.Text = "Imprimir..." Then
            '7º Abrindo o excel
            Excell.Visible = True
            '8º Salvando a Planilha
            Documento_xlsx.Save()
            'imprimircom dialogo
            printview()
        Else
            '7º Abrindo o excel
            Excell.Visible = False
            '8º Salvando a Planilha
            Documento_xlsx.Save()
            'imprimircom dialogo
        End If



        '9º encerra os processos EXCEL.EXE no gerenciador de tarefas do windows 
ExitHere:
        Excell.Quit()
        Marshal.ReleaseComObject(Documento_xlsx)
        Marshal.ReleaseComObject(Excell)
        Excell = Nothing
        Exit Sub
ErrHandler:
        MsgBox(Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source, vbCritical, "Erro 86 ")
        Resume ExitHere

    End Sub
    Sub printview()
        Documento_xlsx.PrintPreview()
    End Sub
    Private Sub btImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btImprimir.Click
        Try
            ' btImprimir.Text = "Imprimir RNC..."

            If lblRNC.Text = "*" Or lblRNC.Text = "" Then
                MsgBox("Selecione uma RNC na tabela abaixo", , "RNC")
            Else
                btImprimir.Text = "Imprimir..."
                Call ImprimirRNC()
                btImprimir.Text = "Imprimir RNC"
            End If
        Catch ex As Exception
            MsgBox("Erro 87 " & ex.Message)
            btImprimir.Text = "Imprimir RNC"
        End Try
        btImprimir.Text = "Imprimir RNC"
    End Sub

    Sub email()
        Dim OutlookMessage As Outlook.MailItem
        Dim AppOutlook As New Microsoft.Office.Interop.Outlook.Application
        Try

            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            'Dim Recipents As Outlook.Recipients = OutlookMessage.Recipients
            OutlookMessage.To = "cidevangelista@hotmail.com" ' Criar um grupo no outlook chamado RNC
            'Recipents.Add("cidmevb@gmail.com; inspetor1@mondicap.com.br; rafael.pedroso@mondicap.com.br; recebimento@mondicap.com.br")
            If btInserir.Text = "Aplicar" Then
                OutlookMessage.Subject = "Segue uma Nova RNC: " & Today.Year - 2000 & "/" & lblRNC.Text
            ElseIf btAlterar.Text = "Aplicar" Then
                OutlookMessage.Subject = "Segue a Alteração da RNC: " & Today.Year - 2000 & "/" & lblRNC.Text
            ElseIf btExcluir.Text = "Aplicar" Then
                OutlookMessage.Body = "Segue a Exclusão da RNC"
            Else
                OutlookMessage.Subject = "Segue o Reenvio da RNC: " & Today.Year - 2000 & "/" & lblRNC.Text
            End If
            If txtQtCaixasReprovada1.Text = "" Then
                txtQtCaixasReprovada1.Text = 0
            End If
            If txtQtCaixasReprovada2.Text = "" Then
                txtQtCaixasReprovada2.Text = 0
            End If
            If txtQtCaixasReprovada3.Text = "" Then
                txtQtCaixasReprovada3.Text = 0
            End If
            If txtQtCaixasReprovada4.Text = "" Then
                txtQtCaixasReprovada4.Text = 0
            End If
            If txtQtCaixasReprovada5.Text = "" Then
                txtQtCaixasReprovada5.Text = 0
            End If
            If txtQtCaixasReprovada6.Text = "" Then
                txtQtCaixasReprovada6.Text = 0
            End If
            If txtQtCaixasReprovada7.Text = "" Then
                txtQtCaixasReprovada7.Text = 0
            End If
            If txtQtCaixasReprovada8.Text = "" Then
                txtQtCaixasReprovada8.Text = 0
            End If
            If txtQtCaixasReprovada9.Text = "" Then
                txtQtCaixasReprovada9.Text = 0
            End If
            If txtQtCaixasReprovada10.Text = "" Then
                txtQtCaixasReprovada10.Text = 0
            End If
            Dim Descricaox As String
            Dim Turnox As String = "" & cb1Turno.Text & " " & cb2Turnos.Text & " " & cb3Turnos.Text & " " & cb4Turnos.Text & " " & cb5Turno.Text & " " & cb6Turno.Text & " " & cb7Turno.Text & " " & cb8Turno.Text & " " & cb9Turno.Text & " " & cb10Turno.Text & " "
            If btInserir.Text = "Aplicar" Then
                Descricaox = Defeito1 & " " & txtDescricaoRNC1.Text & ", " & Defeito2 & " " & txtDescricaoRNC2.Text & ", " & Defeito3 & " " & txtDescricaoRNC3.Text & ", " & Defeito4 & " " & txtDescricaoRNC4.Text & ", " & Defeito5 & " " & txtDescricaoRNC5.Text & ", " & Defeito6 & " " & txtDescricaoRNC6.Text & ", " & Defeito7 & " " & txtDescricaoRNC7.Text & ", " & Defeito8 & " " & txtDescricaoRNC8.Text & ", " & Defeito9 & " " & txtDescricaoRNC9.Text & ", " & Defeito10 & " " & txtDescricaoRNC10.Text & ""
            Else
                Descricaox = txtDescricaoRNC1.Text & ", " & txtDescricaoRNC2.Text & ", " & txtDescricaoRNC3.Text & ", " & txtDescricaoRNC4.Text & ", " & txtDescricaoRNC5.Text & ", " & txtDescricaoRNC6.Text & ", " & txtDescricaoRNC7.Text & ", " & txtDescricaoRNC8.Text & ", " & txtDescricaoRNC9.Text & ", " & txtDescricaoRNC10.Text & ""
            End If
            Dim Totalx As Integer = Integer.Parse(txtQtCaixasReprovada1.Text + Integer.Parse(txtQtCaixasReprovada2.Text + Integer.Parse(txtQtCaixasReprovada3.Text + Integer.Parse(txtQtCaixasReprovada4.Text + Integer.Parse(txtQtCaixasReprovada5.Text + Integer.Parse(txtQtCaixasReprovada6.Text + Integer.Parse(txtQtCaixasReprovada7.Text + Integer.Parse(txtQtCaixasReprovada8.Text + Integer.Parse(txtQtCaixasReprovada9.Text + Integer.Parse(txtQtCaixasReprovada10.Text))))))))))

            If 1 = 1 Then

                OutlookMessage.Body = "OP.........................: " & txtOP.Text & "" _
                     & Chr(13) _
                     & "Data......................: " & lblData.Text & "" _
                     & Chr(13) _
                     & "Hora......................: " & lblHora.Text & "" _
                     & Chr(13) _
                     & "Produto.................: " & lblProduto.Text & "" _
                     & Chr(13) _
                     & "Código...................: " & lblCodProduto.Text & "" _
                     & Chr(13) _
                     & "Contenção.............: " & cbDetectado.Text & "" _
                     & Chr(13) _
                     & "Máquina................: " & txtMaquina.Text & "" _
                     & Chr(13) _
                     & "Turno(s).................: " & Turnox & "" _
                     & Chr(13) _
                     & "Descrição(ões).......: " & Descricaox & "" _
                     & Chr(13) _
                     & "Quantidade Total...: " & lblTotalPecas.Text & "" _
                     & Chr(13) _
                     & "Total de Caixas.......: " & Totalx & "" _
                     & Chr(13) _
                     & "Observação...........: " & txtOBS.Text & "" _
                     & Chr(13) _
                     & "Enviado Por...........: " & txtInspetor.Text & "" _
                     & Chr(13) _
                     & "Turno....................: " & cbTurno.Text & ""

            End If
            btAlterar.Text = "Alterar"
            btInserir.Text = "Inserir"
            ImprimirRNC()
            System.Threading.Thread.Sleep(5000)
            OutlookMessage.Attachments.Add("C:\Users\Cid\Documents\Projetos\BancoDados\RNCDoc.xlsx")
            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatRichText

            If (MsgBox("O E-mail está pronto para ser enviado. Deseja Enviar?" _
                       & Chr(13) _
                       & Chr(13) _
                       & "'Sim' = Enviar" _
                       & Chr(13) _
                       & "'Não' = Alterar", vbYesNo, "E-mail") = vbYes) Then
                OutlookMessage.Save()
                OutlookMessage.Send()
            Else
                OutlookMessage.Display()
                OutlookMessage.Save()
            End If
        Catch ex As Exception
            MessageBox.Show("Erro 88 " & ex.Message) 'if you dont want this message, simply delete this line 
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try

    End Sub

    Private Sub cb1Turno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoRNC1.KeyPress, cbTurno.KeyPress, cb1Turno.KeyPress, cb2Turnos.KeyPress, cb3Turnos.KeyPress, cb4Turnos.KeyPress, cb5Turno.KeyPress, cb6Turno.KeyPress, cb7Turno.KeyPress, cb8Turno.KeyPress, cb9Turno.KeyPress, cb10Turno.KeyPress, cbDetectado.KeyPress
        e.Handled = True
    End Sub
    Private Sub Codigo(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoRNC1.KeyPress, txtCodigoRNC2.KeyPress, txtCodigoRNC3.KeyPress, txtCodigoRNC4.KeyPress, txtCodigoRNC5.KeyPress, txtCodigoRNC6.KeyPress, txtCodigoRNC7.KeyPress, txtCodigoRNC8.KeyPress, txtCodigoRNC9.KeyPress, txtCodigoRNC10.KeyPress
        e.Handled = True
    End Sub

    Private Sub cbTurno_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTurno.LostFocus
        Try
            If btInserir.Text = "Aplicar" Then
                btInserir.Focus()
            ElseIf btAlterar.Text = "Aplicar" Then
                btAlterar.Focus()
            ElseIf btExcluir.Text = "Aplicar" Then
                btExcluir.Focus()
            Else
                txtOP.Focus()
            End If
        Catch ex As Exception
            MsgBox("Erro 89 " & ex.Message)
        End Try
    End Sub

    Private Sub btEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEmail.Click
        Try
            If lblRNC.Text = "0" Or lblRNC.Text = "" Then
                MsgBox("Selecione um RNC na tabela abaixo", , "Selecione uma RNC")
            Else
                email()
            End If
        Catch exc As Exception
            MsgBox("Erro 90 " & exc.Message)
        End Try
    End Sub

    Sub TesteAbertoConsultaOP()
        Try
            Dim Consulta_OP As Boolean
            Consulta_OP = Test("C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.accdb")
            If Consulta_OP = True Then
                Dim OPConvertida As Integer = 0
                For OPConvertida = 5 To 20
                    Consulta_OP = Test("C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.accdb")
                    If Consulta_OP = True Then
                        OPConvertida = 5
                        If (MsgBox("O Arquivo 'Consulta_OP.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "Consulta_OP.accdb")) = vbRetry Then
                        Else
                            Close()
                            Exit For
                        End If
                    ElseIf Consulta_OP = False Then
                        OPConvertida = 20
                    End If
                Next
            End If
        Catch e As Exception
            MsgBox(e.Message)
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

    Sub testeAbertoMaquina()
        Dim RNC_Maquina As Boolean
        RNC_Maquina = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Maquina.accdb")
        If RNC_Maquina = True Then
            Dim RNCMaquina As Integer = 0
            For RNCMaquina = 5 To 20
                RNC_Maquina = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Maquina.accdb")
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

    Sub TesteAbertoPecasVolume()
        Dim RNC_PecasVolume As Boolean
        RNC_PecasVolume = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_PecasVolume.accdb")
        If RNC_PecasVolume = True Then
            Dim RNCPecasVolume As Integer = 0
            For RNCPecasVolume = 5 To 20
                RNC_PecasVolume = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_PecasVolume.accdb")
                If RNC_PecasVolume = True Then
                    RNCPecasVolume = 5
                    If (MsgBox("O Arquivo 'RNC_PecasVolume.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_PecasVolume.accdb")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNC_PecasVolume = False Then
                    RNCPecasVolume = 20
                End If
            Next
        End If
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

    Sub TesteAbertoRNC()
        Dim RNC_RNC As Boolean
        RNC_RNC = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb")
        If RNC_RNC = True Then
            Dim RNCRNC As Integer = 0
            For RNCRNC = 5 To 20
                RNC_RNC = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb")
                If RNC_RNC = True Then
                    RNCRNC = 5
                    If (MsgBox("O Arquivo 'RNC_RNC.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_RNC.accdb")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNC_RNC = False Then
                    RNCRNC = 20
                End If
            Next
        End If
    End Sub

    Sub TesteAbertoDoc()
        Dim RNCDoc As Boolean
        RNCDoc = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCDoc.xlsx")
        If RNCDoc = True Then
            Dim RNC_Doc As Integer = 0
            For RNC_Doc = 5 To 20
                RNCDoc = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCDoc.xlsx")
                If RNCDoc = True Then
                    RNC_Doc = 5
                    If (MsgBox("O Arquivo 'RNCDoc.xlsx' está aberto, Feche-o para para continuar", vbRetryCancel, "RNCDoc.xlsx")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNCDoc = False Then
                    RNC_Doc = 20
                End If
            Next
        End If

    End Sub

    Sub TesteAbertoEtiqueta()
        Dim RNCEtiqueta As Boolean
        RNCEtiqueta = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCEtiqueta.xlsx")
        If RNCEtiqueta = True Then
            Dim RNC_Etiqueta As Integer = 0
            For RNC_Etiqueta = 5 To 20
                RNCEtiqueta = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCEtiqueta.xlsx")
                If RNCEtiqueta = True Then
                    RNC_Etiqueta = 5
                    If (MsgBox("O Arquivo 'RNCEtiqueta.xlsx' está aberto, Feche-o para para continuar", vbRetryCancel, "RNCEtiqueta.xlsx")) = vbRetry Then
                    Else
                        Close()
                        Exit For
                    End If
                ElseIf RNCEtiqueta = False Then
                    RNC_Etiqueta = 20
                End If
            Next
        End If
    End Sub

    Sub Teste_Aberto()
        Try
            Dim Consulta_OP As Boolean
            Dim RNC_Defeito As Boolean
            Dim RNC_Maquina As Boolean
            Dim RNC_PecasVolume As Boolean
            Dim RNC_RE As Boolean
            Dim RNC_RNC As Boolean
            Dim RNCDoc As Boolean
            Dim RNCEtiqueta As Boolean

            Consulta_OP = Test("C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.accdb")
            RNC_Defeito = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Defeito.accdb")
            RNC_Maquina = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Maquina.accdb")
            RNC_PecasVolume = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_PecasVolume.accdb")
            RNC_RE = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RE.accdb")
            RNC_RNC = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb")
            RNCDoc = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCDoc.xlsx")
            RNCEtiqueta = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCEtiqueta.xlsx")


            If Consulta_OP = True Then
                Dim OPConvertida As Integer = 0
                For OPConvertida = 5 To 20
                    Consulta_OP = Test("C:\Users\Cid\Documents\Projetos\BancoDados\Consulta_OP.accdb")
                    If Consulta_OP = True Then
                        OPConvertida = 5
                        If (MsgBox("O Arquivo 'Consulta_OP.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "Consulta_OP.accdb")) = vbRetry Then
                        Else
                            Close()
                            Exit For
                        End If
                    ElseIf Consulta_OP = False Then
                        OPConvertida = 20
                    End If
                Next

            ElseIf RNC_Defeito = True Then
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

            ElseIf RNC_Maquina = True Then
                Dim RNCMaquina As Integer = 0
                For RNCMaquina = 5 To 20
                    RNC_Maquina = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_Maquina.accdb")
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

            ElseIf RNC_PecasVolume = True Then
                Dim RNCPecasVolume As Integer = 0
                For RNCPecasVolume = 5 To 20
                    RNC_PecasVolume = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_PecasVolume.accdb")
                    If RNC_PecasVolume = True Then
                        RNCPecasVolume = 5
                        If (MsgBox("O Arquivo 'RNC_PecasVolume.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_PecasVolume.accdb")) = vbRetry Then
                        Else
                            Close()
                            Exit For
                        End If
                    ElseIf RNC_PecasVolume = False Then
                        RNCPecasVolume = 20
                    End If
                Next

            ElseIf RNC_RE = True Then
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

            ElseIf RNC_RNC = True Then
                Dim RNCRNC As Integer = 0
                For RNCRNC = 5 To 20
                    RNC_RNC = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb")
                    If RNC_RNC = True Then
                        RNCRNC = 5
                        If (MsgBox("O Arquivo 'RNC_RNC.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "RNC_RNC.accdb")) = vbRetry Then
                        Else
                            Close()
                            Exit For
                        End If
                    ElseIf RNC_RNC = False Then
                        RNCRNC = 20
                    End If
                Next

            ElseIf RNCDoc = True Then
                Dim RNC_Doc As Integer = 0
                For RNC_Doc = 5 To 20
                    RNCDoc = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCDoc.xlsx")
                    If RNCDoc = True Then
                        RNC_Doc = 5
                        If (MsgBox("O Arquivo 'RNCDoc.xlsx' está aberto, Feche-o para para continuar", vbRetryCancel, "RNCDoc.xlsx")) = vbRetry Then
                        Else
                            Close()
                            Exit For
                        End If
                    ElseIf RNCDoc = False Then
                        RNC_Doc = 20
                    End If
                Next

            ElseIf RNCEtiqueta = True Then
                Dim RNC_Etiqueta As Integer = 0
                For RNC_Etiqueta = 5 To 20
                    RNCEtiqueta = Test("C:\Users\Cid\Documents\Projetos\BancoDados\RNCEtiqueta.xlsx")
                    If RNCEtiqueta = True Then
                        RNC_Etiqueta = 5
                        If (MsgBox("O Arquivo 'RNCEtiqueta.xlsx' está aberto, Feche-o para para continuar", vbRetryCancel, "RNCEtiqueta.xlsx")) = vbRetry Then
                        Else
                            Close()
                            Exit For
                        End If
                    ElseIf RNCEtiqueta = False Then
                        RNC_Etiqueta = 20
                    End If
                Next

            End If
        Catch e As Exception
            MsgBox(e.Message)
        End Try

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
    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        frmMaquinaConsulta.ShowDialog()
        txtMaquina.Text = txtMMaquina
        rb1T.Checked = True
        rb1T.Focus()
    End Sub
    Private Sub Label5_Clck(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.MouseEnter
        Cursor = Cursors.Hand
    End Sub
    Private Sub Label5_Clic(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        Cursor = Cursors.Default
    End Sub

    Private Sub txtOPRetrabalho_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOPRetrabalho.LostFocus
        Try

            If btAlterarStatus.Text = "Aplicar" Then
                If txtOPRetrabalho.Text = "" Or txtOPRetrabalho.Text = "0" Or txtOPRetrabalho.Text = "00" Or txtOPRetrabalho.Text = "000" Or txtOPRetrabalho.Text = "0000" Or txtOPRetrabalho.Text = "00000" Or txtOPRetrabalho.Text = "000000" Then
                    MsgBox("Insira uma 'OP de Retrabalho' válida", , "OP de Retrabalho")
                    txtOPRetrabalho.Focus()
                Else

                    Dim da10 As New OleDbDataAdapter
                    Dim ds10 As New DataSet
                    Dim dt10 As New DataTable
                    Dim cb10 As New OleDbCommandBuilder
                    conConsulta_OP.Open()
                    Dim sel12 As String = "SELECT top 1 Cod_Mondicap FROM tblOP where OP = " & txtOPRetrabalho.Text & "  "
                    da10 = New OleDbDataAdapter(sel12, conConsulta_OP)
                    ds10.Clear()
                    dt10.Clear()
                    da10.Fill(dt10)
                    If dt10.Rows.Count = 0 Then
                        conConsulta_OP.Close()
                        MsgBox("A OP não existe")
                        txtOPRetrabalho.Focus()
                    Else

                        da10.Fill(ds10, "tblOP")
                        compara = Int64.Parse(ds10.Tables("tblOP").Rows(0)("Cod_Mondicap"))
                        conConsulta_OP.Close()
                        lblCarregada.Text = "Carregada"
                        If lblCodProduto.Text = compara Then
                        Else
                            MsgBox("As OPs 'Reprovada vs Retrabalho' não se tratam do mesmo Produto, favor verificar!", , "Divergência de OPs")
                            txtOPRetrabalho.Clear()
                            txtOPRetrabalho.Focus()
                            lblCarregada.Text = "*"
                            compara = 0
                        End If
                    End If
                End If
            Else
            End If
        Catch ex As Exception
            MsgBox("Erro 100 " & ex.Message)
            conConsulta_OP.Close()
        End Try


    End Sub

    Private Sub rbRT1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT1.CheckedChanged, rbRF1.CheckedChanged, rbLC1.CheckedChanged
        If rbRT1.Checked = True Then
            txtQTReprovado1.Clear()
            txtQTAprovado1.Clear()
            txtQTReprovado1.Enabled = True
            txtQTAprovado1.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho1 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF1.Checked = True Then
            txtQTAprovado1.Text = "0"
            txtQTReprovado1.Enabled = False
            txtQTAprovado1.Enabled = False
            txtQTReprovado1.Text = lblQtPorTurno1.Text
            OPRetrabalho1 = 0
        ElseIf rbLC1.Checked = True Then
            txtQTReprovado1.Text = "0"
            txtQTReprovado1.Enabled = False
            txtQTAprovado1.Enabled = False
            txtQTAprovado1.Text = lblQtPorTurno1.Text
            OPRetrabalho1 = 0
        End If
    End Sub

    Private Sub rbRT2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT2.CheckedChanged, rbRF2.CheckedChanged, rbLC2.CheckedChanged
        If rbRT2.Checked = True Then
            txtQTReprovado2.Clear()
            txtQTAprovado2.Clear()
            txtQTReprovado2.Enabled = True
            txtQTAprovado2.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho2 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF2.Checked = True Then
            txtQTAprovado2.Text = "0"
            txtQTReprovado2.Enabled = False
            txtQTAprovado2.Enabled = False
            txtQTReprovado2.Text = lblQtPorTurno2.Text
            OPRetrabalho2 = 0
        ElseIf rbLC2.Checked = True Then
            txtQTReprovado2.Text = "0"
            txtQTReprovado2.Enabled = False
            txtQTAprovado2.Enabled = False
            txtQTAprovado2.Text = lblQtPorTurno2.Text
            OPRetrabalho2 = 0
        End If
    End Sub

    Private Sub rbRT3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT3.CheckedChanged, rbRF3.CheckedChanged, rbLC3.CheckedChanged
        If rbRT3.Checked = True Then
            txtQTReprovado3.Clear()
            txtQTAprovado3.Clear()
            txtQTReprovado3.Enabled = True
            txtQTAprovado3.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho3 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF3.Checked = True Then
            txtQTAprovado3.Text = "0"
            txtQTReprovado3.Enabled = False
            txtQTAprovado3.Enabled = False
            txtQTReprovado3.Text = lblQtPorTurno3.Text
            OPRetrabalho3 = 0
        ElseIf rbLC3.Checked = True Then
            txtQTReprovado3.Text = "0"
            txtQTReprovado3.Enabled = False
            txtQTAprovado3.Enabled = False
            txtQTAprovado3.Text = lblQtPorTurno3.Text
            OPRetrabalho3 = 0
        End If
    End Sub

    Private Sub rbRT4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT4.CheckedChanged, rbRF4.CheckedChanged, rbLC4.CheckedChanged
        If rbRT4.Checked = True Then
            txtQTReprovado4.Clear()
            txtQTAprovado4.Clear()
            txtQTReprovado4.Enabled = True
            txtQTAprovado4.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho4 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF4.Checked = True Then
            txtQTAprovado4.Text = "0"
            txtQTReprovado4.Enabled = False
            txtQTAprovado4.Enabled = False
            txtQTReprovado4.Text = lblQtPorTurno4.Text
            OPRetrabalho4 = 0
        ElseIf rbLC4.Checked = True Then
            txtQTReprovado4.Text = "0"
            txtQTReprovado4.Enabled = False
            txtQTAprovado4.Enabled = False
            txtQTAprovado4.Text = lblQtPorTurno4.Text
            OPRetrabalho4 = 0
        End If
    End Sub

    Private Sub rbRT5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT5.CheckedChanged, rbRF5.CheckedChanged, rbLC5.CheckedChanged
        If rbRT5.Checked = True Then
            txtQTReprovado5.Clear()
            txtQTAprovado5.Clear()
            txtQTReprovado5.Enabled = True
            txtQTAprovado5.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho5 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF5.Checked = True Then
            txtQTAprovado5.Text = "0"
            txtQTReprovado5.Enabled = False
            txtQTAprovado5.Enabled = False
            txtQTReprovado5.Text = lblQtPorTurno5.Text
            OPRetrabalho5 = 0
        ElseIf rbLC5.Checked = True Then
            txtQTReprovado5.Text = "0"
            txtQTReprovado5.Enabled = False
            txtQTAprovado5.Enabled = False
            txtQTAprovado5.Text = lblQtPorTurno5.Text
            OPRetrabalho5 = 0
        End If
    End Sub

    Private Sub rbRT6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT6.CheckedChanged, rbRF6.CheckedChanged, rbLC6.CheckedChanged
        If rbRT6.Checked = True Then
            txtQTReprovado6.Clear()
            txtQTAprovado6.Clear()
            txtQTReprovado6.Enabled = True
            txtQTAprovado6.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho6 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF6.Checked = True Then
            txtQTAprovado6.Text = "0"
            txtQTReprovado6.Enabled = False
            txtQTAprovado6.Enabled = False
            txtQTReprovado6.Text = lblQtPorTurno6.Text
            OPRetrabalho6 = 0
        ElseIf rbLC6.Checked = True Then
            txtQTReprovado6.Text = "0"
            txtQTReprovado6.Enabled = False
            txtQTAprovado6.Enabled = False
            txtQTAprovado6.Text = lblQtPorTurno6.Text
            OPRetrabalho6 = 0
        End If
    End Sub

    Private Sub rbRT7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT7.CheckedChanged, rbRF7.CheckedChanged, rbLC7.CheckedChanged
        If rbRT7.Checked = True Then
            txtQTReprovado7.Clear()
            txtQTAprovado7.Clear()
            txtQTReprovado7.Enabled = True
            txtQTAprovado7.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho7 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF7.Checked = True Then
            txtQTAprovado7.Text = "0"
            txtQTReprovado7.Enabled = False
            txtQTAprovado7.Enabled = False
            txtQTReprovado7.Text = lblQtPorTurno7.Text
            OPRetrabalho7 = 0
        ElseIf rbLC7.Checked = True Then
            txtQTReprovado7.Text = "0"
            txtQTReprovado7.Enabled = False
            txtQTAprovado7.Enabled = False
            txtQTAprovado7.Text = lblQtPorTurno7.Text
            OPRetrabalho7 = 0
        End If
    End Sub

    Private Sub rbRT8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT8.CheckedChanged, rbRF8.CheckedChanged, rbLC8.CheckedChanged
        If rbRT8.Checked = True Then
            txtQTReprovado8.Clear()
            txtQTAprovado8.Clear()
            txtQTReprovado8.Enabled = True
            txtQTAprovado8.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho8 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF8.Checked = True Then
            txtQTAprovado8.Text = "0"
            txtQTReprovado8.Enabled = False
            txtQTAprovado8.Enabled = False
            txtQTReprovado8.Text = lblQtPorTurno8.Text
            OPRetrabalho8 = 0
        ElseIf rbLC8.Checked = True Then
            txtQTReprovado8.Text = "0"
            txtQTReprovado8.Enabled = False
            txtQTAprovado8.Enabled = False
            txtQTAprovado8.Text = lblQtPorTurno8.Text
            OPRetrabalho8 = 0
        End If
    End Sub

    Private Sub rbRT9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT9.CheckedChanged, rbRF9.CheckedChanged, rbLC9.CheckedChanged
        If rbRT9.Checked = True Then
            txtQTReprovado9.Clear()
            txtQTAprovado9.Clear()
            txtQTReprovado9.Enabled = True
            txtQTAprovado9.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho9 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF9.Checked = True Then
            txtQTAprovado9.Text = "0"
            txtQTReprovado9.Enabled = False
            txtQTAprovado9.Enabled = False
            txtQTReprovado9.Text = lblQtPorTurno9.Text
            OPRetrabalho9 = 0
        ElseIf rbLC9.Checked = True Then
            txtQTReprovado9.Text = "0"
            txtQTReprovado9.Enabled = False
            txtQTAprovado9.Enabled = False
            txtQTAprovado9.Text = lblQtPorTurno9.Text
            OPRetrabalho9 = 0
        End If
    End Sub

    Private Sub rbRT10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRT10.CheckedChanged, rbRF10.CheckedChanged, rbLC10.CheckedChanged
        If rbRT10.Checked = True Then
            txtQTReprovado10.Clear()
            txtQTAprovado10.Clear()
            txtQTReprovado10.Enabled = True
            txtQTAprovado10.Enabled = True
            If txtOPRetrabalho.TextLength = 0 Then
            Else
                OPRetrabalho10 = txtOPRetrabalho.Text
            End If
        ElseIf rbRF10.Checked = True Then
            txtQTAprovado10.Text = "0"
            txtQTReprovado10.Enabled = False
            txtQTAprovado10.Enabled = False
            txtQTReprovado10.Text = lblQtPorTurno10.Text
            OPRetrabalho10 = 0
        ElseIf rbLC10.Checked = True Then
            txtQTReprovado10.Text = "0"
            txtQTReprovado10.Enabled = False
            txtQTAprovado10.Enabled = False
            txtQTAprovado10.Text = lblQtPorTurno10.Text
            OPRetrabalho10 = 0
        End If
    End Sub
    Sub Permicao()
        TesteAbertoRNC()
        Try

            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet

            conRNC.Open()
            Dim sel As String = "SELECT TOP 10 RNC, Disposicao FROM tblRNC WHERE RNC = " & lblRNC.Text & " and Disposicao like 'Sem Disposição'"
            da = New OleDbDataAdapter(sel, conRNC)
            ds.Clear()
            da.Fill(ds, "tblRNC")
            If ds.Tables("tblRNC").Rows.Count = 0 Then
                Posicao = ""
            Else
                Posicao = ds.Tables("tblRNC").Rows(0).Item("Disposicao")
            End If
            conRNC.Close()
        Catch ex As Exception
            Beep()
            MsgBox("Erro 1frt " & ex.Message)
            conRNC.Close()
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterarStatus.Click

        If txtOP.Text = "" Or txtOP.Text = "00" Or txtOP.Text = "000" Or txtOP.Text = "0000" Or txtOP.Text = "00000" Or txtOP.Text = "000000" Then
            MsgBox("Selecione uma RNC Válida na tabela abaixo", , "Selecionar RNC")
        Else
            If btAlterarStatus.Text = "Alterar Status" Then
                Permicao()
                If Posicao <> "Sem Disposição" Then

                    If (MsgBox("Deseja Alterar o Status da OP?", vbYesNo, "Alteração do Status")) = vbYes Then

                        btAlterarStatus.Text = "Aplicar"
                        txtOPRetrabalho.Focus()
                        Valor1 = 0 And Valor2 = 0 And Valor3 = 0 And Valor4 = 0 And Valor5 = 0 And Valor6 = 0 And Valor7 = 0 And Valor8 = 0 And Valor9 = 0 And Valor10 = 0
                        ValorX1 = 0 And ValorX2 = 0 And ValorX3 = 0 And ValorX4 = 0 And ValorX5 = 0 And ValorX6 = 0 And ValorX7 = 0 And ValorX8 = 0 And ValorX9 = 0 And ValorX10 = 0
                    End If
                Else
                    MsgBox("Não é permitido alterar um Státus sem a disposição da Supervisão!", , "Disposição da Supervisão")
                End If

            Else

                If rbRT1.Checked = False And rbRT2.Checked = False And rbRT3.Checked = False And rbRT4.Checked = False And rbRT5.Checked = False And rbRT6.Checked = False And rbRT7.Checked = False And rbRT8.Checked = False And rbRT9.Checked = False And rbRT10.Checked = False Then
                    ChecarVerificarAlterar()
                    Atualizar()
                Else
                    ChecarCampos()
                    If ValorXT > 0 Then
                        If ValorT <= 100 Then
                            If lblCodProduto.Text = compara Then
                                ChecarVerificarAlterar()
                                Atualizar()
                            Else
                                MsgBox("As OPs 'Reprovada vs Retrabalho' não se tratam do mesmo Produto, favor verificar!", , "Divergência de OPs")
                            End If
                        ElseIf ValorT >= 200 Then
                            If lblCodProduto.Text = compara Then
                            Else
                                MsgBox("As OPs 'Reprovada vs Retrabalho' não se tratam do mesmo Produto, favor verificar!", , "Divergência de OPs")
                            End If
                        End If
                    Else
                    End If
                End If
            End If
        End If
    End Sub

    Sub ChecarCampos()

        If txtQTReprovado1.Text = "" Then
            txtQTReprovado1.Text = 0
        End If
        If txtQTAprovado1.Text = "" Then
            txtQTAprovado1.Text = 0
        End If

        If txtQTReprovado2.Text = "" Then
            txtQTReprovado2.Text = 0
        End If
        If txtQTAprovado2.Text = "" Then
            txtQTAprovado2.Text = 0
        End If

        If txtQTReprovado3.Text = "" Then
            txtQTReprovado3.Text = 0
        End If
        If txtQTAprovado3.Text = "" Then
            txtQTAprovado3.Text = 0
        End If

        If txtQTReprovado4.Text = "" Then
            txtQTReprovado4.Text = 0
        End If
        If txtQTAprovado4.Text = "" Then
            txtQTAprovado4.Text = 0
        End If

        If txtQTReprovado5.Text = "" Then
            txtQTReprovado5.Text = 0
        End If
        If txtQTAprovado5.Text = "" Then
            txtQTAprovado5.Text = 0
        End If

        If txtQTReprovado6.Text = "" Then
            txtQTReprovado6.Text = 0
        End If
        If txtQTAprovado6.Text = "" Then
            txtQTAprovado6.Text = 0
        End If

        If txtQTReprovado7.Text = "" Then
            txtQTReprovado7.Text = 0
        End If
        If txtQTAprovado7.Text = "" Then
            txtQTAprovado7.Text = 0
        End If

        If txtQTReprovado8.Text = "" Then
            txtQTReprovado8.Text = 0
        End If
        If txtQTAprovado8.Text = "" Then
            txtQTAprovado8.Text = 0
        End If

        If txtQTReprovado9.Text = "" Then
            txtQTReprovado9.Text = 0
        End If
        If txtQTAprovado9.Text = "" Then
            txtQTAprovado9.Text = 0
        End If

        If txtQTReprovado10.Text = "" Then
            txtQTReprovado10.Text = 0
        End If
        If txtQTAprovado10.Text = "" Then
            txtQTAprovado10.Text = 0
        End If


        If rb1T.Checked = True Then
            ValorXT = 10
        End If
        If rb2T.Checked = True Then
            ValorXT = 20
        End If
        If rb3T.Checked = True Then
            ValorXT = 30
        End If
        If rb4T.Checked = True Then
            ValorXT = 40
        End If
        If rb5T.Checked = True Then
            ValorXT = 50
        End If
        If rb6T.Checked = True Then
            ValorXT = 60
        End If
        If rb7T.Checked = True Then
            ValorXT = 70
        End If
        If rb8T.Checked = True Then
            ValorXT = 80
        End If
        If rb9T.Checked = True Then
            ValorXT = 90
        End If
        If rb10T.Checked = True Then
            ValorXT = 100
        End If


        If lblQtPorTurno1.Text <> Integer.Parse(txtQTReprovado1.Text + Integer.Parse(txtQTAprovado1.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(1)' e 'Aprovado(1)' não somam a Quantidade Reclamada(1). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor1 = 10 para não reparar e avançar
                Valor1 = 10
            Else
                'valor1 = 200 para reparar e parar
                Valor1 = 200
            End If
        Else
            'valorx1 = 10 para continuar
            ValorX1 = 10
        End If

        If lblQtPorTurno2.Text <> Integer.Parse(txtQTReprovado2.Text + Integer.Parse(txtQTAprovado2.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(2)' e 'Aprovado(2)' não somam a Quantidade Reclamada(2). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor2 = 20 para não reparar e avançar
                Valor2 = 10
            Else
                'valor2 = 200 para reparar e parar
                Valor2 = 200
            End If
        Else
            'valorx2 = 20 para continuar
            ValorX2 = 10
        End If

        If lblQtPorTurno3.Text <> Integer.Parse(txtQTReprovado3.Text + Integer.Parse(txtQTAprovado3.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(3)' e 'Aprovado(3)' não somam a Quantidade Reclamada(3). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor3 = 30 para não reparar e avançar
                Valor3 = 10
            Else
                'valor3 = 200 para reparar e parar
                Valor3 = 200
            End If
        Else
            'valorx3 = 30 para continuar
            ValorX3 = 10
        End If

        If lblQtPorTurno4.Text <> Integer.Parse(txtQTReprovado4.Text + Integer.Parse(txtQTAprovado4.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(4)' e 'Aprovado(4)' não somam a Quantidade Reclamada(4). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor4 = 40 para não reparar e avançar
                Valor4 = 10
            Else
                'valor4 = 200 para reparar e parar
                Valor4 = 200
            End If
        Else
            'valorx4 = 40 para continuar
            ValorX4 = 10
        End If

        If lblQtPorTurno5.Text <> Integer.Parse(txtQTReprovado5.Text + Integer.Parse(txtQTAprovado5.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(5)' e 'Aprovado(5)' não somam a Quantidade Reclamada(5). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor5 = 50 para não reparar e avançar
                Valor5 = 10
            Else
                'valor5 = 200 para reparar e parar
                Valor5 = 200
            End If
        Else
            'valorx5 = 50 para continuar
            ValorX5 = 10
        End If

        If lblQtPorTurno6.Text <> Integer.Parse(txtQTReprovado6.Text + Integer.Parse(txtQTAprovado6.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(6)' e 'Aprovado(6)' não somam a Quantidade Reclamada(6). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor6 = 60 para não reparar e avançar
                Valor6 = 10
            Else
                'valor6 = 200 para reparar e parar
                Valor6 = 200
            End If
        Else
            'valorx6 = 60 para continuar
            ValorX6 = 10
        End If

        If lblQtPorTurno7.Text <> Integer.Parse(txtQTReprovado7.Text + Integer.Parse(txtQTAprovado7.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(7)' e 'Aprovado(7)' não somam a Quantidade Reclamada(7). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor7 = 70 para não reparar e avançar
                Valor7 = 10
            Else
                'valor7 = 200 para reparar e parar
                Valor7 = 200
            End If
        Else
            'valorx7 = 70 para continuar
            ValorX7 = 10
        End If

        If lblQtPorTurno8.Text <> Integer.Parse(txtQTReprovado8.Text + Integer.Parse(txtQTAprovado8.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(8)' e 'Aprovado(8)' não somam a Quantidade Reclamada(8). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor8 = 80 para não reparar e avançar
                Valor8 = 10
            Else
                'valor8 = 200 para reparar e parar
                Valor8 = 200
            End If
        Else
            'valorx8 = 80 para continuar
            ValorX8 = 10
        End If

        If lblQtPorTurno9.Text <> Integer.Parse(txtQTReprovado9.Text + Integer.Parse(txtQTAprovado9.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(9)' e 'Aprovado(9)' não somam a Quantidade Reclamada(9). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor9 = 90 para não reparar e avançar
                Valor9 = 10
            Else
                'valor9 = 200 para reparar e parar
                Valor9 = 200
            End If
        Else
            'valorx9 = 90 para continuar
            ValorX9 = 10
        End If

        If lblQtPorTurno10.Text <> Integer.Parse(txtQTReprovado10.Text + Integer.Parse(txtQTAprovado10.Text)) Then
            If (MsgBox("As Quantidades de peças 'Reprovado(10)' e 'Aprovado(10)' não somam a Quantidade Reclamada(10). Deseja Reparar? ", vbYesNo, "")) = vbNo Then
                'valor10 = 100 para não reparar e avançar 
                Valor10 = 10
            Else
                'valor10 = 200 para reparar e parar
                Valor10 = 200
            End If
        Else
            'valorx10 = 100 para continuar
            ValorX10 = 10
        End If




        ValorT = Valor1 + Valor2 + Valor3 + Valor4 + Valor5 + Valor6 + Valor7 + Valor8 + Valor9 + Valor10
        Dim x As Int16 = ValorX1 + ValorX2 + ValorX3 + ValorX4 + ValorX5 + ValorX6 + ValorX7 + ValorX8 + ValorX9 + ValorX10
        If x = 0 Then
        Else
            ValorXT = (ValorX1 + ValorX2 + ValorX3 + ValorX4 + ValorX5 + ValorX6 + ValorX7 + ValorX8 + ValorX9 + ValorX10) / ValorXT
        End If



    End Sub

    Sub ChecarVerificarAlterar()
        conRNC.Open()
        If rb1T.Checked = True Then
            VerificacaoStatus1()
            StatusAll = Status1
            AlterarStatus1()
        ElseIf rb2T.Checked = True Then
            VerificacaoStatus2()
            If Status1 = "Fechada" And Status2 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus2()
        ElseIf rb3T.Checked = True Then
            VerificacaoStatus3()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus3()
        ElseIf rb4T.Checked = True Then
            VerificacaoStatus4()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus4()
        ElseIf rb5T.Checked = True Then
            VerificacaoStatus5()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" And Status5 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus5()
        ElseIf rb6T.Checked = True Then
            VerificacaoStatus6()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" And Status5 = "Fechada" And Status6 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus6()
        ElseIf rb7T.Checked = True Then
            VerificacaoStatus7()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" And Status5 = "Fechada" And Status6 = "Fechada" And Status7 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus7()
        ElseIf rb8T.Checked = True Then
            VerificacaoStatus8()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" And Status5 = "Fechada" And Status6 = "Fechada" And Status7 = "Fechada" And Status8 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus8()
        ElseIf rb9T.Checked = True Then
            VerificacaoStatus9()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" And Status5 = "Fechada" And Status6 = "Fechada" And Status7 = "Fechada" And Status8 = "Fechada" And Status9 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus9()
        ElseIf rb10T.Checked = True Then
            VerificacaoStatus10()
            If Status1 = "Fechada" And Status2 = "Fechada" And Status3 = "Fechada" And Status4 = "Fechada" And Status5 = "Fechada" And Status6 = "Fechada" And Status7 = "Fechada" And Status8 = "Fechada" And Status9 = "Fechada" And Status10 = "Fechada" Then
                StatusAll = "Fechada"
            Else
                StatusAll = "Pendente"
            End If
            AlterarStatus10()
        End If
        conRNC.Close()
        MsgBox("Os dados foram alterados com sucesso. O status é: " & StatusAll & "", , "Status e Disposição")
        LimparDisposicao()
    End Sub

    Sub AlterarStatus1()
        If txtQTReprovado1.TextLength = 0 Then
            txtQTReprovado1.Text = 0
        End If
        If txtQTAprovado1.TextLength = 0 Then
            txtQTAprovado1.Text = 0
        End If
        If rbRT1.Checked = True Then
            OPRetrabalho1 = txtOPRetrabalho.Text
        Else
            OPRetrabalho1 = txtOP.Text
        End If
        Dim da100 As New OleDbDataAdapter
        Dim ds100 As New DataSet
        ds100 = New DataSet
        da100 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L1 & "', OP_Retrabalho = " & OPRetrabalho1 & ", QT_ReprovadoR = " & txtQTReprovado1.Text & ", QT_AprovadoR = " & txtQTAprovado1.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID1 & "", conRNC)
        ds100.Clear()
        da100.Fill(ds100, "tblRNC")
    End Sub

    Sub AlterarStatus2()
        If txtQTReprovado2.TextLength = 0 Then
            txtQTReprovado2.Text = 0
        End If
        If txtQTAprovado2.TextLength = 0 Then
            txtQTAprovado2.Text = 0
        End If
        If rbRT2.Checked = True Then
            OPRetrabalho2 = txtOPRetrabalho.Text
        Else
            OPRetrabalho2 = txtOP.Text
        End If
        Dim da110 As New OleDbDataAdapter
        Dim ds110 As New DataSet
        ds110 = New DataSet
        da110 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L2 & "', OP_Retrabalho = " & OPRetrabalho2 & ", QT_ReprovadoR = " & txtQTReprovado2.Text & ", QT_AprovadoR = " & txtQTAprovado2.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID2 & "", conRNC)
        ds110.Clear()
        da110.Fill(ds110, "tblRNC")
        AlterarStatus1()
    End Sub

    Sub AlterarStatus3()
        If txtQTReprovado3.TextLength = 0 Then
            txtQTReprovado3.Text = 0
        End If
        If txtQTAprovado3.TextLength = 0 Then
            txtQTAprovado3.Text = 0
        End If
        If rbRT3.Checked = True Then
            OPRetrabalho3 = txtOPRetrabalho.Text
        Else
            OPRetrabalho3 = txtOP.Text
        End If
        Dim da120 As New OleDbDataAdapter
        Dim ds120 As New DataSet
        ds120 = New DataSet
        da120 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L3 & "', OP_Retrabalho = " & OPRetrabalho3 & ", QT_ReprovadoR = " & txtQTReprovado3.Text & ", QT_AprovadoR = " & txtQTAprovado3.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID3 & "", conRNC)
        ds120.Clear()
        da120.Fill(ds120, "tblRNC")
        AlterarStatus2()
    End Sub

    Sub AlterarStatus4()
        If txtQTReprovado4.TextLength = 0 Then
            txtQTReprovado4.Text = 0
        End If
        If txtQTAprovado4.TextLength = 0 Then
            txtQTAprovado4.Text = 0
        End If
        If rbRT4.Checked = True Then
            OPRetrabalho4 = txtOPRetrabalho.Text
        Else
            OPRetrabalho4 = txtOP.Text
        End If
        Dim da130 As New OleDbDataAdapter
        Dim ds130 As New DataSet
        ds130 = New DataSet
        da130 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L4 & "', OP_Retrabalho = " & OPRetrabalho4 & ", QT_ReprovadoR = " & txtQTReprovado4.Text & ", QT_AprovadoR = " & txtQTAprovado4.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID4 & "", conRNC)
        ds130.Clear()
        da130.Fill(ds130, "tblRNC")
        AlterarStatus3()
    End Sub

    Sub AlterarStatus5()
        If txtQTReprovado5.TextLength = 0 Then
            txtQTReprovado5.Text = 0
        End If
        If txtQTAprovado5.TextLength = 0 Then
            txtQTAprovado5.Text = 0
        End If
        If rbRT5.Checked = True Then
            OPRetrabalho5 = txtOPRetrabalho.Text
        Else
            OPRetrabalho5 = txtOP.Text
        End If
        Dim da140 As New OleDbDataAdapter
        Dim ds140 As New DataSet
        ds140 = New DataSet
        da140 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L5 & "', OP_Retrabalho = " & OPRetrabalho5 & ", QT_ReprovadoR = " & txtQTReprovado5.Text & ", QT_AprovadoR = " & txtQTAprovado5.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID5 & "", conRNC)
        ds140.Clear()
        da140.Fill(ds140, "tblRNC")
        AlterarStatus4()
    End Sub

    Sub AlterarStatus6()
        If txtQTReprovado6.TextLength = 0 Then
            txtQTReprovado6.Text = 0
        End If
        If txtQTAprovado6.TextLength = 0 Then
            txtQTAprovado6.Text = 0
        End If
        If rbRT6.Checked = True Then
            OPRetrabalho6 = txtOPRetrabalho.Text
        Else
            OPRetrabalho6 = txtOP.Text
        End If
        Dim da150 As New OleDbDataAdapter
        Dim ds150 As New DataSet
        ds150 = New DataSet
        da150 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L6 & "', OP_Retrabalho = " & OPRetrabalho6 & ", QT_ReprovadoR = " & txtQTReprovado6.Text & ", QT_AprovadoR = " & txtQTAprovado6.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID6 & "", conRNC)
        ds150.Clear()
        da150.Fill(ds150, "tblRNC")
        AlterarStatus5()
    End Sub

    Sub AlterarStatus7()
        If txtQTReprovado7.TextLength = 0 Then
            txtQTReprovado7.Text = 0
        End If
        If txtQTAprovado7.TextLength = 0 Then
            txtQTAprovado7.Text = 0
        End If
        If rbRT7.Checked = True Then
            OPRetrabalho7 = txtOPRetrabalho.Text
        Else
            OPRetrabalho7 = txtOP.Text
        End If
        Dim da160 As New OleDbDataAdapter
        Dim ds160 As New DataSet
        ds160 = New DataSet
        da160 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L7 & "', OP_Retrabalho = " & OPRetrabalho7 & ", QT_ReprovadoR = " & txtQTReprovado7.Text & ", QT_AprovadoR = " & txtQTAprovado7.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID7 & "", conRNC)
        ds160.Clear()
        da160.Fill(ds160, "tblRNC")
        AlterarStatus6()
    End Sub

    Sub AlterarStatus8()
        If txtQTReprovado8.TextLength = 0 Then
            txtQTReprovado8.Text = 0
        End If
        If txtQTAprovado8.TextLength = 0 Then
            txtQTAprovado8.Text = 0
        End If
        If rbRT8.Checked = True Then
            OPRetrabalho8 = txtOPRetrabalho.Text
        Else
            OPRetrabalho8 = txtOP.Text
        End If
        Dim da170 As New OleDbDataAdapter
        Dim ds170 As New DataSet
        ds170 = New DataSet
        da170 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L8 & "', OP_Retrabalho = " & OPRetrabalho8 & ", QT_ReprovadoR = " & txtQTReprovado8.Text & ", QT_AprovadoR = " & txtQTAprovado8.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID8 & "", conRNC)
        ds170.Clear()
        da170.Fill(ds170, "tblRNC")
        AlterarStatus7()
    End Sub

    Sub AlterarStatus9()
        If txtQTReprovado9.TextLength = 0 Then
            txtQTReprovado9.Text = 0
        End If
        If txtQTAprovado9.TextLength = 0 Then
            txtQTAprovado9.Text = 0
        End If
        If rbRT9.Checked = True Then
            OPRetrabalho9 = txtOPRetrabalho.Text
        Else
            OPRetrabalho9 = txtOP.Text
        End If
        Dim da180 As New OleDbDataAdapter
        Dim ds180 As New DataSet
        ds180 = New DataSet
        da180 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L9 & "', OP_Retrabalho = " & OPRetrabalho9 & ", QT_ReprovadoR = " & txtQTReprovado9.Text & ", QT_AprovadoR = " & txtQTAprovado9.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID9 & "", conRNC)
        ds180.Clear()
        da180.Fill(ds180, "tblRNC")
        AlterarStatus8()
    End Sub

    Sub AlterarStatus10()
        If txtQTReprovado10.TextLength = 0 Then
            txtQTReprovado10.Text = 0
        End If
        If txtQTAprovado10.TextLength = 0 Then
            txtQTAprovado10.Text = 0
        End If
        If rbRT10.Checked = True Then
            OPRetrabalho10 = txtOPRetrabalho.Text
        Else
            OPRetrabalho10 = txtOP.Text
        End If
        Dim da190 As New OleDbDataAdapter
        Dim ds190 As New DataSet
        ds190 = New DataSet
        da190 = New OleDbDataAdapter("UPDATE tblRNC SET  Status = '" & StatusAll & "', Disposicao = '" & L10 & "', OP_Retrabalho = " & OPRetrabalho10 & ", QT_ReprovadoR = " & txtQTReprovado10.Text & ", QT_AprovadoR = " & txtQTAprovado10.Text & ", Data_Encerramento = '" & Today.ToShortDateString & "' WHERE ID = " & ID10 & "", conRNC)
        ds190.Clear()
        da190.Fill(ds190, "tblRNC")
        AlterarStatus9()
    End Sub

    Sub VerificacaoStatus1()
        If rbRF1.Checked = True Then
            L1 = "Refugar"
            Status1 = "Fechada"
        ElseIf rbRT1.Checked = True Then
            L1 = "Retrabalhar"

            If ValorX1 = 10 Then
                Status1 = "Fechada"
            ElseIf Valor1 = 10 Then
                Status1 = "Fechada"
            Else
                Status1 = "Pendente"
            End If

        ElseIf rbLC1.Checked = True Then
            L1 = "Liberado Condicional"
            Status1 = "Fechada"
        End If
    End Sub

    Sub VerificacaoStatus2()
        If rbRF2.Checked = True Then
            L2 = "Refugar"
            Status2 = "Fechada"
        ElseIf rbRT2.Checked = True Then
            L2 = "Retrabalhar"
            If ValorX2 = 10 Then
                Status2 = "Fechada"
            ElseIf Valor2 = 10 Then
                Status2 = "Fechada"
            Else
                Status2 = "Pendente"
            End If
        ElseIf rbLC2.Checked = True Then
            L2 = "Liberado Condicional"
            Status2 = "Fechada"
        End If
        VerificacaoStatus1()
    End Sub

    Sub VerificacaoStatus3()
        If rbRF3.Checked = True Then
            L3 = "Refugar"
            Status3 = "Fechada"
        ElseIf rbRT3.Checked = True Then
            L3 = "Retrabalhar"
            If ValorX3 = 10 Then
                Status3 = "Fechada"
            ElseIf Valor3 = 10 Then
                Status3 = "Fechada"
            Else
                Status3 = "Pendente"
            End If
        ElseIf rbLC3.Checked = True Then
            L3 = "Liberado Condicional"
            Status3 = "Fechada"
        End If
        VerificacaoStatus2()
    End Sub

    Sub VerificacaoStatus4()
        If rbRF4.Checked = True Then
            L4 = "Refugar"
            Status4 = "Fechada"
        ElseIf rbRT4.Checked = True Then
            L4 = "Retrabalhar"
            If ValorX4 = 10 Then
                Status4 = "Fechada"
            ElseIf Valor4 = 10 Then
                Status4 = "Fechada"
            Else
                Status4 = "Pendente"
            End If
        ElseIf rbLC4.Checked = True Then
            L4 = "Liberado Condicional"
            Status4 = "Fechada"
        End If
        VerificacaoStatus3()
    End Sub

    Sub VerificacaoStatus5()
        If rbRF5.Checked = True Then
            L5 = "Refugar"
            Status5 = "Fechada"
        ElseIf rbRT5.Checked = True Then
            L5 = "Retrabalhar"
            If ValorX5 = 10 Then
                Status5 = "Fechada"
            ElseIf Valor5 = 10 Then
                Status5 = "Fechada"
            Else
                Status5 = "Pendente"
            End If
        ElseIf rbLC5.Checked = True Then
            L5 = "Liberado Condicional"
            Status5 = "Fechada"
        End If
        VerificacaoStatus4()
    End Sub

    Sub VerificacaoStatus6()
        If rbRF6.Checked = True Then
            L6 = "Refugar"
            Status6 = "Fechada"
        ElseIf rbRT6.Checked = True Then
            L6 = "Retrabalhar"
            If ValorX6 = 10 Then
                Status6 = "Fechada"
            ElseIf Valor6 = 10 Then
                Status6 = "Fechada"
            Else
                Status6 = "Pendente"
            End If
        ElseIf rbLC6.Checked = True Then
            L6 = "Liberado Condicional"
            Status6 = "Fechada"
        End If
        VerificacaoStatus5()
    End Sub

    Sub VerificacaoStatus7()
        If rbRF7.Checked = True Then
            L7 = "Refugar"
            Status7 = "Fechada"
        ElseIf rbRT7.Checked = True Then
            L7 = "Retrabalhar"
            If ValorX7 = 10 Then
                Status7 = "Fechada"
            ElseIf Valor7 = 10 Then
                Status7 = "Fechada"
            Else
                Status7 = "Pendente"
            End If
        ElseIf rbLC7.Checked = True Then
            L7 = "Liberado Condicional"
            Status7 = "Fechada"
        End If
        VerificacaoStatus6()
    End Sub

    Sub VerificacaoStatus8()
        If rbRF8.Checked = True Then
            L8 = "Refugar"
            Status8 = "Fechada"
        ElseIf rbRT8.Checked = True Then
            L8 = "Retrabalhar"
            If ValorX8 = 10 Then
                Status8 = "Fechada"
            ElseIf Valor8 = 10 Then
                Status8 = "Fechada"
            Else
                Status8 = "Pendente"
            End If
        ElseIf rbLC8.Checked = True Then
            L8 = "Liberado Condicional"
            Status8 = "Fechada"
        End If
        VerificacaoStatus7()
    End Sub

    Sub VerificacaoStatus9()
        If rbRF9.Checked = True Then
            L9 = "Refugar"
            Status9 = "Fechada"
        ElseIf rbRT9.Checked = True Then
            L9 = "Retrabalhar"
            If ValorX9 = 10 Then
                Status9 = "Fechada"
            ElseIf Valor9 = 10 Then
                Status9 = "Fechada"
            Else
                Status9 = "Pendente"
            End If
        ElseIf rbLC9.Checked = True Then
            L9 = "Liberado Condicional"
            Status9 = "Fechada"
        End If
        VerificacaoStatus8()
    End Sub

    Sub VerificacaoStatus10()
        If rbRF10.Checked = True Then
            L10 = "Refugar"
            Status10 = "Fechada"
        ElseIf rbRT10.Checked = True Then
            L10 = "Retrabalhar"
            If ValorX10 = 10 Then
                Status10 = "Fechada"
            ElseIf Valor10 = 10 Then
                Status10 = "Fechada"
            Else
                Status10 = "Pendente"
            End If
        ElseIf rbLC10.Checked = True Then
            L10 = "Liberado Condicional"
            Status10 = "Fechada"
        End If
        VerificacaoStatus9()
    End Sub

    Sub LimparDisposicao()

        btAlterarStatus.Text = "Alterar Status"
        txtOPRetrabalho.Clear()
        lblStatus.Text = "*"

        txtQTReprovado1.Clear()
        txtQTAprovado1.Clear()
        txtQTReprovado1.Enabled = True
        txtQTAprovado1.Enabled = True
        rbRF1.Checked = False
        rbRT1.Checked = False
        rbLC1.Checked = False

        txtQTReprovado2.Clear()
        txtQTAprovado2.Clear()
        txtQTReprovado2.Enabled = True
        txtQTAprovado2.Enabled = True
        rbRF2.Checked = False
        rbRT2.Checked = False
        rbLC2.Checked = False

        txtQTReprovado3.Clear()
        txtQTAprovado3.Clear()
        txtQTReprovado3.Enabled = True
        txtQTAprovado3.Enabled = True
        rbRF3.Checked = False
        rbRT3.Checked = False
        rbLC3.Checked = False

        txtQTReprovado4.Clear()
        txtQTAprovado4.Clear()
        txtQTReprovado4.Enabled = True
        txtQTAprovado4.Enabled = True
        rbRF4.Checked = False
        rbRT4.Checked = False
        rbLC4.Checked = False

        txtQTReprovado5.Clear()
        txtQTAprovado5.Clear()
        txtQTReprovado5.Enabled = True
        txtQTAprovado5.Enabled = True
        rbRF5.Checked = False
        rbRT5.Checked = False
        rbLC5.Checked = False

        txtQTReprovado6.Clear()
        txtQTAprovado6.Clear()
        txtQTReprovado6.Enabled = True
        txtQTAprovado6.Enabled = True
        rbRF6.Checked = False
        rbRT6.Checked = False
        rbLC6.Checked = False

        txtQTReprovado7.Clear()
        txtQTAprovado7.Clear()
        txtQTReprovado7.Enabled = True
        txtQTAprovado7.Enabled = True
        rbRF7.Checked = False
        rbRT7.Checked = False
        rbLC7.Checked = False

        txtQTReprovado8.Clear()
        txtQTAprovado8.Clear()
        txtQTReprovado8.Enabled = True
        txtQTAprovado8.Enabled = True
        rbRF8.Checked = False
        rbRT8.Checked = False
        rbLC8.Checked = False

        txtQTReprovado9.Clear()
        txtQTAprovado9.Clear()
        txtQTReprovado9.Enabled = True
        txtQTAprovado9.Enabled = True
        rbRF9.Checked = False
        rbRT9.Checked = False
        rbLC9.Checked = False

        txtQTReprovado10.Clear()
        txtQTAprovado10.Clear()
        txtQTReprovado10.Enabled = True
        txtQTAprovado10.Enabled = True
        rbRF10.Checked = False
        rbRT10.Checked = False
        rbLC10.Checked = False

    End Sub

    Private Sub Label42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LimparDisposicao()
    End Sub

    Private Sub Label45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        rbRT1.Checked = True
        rbRT2.Checked = True
        rbRT3.Checked = True
        rbRT4.Checked = True
        rbRT5.Checked = True
        rbRT6.Checked = True
        rbRT7.Checked = True
        rbRT8.Checked = True
        rbRT9.Checked = True
        rbRT10.Checked = True

    End Sub

    Private Sub Label44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        rbRF1.Checked = True
        rbRF2.Checked = True
        rbRF3.Checked = True
        rbRF4.Checked = True
        rbRF5.Checked = True
        rbRF6.Checked = True
        rbRF7.Checked = True
        rbRF8.Checked = True
        rbRF9.Checked = True
        rbRF10.Checked = True

    End Sub

    Private Sub Label46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        rbLC1.Checked = True
        rbLC2.Checked = True
        rbLC3.Checked = True
        rbLC4.Checked = True
        rbLC6.Checked = True
        rbLC7.Checked = True
        rbLC8.Checked = True
        rbLC9.Checked = True
        rbLC10.Checked = True

    End Sub

    Private Sub RT(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Cursor = Cursors.Hand
    End Sub

    Private Sub sair(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.MouseEnter
        Cursor = Cursors.Default
    End Sub

    Private Sub Defeito(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btDefeito.Click
        frmDefeito.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        frmMaquina.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRE.Click
        frmRE.ShowDialog()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        frmPecasVolume.ShowDialog()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btImprimirEtiqueta.Click
        Try
            If lblRNC.Text = "*" Or lblRNC.Text = "" Then
                MsgBox("Selecione uma RNC na tabela abaixo", , "Impressão de Etiquetas")
            Else
                ImprimirEtiqueta()
            End If
        Catch ex As Exception
            MsgBox("Erro 187 " & ex.Message)
        End Try

    End Sub

    Private Sub btExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles btExport.Click
        'TesteAbertoRNC()
        'TesteAbertoPlanilhaRNC_xlsx()

        Dim ds1, ds2 As New DataSet
        Dim da2 As New OleDbDataAdapter
        'carregar o excel num dataset
        Dim _conn As String = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\Gerenciamento de RNC 2013 - INTERNO.xlsx;Extended Properties=Excel 8.0")
        Dim _connection As OleDbConnection = New OleDbConnection(_conn)
        Dim da1 As OleDbDataAdapter = New OleDbDataAdapter()
        Dim _command As OleDbCommand = New OleDbCommand()
        _command.Connection = _connection
        _command.CommandText = "SELECT top 10 * FROM [Entrada2$] order by ID asc "
        da1.SelectCommand = _command
        da1.Fill(ds1, "Entrada2")
        _connection.Close()

        MsgBox("Linhas " & ds1.Tables(0).Rows.Count & " Colunas " & ds1.Tables(0).Columns.Count, , "")


        conRNC.Open()
        Dim sel As String = "Select top 10 * from tblRNC order by ID asc"
        'Dim sel As String = "select Contador, count (*) from tblRNC group by Contador order by contador desc" 'conta quantas RNCs exitem
        da2 = New OleDbDataAdapter(sel, conRNC)
        ds2.Clear()
        da2.Fill(ds2, "tblRNC")
        conRNC.Close()
        Dim Valor As Int16
        Valor = Int16.Parse(ds2.Tables(0).Rows(9).Item(0)) - Int16.Parse(ds1.Tables(0).Rows(9).Item(0))
        MsgBox(Valor, , "Linhas Acrecentar")



        Dim Excell As New Microsoft.Office.Interop.Excel.Application
        Dim Documento_xlsx As Microsoft.Office.Interop.Excel.Workbook
        Dim Planilha_do_Documento_xlsx As Microsoft.Office.Interop.Excel.Worksheet

        Dim RNC As Microsoft.Office.Interop.Excel.Range


        On Error GoTo ErrHandler

        '3º Abrir o arquivo Excel
        Documento_xlsx = Excell.Workbooks.Open("C:\Users\Cid\Documents\Projetos\BancoDados\Gerenciamento de RNC 2013 - INTERNO.xlsx")

        '4º Abrir a planilha para inserir texto
        Planilha_do_Documento_xlsx = Documento_xlsx.Sheets.Item("Entrada2")

        '5º Atribuir uma célula na planilha
        Dim i As Int16 = 0
        For Each dr In ds2.Tables(0).Rows

            Dim A As Microsoft.Office.Interop.Excel.Range
            A = Planilha_do_Documento_xlsx.Cells(2 + i, 1) ' = A2
            A.Value = dr(0)

            i = i + 1

        Next


        '7º Abrindo o excel
        Excell.Visible = False
        '8º Salvando a Planilha
        Documento_xlsx.Save()


        '9º encerra os processos EXCEL.EXE no gerenciador de tarefas do windows 
ExitHere:
        Excell.Quit()
        Exit Sub
ErrHandler:
        MsgBox(Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source, vbCritical, "Erro 86 ")
        Resume ExitHere

    End Sub


    Private Sub btSupervisao_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSupervisao.Click
        frmDisposicao.ShowDialog()
    End Sub

    Private Sub btCancelarRNC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelarRNC.Click

        TesteAbertoRNC()
        Try
            Dim da3 As New OleDbDataAdapter
            Dim ds3 As New DataSet
            Dim sel3 As String = "Select top 100 * from tblRNC where Cancelar = 'Cancelar' order by ID desc"
            da3 = New OleDbDataAdapter(sel3, conRNC)
            ds3.Clear()
            da3.Fill(ds3, "tblRNC")
            conRNC.Close()
            Me.DataGridView1.DataSource = ds3
            Me.DataGridView1.DataMember = "tblRNC"
            FormatacaoGrid()
            Call Limpar()
        Catch ex As Exception
            MsgBox("Erro 14 " & ex.Message)
        End Try
    End Sub

    Private Sub btAtualizar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAtualizar.Click
        Try
            Dim da As New OleDbDataAdapter
            Dim ds As New DataSet

            conRNC.Open()
            Dim sel As String = "Select top 100 * from tblRNC where Status = 'Pendente' and Disposicao <> 'Sem Disposição' order by ID desc"
            'Dim sel As String = "select Contador, count (*) from tblRNC group by Contador order by contador desc" 'conta quantas RNCs exitem
            da = New OleDbDataAdapter(sel, conRNC)
            ds.Clear()
            da.Fill(ds, "tblRNC")
            Me.DataGridView1.DataSource = ds
            Me.DataGridView1.DataMember = "tblRNC"
            FormatacaoGrid()
            lblData.Text = Today
            lblHora.Text = TimeOfDay.ToShortTimeString
            conRNC.Close()
        Catch ex As Exception
            Beep()
            MsgBox("Erro 1fim " & ex.Message)
        End Try
    End Sub


    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        frmEstatistica.ShowDialog()
    End Sub
End Class
