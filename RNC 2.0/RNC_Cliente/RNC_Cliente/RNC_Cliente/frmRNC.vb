Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Object
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports RNC_Cliente .Module1

Public Class frmRNC
    Dim conConsulta_OP As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conDefeito As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\RNC_Defeito.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conMaquina As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conPecasVolume As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conRE As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\RNC_RE.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conRNC As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\RNC_RNC.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conABRIR As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\ABRIR.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim cs As ConnectionState
    Dim Mes_ As String '
    Dim Cliente As String
    Dim cliente2 As String
    Dim Celula As String
    Dim Defeito1 As String
    Dim Defeito2 As String
    Dim Defeito3 As String
    Dim Defeito4 As String
    Dim Defeito5 As String
    Dim Defeito6 As String
    Dim Defeito7 As String
    Dim Defeito8 As String
    Dim Defeito9 As String
    Dim Defeito10 As String
    Dim Alteradu As String
    Dim seleccion3 As String
    Dim ID1, ID2, ID3, ID4, ID5, ID6, ID7, ID8, ID9, ID10 As Integer
    Dim Limpo As String
    Dim SMC As Int64 = 0
    'variaveis do fromto
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
    Dim AbrirFT As Boolean
    Dim RNC_RNC2 As Boolean

    Private Sub frmRNC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'RNC_MaquinaDataSet.tblMaquina' table. You can move, or remove it, as needed.
        If Today > "08/05/2015" Then
            'MsgBox("Contate o Programador: Cid (15) 981797980 - cidevangelista@hotmail.com")
            Close()
        Else
            Call Teste_AbertoFT()
            conABRIR.Open()
            Call PriMeiro_Passo()
            TesteAbertoRNC()
            Try
                Dim da As New OleDbDataAdapter
                Dim ds As New DataSet

                conRNC.Open()
                Dim sel As String = "Select top 100 * from tblRNC Where Data_Abertura like '" & Today.ToShortDateString & "' and Cancelar IS Null order by ID desc" 'porque a coluna CANCELAR??????????????????????
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
                conRNC.Close()
            Finally
                conRNC.Close()
            End Try
        End If
    End Sub
    'inicio FromTo
    Sub PriMeiro_Passo() 'Handles MyBase.Shown
        Try

            _connFT = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.xlsx;Extended Properties=Excel 8.0")
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
    '   Close()
    ' End Sub
    Sub Teste_AbertoFT()
        AbrirFT = TestFT("f:\Receb.Mat.Prima\Banco_Dados\ABRIR.accdb")
        ExcelFT = TestFT("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.xlsx")
        AccessFT = TestFT("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb")
        If ExcelFT = True Then
            MsgBox("O Arquivo Excel de importação está aberto, Feche-o para para continuar")
            Close()
        ElseIf AccessFT = True Then
            MsgBox("O Arquivo Access de importação está aberto, Feche-o para para continuar")
            Close()
        ElseIf AbrirFT = True Then
            MsgBox("NÃO É PERMITIDO ABRIR ATÉ QUE OUTRO USUÁRIO FECHE O PROGRAMA")
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
    ''fim FromTo
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
                    'verificação generica
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

                Else
                End If
                If btInserir.Text = "Aplicar" Then
                Else
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 2 " & ex.Message)
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

        Try

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
        Catch ex As Exception
            MsgBox("Erro JPW35 " & ex.Message)
        End Try
        SMC = 0
        SMC = CAIXA1 + CAIXA2 + CAIXA3 + CAIXA4 + CAIXA5 + CAIXA6 + CAIXA7 + CAIXA8 + CAIXA9 + CAIXA10
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
                'ElseIf lbldescricaornc10.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc10.Focus()
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
                'ElseIf lbldescricaornc9.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc9.Focus()
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
                ' ElseIf lbldescricaornc8.Text = "" Then
                '    MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '   lbldescricaornc8.Focus()
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
                'ElseIf lbldescricaornc7.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc7.Focus()
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
                'ElseIf lbldescricaornc6.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc6.Focus()
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
                'ElseIf lbldescricaornc5.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc5.Focus()
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
                'ElseIf lbldescricaornc4.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc4.Focus()
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
                'ElseIf lbldescricaornc3.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc3.Focus()
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
                ' ElseIf lbldescricaornc2.Text = "" Then
                '    MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '   lbldescricaornc2.Focus()
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
            ElseIf lblCodProduto.Text = "*" Or lblProduto.Text = "*" Then
                MsgBox("A OP não existe, peça que atualize o Banco de Dado", , "OP Reprovada")
                txtOP.Focus()
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
                'ElseIf lbldescricaornc1.Text = "" Then
                '   MsgBox("O campo 'Descrição da RNC' está vazio", , "Descrição da RNC")
                '  lbldescricaornc1.Focus()
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
                    'Call Atualizar()
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
            Dim sel3 As String = "Select top 100 * from tblRNC where  Data_Abertura like '" & Today.ToShortDateString & "' and Cancelar IS Null order by ID desc"
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
    Sub Inserir1()
        Try
            Dim da4 As New OleDbDataAdapter
            Dim ds4 As New DataSet
            Call Mes()
            Call Clientex()
            Call Celulax()
            ds4 = New DataSet
            da4 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb1Turno.Text & "', '" & txtCaixas1Turno.Text & "', " & txtQtCaixasReprovada1.Text & ", " & txtQTPorTurno1.Text & ", " & txtCodigoRNC1.Text & ", '" & lblDescricaoRNC1.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da5 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb2Turnos.Text & "', '" & txtCaixas2Turno.Text & "', " & txtQtCaixasReprovada2.Text & ", '" & txtQTPorTurno2.Text & "', " & txtCodigoRNC2.Text & ", '" & lblDescricaoRNC2.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da6 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb3Turnos.Text & "', '" & txtCaixas3Turno.Text & "', " & txtQtCaixasReprovada3.Text & ", '" & txtQTPorTurno3.Text & "', " & txtCodigoRNC3.Text & ", '" & lblDescricaoRNC3.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "','" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da7 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb4Turnos.Text & "', '" & txtCaixas4Turno.Text & "', " & txtQtCaixasReprovada4.Text & ", '" & txtQTPorTurno4.Text & "', " & txtCodigoRNC4.Text & ", '" & lblDescricaoRNC4.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da8 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb5Turno.Text & "', '" & txtCaixas5Turno.Text & "', " & txtQtCaixasReprovada5.Text & ", '" & txtQTPorTurno5.Text & "', " & txtCodigoRNC5.Text & ", '" & lblDescricaoRNC5.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da8 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb6Turno.Text & "', '" & txtCaixas6Turno.Text & "', " & txtQtCaixasReprovada6.Text & ", '" & txtQTPorTurno6.Text & "', " & txtCodigoRNC6.Text & ", '" & lblDescricaoRNC6.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da8 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb7Turno.Text & "', '" & txtCaixas7Turno.Text & "', " & txtQtCaixasReprovada7.Text & ", '" & txtQTPorTurno7.Text & "', " & txtCodigoRNC7.Text & ", '" & lblDescricaoRNC7.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da9 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb8Turno.Text & "', '" & txtCaixas8Turno.Text & "', " & txtQtCaixasReprovada8.Text & ", '" & txtQTPorTurno8.Text & "', " & txtCodigoRNC8.Text & ", '" & lblDescricaoRNC8.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da10 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb9Turno.Text & "', '" & txtCaixas9Turno.Text & "', " & txtQtCaixasReprovada9.Text & ", '" & txtQTPorTurno9.Text & "', " & txtCodigoRNC9.Text & ", '" & lblDescricaoRNC9.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
            da11 = New OleDbDataAdapter("INSERT INTO tblRNC (RNC, Status, Origem, Data_Abertura, Hora, Mes, Cod_Produto, Cliente, Produto, OP_Reprovado, Turno, NúmerosCaixas, QT_Caixas, QT_Reprovado, Cod_Defeito, Nao_Conformidade, Maquina, Celula, Observacao, RE, Inspetor, Setor, TurnoDetector) Values (" & lblRNC.Text & ", 'Pendente', '" & cbDetectado.Text & "', '" & lblData.Text & "', '" & lblHora.Text & "', '" & Mes_ & "', " & lblCodProduto.Text & ", '" & Cliente & "', '" & lblProduto.Text & "', " & txtOP.Text & ", '" & cb10Turno.Text & "', '" & txtCaixas10Turno.Text & "', " & txtQtCaixasReprovada10.Text & ", '" & txtQTPorTurno10.Text & "', " & txtCodigoRNC10.Text & ", '" & lblDescricaoRNC10.Text & "', '" & txtMaquina.Text & "', '" & Celula & "', '" & txtOBS.Text & "', " & txtRE.Text & ", '" & txtInspetor.Text & "', '" & txtSetor.Text & "', '" & cbTurno.Text & "') ", conRNC)
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
                cliente2 = Cliente
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
                lblCelula.Text = ds9.Tables("tblMaquina").Rows(0)("Celula")
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

            txtQTPorTurno1.Text = 0
            txtQTPorTurno2.Text = 0
            txtQTPorTurno3.Text = 0
            txtQTPorTurno4.Text = 0
            txtQTPorTurno5.Text = 0
            txtQTPorTurno6.Text = 0
            txtQTPorTurno7.Text = 0
            txtQTPorTurno8.Text = 0
            txtQTPorTurno9.Text = 0
            txtQTPorTurno10.Text = 0

            txtPecasPorVolume.Clear()
            lblTotalPecas.Text = 0
            txtCodigoRNC1.Clear()
            lblDescricaoRNC1.Text = ""
            txtCodigoRNC2.Clear()
            lblDescricaoRNC2.Text = ""
            txtCodigoRNC3.Clear()
            lblDescricaoRNC3.Text = ""
            txtCodigoRNC4.Clear()
            lblDescricaoRNC4.Text = ""
            txtCodigoRNC5.Clear()
            lblDescricaoRNC5.Text = ""
            txtCodigoRNC6.Clear()
            lblDescricaoRNC6.Text = ""
            txtCodigoRNC7.Clear()
            lblDescricaoRNC7.Text = ""
            txtCodigoRNC8.Clear()
            lblDescricaoRNC8.Text = ""
            txtCodigoRNC9.Clear()
            lblDescricaoRNC9.Text = ""
            txtCodigoRNC10.Clear()
            lblDescricaoRNC10.Text = ""


            txtOBS.Clear()
            txtRE.Clear()
            txtInspetor.Clear()
            btInserir.Text = "Inserir"
            btInserir.Enabled = True
            btAlterar.Text = "Alterar"
            'btAlterar.Enabled = true
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
            txtQTPorTurno1.Visible = True
            txtCodigoRNC1.Visible = True
            lblDescricaoRNC1.Visible = True
        Catch ex As Exception
            MsgBox("Erro 32 " & ex.Message)
        End Try
    End Sub
    Sub rb2v()
        Try
            cb2Turnos.Visible = True
            txtCaixas2Turno.Visible = True
            txtQtCaixasReprovada2.Visible = True
            txtQTPorTurno2.Visible = True
            txtCodigoRNC2.Visible = True
            lblDescricaoRNC2.Visible = True
        Catch ex As Exception
            MsgBox("Erro 33 " & ex.Message)
        End Try
    End Sub
    Sub rb3v()
        Try
            cb3Turnos.Visible = True
            txtCaixas3Turno.Visible = True
            txtQtCaixasReprovada3.Visible = True
            txtQTPorTurno3.Visible = True
            txtCodigoRNC3.Visible = True
            lblDescricaoRNC3.Visible = True
        Catch ex As Exception
            MsgBox("Erro 34 " & ex.Message)
        End Try
    End Sub
    Sub rb4v()
        Try
            cb4Turnos.Visible = True
            txtCaixas4Turno.Visible = True
            txtQtCaixasReprovada4.Visible = True
            txtQTPorTurno4.Visible = True
            txtCodigoRNC4.Visible = True
            lblDescricaoRNC4.Visible = True
        Catch ex As Exception
            MsgBox("Erro 35 " & ex.Message)
        End Try
    End Sub
    Sub rb5v()
        Try
            cb5Turno.Visible = True
            txtCaixas5Turno.Visible = True
            txtQtCaixasReprovada5.Visible = True
            txtQTPorTurno5.Visible = True
            txtCodigoRNC5.Visible = True
            lblDescricaoRNC5.Visible = True
        Catch ex As Exception
            MsgBox("Erro 36 " & ex.Message)
        End Try
    End Sub
    Sub rb6v()
        Try
            cb6Turno.Visible = True
            txtCaixas6Turno.Visible = True
            txtQtCaixasReprovada6.Visible = True
            txtQTPorTurno6.Visible = True
            txtCodigoRNC6.Visible = True
            lblDescricaoRNC6.Visible = True
        Catch ex As Exception
            MsgBox("Erro 37 " & ex.Message)
        End Try
    End Sub
    Sub rb7v()
        Try
            cb7Turno.Visible = True
            txtCaixas7Turno.Visible = True
            txtQtCaixasReprovada7.Visible = True
            txtQTPorTurno7.Visible = True
            txtCodigoRNC7.Visible = True
            lblDescricaoRNC7.Visible = True
        Catch ex As Exception
            MsgBox("Erro 38 " & ex.Message)
        End Try
    End Sub
    Sub rb8v()
        Try
            cb8Turno.Visible = True
            txtCaixas8Turno.Visible = True
            txtQtCaixasReprovada8.Visible = True
            txtQTPorTurno8.Visible = True
            txtCodigoRNC8.Visible = True
            lblDescricaoRNC8.Visible = True
        Catch ex As Exception
            MsgBox("Erro 39 " & ex.Message)
        End Try
    End Sub
    Sub rb9v()
        Try
            cb9Turno.Visible = True
            txtCaixas9Turno.Visible = True
            txtQtCaixasReprovada9.Visible = True
            txtQTPorTurno9.Visible = True
            txtCodigoRNC9.Visible = True
            lblDescricaoRNC9.Visible = True
        Catch ex As Exception
            MsgBox("Erro 40 " & ex.Message)
        End Try
    End Sub
    Sub rb10v()
        Try
            cb10Turno.Visible = True
            txtCaixas10Turno.Visible = True
            txtQtCaixasReprovada10.Visible = True
            txtQTPorTurno10.Visible = True
            txtCodigoRNC10.Visible = True
            lblDescricaoRNC10.Visible = True
        Catch ex As Exception
            MsgBox("Erro 41 " & ex.Message)
        End Try
    End Sub
    Sub rb1f()
        Try
            cb1Turno.Visible = False
            txtCaixas1Turno.Visible = False
            txtQtCaixasReprovada1.Visible = False
            txtQTPorTurno1.Visible = False
            txtCodigoRNC1.Visible = False
            lblDescricaoRNC1.Visible = False
        Catch ex As Exception
            MsgBox("Erro 42 " & ex.Message)
        End Try
    End Sub
    Sub rb2f()
        Try
            cb2Turnos.Visible = False
            txtCaixas2Turno.Visible = False
            txtQtCaixasReprovada2.Visible = False
            txtQTPorTurno2.Visible = False
            txtCodigoRNC2.Visible = False
            lblDescricaoRNC2.Visible = False
        Catch ex As Exception
            MsgBox("Erro 43 " & ex.Message)
        End Try
    End Sub
    Sub rb3f()
        Try
            cb3Turnos.Visible = False
            txtCaixas3Turno.Visible = False
            txtQtCaixasReprovada3.Visible = False
            txtQTPorTurno3.Visible = False
            txtCodigoRNC3.Visible = False
            lblDescricaoRNC3.Visible = False
        Catch ex As Exception
            MsgBox("Erro 44 " & ex.Message)
        End Try
    End Sub
    Sub rb4f()
        Try
            cb4Turnos.Visible = False
            txtCaixas4Turno.Visible = False
            txtQtCaixasReprovada4.Visible = False
            txtQTPorTurno4.Visible = False
            txtCodigoRNC4.Visible = False
            lblDescricaoRNC4.Visible = False
        Catch ex As Exception
            MsgBox("Erro 45 " & ex.Message)
        End Try
    End Sub
    Sub rb5f()
        Try
            cb5Turno.Visible = False
            txtCaixas5Turno.Visible = False
            txtQtCaixasReprovada5.Visible = False
            txtQTPorTurno5.Visible = False
            txtCodigoRNC5.Visible = False
            lblDescricaoRNC5.Visible = False
        Catch ex As Exception
            MsgBox("Erro 46 " & ex.Message)
        End Try
    End Sub
    Sub rb6f()
        Try
            cb6Turno.Visible = False
            txtCaixas6Turno.Visible = False
            txtQtCaixasReprovada6.Visible = False
            txtQTPorTurno6.Visible = False
            txtCodigoRNC6.Visible = False
            lblDescricaoRNC6.Visible = False
        Catch ex As Exception
            MsgBox("Erro 47 " & ex.Message)
        End Try
    End Sub
    Sub rb7f()
        Try
            cb7Turno.Visible = False
            txtCaixas7Turno.Visible = False
            txtQtCaixasReprovada7.Visible = False
            txtQTPorTurno7.Visible = False
            txtCodigoRNC7.Visible = False
            lblDescricaoRNC7.Visible = False
        Catch ex As Exception
            MsgBox("Erro 48 " & ex.Message)
        End Try
    End Sub
    Sub rb8f()
        Try
            cb8Turno.Visible = False
            txtCaixas8Turno.Visible = False
            txtQtCaixasReprovada8.Visible = False
            txtQTPorTurno8.Visible = False
            txtCodigoRNC8.Visible = False
            lblDescricaoRNC8.Visible = False
        Catch ex As Exception
            MsgBox("Erro 49 " & ex.Message)
        End Try
    End Sub
    Sub rb9f()
        Try
            cb9Turno.Visible = False
            txtCaixas9Turno.Visible = False
            txtQtCaixasReprovada9.Visible = False
            txtQTPorTurno9.Visible = False
            txtCodigoRNC9.Visible = False
            lblDescricaoRNC9.Visible = False
        Catch ex As Exception
            MsgBox("Erro 50 " & ex.Message)
        End Try
    End Sub
    Sub rb10f()
        Try
            cb10Turno.Visible = False
            txtCaixas10Turno.Visible = False
            txtQtCaixasReprovada10.Visible = False
            txtQTPorTurno10.Visible = False
            txtCodigoRNC10.Visible = False
            lblDescricaoRNC10.Visible = False
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

        If InStr("1234567890-atée=cixsplt,", Chr(Keyascii)) = 0 Then
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
    Function SEM_TREMA(ByVal Keyascii As Short) As Short

        If InStr("1234567890- QWERTYUIOP´`ASDFGHJKLÇ~^ZXCVBNMqwertyuiopasdfghjklzçxcvbnm,.;:/?*-+!()_", Chr(Keyascii)) = 0 Then
            SEM_TREMA = 0
        Else
            SEM_TREMA = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                SEM_TREMA = Keyascii
            Case 13
                SEM_TREMA = Keyascii
            Case 32 'permite espaço
                SEM_TREMA = Keyascii
        End Select
    End Function
    Private Sub SEM_TREMA(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDadoColuna.KeyPress, txtMaquina.KeyPress, txtOBS.KeyPress 'lbldescricaornc1.KeyPress  até o 10

        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(SEM_TREMA(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 599gh " & ex.Message)
        End Try
    End Sub

    '-------------
    Function SO_LETRAS(ByVal Keyascii As Short) As Short

        If InStr("QWERTYUIOPASDFGHJKLÇZXCVBNMqwertyuiopasdfghjklçzxcvbnm", Chr(Keyascii)) = 0 Then
            SO_LETRAS = 0
        Else
            SO_LETRAS = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                SO_LETRAS = Keyascii
            Case 13
                SO_LETRAS = Keyascii
            Case 32 'permite espaço
                SO_LETRAS = Keyascii
        End Select
    End Function
    Private Sub SO_LETRAS(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSetor.KeyPress, txtInspetor.KeyPress

        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(SO_LETRAS(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 59889gh " & ex.Message)
        End Try
    End Sub
    '---------------

    Private Sub Quantidades4(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOP.KeyPress, txtQTPorTurno1.KeyPress, txtQTPorTurno2.KeyPress, txtQTPorTurno3.KeyPress, txtQTPorTurno4.KeyPress, txtQTPorTurno5.KeyPress, txtQTPorTurno6.KeyPress, txtQTPorTurno7.KeyPress, txtQTPorTurno8.KeyPress, txtQTPorTurno9.KeyPress, txtQTPorTurno10.KeyPress, txtLinhas.KeyPress
        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(Numero4(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 55gh " & ex.Message)
        End Try
    End Sub
    Function Numero4(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            Numero4 = 0
        Else
            Numero4 = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                Numero4 = Keyascii
            Case 13
                Numero4 = Keyascii
                'Case 32 'permite espaço
                '   SoNumeros = Keyascii
        End Select
    End Function
    Private Sub Quantidades2(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCodigoRNC1.KeyPress, txtCodigoRNC2.KeyPress, txtCodigoRNC3.KeyPress, txtCodigoRNC4.KeyPress, txtCodigoRNC5.KeyPress, txtCodigoRNC6.KeyPress, txtCodigoRNC7.KeyPress, txtCodigoRNC8.KeyPress, txtCodigoRNC9.KeyPress, txtCodigoRNC10.KeyPress, txtRE.KeyPress, txtQtCaixasReprovada1.KeyPress, txtQtCaixasReprovada2.KeyPress, txtQtCaixasReprovada3.KeyPress, txtQtCaixasReprovada4.KeyPress, txtQtCaixasReprovada5.KeyPress, txtQtCaixasReprovada6.KeyPress, txtQtCaixasReprovada7.KeyPress, txtQtCaixasReprovada8.KeyPress, txtQtCaixasReprovada9.KeyPress, txtQtCaixasReprovada10.KeyPress
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
    Private Sub Quantidades3(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPecasPorVolume.KeyPress
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
        If InStr("1234567890,", Chr(Keyascii)) = 0 Then
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
                txtQTPorTurno1.Text = Double.Parse(txtQtCaixasReprovada1.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno1.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada2.TextLength > 0 Then
                txtQTPorTurno2.Text = Double.Parse(txtQtCaixasReprovada2.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno2.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada3.TextLength > 0 Then
                txtQTPorTurno3.Text = Double.Parse(txtQtCaixasReprovada3.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno3.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada4.TextLength > 0 Then
                txtQTPorTurno4.Text = Double.Parse(txtQtCaixasReprovada4.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno4.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada5.TextLength > 0 Then
                txtQTPorTurno5.Text = Double.Parse(txtQtCaixasReprovada5.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno5.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada6.TextLength > 0 Then
                txtQTPorTurno6.Text = Double.Parse(txtQtCaixasReprovada6.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno6.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada7.TextLength > 0 Then
                txtQTPorTurno7.Text = Double.Parse(txtQtCaixasReprovada7.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno7.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada8.TextLength > 0 Then
                txtQTPorTurno8.Text = Double.Parse(txtQtCaixasReprovada8.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno8.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada9.TextLength > 0 Then
                txtQTPorTurno9.Text = Double.Parse(txtQtCaixasReprovada9.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno9.Text = 0
            End If
            If txtPecasPorVolume.TextLength > 0 And txtQtCaixasReprovada10.TextLength > 0 Then
                txtQTPorTurno10.Text = Double.Parse(txtQtCaixasReprovada10.Text * txtPecasPorVolume.Text)
            Else
                txtQTPorTurno10.Text = 0
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

            x1 = txtQTPorTurno1.Text
            x2 = txtQTPorTurno2.Text
            x3 = txtQTPorTurno3.Text
            x4 = txtQTPorTurno4.Text
            x5 = txtQTPorTurno5.Text
            x6 = txtQTPorTurno6.Text
            x7 = txtQTPorTurno7.Text
            x8 = txtQTPorTurno8.Text
            x9 = txtQTPorTurno9.Text
            x10 = txtQTPorTurno10.Text



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

    Private Sub txtCodigoRNC1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC1.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC1.Text = Mcod1

        Defeito1 = ""
        Defeito1 = MRNC1
        lblDescricaoRNC1.Text = MRNC1 'incluir


    End Sub

    Private Sub txtCodigoRNC2_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC2.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC2.Text = Mcod2

        Defeito2 = ""
        Defeito2 = MRNC2
        lblDescricaoRNC2.Text = MRNC2 'incluir
        
    End Sub
    Private Sub txtCodigoRNC3_TextChanged_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC3.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC3.Text = Mcod3

        Defeito3 = ""
        Defeito3 = MRNC3
        lblDescricaoRNC3.Text = MRNC3 'incluir

    End Sub
    Private Sub txtCodigoRNC4_TextChanged_4(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC4.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC4.Text = Mcod4

        Defeito4 = ""
        Defeito4 = MRNC4
        lblDescricaoRNC4.Text = MRNC4 'incluir

    End Sub
    Private Sub txtCodigoRNC5_TextChanged_5(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC5.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC5.Text = Mcod5

        Defeito5 = ""
        Defeito5 = MRNC5
        lblDescricaoRNC5.Text = MRNC5 'incluir

    End Sub
    Private Sub txtCodigoRNC6_TextChanged_6(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC6.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC6.Text = Mcod6

        Defeito6 = ""
        Defeito6 = MRNC6
        lblDescricaoRNC6.Text = MRNC6 'incluir

    End Sub
    Private Sub txtCodigoRNC7_TextChanged_7(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC7.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC7.Text = Mcod7

        Defeito7 = ""
        Defeito7 = MRNC7
        lblDescricaoRNC7.Text = MRNC7 'incluir

    End Sub
    Private Sub txtCodigoRNC8_TextChanged_8(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC8.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC8.Text = Mcod8

        Defeito8 = ""
        Defeito8 = MRNC8
        lblDescricaoRNC8.Text = MRNC8 'incluir

    End Sub
    Private Sub txtCodigoRNC9_TextChanged_9(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC9.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC9.Text = Mcod9

        Defeito9 = ""
        Defeito9 = MRNC9
        lblDescricaoRNC9.Text = MRNC9 'incluir

    End Sub
    Private Sub txtCodigoRNC10_TextChanged_10(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoRNC10.MouseClick
        frmCodigos.ShowDialog()
        txtCodigoRNC10.Text = Mcod10

        Defeito10 = ""
        Defeito10 = MRNC10
        lblDescricaoRNC10.Text = MRNC10 'incluir

    End Sub

    Private Sub txtCodigoRNC1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC1.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC1.TextLength = 0 Then
                txtCodigoRNC1.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC1.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC1.Text = "0"
                    lblDescricaoRNC1.Focus()
                Else
                    txtCodigoRNC1.Clear()
                    txtCodigoRNC1.Focus()
                End If
            Else
                Defeito1 = ""
                Defeito1 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC1.Text = Defeito1 & " - "
                lblDescricaoRNC1.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 59 " & ex.Message)
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
    Private Sub txtCodigoRNC2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC2.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da16 As New OleDbDataAdapter
            Dim ds16 As New DataSet
            If txtCodigoRNC2.TextLength = 0 Then
                txtCodigoRNC2.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC2.Text & " "
            da16 = New OleDbDataAdapter(sel9, conDefeito)
            ds16.Clear()
            da16.Fill(ds16, "tblDefeitos")

            If ds16.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC2.Text = "0"
                    lblDescricaoRNC2.Focus()
                Else
                    txtCodigoRNC2.Clear()
                    txtCodigoRNC2.Focus()
                End If
            Else
                Defeito2 = ""
                Defeito2 = ds16.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC2.Text = Defeito2 & " - "
                lblDescricaoRNC2.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 61 " & ex.Message)
        End Try
    End Sub
    Private Sub txtCodigoRNC3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles txtCodigoRNC3.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da17 As New OleDbDataAdapter
            Dim ds17 As New DataSet
            If txtCodigoRNC3.TextLength = 0 Then
                txtCodigoRNC3.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC3.Text & " "
            da17 = New OleDbDataAdapter(sel9, conDefeito)
            ds17.Clear()
            da17.Fill(ds17, "tblDefeitos")

            If ds17.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC3.Text = "0"
                    lblDescricaoRNC3.Focus()
                Else
                    txtCodigoRNC3.Clear()
                    txtCodigoRNC3.Focus()
                End If
            Else
                Defeito3 = ""
                Defeito3 = ds17.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC3.Text = Defeito3 & " - "
                lblDescricaoRNC3.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 62 " & ex.Message)
        End Try
    End Sub
    Private Sub txtCodigoRNC4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC4.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da18 As New OleDbDataAdapter
            Dim ds18 As New DataSet
            If txtCodigoRNC4.TextLength = 0 Then
                txtCodigoRNC4.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC4.Text & " "
            da18 = New OleDbDataAdapter(sel9, conDefeito)
            ds18.Clear()
            da18.Fill(ds18, "tblDefeitos")

            If ds18.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC4.Text = "0"
                    lblDescricaoRNC4.Focus()
                Else
                    txtCodigoRNC4.Clear()
                    txtCodigoRNC4.Focus()
                End If
            Else
                Defeito4 = ""
                Defeito4 = ds18.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC4.Text = Defeito4 & " - "
                lblDescricaoRNC4.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 63 " & ex.Message)
        End Try
    End Sub
    Private Sub txtCodigoRNC5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC5.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC5.TextLength = 0 Then
                txtCodigoRNC5.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC5.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC5.Text = "0"
                    lblDescricaoRNC5.Focus()
                Else
                    txtCodigoRNC5.Clear()
                    txtCodigoRNC5.Focus()
                End If
            Else
                Defeito5 = ""
                Defeito5 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC5.Text = Defeito5 & " - "
                lblDescricaoRNC5.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 64 " & ex.Message)
        End Try
    End Sub

    Private Sub txtCodigoRNC6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles txtCodigoRNC6.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC6.TextLength = 0 Then
                txtCodigoRNC6.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC6.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC6.Text = "0"
                    lblDescricaoRNC6.Focus()
                Else
                    txtCodigoRNC6.Clear()
                    txtCodigoRNC6.Focus()
                End If
            Else
                Defeito6 = ""
                Defeito6 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC6.Text = Defeito6 & " - "
                lblDescricaoRNC6.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 65 " & ex.Message)
        End Try
    End Sub

    Private Sub txtCodigoRNC7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC7.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC7.TextLength = 0 Then
                txtCodigoRNC7.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC7.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC7.Text = "0"
                    lblDescricaoRNC7.Focus()
                Else
                    txtCodigoRNC7.Clear()
                    txtCodigoRNC7.Focus()
                End If
            Else
                Defeito7 = ""
                Defeito7 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC7.Text = Defeito7 & " - "
                lblDescricaoRNC7.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 66 " & ex.Message)
        End Try
    End Sub

    Private Sub txtCodigoRNC8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC8.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC8.TextLength = 0 Then
                txtCodigoRNC8.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC8.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC8.Text = "0"
                    lblDescricaoRNC8.Focus()
                Else
                    txtCodigoRNC8.Clear()
                    txtCodigoRNC8.Focus()
                End If
            Else
                Defeito8 = ""
                Defeito8 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC8.Text = Defeito8 & " - "
                lblDescricaoRNC8.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 67 " & ex.Message)
        End Try
    End Sub

    Private Sub txtCodigoRNC9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) ' Handles txtCodigoRNC9.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC9.TextLength = 0 Then
                txtCodigoRNC9.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC9.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC9.Text = "0"
                    lblDescricaoRNC9.Focus()
                Else
                    txtCodigoRNC9.Clear()
                    txtCodigoRNC9.Focus()
                End If
            Else
                Defeito9 = ""
                Defeito9 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lblDescricaoRNC9.Text = Defeito9 & " - "
                lblDescricaoRNC9.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 68 " & ex.Message)
        End Try
    End Sub

    Private Sub txtCodigoRNC10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles txtCodigoRNC10.LostFocus
        Try
            TesteAbertoDefeito()
            Dim da14 As New OleDbDataAdapter
            Dim ds14 As New DataSet
            If txtCodigoRNC10.TextLength = 0 Then
                txtCodigoRNC10.Text = "0"
            End If
            conDefeito.Open()
            Dim sel9 As String = "SELECT Nao_Conformidade FROM tblDefeitos where Codigo = " & txtCodigoRNC10.Text & " "
            da14 = New OleDbDataAdapter(sel9, conDefeito)
            ds14.Clear()

            da14.Fill(ds14, "tblDefeitos")
            If ds14.Tables("tblDefeitos").Rows.Count = 0 Then
                conDefeito.Close()
                If (MsgBox("O 'Código' da RNC não existe, deseja inserir um valor padrão?", vbYesNo, "Código da RNC")) = vbYes Then
                    txtCodigoRNC10.Text = "0"
                    lbldescricaornc10.Focus()
                Else
                    txtCodigoRNC10.Clear()
                    txtCodigoRNC10.Focus()
                End If
            Else
                Defeito10 = ""
                Defeito10 = ds14.Tables("tblDefeitos").Rows(0)("Nao_Conformidade")
                lbldescricaornc10.Text = Defeito10 & " - "
                lbldescricaornc10.Focus()
                conDefeito.Close()
            End If
        Catch ex As Exception
            MsgBox("Erro 69 " & ex.Message)
        End Try
    End Sub


    Private Sub btCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelar.Click
        Call Limpar()
        Atualizar()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            TesteAbertoRNC()
            Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

            Dim ID = row.Cells(0)

            Dim RNC = row.Cells(1)
            If lblRNC.Text = "*" Then
                Me.lblRNC.Text = RNC.Value
                AlterarCarregar()
            ElseIf RNC.Value = lblRNC.Text Then
                Me.lblRNC.Text = RNC.Value
            Else
                Me.lblRNC.Text = RNC.Value
                AlterarCarregar()
            End If

            Dim Origem = row.Cells(3)
            Dim Data_Abertura = row.Cells(4)
            Dim Hora = row.Cells(5)
            'Dim Cod_Produto = row.Cells(7)
            Dim Cliente = row.Cells(8)
            'Dim Produto = row.Cells(9)
            Dim OP_Reprovado = row.Cells(10)

            'Dim Turno = row.Cells(11)
            'Dim NúmerosCaixas = row.Cells(12)
            'Dim QT_Caixas = row.Cells(13)
            'Dim QT_Reprovado = row.Cells(14)
            'Dim Cod_Defeito = row.Cells(15)
            'Dim Nao_Conformidade = row.Cells(16)

            Dim Maquina = row.Cells(17)
            Dim Observacao = row.Cells(25)
            Dim RE = row.Cells(26)
            Dim Inspetor = row.Cells(27)
            Dim Setor = row.Cells(28)
            Dim TurnoDetector = row.Cells(29)
            Dim Alterado = row.Cells(30)


            Me.lblID.Text = ID.Value
            Me.txtOP.Text = OP_Reprovado.Value
            Me.cbDetectado.Text = Origem.Value
            Me.lblData.Text = Data_Abertura.Value
            Me.lblHora.Text = Hora.Value
            'Me.lblCodProduto.Text = Cod_Produto.Value
            'Me.lblProduto.Text = Produto.Value


            'Me.cb1Turno.Text = Turno.Value
            'Me.txtCaixas1Turno.Text = NúmerosCaixas.Value
            'Me.txtQtCaixasReprovada1.Text = QT_Caixas.Value
            'Me.txtQtPorTurno1.Text = QT_Reprovado.Value
            'If Me.txtQtCaixasReprovada1.Text = 0 Then
            'Me.txtPecasPorVolume.Text = ""
            'Else
            'Me.txtPecasPorVolume.Text = QT_Reprovado.Value / QT_Caixas.Value
            'End If
            'Me.txtCodigoRNC1.Text = Cod_Defeito.Value
            'Me.lbldescricaornc1.Text = Nao_Conformidade.Value

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
        Catch ex As Exception
            MsgBox("Erro 70 " & ex.Message)
        End Try
        ContarCaixas()

        txtOP.Focus()
        txtOBS.Focus()

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

    Private Sub btAlterar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles btAlterar.Click
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


                    'verificação generica

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
        Catch ex As Exception
            MsgBox("Erro 72 " & ex.Message)
        End Try
    End Sub
    Sub Alterar1()
        Try
            Call codRNC1()
            Call Celulax()
            conRNC.Open()
            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet
            ds20 = New DataSet
            da20 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb1Turno.Text & "', NúmerosCaixas = '" & txtCaixas1Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada1.Text & ", QT_Reprovado = '" & txtQTPorTurno1.Text & "', Cod_Defeito = " & txtCodigoRNC1.Text & ", Nao_Conformidade = '" & lblDescricaoRNC1.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID1 & "", conRNC)
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
            da20_2 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb2Turnos.Text & "', NúmerosCaixas = '" & txtCaixas2Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada2.Text & ", QT_Reprovado = '" & txtQTPorTurno2.Text & "', Cod_Defeito = " & txtCodigoRNC2.Text & ", Nao_Conformidade = '" & lblDescricaoRNC2.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID2 & "", conRNC)
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
            da20_3 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb3Turnos.Text & "', NúmerosCaixas = '" & txtCaixas3Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada3.Text & ", QT_Reprovado = '" & txtQTPorTurno3.Text & "', Cod_Defeito = " & txtCodigoRNC3.Text & ", Nao_Conformidade = '" & lblDescricaoRNC3.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID3 & "", conRNC)
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
            da20_4 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb4Turnos.Text & "', NúmerosCaixas = '" & txtCaixas4Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada4.Text & ", QT_Reprovado = '" & txtQTPorTurno4.Text & "', Cod_Defeito = " & txtCodigoRNC4.Text & ", Nao_Conformidade = '" & lblDescricaoRNC4.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID4 & "", conRNC)
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
            da20_5 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb5Turno.Text & "', NúmerosCaixas = '" & txtCaixas5Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada5.Text & ", QT_Reprovado = '" & txtQTPorTurno5.Text & "', Cod_Defeito = " & txtCodigoRNC5.Text & ", Nao_Conformidade = '" & lblDescricaoRNC5.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID5 & "", conRNC)
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
            da20_6 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb6Turno.Text & "', NúmerosCaixas = '" & txtCaixas6Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada6.Text & ", QT_Reprovado = '" & txtQTPorTurno6.Text & "', Cod_Defeito = " & txtCodigoRNC6.Text & ", Nao_Conformidade = '" & lblDescricaoRNC6.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID6 & "", conRNC)
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
            da20_7 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb7Turno.Text & "', NúmerosCaixas = '" & txtCaixas7Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada7.Text & ", QT_Reprovado = '" & txtQTPorTurno7.Text & "', Cod_Defeito = " & txtCodigoRNC7.Text & ", Nao_Conformidade = '" & lblDescricaoRNC7.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID7 & "", conRNC)
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
            da20_8 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb8Turno.Text & "', NúmerosCaixas = '" & txtCaixas8Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada8.Text & ", QT_Reprovado = '" & txtQTPorTurno8.Text & "', Cod_Defeito = " & txtCodigoRNC8.Text & ", Nao_Conformidade = '" & lblDescricaoRNC8.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID8 & "", conRNC)
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
            da20_9 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb9Turno.Text & "', NúmerosCaixas = '" & txtCaixas9Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada9.Text & ", QT_Reprovado = '" & txtQTPorTurno9.Text & "', Cod_Defeito = " & txtCodigoRNC9.Text & ", Nao_Conformidade = '" & lblDescricaoRNC9.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID9 & "", conRNC)
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
            da20_10 = New OleDbDataAdapter("UPDATE tblRNC SET  Origem = '" & cbDetectado.Text & "', Turno = '" & cb10Turno.Text & "', NúmerosCaixas = '" & txtCaixas10Turno.Text & "', QT_Caixas = " & txtQtCaixasReprovada10.Text & ", QT_Reprovado = '" & txtQTPorTurno10.Text & "', Cod_Defeito = " & txtCodigoRNC10.Text & ", Nao_Conformidade = '" & lbldescricaornc10.Text & "', Maquina = '" & txtMaquina.Text & "', Celula = '" & Celula & "', Observacao = '" & txtOBS.Text & "', RE = " & txtRE.Text & ", Inspetor = '" & txtInspetor.Text & "', Setor = '" & txtSetor.Text & "', TurnoDetector = '" & cbTurno.Text & "', Data_Hora_Alteracao = '" & Today & " " & TimeOfDay.ToShortTimeString & "' WHERE ID = " & ID10 & "", conRNC)
            ds20_10.Clear()
            da20_10.Fill(ds20_10, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 82 " & ex.Message)
        End Try
    End Sub
    Sub limpar2()

        cb2Turnos.Text = ""
        txtCaixas2Turno.Clear()
        txtQtCaixasReprovada2.Clear()
        txtQTPorTurno2.Text = ""
        txtCodigoRNC2.Clear()

        'lblDescricaoRNC1.Text = "*"
        lblDescricaoRNC2.Text = "*"

        cb3Turnos.Text = ""
        txtCaixas3Turno.Clear()
        txtQtCaixasReprovada3.Clear()
        txtQTPorTurno3.Text = ""
        txtCodigoRNC3.Clear()
        lblDescricaoRNC3.Text = "*"

        cb4Turnos.Text = ""
        txtCaixas4Turno.Clear()
        txtQtCaixasReprovada4.Clear()
        txtQTPorTurno4.Text = ""
        txtCodigoRNC4.Clear()
        lblDescricaoRNC4.Text = "*"

        cb5Turno.Text = ""
        txtCaixas5Turno.Clear()
        txtQtCaixasReprovada5.Clear()
        txtQTPorTurno5.Text = ""
        txtCodigoRNC5.Clear()
        lblDescricaoRNC5.Text = "*"

        cb6Turno.Text = ""
        txtCaixas6Turno.Clear()
        txtQtCaixasReprovada6.Clear()
        txtQTPorTurno6.Text = ""
        txtCodigoRNC6.Clear()
        lblDescricaoRNC6.Text = "*"

        cb7Turno.Text = ""
        txtCaixas7Turno.Clear()
        txtQtCaixasReprovada7.Clear()
        txtQTPorTurno7.Text = ""
        txtCodigoRNC7.Clear()
        lblDescricaoRNC7.Text = "*"

        cb8Turno.Text = ""
        txtCaixas8Turno.Clear()
        txtQtCaixasReprovada8.Clear()
        txtQTPorTurno8.Text = ""
        txtCodigoRNC8.Clear()
        lblDescricaoRNC8.Text = "*"

        cb9Turno.Text = ""
        txtCaixas9Turno.Clear()
        txtQtCaixasReprovada9.Clear()
        txtQTPorTurno9.Text = ""
        txtCodigoRNC9.Clear()
        lblDescricaoRNC9.Text = "*"

        cb10Turno.Text = ""
        txtCaixas10Turno.Clear()
        txtQtCaixasReprovada10.Clear()
        txtQTPorTurno10.Text = ""
        txtCodigoRNC10.Clear()
        lblDescricaoRNC10.Text = "*"


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
            txtQTPorTurno1.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_Reprovado")
            txtCodigoRNC1.Text = dsPRINT.Tables("tblRNC").Rows(0)("Cod_Defeito")
            lblDescricaoRNC1.Text = dsPRINT.Tables("tblRNC").Rows(0)("Nao_Conformidade")

            limpar2()
            If dtPrint.Rows.Count >= 2 Then
                rb2T.Checked = True
                ID2 = dsPRINT.Tables("tblRNC").Rows(1)("ID")
                cb2Turnos.Text = dsPRINT.Tables("tblRNC").Rows(1)("Turno")
                txtCaixas2Turno.Text = dsPRINT.Tables("tblRNC").Rows(1)("NúmerosCaixas")
                txtQtCaixasReprovada2.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_Caixas")
                txtQTPorTurno2.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_Reprovado")
                txtCodigoRNC2.Text = dsPRINT.Tables("tblRNC").Rows(1)("Cod_Defeito")
                lblDescricaoRNC2.Text = dsPRINT.Tables("tblRNC").Rows(1)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 3 Then
                ID3 = dsPRINT.Tables("tblRNC").Rows(2)("ID")
                rb3T.Checked = True
                cb3Turnos.Text = dsPRINT.Tables("tblRNC").Rows(2)("Turno")
                txtCaixas3Turno.Text = dsPRINT.Tables("tblRNC").Rows(2)("NúmerosCaixas")
                txtQtCaixasReprovada3.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_Caixas")
                txtQTPorTurno3.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_Reprovado")
                txtCodigoRNC3.Text = dsPRINT.Tables("tblRNC").Rows(2)("Cod_Defeito")
                lblDescricaoRNC3.Text = dsPRINT.Tables("tblRNC").Rows(2)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 4 Then
                ID4 = dsPRINT.Tables("tblRNC").Rows(3)("ID")
                rb4T.Checked = True
                cb4Turnos.Text = dsPRINT.Tables("tblRNC").Rows(3)("Turno")
                txtCaixas4Turno.Text = dsPRINT.Tables("tblRNC").Rows(3)("NúmerosCaixas")
                txtQtCaixasReprovada4.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_Caixas")
                txtQTPorTurno4.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_Reprovado")
                txtCodigoRNC4.Text = dsPRINT.Tables("tblRNC").Rows(3)("Cod_Defeito")
                lblDescricaoRNC4.Text = dsPRINT.Tables("tblRNC").Rows(3)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 5 Then
                ID5 = dsPRINT.Tables("tblRNC").Rows(4)("ID")
                rb5T.Checked = True
                cb5Turno.Text = dsPRINT.Tables("tblRNC").Rows(4)("Turno")
                txtCaixas5Turno.Text = dsPRINT.Tables("tblRNC").Rows(4)("NúmerosCaixas")
                txtQtCaixasReprovada5.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_Caixas")
                txtQTPorTurno5.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_Reprovado")
                txtCodigoRNC5.Text = dsPRINT.Tables("tblRNC").Rows(4)("Cod_Defeito")
                lblDescricaoRNC5.Text = dsPRINT.Tables("tblRNC").Rows(4)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 6 Then
                ID6 = dsPRINT.Tables("tblRNC").Rows(5)("ID")
                rb6T.Checked = True
                cb6Turno.Text = dsPRINT.Tables("tblRNC").Rows(5)("Turno")
                txtCaixas6Turno.Text = dsPRINT.Tables("tblRNC").Rows(5)("NúmerosCaixas")
                txtQtCaixasReprovada6.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_Caixas")
                txtQTPorTurno6.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_Reprovado")
                txtCodigoRNC6.Text = dsPRINT.Tables("tblRNC").Rows(5)("Cod_Defeito")
                lblDescricaoRNC6.Text = dsPRINT.Tables("tblRNC").Rows(5)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 7 Then
                ID7 = dsPRINT.Tables("tblRNC").Rows(6)("ID")
                rb7T.Checked = True
                cb7Turno.Text = dsPRINT.Tables("tblRNC").Rows(6)("Turno")
                txtCaixas7Turno.Text = dsPRINT.Tables("tblRNC").Rows(6)("NúmerosCaixas")
                txtQtCaixasReprovada7.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_Caixas")
                txtQTPorTurno7.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_Reprovado")
                txtCodigoRNC7.Text = dsPRINT.Tables("tblRNC").Rows(6)("Cod_Defeito")
                lblDescricaoRNC7.Text = dsPRINT.Tables("tblRNC").Rows(6)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 8 Then
                ID8 = dsPRINT.Tables("tblRNC").Rows(7)("ID")
                rb8T.Checked = True
                cb8Turno.Text = dsPRINT.Tables("tblRNC").Rows(7)("Turno")
                txtCaixas8Turno.Text = dsPRINT.Tables("tblRNC").Rows(7)("NúmerosCaixas")
                txtQtCaixasReprovada8.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_Caixas")
                txtQTPorTurno8.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_Reprovado")
                txtCodigoRNC8.Text = dsPRINT.Tables("tblRNC").Rows(7)("Cod_Defeito")
                lblDescricaoRNC8.Text = dsPRINT.Tables("tblRNC").Rows(7)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count >= 9 Then
                ID9 = dsPRINT.Tables("tblRNC").Rows(8)("ID")
                rb9T.Checked = True
                cb9Turno.Text = dsPRINT.Tables("tblRNC").Rows(8)("Turno")
                txtCaixas9Turno.Text = dsPRINT.Tables("tblRNC").Rows(8)("NúmerosCaixas")
                txtQtCaixasReprovada9.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_Caixas")
                txtQTPorTurno9.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_Reprovado")
                txtCodigoRNC9.Text = dsPRINT.Tables("tblRNC").Rows(8)("Cod_Defeito")
                lblDescricaoRNC9.Text = dsPRINT.Tables("tblRNC").Rows(8)("Nao_Conformidade")
            End If
            If dtPrint.Rows.Count = 10 Then
                ID10 = dsPRINT.Tables("tblRNC").Rows(9)("ID")
                rb10T.Checked = True
                cb10Turno.Text = dsPRINT.Tables("tblRNC").Rows(9)("Turno")
                txtCaixas10Turno.Text = dsPRINT.Tables("tblRNC").Rows(9)("NúmerosCaixas")
                txtQtCaixasReprovada10.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_Caixas")
                txtQTPorTurno10.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_Reprovado")
                txtCodigoRNC10.Text = dsPRINT.Tables("tblRNC").Rows(9)("Cod_Defeito")
                lblDescricaoRNC10.Text = dsPRINT.Tables("tblRNC").Rows(9)("Nao_Conformidade")
            End If
        Catch ex As Exception
            MsgBox("Erro 83 " & ex.Message)
        End Try
    End Sub
    Private Sub btExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExcluir.Click
        Try

            Dim hora1, hora2 As Date
            hora1 = Date.Parse(lblHora.Text)
            hora2 = TimeOfDay
            If lblRNC.Text = "*" Or lblRNC.Text = "" Then
                MsgBox("Selecione um RNC na tabela abaixo", , "Selecione uma RNC")
            Else
                If (lblData.Text = Today) And (hora1.Hour = hora2.Hour Or hora1.Hour = hora2.Hour - 1 Or hora1.Hour = hora2.Hour - 2) Then
                    Dim da21 As New OleDbDataAdapter
                    Dim ds21 As New DataSet
                    If btExcluir.Text = "Excluir" Then
                        If MsgBox("Deseja Excluir uma RNC?", vbYesNo, "Excluir RNC") = vbYes Then
                            btExcluir.Text = "Aplicar"
                            btInserir.Enabled = False
                            ' btAlterar.Enabled = False
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
                        Else
                        End If
                    ElseIf txtOBS.TextLength = 0 Then
                        MsgBox("Favor informe o motivo do pedido de exclusão no campo de observação!")
                        txtOBS.Focus()
                    Else
                        TesteAbertoRNC()
                        conRNC.Open()
                        Dim da20 As New OleDbDataAdapter
                        Dim ds20 As New DataSet
                        ds20 = New DataSet
                        da20 = New OleDbDataAdapter("UPDATE tblRNC SET Cancelar = 'Cancelar', Observacao = 'Motivo : " & txtOBS.Text & "' WHERE RNC = " & lblRNC.Text & "", conRNC)
                        ds20.Clear()
                        da20.Fill(ds20, "tblRNC")
                        conRNC.Close()
                        email_Excluir()
                        Call Atualizar()
                        Limpar()
                        MsgBox("Pedido de exclusão registrado com sucesso!")
                    End If
                Else
                    MsgBox("Você não pode mais pedir para cancelar esta RNC, solicite verbalmente!")
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 84 " & ex.Message)
        End Try
    End Sub
    Sub VerificaAsterisco1()
        If txtCodigoRNC1.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC1.Text = Mcod1
                lblDescricaoRNC1.Text = MRNC1
                If txtCodigoRNC1.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco2()
        If txtCodigoRNC2.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC2.Text = Mcod2
                lblDescricaoRNC2.Text = MRNC2
                If txtCodigoRNC2.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco3()
        If txtCodigoRNC3.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC3.Text = Mcod3
                lblDescricaoRNC3.Text = MRNC3
                If txtCodigoRNC3.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco4()
        If txtCodigoRNC4.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC4.Text = Mcod4
                lblDescricaoRNC4.Text = MRNC4
                If txtCodigoRNC4.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco5()
        If txtCodigoRNC5.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC5.Text = Mcod5
                lblDescricaoRNC5.Text = MRNC5
                If txtCodigoRNC5.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco6()
        If txtCodigoRNC6.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC6.Text = Mcod6
                lblDescricaoRNC6.Text = MRNC6
                If txtCodigoRNC6.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco7()
        If txtCodigoRNC7.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC7.Text = Mcod7
                lblDescricaoRNC7.Text = MRNC7
                If txtCodigoRNC7.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco8()
        If txtCodigoRNC8.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC8.Text = Mcod8
                lblDescricaoRNC8.Text = MRNC8
                If txtCodigoRNC8.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco9()
        If txtCodigoRNC9.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC9.Text = Mcod9
                lblDescricaoRNC9.Text = MRNC9
                If txtCodigoRNC9.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub

    Sub VerificaAsterisco10()
        If txtCodigoRNC10.Text = "*" Then
            Dim i As Integer = 0
            For i = 0 To 10 Step 1
                frmCodigos.ShowDialog()
                txtCodigoRNC10.Text = Mcod10
                lblDescricaoRNC10.Text = MRNC10
                If txtCodigoRNC10.Text = "*" Then
                    i = 0
                Else
                    i = 10
                End If
            Next
        End If
    End Sub
    Dim rB As Integer
    Sub VerificarRadioButon()
        If rb1T.Checked = True Then
            rB = 1
        ElseIf rb2T.Checked = True Then
            rB = 2
        ElseIf rb3T.Checked = True Then
            rB = 3
        ElseIf rb4T.Checked = True Then
            rB = 4
        ElseIf rb5T.Checked = True Then
            rB = 5
        ElseIf rb6T.Checked = True Then
            rB = 6
        ElseIf rb7T.Checked = True Then
            rB = 7
        ElseIf rb8T.Checked = True Then
            rB = 8
        ElseIf rb9T.Checked = True Then
            rB = 9
        ElseIf rb10T.Checked = True Then
            rB = 10
        End If
    End Sub

    Public Sub ImprimirEtiqueta()

        VerificarRadioButon()

        Select Case rB
            Case 1
                VerificaAsterisco1()
            Case 2
                VerificaAsterisco1()
                VerificaAsterisco2()
            Case 3
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
            Case 4
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
            Case 5
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
                VerificaAsterisco5()
            Case 6
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
                VerificaAsterisco5()
                VerificaAsterisco6()
            Case 7
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
                VerificaAsterisco5()
                VerificaAsterisco6()
                VerificaAsterisco7()

            Case 8
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
                VerificaAsterisco5()
                VerificaAsterisco6()
                VerificaAsterisco7()
                VerificaAsterisco8()
            Case 9
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
                VerificaAsterisco5()
                VerificaAsterisco6()
                VerificaAsterisco7()
                VerificaAsterisco8()
                VerificaAsterisco9()
            Case 10
                VerificaAsterisco1()
                VerificaAsterisco2()
                VerificaAsterisco3()
                VerificaAsterisco4()
                VerificaAsterisco5()
                VerificaAsterisco6()
                VerificaAsterisco7()
                VerificaAsterisco8()
                VerificaAsterisco9()
                VerificaAsterisco10()
        End Select

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
        Documento_xlsx_ETQ = Excell_ETQ.Workbooks.Open("f:\Receb.Mat.Prima\Banco_Dados\RNCEtiqueta.xlsx")

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
        'If btInserir.Text = "Aplicar" Then
        'Descricao.Value = "Descrição: " & Defeito1 & " " & lblDescricaoRNC1.Text & ", " & Defeito2 & " " & lblDescricaoRNC2.Text & ", " & Defeito3 & " " & lblDescricaoRNC3.Text & ", " & Defeito4 & " " & lblDescricaoRNC4.Text & ", " & Defeito5 & " " & lblDescricaoRNC5.Text & ", " & Defeito6 & " " & lblDescricaoRNC6.Text & ", " & Defeito7 & " " & lblDescricaoRNC7.Text & ", " & Defeito8 & " " & lblDescricaoRNC8.Text & ", " & Defeito9 & " " & lblDescricaoRNC9.Text & ", " & Defeito10 & " " & lblDescricaoRNC10.Text
        'Else
        Descricao.Value = "Descrição: " & lblDescricaoRNC1.Text & ", " & lblDescricaoRNC2.Text & ", " & lblDescricaoRNC3.Text & ", " & lblDescricaoRNC4.Text & ", " & lblDescricaoRNC5.Text & ", " & lblDescricaoRNC6.Text & ", " & lblDescricaoRNC7.Text & ", " & lblDescricaoRNC8.Text & ", " & lblDescricaoRNC9.Text & ", " & lblDescricaoRNC10.Text
        'End If
        Maquina.Value = "Maquina: " & txtMaquina.Text
        'RE.Value = "RE: " & txtRE.Text
        Inspetor.Value = "Inspetor: " & txtInspetor.Text
        TurnoDetector.Value = "T Detector: " & cbTurno.Text

        'Dim QT_Informada As Int16 = InputBox("Informe a Quantidade de Etiquetas!!")
        Dim QT_Final As Int16
        If SMC <= 8 Then
            MsgBox("Prepare: 1 Folha na Impressora," _
               & Chr(13) _
               & "E informe este valor no dialogo de Impressão", , "Imprimir Etiquetas")
        Else


            QT_Final = SMC / 8

            Dim resto As Int16 = SMC Mod 8
            If resto = 0 Then
                'não inclui nada
            Else
                QT_Final = QT_Final + 1 'soma 1 no QT_Final
            End If


            MsgBox("Prepare: " & QT_Final & " Folhas na Impressora," _
                   & Chr(13) _
                   & "E informe este valor no dialogo de Impressão", , "Imprimir Etiquetas")
        End If
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

    Public Sub ImprimirRNC()
        Dim Excell As New Microsoft.Office.Interop.Excel.Application
        On Error GoTo ErrHandler

        If btImprimir.Text = "Imprimir..." Then
            RNC_RNC2 = Test("F:\RECEB.MAT.PRIMA\Banco_Dados\Documentos_RNC\RNC_" & lblRNC.Text & ".xlsx")
            If RNC_RNC2 = False Then
                Documento_xlsx = Excell.Workbooks.Open("F:\RECEB.MAT.PRIMA\Banco_Dados\Documentos_RNC\" & "RNC_" & lblRNC.Text & ".xlsx")
                Documento_xlsx.PrintOutEx(1, 2, 1)
                Documento_xlsx.Save()
                Resume ExitHere
            Else

            End If
        End If
        'If MsgBox("Deseja Imprimir o relatório 2 com as 3 vias? - Você poderá alterar a quantidade!", MsgBoxStyle.DefaultButton2, vbYesNo) = vbYes Then
        'Dim quantos As Integer = InputBox("Informe quantas impressões", "Impressões", 3, 500, 500)
        'Documento_xlsx.PrintOutEx(2, 2, quantos)
        'Documento_xlsx.Save()
        'Resume ExitHere
        'Else
        'Resume ExitHere
        'End If



        TesteAbertoDoc()

        Dim Planilha_do_Documento_xlsx As Microsoft.Office.Interop.Excel.Worksheet

        Dim RNC As Microsoft.Office.Interop.Excel.Range
        Dim OP As Microsoft.Office.Interop.Excel.Range
        Dim CodProduto As Microsoft.Office.Interop.Excel.Range
        Dim Produto As Microsoft.Office.Interop.Excel.Range
        Dim Data As Microsoft.Office.Interop.Excel.Range
        Dim Hora As Microsoft.Office.Interop.Excel.Range
        Dim Deteccao As Microsoft.Office.Interop.Excel.Range
        Dim Maquina As Microsoft.Office.Interop.Excel.Range
        Dim Cliente As Microsoft.Office.Interop.Excel.Range
        Dim Descricao As Microsoft.Office.Interop.Excel.Range

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


        '3º Abrir o arquivo Excel
        Documento_xlsx = Excell.Workbooks.Open("f:\Receb.Mat.Prima\Banco_Dados\RNCDoc.xlsx")

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
        Quantidade1.Value = Double.Parse(txtQTPorTurno1.Text)
        CodRNC1.Value = txtCodigoRNC1.Text
        DescricaoRNC1.Value = lblDescricaoRNC1.Text ' & Defeito1


        Turno2.Value = cb2Turnos.Text
        Caixas2.Value = txtCaixas2Turno.Text
        QT_Caixa2.Value = txtQtCaixasReprovada2.Text
        If txtQTPorTurno2.Text = "" Then
            Quantidade2.Value = 0
        Else
            Quantidade2.Value = Double.Parse(txtQTPorTurno2.Text)
        End If
        CodRNC2.Value = txtCodigoRNC2.Text
        DescricaoRNC2.Value = lblDescricaoRNC2.Text  ' & Defeito2

        Turno3.Value = cb3Turnos.Text
        Caixas3.Value = txtCaixas3Turno.Text
        QT_Caixa3.Value = txtQtCaixasReprovada3.Text
        If txtQTPorTurno3.Text = "" Then
            Quantidade3.Value = 0
        Else
            Quantidade3.Value = Double.Parse(txtQTPorTurno3.Text)
        End If
        CodRNC3.Value = txtCodigoRNC3.Text
        DescricaoRNC3.Value = lblDescricaoRNC3.Text ' & Defeito3

        Turno4.Value = cb4Turnos.Text
        Caixas4.Value = txtCaixas4Turno.Text
        QT_Caixa4.Value = txtQtCaixasReprovada4.Text
        If txtQTPorTurno4.Text = "" Then
            Quantidade4.Value = 0
        Else
            Quantidade4.Value = Double.Parse(txtQTPorTurno4.Text)
        End If
        CodRNC4.Value = txtCodigoRNC4.Text
        DescricaoRNC4.Value = lblDescricaoRNC4.Text ' & Defeito4

        Turno5.Value = cb5Turno.Text
        Caixas5.Value = txtCaixas5Turno.Text
        QT_Caixa5.Value = txtQtCaixasReprovada5.Text
        If txtQTPorTurno5.Text = "" Then
            Quantidade5.Value = 0
        Else
            Quantidade5.Value = Double.Parse(txtQTPorTurno5.Text)
        End If
        CodRNC5.Value = txtCodigoRNC5.Text
        DescricaoRNC5.Value = lblDescricaoRNC5.Text ' & Defeito5

        Turno6.Value = cb6Turno.Text
        Caixas6.Value = txtCaixas6Turno.Text
        QT_Caixa6.Value = txtQtCaixasReprovada6.Text
        If txtQTPorTurno6.Text = "" Then
            Quantidade6.Value = 0
        Else
            Quantidade6.Value = Double.Parse(txtQTPorTurno6.Text)
        End If
        CodRNC6.Value = txtCodigoRNC6.Text
        DescricaoRNC6.Value = lblDescricaoRNC6.Text ' & Defeito6

        Turno7.Value = cb7Turno.Text
        Caixas7.Value = txtCaixas7Turno.Text
        QT_Caixa7.Value = txtQtCaixasReprovada7.Text
        If txtQTPorTurno7.Text = "" Then
            Quantidade7.Value = 0
        Else
            Quantidade7.Value = Double.Parse(txtQTPorTurno7.Text)
        End If
        CodRNC7.Value = txtCodigoRNC7.Text
        DescricaoRNC7.Value = lblDescricaoRNC7.Text ' & Defeito7

        Turno8.Value = cb8Turno.Text
        Caixas8.Value = txtCaixas8Turno.Text
        QT_Caixa8.Value = txtQtCaixasReprovada8.Text
        If txtQTPorTurno8.Text = "" Then
            Quantidade8.Value = 0
        Else
            Quantidade8.Value = Double.Parse(txtQTPorTurno8.Text)
        End If
        CodRNC8.Value = txtCodigoRNC8.Text
        DescricaoRNC8.Value = lblDescricaoRNC8.Text ' & Defeito8

        Turno9.Value = cb9Turno.Text
        Caixas9.Value = txtCaixas9Turno.Text
        QT_Caixa9.Value = txtQtCaixasReprovada9.Text
        If txtQTPorTurno9.Text = "" Then
            Quantidade9.Value = 0
        Else
            Quantidade9.Value = Double.Parse(txtQTPorTurno9.Text)
        End If
        CodRNC9.Value = txtCodigoRNC9.Text
        DescricaoRNC9.Value = lblDescricaoRNC9.Text ' & Defeito9

        Turno10.Value = cb10Turno.Text
        Caixas10.Value = txtCaixas10Turno.Text
        QT_Caixa10.Value = txtQtCaixasReprovada10.Text
        If txtQTPorTurno10.Text = "" Then
            Quantidade10.Value = 0
        Else
            Quantidade10.Value = Double.Parse(txtQTPorTurno10.Text)
        End If
        CodRNC10.Value = txtCodigoRNC10.Text
        DescricaoRNC10.Value = lblDescricaoRNC10.Text ' & Defeito10
        RE.Value = "RE: " & txtRE.Text
        Inspetor.Value = "Nome: " & txtInspetor.Text
        Setor.Value = "Setor: " & txtSetor.Text
        TurnoDetector.Value = "Turno: " & cbTurno.Text
        Obs.Value = txtOBS.Text

        'Documento_xlsx.PrintOutEx() ' imprime direto

        If btInserir.Text = "Aplicar" Then
            '7º Abrindo o excel
            Excell.Visible = False
            '8º Salvando a Planilha
            Documento_xlsx.Save()
            'imprimirsemm dialogo
            'Documento_xlsx.PrintOutEx(1, 2, 1)



            'If MsgBox("Deseja Imprimir o relatório 2 com as 3 vias? Você poderá alterar a quantidade!", vbYesNo, "Nova RNC") = vbYes Then
            'Dim quantos As Integer = InputBox("Informe quantas impressões", "Impressões", 3, 500, 500)
            '4º Abrir a planilha para inserir texto

            Planilha_do_Documento_xlsx = Documento_xlsx.Sheets.Item("RNCForm3")
            '5º Atribuir uma célula na planilha

            '5º Atribuir uma célula na planilha

            Cliente = Planilha_do_Documento_xlsx.Range("D5")
            Descricao = Planilha_do_Documento_xlsx.Range("A10")
            Cliente.Value = cliente2.ToString
            Descricao.Value = lblDescricaoRNC1.Text & ", " & lblDescricaoRNC2.Text & ", " & lblDescricaoRNC3.Text & ", " & lblDescricaoRNC4.Text & ", " & lblDescricaoRNC5.Text & ", " & lblDescricaoRNC6.Text & ", " & lblDescricaoRNC7.Text & ", " & lblDescricaoRNC8.Text & ", " & lblDescricaoRNC9.Text & ", " & lblDescricaoRNC10.Text
            'Descricao.Value = Defeito1 & ", " & Defeito2 & ", " & Defeito3 & ", " & Defeito4 & ", " & Defeito5 & ", " & Defeito6 & ", " & Defeito7 & ", " & Defeito8 & ", " & Defeito9 & ", " & Defeito10
            Documento_xlsx.SaveAs("F:\RECEB.MAT.PRIMA\Banco_Dados\Documentos_RNC\" & "RNC_" & lblRNC.Text & ".xlsx")
            'Documento_xlsx.PrintOutEx(2, 2, quantos)
            'End If

            'imprimirsemm dialogo
            Documento_xlsx.PrintOutEx(1, 2, 1)

        ElseIf btImprimir.Text = "Imprimir..." Or btEmail.Text = "Email..." Then
            '7º Abrindo o excel
            Excell.Visible = False
            '8º Salvando a Planilha
            Documento_xlsx.Save()
            'imprimirsemm dialogo

           

            'If MsgBox("Deseja Imprimir o relatório 2 com as 3 vias? Você poderá alterar a quantidade!", vbYesNo, "Nova RNC") = vbYes Then
            'Dim quantos As Integer = InputBox("Informe quantas impressões", "Impressões", 3, 500, 500)
            '4º Abrir a planilha para inserir texto

            Planilha_do_Documento_xlsx = Documento_xlsx.Sheets.Item("RNCForm3")
            '5º Atribuir uma célula na planilha

            '5º Atribuir uma célula na planilha

            Cliente = Planilha_do_Documento_xlsx.Range("D5")
            Descricao = Planilha_do_Documento_xlsx.Range("A10")
            Clientex()
            Cliente.Value = cliente2.ToString
            'Descricao.Value = Defeito1 & ", " & Defeito2 & ", " & Defeito3 & ", " & Defeito4 & ", " & Defeito5 & ", " & Defeito6 & ", " & Defeito7 & ", " & Defeito8 & ", " & Defeito9 & ", " & Defeito10
            Descricao.Value = lblDescricaoRNC1.Text & ", " & lblDescricaoRNC2.Text & ", " & lblDescricaoRNC3.Text & ", " & lblDescricaoRNC4.Text & ", " & lblDescricaoRNC5.Text & ", " & lblDescricaoRNC6.Text & ", " & lblDescricaoRNC7.Text & ", " & lblDescricaoRNC8.Text & ", " & lblDescricaoRNC9.Text & ", " & lblDescricaoRNC10.Text
            Documento_xlsx.SaveAs("F:\RECEB.MAT.PRIMA\Banco_Dados\Documentos_RNC\" & "RNC_" & lblRNC.Text & ".xlsx")
            If btImprimir.Text = "Imprimir..." Then
                Documento_xlsx.PrintOutEx(1, 2, 1)
            End If
        End If


        '9º encerra os processos EXCEL.EXE no gerenciador de tarefas do windows 
ExitHere:
        Excell.Quit()
        Marshal.ReleaseComObject(Documento_xlsx)
        Marshal.ReleaseComObject(Excell)
        Excell = Nothing
        Exit Sub
ErrHandler:
        ' MsgBox(Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source, vbCritical, "Erro xx6 ")
        Resume ExitHere

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

    Sub email_Excluir()

        Dim OutlookMessage As Outlook.MailItem
        Dim AppOutlook As New Microsoft.Office.Interop.Outlook.Application
        Try

            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            'Dim Recipents As Outlook.Recipients = OutlookMessage.Recipients
            OutlookMessage.To = "RNC_Excluir" ' Criar um grupo no outlook chamado RNC_Excluir
            'Recipents.Add("cidmevb@gmail.com; inspetor1@mondicap.com.br; rafael.pedroso@mondicap.com.br; recebimento@mondicap.com.br")
            OutlookMessage.Subject = "Exclusão da  RNC: " & lblRNC.Text


            OutlookMessage.Body = "Favor excluir a RNC : " & lblRNC.Text & "" _
                    & Chr(13) _
                    & "Motivo: " & txtOBS.Text & ""

            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML

            'If (MsgBox("O E-mail está pronto para ser enviado. Deseja Enviar?" _
            '         & Chr(13) _
            '        & Chr(13) _
            '       & "'Sim' = Enviar" _
            '      & Chr(13) _
            '     & "'Não' = Alterar", vbYesNo, "E-mail") = vbYes) Then
            OutlookMessage.Save()
            OutlookMessage.Send()
            'Else
            'OutlookMessage.Display()
            'OutlookMessage.Save()
            'End If

        Catch ex As Exception
            MessageBox.Show("Erro 88w " & ex.Message) 'if you dont want this message, simply delete this line 
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try

    End Sub

    Sub email()
        Dim OutlookMessage As Outlook.MailItem
        Dim AppOutlook As New Microsoft.Office.Interop.Outlook.Application
        Try

            OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
            'Dim Recipent As Outlook.Recipients = OutlookMessage.Recipients
            OutlookMessage.To = "RNC_RNC" ' Criar um grupo no outlook chamado RNC
            'Recipent.Add("RNC_RNC")
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
            'If btInserir.Text = "Aplicar" Then
            'Descricaox = Defeito1 & " " & lblDescricaoRNC1.Text & ", " & Defeito2 & " " & lblDescricaoRNC2.Text & ", " & Defeito3 & " " & lblDescricaoRNC3.Text & ", " & Defeito4 & " " & lblDescricaoRNC4.Text & ", " & Defeito5 & " " & lblDescricaoRNC5.Text & ", " & Defeito6 & " " & lblDescricaoRNC6.Text & ", " & Defeito7 & " " & lblDescricaoRNC7.Text & ", " & Defeito8 & " " & lblDescricaoRNC8.Text & ", " & Defeito9 & " " & lblDescricaoRNC9.Text & ", " & Defeito10 & " " & lblDescricaoRNC10.Text & ""
            'Else
            Descricaox = lblDescricaoRNC1.Text & ", " & lblDescricaoRNC2.Text & ", " & lblDescricaoRNC3.Text & ", " & lblDescricaoRNC4.Text & ", " & lblDescricaoRNC5.Text & ", " & lblDescricaoRNC6.Text & ", " & lblDescricaoRNC7.Text & ", " & lblDescricaoRNC8.Text & ", " & lblDescricaoRNC9.Text & ", " & lblDescricaoRNC10.Text & ""
            'End If
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
            RNC_RNC2 = Test("F:\RECEB.MAT.PRIMA\Banco_Dados\Documentos_RNC\RNC_" & lblRNC.Text & ".xlsx")
            If RNC_RNC2 = True Then 'se não existe
                ImprimirRNC()
            End If
            'System.Threading.Thread.Sleep(5000)
            OutlookMessage.Attachments.Add("F:\RECEB.MAT.PRIMA\Banco_Dados\Documentos_RNC\" & "RNC_" & lblRNC.Text & ".xlsx")
            OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatRichText

            'If (MsgBox("O E-mail está pronto para ser enviado. Deseja Enviar?" _
            '         & Chr(13) _
            '        & Chr(13) _
            '       & "'Sim' = Enviar" _
            '      & Chr(13) _
            '     & "'Não' = Alterar", vbYesNo, "E-mail") = vbYes) Then
            OutlookMessage.Save()
            OutlookMessage.Send()
            'Else
            'OutlookMessage.Display()
            'OutlookMessage.Save()
            'End If

        Catch ex As Exception
            MessageBox.Show("Erro 88 " & ex.Message) 'if you dont want this message, simply delete this line 
        Finally
            OutlookMessage = Nothing
            AppOutlook = Nothing
        End Try

    End Sub

    Private Sub cb1Turno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles cbColuna.KeyPress, cbOrdenadoPor.KeyPress, txtCodigoRNC1.KeyPress, cbTurno.KeyPress, cb1Turno.KeyPress, cb2Turnos.KeyPress, cb3Turnos.KeyPress, cb4Turnos.KeyPress, cb5Turno.KeyPress, cb6Turno.KeyPress, cb7Turno.KeyPress, cb8Turno.KeyPress, cb9Turno.KeyPress, cb10Turno.KeyPress, cbDetectado.KeyPress
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
                btEmail.Text = "Email..."
                email()
                btEmail.Text = "Email"
            End If
        Catch exc As Exception
            MsgBox("Erro 90 " & exc.Message)
        End Try
    End Sub
    Sub TesteAbertoConsultaOP()
        Try
            Dim Consulta_OP As Boolean
            Consulta_OP = Test("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb")
            If Consulta_OP = True Then
                Dim OPConvertida As Integer = 0
                For OPConvertida = 5 To 20
                    Consulta_OP = Test("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb")
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
        RNC_Defeito = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Defeito.accdb")
        If RNC_Defeito = True Then
            Dim RNCDefeito As Integer = 0
            For RNCDefeito = 5 To 20
                RNC_Defeito = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Defeito.accdb")
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
        RNC_Maquina = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb")
        If RNC_Maquina = True Then
            Dim RNCMaquina As Integer = 0
            For RNCMaquina = 5 To 20
                RNC_Maquina = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb")
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
        RNC_PecasVolume = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb")
        If RNC_PecasVolume = True Then
            Dim RNCPecasVolume As Integer = 0
            For RNCPecasVolume = 5 To 20
                RNC_PecasVolume = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb")
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
        RNC_RE = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RE.accdb")
        If RNC_RE = True Then
            Dim RNCRE As Integer = 0
            For RNCRE = 5 To 20
                RNC_RE = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RE.accdb")
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
        RNC_RNC = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RNC.accdb")
        If RNC_RNC = True Then
            Dim RNCRNC As Integer = 0
            For RNCRNC = 5 To 20
                RNC_RNC = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RNC.accdb")
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
        RNCDoc = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCDoc.xlsx")
        If RNCDoc = True Then
            Dim RNC_Doc As Integer = 0
            For RNC_Doc = 5 To 20
                RNCDoc = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCDoc.xlsx")
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
        RNCEtiqueta = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCEtiqueta.xlsx")
        If RNCEtiqueta = True Then
            Dim RNC_Etiqueta As Integer = 0
            For RNC_Etiqueta = 5 To 20
                RNCEtiqueta = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCEtiqueta.xlsx")
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

            Consulta_OP = Test("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb")
            RNC_Defeito = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Defeito.accdb")
            RNC_Maquina = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb")
            RNC_PecasVolume = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb")
            RNC_RE = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RE.accdb")
            RNC_RNC = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RNC.accdb")
            RNCDoc = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCDoc.xlsx")
            RNCEtiqueta = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCEtiqueta.xlsx")


            If Consulta_OP = True Then
                Dim OPConvertida As Integer = 0
                For OPConvertida = 5 To 20
                    Consulta_OP = Test("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb")
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
                    RNC_Defeito = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Defeito.accdb")
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
                    RNC_Maquina = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_Maquina.accdb")
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
                    RNC_PecasVolume = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_PecasVolume.accdb")
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
                    RNC_RE = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RE.accdb")
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
                    RNC_RNC = Test("f:\Receb.Mat.Prima\Banco_Dados\RNC_RNC.accdb")
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
                    RNCDoc = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCDoc.xlsx")
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
                    RNCEtiqueta = Test("f:\Receb.Mat.Prima\Banco_Dados\RNCEtiqueta.xlsx")
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
    Function Test2(ByVal pathfile As String) As Boolean
        Dim ff As Integer
        If System.IO.File.Exists(pathfile) Then
            Try
                ff = FreeFile()
                'Microsoft.VisualBasic.FileOpen(ff, pathfile, OpenMode.Binary, OpenAccess.Read, OpenShare.LockReadWrite)
                Microsoft.VisualBasic.FileOpen(ff, pathfile, OpenMode.Binary, OpenAccess.Default, OpenShare.LockReadWrite, RecordLength:=-1)
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
        frmMaquina.ShowDialog()
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

    Sub FormatacaoGrid()

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
        DataGridView1.Columns(14).HeaderText = "QT Reprovada"
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
        DataGridView1.Columns(14).Width = 85
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

        '3 - faz a coluna ajustar no resto do grid
        'DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

        'lblCodProduto.Text = DataGridView1.RowCount 'conta quantas RNCs exitem
    End Sub

    Private Sub btImprimirEtiqueta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btImprimirEtiqueta.Click
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

    Private Sub txtOP_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP.LostFocus

        Try
            If txtOP.Text = "" Or txtOP.Text = "0" Or txtOP.Text = "00" Or txtOP.Text = "000" Or txtOP.Text = "0000" Or txtOP.Text = "00000" Or txtOP.Text = "000000" Then
                If Limpo = "" Then
                    MsgBox("Insira uma 'OP' válida", , "OP")
                    txtOP.Focus()
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

   
End Class
