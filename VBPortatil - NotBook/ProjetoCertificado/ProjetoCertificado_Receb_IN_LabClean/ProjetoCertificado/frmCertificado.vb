Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Object
Imports System.Runtime.InteropServices
Imports System.IO
Imports Microsoft.Office.Interop


Public Class frmCertificado
    Dim conCertificado As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\dbCertificado.accdb;Jet OLEDB:Database Password=projetocertificado;")
    Dim conConsulta_OP As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim conPecasVolume As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\ProjetoCertificado\dbVolume.accdb;Jet OLEDB:Database Password=projetocertificado;")
    Dim conRE As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\RNC_RE.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim da As New OleDbDataAdapter
    Dim ds, ds12 As New DataSet
    Dim verRB, terRB As Byte
    'do controle()
    Dim txtOPx As New TextBox
    Dim txtVolumex As New TextBox
    Dim txtQuantidadex As New TextBox
    Dim dtpDex As New DateTimePicker
    Dim dtpAtex As New DateTimePicker
    Dim txtPecasPorVolumex As New TextBox
    Dim lblProdutox As New Label
    Dim txtClientex As New Label
    Dim lblCodigox As New TextBox
    Dim Obs1, Obs2, Obs3, Obs4, Obs5, Obs6, Obs7, Obs8, Obs9, Obs10, Obsx As Object
    'quantos anexos deve-se enviar
    Dim arrei_anexos(29) As Integer
    Dim val As Integer = 0
    Dim idx As Integer = 0
    Dim anex As Integer = 0
    Dim incluirEmail, interromper As String

    'OK
    Private Sub VerificarOP_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Today < "08/05/2016" Then

            Teste_dbCertificado()

            Teste_ConsultaOP()

            Teste_RNC_RE()

            AtualizarGrid()
        Else
            Me.Close()
        End If
    End Sub
    'OK
    Sub AtualizarGrid()

        Try

            ds.Clear()

            conCertificado.Open()

            Dim sel As String = "Select top 100 * from tblCertificado Where NotaFiscal = '' or Notafiscal is null order by ID desc "

            da = New OleDbDataAdapter(sel, conCertificado)

            da.Fill(ds, "tblCertificado")

            conCertificado.Close()

            Me.DataGridView1.DataSource = ds

            Me.DataGridView1.DataMember = "tblCertificado"

            FormatacaoGrid()

            lblData.Text = Today

            lblHora.Text = TimeOfDay

            txtInvoice.Enabled = False

            txtNotaFiscal.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    'OK
    Sub FormatacaoGrid()

        '1 - Coloca o Cabeçalho na coluna 
        DataGridView1.Columns(0).HeaderText = "ID"

        DataGridView1.Columns(1).HeaderText = "Pedido"

        DataGridView1.Columns(2).HeaderText = "Nota Fiscal"

        DataGridView1.Columns(3).HeaderText = "Produto"

        DataGridView1.Columns(4).HeaderText = "Código"

        DataGridView1.Columns(5).HeaderText = "Invoice"

        DataGridView1.Columns(6).HeaderText = "OP"

        DataGridView1.Columns(7).HeaderText = "Volume"

        DataGridView1.Columns(8).HeaderText = "Quantidade"

        DataGridView1.Columns(9).HeaderText = "Data"

        DataGridView1.Columns(10).HeaderText = "Hora"

        DataGridView1.Columns(11).HeaderText = "Data de Fabricação - Início"

        DataGridView1.Columns(12).HeaderText = "Data de Fabricação - Fim"

        DataGridView1.Columns(13).HeaderText = "Observação"

        DataGridView1.Columns(14).HeaderText = "Data Alterado"

        DataGridView1.Columns(15).HeaderText = "Hora Alterado"

        DataGridView1.Columns(16).HeaderText = "Cliente"

        DataGridView1.Columns(17).HeaderText = "Inspetor"


        '2 - Acerta a largura da coluna em pixels
        'DataGridView1.Columns(0).Width = 80

        '3 - faz a coluna ajustar no resto do grid
        DataGridView1.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(2).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(5).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(6).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(7).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(8).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(9).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(10).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(11).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(12).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(13).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(14).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(15).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(16).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        DataGridView1.Columns(17).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

        'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

        'lblCodProduto.Text = DataGridView1.RowCount 'conta quantas RNCs exitem
    End Sub
    'OK
    Sub FormatacaoGrid2()

        '1 - Coloca o Cabeçalho na coluna 
        DataGridView2.Columns(0).HeaderText = "Código"

        DataGridView2.Columns(1).HeaderText = "Produto"

        DataGridView2.Columns(2).HeaderText = "Cliente"

        DataGridView2.Columns(3).HeaderText = "Quantidade"

        '2 - Acerta a largura da coluna em pixels
        DataGridView2.Columns(0).Width = 0

        DataGridView2.Columns(1).Width = 0

        DataGridView2.Columns(2).Width = 300

        DataGridView2.Columns(3).Width = 0

        '3 - faz a coluna ajustar no resto do grid
        'DataGridView2.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill


        'Detalhe que isto pode ser feito para qualquer coluna, basta informar o respectivo indice

        'lblCodProduto.Text = DataGridView1.RowCount 'conta quantas RNCs exitem
    End Sub
    'OK
    Private Sub rbSim_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSim.CheckedChanged

        txtInvoice.Enabled = True

        txtNotaFiscal.Enabled = True

    End Sub
    'OK
    Private Sub rbNao_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNao.CheckedChanged

        txtInvoice.Enabled = False

        txtNotaFiscal.Enabled = False

        txtInvoice.Clear()

        txtNotaFiscal.Clear()


    End Sub
    'OK
    Private Sub btPesquisar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPesquisar.Click
        Try
            pb.Value = 0
            pb.Minimum = 0
            pb.Maximum = 110
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            If cbPesquisar.Text = "" Or txtPesquisar.TextLength = 0 Then
                MsgBox("Há Campos de Pesquisa Vazio")
            Else 'If cbColuna.Text = "RNC" Or cbColuna.Text = "Origem" Or cbColuna.Text = "Data_Abertura" Or cbColuna.Text = "Cod_Produto" Or cbColuna.Text = "Produto" Or cbColuna.Text = "OP_Reprovado" Or cbColuna.Text = "Turno" Or cbColuna.Text = "NúmerosCaixas" Or cbColuna.Text = "QT_Caixas" Or cbColuna.Text = "QT_P_Caixa" Or cbColuna.Text = "QT_Reprovado" Or cbColuna.Text = "Cod_Defeito" Or cbColuna.Text = "Nao_Conformidade" Or cbColuna.Text = "Maquina" Or cbColuna.Text = "Observacao" Or cbColuna.Text = "RE" Or cbColuna.Text = "Inspetor" Then
                DataGridView1.DataSource.clear()

                conCertificado.Open()

                Dim sel_ As String = "SELECT * FROM tblCertificado WHERE " & cbPesquisar.Text & " LIKE '" & "%" & txtPesquisar.Text & "%" & "' ORDER BY ID DESC "

                da19 = New OleDbDataAdapter(sel_, conCertificado)

                ds19.Clear()

                da19.Fill(ds19, "tblCertificado")

                Me.DataGridView1.DataSource = ds19

                Me.DataGridView1.DataMember = "tblCertificado"
                FormatacaoGrid()

                conCertificado.Close()

            End If
        Catch ex As Exception
            MsgBox("Erro 71 " & ex.Message)
            conCertificado.Close()
        Finally
            conCertificado.Close()

        End Try
    End Sub
    'OK
    Private Sub btCriar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCriar.Click
        pb.Value = 0
        pb.Minimum = 0
        pb.Maximum = 10000

        If btCriar.Text = "Criar" Then

            'If MsgBox("Deseja criar certificado(s) para um Pedido?", vbYesNo, "Novo(s) Certificado(s)") = vbYes Then

            Email_E_Impressao()

            LimparTudo()

            rbSim.Checked = True

            txtPedido.Focus()

            btCriar.Text = "Aplicar"

            btAlterar.Enabled = False

            btExcluir.Enabled = False

            btImprimir.Enabled = False

            btEmail.Enabled = False

            lblData.Text = Today

            lblHora.Text = TimeOfDay.ToShortTimeString

            DataGridView1.Enabled = False

            ProximoID()

            GroupBox4.Enabled = False

            'Else

            'End If
            'se botão Criar for  = Aplicar
        Else
            interromper = "Não"

            CriarEnviar()

            If interromper = "Não" Then

                LimparTudo()

            End If

            AtualizarGrid()

        End If

    End Sub
    'OK
    Sub LimparTudo()
        imprimirx = ""

        btImprimir.Text = "Imprimir"

        DataGridView1.Enabled = True

        lblID.Text = "*"

        txtPedido.Clear()

        lblData.Text = Today

        lblHora.Text = TimeOfDay

        rbNao.Checked = True

        txtNotaFiscal.Clear()

        txtInvoice.Clear()

        lblDataAlterado.Text = Today

        lblHoraAlterado.Text = TimeOfDay

        txtRE.Clear()

        lblInspetor.Text = ""

        btCriar.Text = "Criar"

        btCriar.Enabled = True

        btAlterar.Text = "Alterar"

        btAlterar.Enabled = True

        btAlterarIndividual.Text = "Alterar"

        btAlterarIndividual.Enabled = True

        btImprimirIndividual.Enabled = True

        btImprimirIndividual.Text = "Imprimir"

        btEmailIndividual.Enabled = True

        btExcluir.Text = "Excluir"

        btExcluir.Enabled = True

        btImprimir.Enabled = True

        btEmail.Enabled = True

        btEmailIndividual.Text = "Email"

        conCertificado.Close()

        lblTotal.Text = "0"

        rb1T.Checked = True

        verRB = 0

        terRB = 0

        btEmail.Text = "Email"

        GroupBox2.Enabled = True

        LimparContinua()


        arrei_anexos.Initialize()

        val = 0

        idx = 0

        anex = 0

        incluirEmail = ""

        lblTotal.Text = "0"

        GroupBox4.Enabled = True

        ds12.Clear()
        Me.DataGridView2.DataSource = ds12

    End Sub
    'OK
    Sub LimparContinua()

        CheckBox1.Checked = False

        CheckBox2.Checked = False

        CheckBox3.Checked = False

        CheckBox4.Checked = False

        CheckBox5.Checked = False

        CheckBox6.Checked = False

        CheckBox7.Checked = False

        CheckBox8.Checked = False

        CheckBox9.Checked = False

        CheckBox10.Checked = False

        Obsx = ""

        Obs1 = ""

        Obs2 = ""

        Obs3 = ""

        Obs4 = ""

        Obs5 = ""

        Obs6 = ""

        Obs7 = ""

        Obs8 = ""

        Obs9 = ""

        Obs10 = ""


        txtOP1.Clear()

        txtOP2.Clear()

        txtOP3.Clear()

        txtOP4.Clear()

        txtOP5.Clear()

        txtOP6.Clear()

        txtOP7.Clear()

        txtOP8.Clear()

        txtOP9.Clear()

        txtOP10.Clear()


        lblProduto1.Text = "*"

        lblProduto2.Text = "*"

        lblProduto3.Text = "*"

        lblProduto4.Text = "*"

        lblProduto5.Text = "*"

        lblProduto6.Text = "*"

        lblProduto7.Text = "*"

        lblProduto8.Text = "*"

        lblProduto9.Text = "*"

        lblProduto10.Text = "*"

        lblCodigo1.Text = "*"

        lblCodigo2.Text = "*"

        lblCodigo3.Text = "*"

        lblCodigo4.Text = "*"

        lblCodigo5.Text = "*"

        lblCodigo6.Text = "*"

        lblCodigo7.Text = "*"

        lblCodigo8.Text = "*"

        lblCodigo9.Text = "*"

        lblCodigo10.Text = "*"


        txtCliente1.Clear()

        txtCliente2.Clear()

        txtCliente3.Clear()

        txtCliente4.Clear()

        txtCliente5.Clear()

        txtCliente6.Clear()

        txtCliente7.Clear()

        txtCliente8.Clear()

        txtCliente9.Clear()

        txtCliente10.Clear()


        txtVolume1.Text = "0"

        txtVolume2.Text = "0"

        txtVolume3.Text = "0"

        txtVolume4.Text = "0"

        txtVolume5.Text = "0"

        txtVolume6.Text = "0"

        txtVolume7.Text = "0"

        txtVolume8.Text = "0"

        txtVolume9.Text = "0"
        txtVolume10.Text = "0"


        txtQuantidade1.Text = "0"

        txtQuantidade2.Text = "0"

        txtQuantidade3.Text = "0"

        txtQuantidade4.Text = "0"

        txtQuantidade5.Text = "0"

        txtQuantidade6.Text = "0"

        txtQuantidade7.Text = "0"

        txtQuantidade8.Text = "0"

        txtQuantidade9.Text = "0"

        txtQuantidade10.Text = "0"

        txtPecasPorVolume1.Text = "0"

        txtPecasPorVolume2.Text = "0"

        txtPecasPorVolume3.Text = "0"

        txtPecasPorVolume4.Text = "0"

        txtPecasPorVolume5.Text = "0"

        txtPecasPorVolume6.Text = "0"

        txtPecasPorVolume7.Text = "0"

        txtPecasPorVolume8.Text = "0"

        txtPecasPorVolume9.Text = "0"

        txtPecasPorVolume10.Text = "0"


        dtpDe1.Value = Today.ToShortDateString()

        dtpDe2.Value = Today.ToShortDateString()

        dtpDe3.Value = Today.ToShortDateString()

        dtpDe4.Value = Today.ToShortDateString()

        dtpDe5.Value = Today.ToShortDateString()

        dtpDe6.Value = Today.ToShortDateString()

        dtpDe7.Value = Today.ToShortDateString()

        dtpDe8.Value = Today.ToShortDateString()

        dtpDe9.Value = Today.ToShortDateString()

        dtpDe10.Value = Today.ToShortDateString()


        dtpAte1.Value = Today.ToShortDateString()

        dtpAte2.Value = Today.ToShortDateString()

        dtpAte3.Value = Today.ToShortDateString()

        dtpAte4.Value = Today.ToShortDateString()

        dtpAte5.Value = Today.ToShortDateString()

        dtpAte6.Value = Today.ToShortDateString()

        dtpAte7.Value = Today.ToShortDateString()

        dtpAte8.Value = Today.ToShortDateString()

        dtpAte9.Value = Today.ToShortDateString()

        dtpAte10.Value = Today.ToShortDateString()


    End Sub
    'OK
    Sub ProximoID()

        Try

            Dim da19 As New OleDbDataAdapter

            Dim ds19 As New DataSet

            Dim dt19 As New System.Data.DataTable

            conCertificado.Open()

            Dim sel_ As String = "SELECT TOP 1 ID FROM tblCertificado ORDER BY ID DESC "

            da19 = New OleDbDataAdapter(sel_, conCertificado)

            dt19.Clear()

            da19.Fill(dt19)

            conCertificado.Close()

            lblID.Text = dt19.Rows(0)("ID") + 1

        Catch ex As Exception
            MsgBox("Erro 71 " & ex.Message)
            conCertificado.Close()
        Finally

            conCertificado.Close()

        End Try

    End Sub
    'OK

    Sub CriarEnviar()

        VerPadrao()

        If interromper = "Sim" Then

        ElseIf interromper = "Não" Then

            Criar()

        End If
    End Sub
    'OK
    Sub VerificarRB()

        If rb1T.Checked = True Then

            verRB = 1

            terRB = 1

        ElseIf rb2T.Checked = True Then

            verRB = 2

            terRB = 2

        ElseIf rb3T.Checked = True Then

            verRB = 3

            terRB = 3

        ElseIf rb4T.Checked = True Then
            verRB = 4

            terRB = 4

        ElseIf rb5T.Checked = True Then

            verRB = 5

            terRB = 5

        ElseIf rb6T.Checked = True Then

            verRB = 6

            terRB = 6

        ElseIf rb7T.Checked = True Then

            verRB = 7

            terRB = 7

        ElseIf rb8T.Checked = True Then

            verRB = 8

            terRB = 8

        ElseIf rb9T.Checked = True Then

            verRB = 9

            terRB = 9

        ElseIf rb10T.Checked = True Then

            verRB = 10

            terRB = 10

        Else
            MsgBox("Selecione a quantidade de OP´s")
        End If

    End Sub
    'OK
    Sub VerPadrao()

        If txtPedido.TextLength >= 3 Then

        Else
            MsgBox("Insira o 'Pedido!'", MsgBoxStyle.Exclamation, "Pedido")

            txtPedido.Focus()

            interromper = "Sim"

            Return
        End If

        If rbSim.Checked = True Then

            If txtNotaFiscal.TextLength >= 4 Then


            Else

                MsgBox("Insira a 'Nota Fiscal!'", MsgBoxStyle.Exclamation, "Nota Fiscal")

                txtNotaFiscal.Focus()

                interromper = "Sim"

                Return
            End If

        End If

        If txtRE.TextLength >= 3 Then


        Else

            MsgBox("Insira o 'RE!'", MsgBoxStyle.Exclamation, "RE")

            txtRE.Focus()

            interromper = "Sim"

            Return
        End If

        If lblCodigo1.Text.Remove(4) = "3007" Then

            If txtInvoice.TextLength >= 2 Then


            Else
                MsgBox("Insira a 'Ivoice!'", MsgBoxStyle.Exclamation, "Invoice")

                txtInvoice.Focus()

                interromper = "Sim"

                Return
            End If

        Else

            txtInvoice.Clear()

        End If

        ' continua no ver campos
        interromper = "Não"

        VerCampos()

    End Sub
    'OK
    Sub VerCampos()

        VerificarRB()

        Select Case verRB
            Case 1

                VerCampos1()

            Case 2

                VerCampos1()

                VerCampos2()

            Case 3

                VerCampos1()

                VerCampos2()

                VerCampos3()

            Case 4

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

            Case 5

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

                VerCampos5()

            Case 6

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

                VerCampos5()

                VerCampos6()

            Case 7

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

                VerCampos5()

                VerCampos6()

                VerCampos7()

            Case 8

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

                VerCampos5()

                VerCampos6()

                VerCampos7()

                VerCampos8()

            Case 9

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

                VerCampos5()

                VerCampos6()

                VerCampos7()

                VerCampos8()

                VerCampos9()

            Case 10

                VerCampos1()

                VerCampos2()

                VerCampos3()

                VerCampos4()

                VerCampos5()

                VerCampos6()

                VerCampos7()

                VerCampos8()

                VerCampos9()

                VerCampos10()

        End Select

    End Sub
    'OK
    Sub VerCampos1()

        If txtOP1.TextLength >= 5 Then

        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP1.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente1.TextLength >= 4 Then

        Else
            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente1.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume1.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume1.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade1.TextLength >= 1 Then


        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade1.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos2()

        If txtOP2.TextLength >= 5 Then

        Else
            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP2.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente2.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente2.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume2.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume2.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade2.TextLength >= 1 Then


        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade2.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos3()

        If txtOP3.TextLength >= 5 Then


        Else
            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP3.Focus()

            interromper = "Sim"

            Return
        End If


        If txtCliente3.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente3.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume3.TextLength >= 1 Then


        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume3.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade3.TextLength >= 1 Then

        Else
            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade3.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos4()

        If txtOP4.TextLength >= 5 Then


        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP4.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente4.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente4.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume4.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume4.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade4.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade4.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos5()

        If txtOP5.TextLength >= 5 Then


        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP5.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente5.TextLength >= 4 Then

        Else
            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente5.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume5.TextLength >= 1 Then


        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume5.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade5.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade5.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos6()

        If txtOP6.TextLength >= 5 Then

        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP6.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente6.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente6.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume6.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume6.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade6.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade6.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos7()

        If txtOP7.TextLength >= 5 Then

        Else
            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP7.Focus()

            interromper = "Sim"

            Return
        End If


        If txtCliente7.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente7.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume7.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume7.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade7.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade7.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos8()

        If txtOP8.TextLength >= 5 Then

        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP8.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente8.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente8.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume8.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume8.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade8.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade8.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos9()

        If txtOP9.TextLength >= 5 Then

        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP9.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente9.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente9.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume9.TextLength >= 1 Then

        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume9.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade9.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade9.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    Sub VerCampos10()

        If txtOP10.TextLength >= 5 Then

        Else

            MsgBox("Insira a 'OP!'", MsgBoxStyle.Exclamation, "OP")

            txtOP10.Focus()

            interromper = "Sim"

            Return
        End If

        If txtCliente10.TextLength >= 4 Then

        Else

            MsgBox("Insira o 'Cliente!'", MsgBoxStyle.Exclamation, "Cliente")

            txtCliente10.Focus()

            interromper = "Sim"

            Return
        End If

        If txtVolume10.TextLength >= 1 Then


        Else

            MsgBox("Insira o 'Volume!'", MsgBoxStyle.Exclamation, "Volume")

            txtVolume10.Focus()

            interromper = "Sim"

            Return
        End If

        If txtQuantidade10.TextLength >= 1 Then

        Else

            MsgBox("Insira a 'Quantidade!'", MsgBoxStyle.Exclamation, "Quantidade")

            txtQuantidade10.Focus()

            interromper = "Sim"

            Return
        End If

    End Sub
    'OK
    Sub Criar()

        VerificarRB()

        Try

            Dim i As Byte = 0

            conCertificado.Open()

            Dim data, hora As Date

            data = Today

            hora = TimeOfDay

            For i = 1 To terRB Step 1

                Controles()

                Dim da4 As New OleDbDataAdapter

                Dim ds4 As New DataSet

                ds4 = New DataSet

                da4 = New OleDbDataAdapter("INSERT INTO tblCertificado (Pedido, NotaFiscal, Produto, Codigo, Invoice, OP, Volume, Quantidade, Data, Hora, DataFab_Inicio, DataFab_Fim, Obs, Cliente, Inspetor) Values ('" & txtPedido.Text & "', '" & txtNotaFiscal.Text & "', '" & lblProduto1.Text & "', '" & lblCodigo1.Text & "', '" & txtInvoice.Text & "', '" & txtOPx.Text & "', '" & txtVolumex.Text & "', '" & txtQuantidadex.Text & "','" & data.ToShortDateString() & "', '" & hora.ToShortTimeString() & "', '" & dtpDex.Value & "', '" & dtpAtex.Value & "', '" & Obsx.ToString() & "', '" & txtCliente1.Text & "', '" & lblInspetor.Text & "') ", conCertificado)

                ds4.Clear()

                da4.Fill(ds4, "tblCertificado")

                conCertificado.Close()

                If rbSim.Checked = True Then

                    LansarNoExcel() ' lança no excell, salva em PDF e Imprime
                    ' adicionando quantos anexos deve-se enviar

                    idx = anex

                    val = Integer.Parse(lblID.Text)

                    arrei_anexos.SetValue(val, idx)

                    anex += 1

                End If

            Next

            If rbSim.Checked = True Then

                If cbEnviarEmail.Checked = True Then

                    Email() 'Armazena os anexos para Email() e depois EneviarEmail()

                End If

            End If

            'MsgBox("Dados inseridos com sucesso")

        Catch ex As Exception
            conCertificado.Close()
            MsgBox("Erro 15 " & ex.Message)
        End Try

    End Sub
    'OK
    Sub Controles()

        Select Case verRB
            Case 1

                txtOPx.Text = txtOP1.Text

                txtVolumex.Text = txtVolume1.Text

                txtQuantidadex.Text = txtQuantidade1.Text

                dtpDex.Value = dtpDe1.Value

                dtpAtex.Value = dtpAte1.Value

                Observacao1()

                Obsx = Obs1

                txtPecasPorVolumex.Text = txtPecasPorVolume1.Text

                lblProdutox.Text = lblProduto1.Text

                txtClientex.Text = txtCliente1.Text

                lblCodigox.Text = lblCodigo1.Text

            Case 2

                txtOPx.Text = txtOP2.Text

                txtVolumex.Text = txtVolume2.Text

                txtQuantidadex.Text = txtQuantidade2.Text

                dtpDex.Value = dtpDe2.Text

                dtpAtex.Value = dtpAte2.Text

                Observacao2()

                Obsx = Obs2

                txtPecasPorVolumex.Text = txtPecasPorVolume2.Text

                lblProdutox.Text = lblProduto2.Text

                txtClientex.Text = txtCliente2.Text

                lblCodigox.Text = lblCodigo2.Text

                verRB -= 1

            Case 3

                txtOPx.Text = txtOP3.Text

                txtVolumex.Text = txtVolume3.Text

                txtQuantidadex.Text = txtQuantidade3.Text

                dtpDex.Value = dtpDe3.Text

                dtpAtex.Value = dtpAte3.Text

                Observacao3()

                Obsx = Obs3

                txtPecasPorVolumex.Text = txtPecasPorVolume3.Text

                lblProdutox.Text = lblProduto3.Text

                txtClientex.Text = txtCliente3.Text

                lblCodigox.Text = lblCodigo3.Text

                verRB -= 1

            Case 4

                txtOPx.Text = txtOP4.Text

                txtVolumex.Text = txtVolume4.Text

                txtQuantidadex.Text = txtQuantidade4.Text

                dtpDex.Value = dtpDe4.Text

                dtpAtex.Value = dtpAte4.Text

                Observacao4()

                Obsx = Obs4

                txtPecasPorVolumex.Text = txtPecasPorVolume4.Text

                lblProdutox.Text = lblProduto4.Text

                txtClientex.Text = txtCliente4.Text

                lblCodigox.Text = lblCodigo4.Text

                verRB -= 1

            Case 5

                txtOPx.Text = txtOP5.Text

                txtVolumex.Text = txtVolume5.Text

                txtQuantidadex.Text = txtQuantidade5.Text

                dtpDex.Value = dtpDe5.Text

                dtpAtex.Value = dtpAte5.Text

                Observacao5()

                Obsx = Obs5

                txtPecasPorVolumex.Text = txtPecasPorVolume5.Text

                lblProdutox.Text = lblProduto5.Text

                txtClientex.Text = txtCliente5.Text

                lblCodigox.Text = lblCodigo5.Text

                verRB -= 1

            Case 6
                txtOPx.Text = txtOP6.Text

                txtVolumex.Text = txtVolume6.Text

                txtQuantidadex.Text = txtQuantidade6.Text

                dtpDex.Value = dtpDe6.Text

                dtpAtex.Value = dtpAte6.Text

                Observacao6()

                Obsx = Obs6

                txtPecasPorVolumex.Text = txtPecasPorVolume6.Text

                lblProdutox.Text = lblProduto6.Text

                txtClientex.Text = txtCliente6.Text

                lblCodigox.Text = lblCodigo6.Text

                verRB -= 1

            Case 7

                txtOPx.Text = txtOP7.Text

                txtVolumex.Text = txtVolume7.Text

                txtQuantidadex.Text = txtQuantidade7.Text

                dtpDex.Value = dtpDe7.Text

                dtpAtex.Value = dtpAte7.Text

                Observacao7()

                Obsx = Obs7

                txtPecasPorVolumex.Text = txtPecasPorVolume7.Text

                lblProdutox.Text = lblProduto7.Text

                txtClientex.Text = txtCliente7.Text

                lblCodigox.Text = lblCodigo7.Text

                verRB -= 1

            Case 8

                txtOPx.Text = txtOP8.Text

                txtVolumex.Text = txtVolume8.Text

                txtQuantidadex.Text = txtQuantidade8.Text

                dtpDex.Value = dtpDe8.Text

                dtpAtex.Value = dtpAte8.Text

                Observacao8()

                Obsx = Obs8

                txtPecasPorVolumex.Text = txtPecasPorVolume8.Text

                lblProdutox.Text = lblProduto8.Text

                txtClientex.Text = txtCliente8.Text

                lblCodigox.Text = lblCodigo8.Text

                verRB -= 1

            Case 9

                txtOPx.Text = txtOP9.Text

                txtVolumex.Text = txtVolume9.Text

                txtQuantidadex.Text = txtQuantidade9.Text

                dtpDex.Value = dtpDe9.Text

                dtpAtex.Value = dtpAte9.Text

                Observacao9()

                Obsx = Obs9

                txtPecasPorVolumex.Text = txtPecasPorVolume9.Text

                lblProdutox.Text = lblProduto9.Text

                txtClientex.Text = txtCliente9.Text

                lblCodigox.Text = lblCodigo9.Text

                verRB -= 1

            Case 10

                txtOPx.Text = txtOP10.Text

                txtVolumex.Text = txtVolume10.Text

                txtQuantidadex.Text = txtQuantidade10.Text

                dtpDex.Value = dtpDe10.Text

                dtpAtex.Value = dtpAte10.Text

                Observacao10()

                Obsx = Obs10

                txtPecasPorVolumex.Text = txtPecasPorVolume10.Text

                lblProdutox.Text = lblProduto10.Text

                txtClientex.Text = txtCliente10.Text

                lblCodigox.Text = lblCodigo10.Text

                verRB -= 1

        End Select

    End Sub
    'OK
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
    'OK
    Private Sub txtOP1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP1.LostFocus

        verRB = 1

        txtOPx.Text = txtOP1.Text

        ConsultaOP()

        dtpDe1.Text = dtpDex.Value

        txtPecasPorVolume1.Text = txtPecasPorVolumex.Text

        lblProduto1.Text = lblProdutox.Text

        txtCliente1.Text = txtClientex.Text

        lblCodigo1.Text = lblCodigox.Text


    End Sub
    Private Sub txtOP2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP2.LostFocus

        verRB = 2

        txtOPx.Text = txtOP2.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto2.Text = "OP com produtos divergentes"

            txtOP2.Clear()

        Else

            dtpDe2.Text = dtpDex.Value

            txtPecasPorVolume2.Text = txtPecasPorVolumex.Text

            lblProduto2.Text = lblProdutox.Text

            txtCliente2.Text = txtClientex.Text

            lblCodigo2.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP3.LostFocus

        verRB = 3

        txtOPx.Text = txtOP3.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto3.Text = "OP com produtos divergentes"

            txtOP3.Clear()

        Else

            dtpDe3.Text = dtpDex.Value

            txtPecasPorVolume3.Text = txtPecasPorVolumex.Text

            lblProduto3.Text = lblProdutox.Text

            txtCliente3.Text = txtClientex.Text

            lblCodigo3.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP4.LostFocus

        verRB = 4

        txtOPx.Text = txtOP4.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto4.Text = "OP com produtos divergentes"

            txtOP4.Clear()

        Else

            dtpDe4.Text = dtpDex.Value

            txtPecasPorVolume4.Text = txtPecasPorVolumex.Text

            lblProduto4.Text = lblProdutox.Text

            txtCliente4.Text = txtClientex.Text

            lblCodigo4.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP5.LostFocus

        verRB = 5

        txtOPx.Text = txtOP5.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto5.Text = "OP com produtos divergentes"

            txtOP5.Clear()

        Else

            dtpDe5.Text = dtpDex.Value

            txtPecasPorVolume5.Text = txtPecasPorVolumex.Text

            lblProduto5.Text = lblProdutox.Text

            txtCliente5.Text = txtClientex.Text

            lblCodigo5.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP6.LostFocus
        verRB = 6
        txtOPx.Text = txtOP6.Text
        ConsultaOP()
        If lblCodigo1.Text <> lblCodigox.Text Then
            lblProduto6.Text = "OP com produtos divergentes"
            txtOP6.Clear()
        Else
            dtpDe6.Text = dtpDex.Value
            txtPecasPorVolume6.Text = txtPecasPorVolumex.Text
            lblProduto6.Text = lblProdutox.Text
            txtCliente6.Text = txtClientex.Text
            lblCodigo6.Text = lblCodigox.Text
        End If
    End Sub
    Private Sub txtOP7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP7.LostFocus

        verRB = 7

        txtOPx.Text = txtOP7.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto7.Text = "OP com produtos divergentes"

            txtOP7.Clear()

        Else

            dtpDe7.Text = dtpDex.Value

            txtPecasPorVolume7.Text = txtPecasPorVolumex.Text

            lblProduto7.Text = lblProdutox.Text

            txtCliente7.Text = txtClientex.Text

            lblCodigo7.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP8.LostFocus

        verRB = 8

        txtOPx.Text = txtOP8.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto8.Text = "OP com produtos divergentes"

            txtOP8.Clear()

        Else

            dtpDe8.Text = dtpDex.Value

            txtPecasPorVolume8.Text = txtPecasPorVolumex.Text

            lblProduto8.Text = lblProdutox.Text

            txtCliente8.Text = txtClientex.Text

            lblCodigo8.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP9.LostFocus
        verRB = 9

        txtOPx.Text = txtOP9.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto9.Text = "OP com produtos divergentes"

            txtOP9.Clear()

        Else

            dtpDe9.Text = dtpDex.Value

            txtPecasPorVolume9.Text = txtPecasPorVolumex.Text

            lblProduto9.Text = lblProdutox.Text

            txtCliente9.Text = txtClientex.Text

            lblCodigo9.Text = lblCodigox.Text

        End If
    End Sub
    Private Sub txtOP10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP10.LostFocus
        verRB = 10

        txtOPx.Text = txtOP10.Text

        ConsultaOP()

        If lblCodigo1.Text <> lblCodigox.Text Then

            lblProduto10.Text = "OP com produtos divergentes"

            txtOP10.Clear()

        Else

            dtpDe10.Text = dtpDex.Value

            txtPecasPorVolume10.Text = txtPecasPorVolumex.Text

            lblProduto10.Text = lblProdutox.Text

            txtCliente10.Text = txtClientex.Text

            lblCodigo10.Text = lblCodigox.Text

        End If
    End Sub
    'OK
    Dim dt10 As New System.Data.DataTable
    Sub ConsultaOP()

        Try

            If txtOPx.Text = "" Or txtOPx.Text = "0" Or txtOPx.Text = "00" Or txtOP1.Text = "000" Or txtOP1.Text = "0000" Or txtOP1.Text = "00000" Or txtOP1.Text = "000000" Then

                MsgBox("Insira uma 'OP' válida", , "OP")

            Else

                Dim da10 As New OleDbDataAdapter

                Dim cb10 As New OleDbCommandBuilder

                conConsulta_OP.Open()

                Dim sel12 As String = "SELECT top 1 Cod_Mondicap, Dt_Abertura FROM tblOP where OP = " & txtOPx.Text & ""

                da10 = New OleDbDataAdapter(sel12, conConsulta_OP)

                dt10.Clear()

                da10.Fill(dt10)


                TesteModeloCertificado()

                If dt10.Rows.Count = 0 Then

                    conConsulta_OP.Close()

                    MsgBox("A OP não existe")

                    lblCodigox.Text = "*"

                    lblProdutox.Text = "*"

                    Return
                Else

                    If dt10.Rows(0)("Cod_Mondicap").ToString.Remove(1) <> 3 Then

                        MsgBox("A OP não é compatível com produto acabado " & dt10.Rows(0)("Cod_Mondicap"))

                        txtOPx.Clear()

                        conConsulta_OP.Close()

                        Return
                    End If

                    lblCodigox.Text = dt10.Rows(0)("Cod_Mondicap")

                    dtpDex.Value = dt10.Rows(0)("Dt_Abertura")

                    conConsulta_OP.Close()

                    Dim da12 As New OleDbDataAdapter

                    Dim dt12 As New System.Data.DataTable

                    Dim ds12 As New DataSet

                    conPecasVolume.Open()

                    Dim sel5 As String = "SELECT * FROM tblVolume where Codigo = '" & lblCodigox.Text & "'"

                    da12 = New OleDbDataAdapter(sel5, conPecasVolume)

                    dt12.Clear()

                    da12.Fill(dt12)

                    If dt12.Rows.Count = 0 Then

                        conPecasVolume.Close()

                        conConsulta_OP.Close()

                        MsgBox("'Peças Por Volume' Não Cadastrado, insira a quantidade manualmente e solicite o cadastro para este item", , "Peças por Volume")

                        txtPecasPorVolumex.Clear()

                    ElseIf dt12.Rows.Count = 1 Then

                        da12.Fill(ds12, "tblVolume")

                        txtPecasPorVolumex.Text = ds12.Tables("tblVolume").Rows(0)("Quantidade")

                        lblProdutox.Text = ds12.Tables("tblVolume").Rows(0)("Produto")

                        txtClientex.Text = ds12.Tables("tblVolume").Rows(0)("Cliente")

                        incluirEmail = Convert.ToString(ds12.Tables("tblVolume").Rows(0)("Email"))

                        conPecasVolume.Close()

                    ElseIf dt12.Rows.Count > 1 Then

                        ds12.Clear()
                        da12.Fill(ds12, "tblVolume")

                        txtPecasPorVolumex.Text = ds12.Tables("tblVolume").Rows(0)("Quantidade")

                        lblProdutox.Text = ds12.Tables("tblVolume").Rows(0)("Produto")

                        Me.DataGridView2.DataSource = ds12

                        Me.DataGridView2.DataMember = "tblVolume"

                        FormatacaoGrid2()

                        conPecasVolume.Close()

                    End If
                    conConsulta_OP.Close()

                End If

            End If

        Catch ex As Exception
            conConsulta_OP.Close()
            conPecasVolume.Close()
            MsgBox("Erro 30 " & ex.Message)
        End Try

    End Sub

    Sub TesteModeloCertificado()

        Try

            Dim modelo_Certificado As Boolean

            modelo_Certificado = Test("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Modelos_Certificados\" & Convert.ToString(dt10.Rows(0)("Cod_Mondicap")) & ".xlsx")

            If modelo_Certificado = True Then

                Dim OPConvertida As Integer = 0

                For OPConvertida = 5 To 20

                    modelo_Certificado = Test("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Modelos_Certificados\" & Convert.ToString(dt10.Rows(0)("Cod_Mondicap")) & ".xlsx")

                    If modelo_Certificado = True Then

                        OPConvertida = 5

                        If (MsgBox("O arquivo " & Convert.ToString(dt10.Rows(0)("Cod_Mondicap")) & ".xlsx não existe ou em uso" _
                                   & Chr(13) _
                                   & Chr(13) _
                                   & "Não será possível avaçar...", vbRetryCancel, "Arquivo inexixtente!")) = vbRetry Then

                        Else

                            Exit For

                        End If

                    ElseIf modelo_Certificado = False Then

                        OPConvertida = 20

                    End If

                Next

            End If

        Catch e As Exception
            MsgBox("Erro T49 " & e.Message)
        End Try

    End Sub
    Sub Teste_dbCertificado()

        Try

            Dim dbCertificado As Boolean

            dbCertificado = Test("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\dbCertificado.accdb")

            If dbCertificado = True Then

                Dim OPConvertida As Integer = 0

                For OPConvertida = 5 To 20

                    dbCertificado = Test("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\dbCertificado.accdb")

                    If dbCertificado = True Then

                        OPConvertida = 5

                        If (MsgBox("O Arquivo 'dbCertificado.accdb' está aberto, Feche-o para para continuar", vbRetryCancel, "dbCertificado.accdb Aberto")) = vbRetry Then

                        Else

                            Close()

                            Exit For

                        End If

                    ElseIf dbCertificado = False Then

                        OPConvertida = 20

                    End If

                Next

            End If

        Catch e As Exception
            MsgBox("Erro T43 " & e.Message)
        End Try

    End Sub
    Sub Teste_ConsultaOP()

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
            MsgBox("Erro T45 " & e.Message)
        End Try

    End Sub
    Sub Teste_RNC_RE()

        Try

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

        Catch e As Exception
            MsgBox("Erro T47 " & e.Message)
        End Try

    End Sub

    Function Test(ByVal pathfile As String) As Boolean

        Dim ff As Integer

        If System.IO.File.Exists(pathfile) Then

            Try

                ff = FreeFile()
                Microsoft.VisualBasic.FileOpen(ff, pathfile, OpenMode.Input)

                Return False

            Catch ex As Exception

                Return True

            Finally

                FileClose(ff)

            End Try

            Return True

        Else

        End If

        Return True

    End Function

    'OK
    Private Sub txtOP1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP1.TextChanged

        CheckBox1.Text = txtOP1.Text

    End Sub
    Private Sub txtOP2_TextChanged_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP2.TextChanged

        CheckBox2.Text = txtOP2.Text

    End Sub
    Private Sub txtOP3_TextChanged_3(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP3.TextChanged

        CheckBox3.Text = txtOP3.Text

    End Sub
    Private Sub txtOP4_TextChanged_4(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP4.TextChanged

        CheckBox4.Text = txtOP4.Text

    End Sub
    Private Sub txtOP5_TextChanged_5(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP5.TextChanged

        CheckBox5.Text = txtOP5.Text

    End Sub
    Private Sub txtOP6_TextChanged_6(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP6.TextChanged

        CheckBox6.Text = txtOP6.Text

    End Sub
    Private Sub txtOP7_TextChanged_7(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP7.TextChanged

        CheckBox7.Text = txtOP7.Text

    End Sub
    Private Sub txtOP8_TextChanged_8(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP8.TextChanged

        CheckBox8.Text = txtOP8.Text

    End Sub
    Private Sub txtOP9_TextChanged_9(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP9.TextChanged

        CheckBox9.Text = txtOP9.Text

    End Sub
    Private Sub txtOP10_TextChanged_10(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOP10.TextChanged

        CheckBox10.Text = txtOP10.Text

    End Sub
    'OK
    Sub Observacao1()

        If CheckBox1.Checked = True Then

            If btAlterarIndividual.Text = "Aplicar" Or btImprimirIndividual.Text = "...Imprimir" Then

            Else

                Obs1 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP1.Text, XPos:=615, YPos:=300)

            End If

        End If

    End Sub
    Sub Observacao2()

        If CheckBox2.Checked = True Then

            Obs2 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP2.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao3()

        If CheckBox3.Checked = True Then

            Obs3 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP3.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao4()

        If CheckBox4.Checked = True Then

            Obs4 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP4.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao5()

        If CheckBox5.Checked = True Then

            Obs5 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP5.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao6()

        If CheckBox6.Checked = True Then

            Obs6 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP6.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao7()

        If CheckBox7.Checked = True Then

            Obs7 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP7.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao8()

        If CheckBox8.Checked = True Then

            Obs8 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP8.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao9()

        If CheckBox9.Checked = True Then

            Obs9 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP9.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    Sub Observacao10()

        If CheckBox10.Checked = True Then

            Obs10 = InputBox(Title:="Observação", Prompt:="Insira a Observação da OP: " & txtOP10.Text, XPos:=615, YPos:=300)

        End If

    End Sub
    'OK
    Private Sub btCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelar.Click

        LimparTudo()

    End Sub
    'OK
    Private Sub txtRE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRE.LostFocus

        If txtRE.Text = "" Or txtRE.Text = "0" Or txtRE.Text = "00" Or txtRE.Text = "000" Or txtRE.Text = "0000" Then

            MsgBox("Insira um 'RE' válido", , "RE")

            txtRE.Clear()

            txtRE.Focus()

        Else

            Try

                Dim da13 As New OleDbDataAdapter

                Dim ds13 As New DataSet

                Dim re As Integer

                conRE.Open()

                If txtRE.TextLength = 0 Then

                    conRE.Close()

                End If

                Dim sel6 As String = "SELECT Inspetor FROM tblRE where RE = " & txtRE.Text & ""

                da13 = New OleDbDataAdapter(sel6, conRE)

                ds13.Clear()

                da13.Fill(ds13, "tblRE")

                lblInspetor.Text = "*"

                re = ds13.Tables("tblRE").Rows.Count

                If re <= 0 Then

                    conRE.Close()

                    MsgBox("'RE' inexitente! Insira um RE válido", , "RE")

                    txtRE.Clear()

                    txtRE.Focus()

                Else

                    conRE.Close()

                    lblInspetor.Text = txtRE.Text & " - " & ds13.Tables("tblRE").Rows(0)("Inspetor")

                End If

            Catch ex As Exception

                conRE.Close()

                MsgBox("Erro 58 " & ex.Message)

            End Try

        End If

    End Sub
    'OK
    Dim Documento_xlsx As Microsoft.Office.Interop.Excel.Workbook
    Dim Excell As New Microsoft.Office.Interop.Excel.Application
    Sub LansarNoExcel()


        Dim Excell As New Microsoft.Office.Interop.Excel.Application

        Dim Planilha_do_Documento_xlsx As Microsoft.Office.Interop.Excel.Worksheet


        Dim ID As Microsoft.Office.Interop.Excel.Range

        Dim Cliente As Microsoft.Office.Interop.Excel.Range

        Dim Inspetor As Microsoft.Office.Interop.Excel.Range

        Dim Pedido As Microsoft.Office.Interop.Excel.Range

        Dim NotaFiscal As Microsoft.Office.Interop.Excel.Range

        Dim Produto As Microsoft.Office.Interop.Excel.Range

        Dim Codigo As Microsoft.Office.Interop.Excel.Range

        Dim Invoice As Microsoft.Office.Interop.Excel.Range

        Dim OP As Microsoft.Office.Interop.Excel.Range

        Dim Volume As Microsoft.Office.Interop.Excel.Range

        Dim Quantidade As Microsoft.Office.Interop.Excel.Range

        Dim Data As Microsoft.Office.Interop.Excel.Range

        Dim Hora As Microsoft.Office.Interop.Excel.Range

        Dim DataFab As Microsoft.Office.Interop.Excel.Range

        Dim Obs As Microsoft.Office.Interop.Excel.Range


        On Error GoTo ErrHandler

        '3º Abrir o arquivo Excel
        Documento_xlsx = Excell.Workbooks.Open("f:\Receb.Mat.Prima\Banco_Dados\ProjetoCertificado\Modelos_Certificados\" & lblCodigox.Text & ".xlsx", , ReadOnly:=True)

        '4º Abrir a planilha para inserir texto
        Planilha_do_Documento_xlsx = Documento_xlsx.Sheets.Item("PLAN1")

        '5º Atribuir uma célula na planilha

        Produto = Planilha_do_Documento_xlsx.Range("P1")

        Codigo = Planilha_do_Documento_xlsx.Range("P2")

        Cliente = Planilha_do_Documento_xlsx.Range("P3")

        Pedido = Planilha_do_Documento_xlsx.Range("P4")

        OP = Planilha_do_Documento_xlsx.Range("P5")

        Volume = Planilha_do_Documento_xlsx.Range("P6")

        Quantidade = Planilha_do_Documento_xlsx.Range("P7")

        NotaFiscal = Planilha_do_Documento_xlsx.Range("P8")

        Invoice = Planilha_do_Documento_xlsx.Range("P9")

        DataFab = Planilha_do_Documento_xlsx.Range("P10")

        ID = Planilha_do_Documento_xlsx.Range("P11")

        Data = Planilha_do_Documento_xlsx.Range("P12")

        Inspetor = Planilha_do_Documento_xlsx.Range("P13")

        Obs = Planilha_do_Documento_xlsx.Range("P14")

        Hora = Planilha_do_Documento_xlsx.Range("P15")


        'conectar com nº da rnc e transferir as rows/ cells para as variaveis abaixo

        Produto.Value = lblProdutox.Text

        Codigo.Value = lblCodigox.Text

        Cliente.Value = txtClientex.Text

        Pedido.Value = txtPedido.Text

        OP.Value = txtOPx.Text

        Volume.Value = txtVolumex.Text

        Quantidade.Value = txtQuantidadex.Text

        NotaFiscal.Value = txtNotaFiscal.Text

        Invoice.Value = txtInvoice.Text

        If dtpDex.Value = dtpAtex.Value Then

            DataFab.Value = dtpDex.Value

        Else

            DataFab.Value = dtpDex.Value & " até " & dtpAtex.Value

        End If

        If btCriar.Text = "Aplicar" Then

            If anex = 0 Then

                ID.Value = lblID.Text

            Else

                ID.Value = Integer.Parse(lblID.Text) + anex

            End If

        ElseIf btAlterar.Text = "Aplicar" Then

            ID.Value = novoiD

        ElseIf btImprimir.Text = "...Imprimir" Then

            ID.Value = novoiD

        ElseIf btAlterarIndividual.Text = "Aplicar" Then

            ID.Value = idIndividual

        ElseIf btImprimirIndividual.Text = "...Imprimir" Then

            ID.Value = idReal

        End If

        If btCriar.Text = "Aplicar" Then

            Data.Value = Today.ToShortDateString() & " - " & TimeOfDay.ToShortTimeString()

        Else

            Data.Value = lblData.Text & " - " & lblHora.Text

        End If


        Inspetor.Value = lblInspetor.Text

        Obs.Value = Obsx.ToString()

        If lblCodigox.Text = "3007000081" Or lblCodigox.Text = "3007000082" Or lblCodigox.Text = "3007000095" Then

            Planilha_do_Documento_xlsx = Documento_xlsx.Sheets.Item("PLAN2")

            Excell.Visible = False

            Planilha_do_Documento_xlsx.Range("A1:BB37").Select()

            Planilha_do_Documento_xlsx.Range("A1:BB37").Copy()

            Planilha_do_Documento_xlsx.Range("A1:BB37").PasteSpecial(Excel.XlPasteType.xlPasteValues)

            Imprimir()
            Documento_xlsx.Close(SaveChanges:=True, Filename:="f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & ID.Value & "-" & Codigo.Value & ".xlsx")



        Else

            Excell.Visible = False

            If btImprimirIndividual.Text = "...Imprimir" Or btImprimir.Text = "...Imprimir" Then

                Imprimir()

            ElseIf txtNotaFiscal.TextLength >= 4 And (btAlterar.Text = "Aplicar" Or btAlterarIndividual.Text = "Aplicar" Or btCriar.Text = "Aplicar") And (lblCodigox.Text <> "3007000081" Or lblCodigox.Text <> "3007000082" Or lblCodigox.Text <> "3007000095") Then

                SalvarPDF()

            End If

            If cbImprimir.Checked = True And (btAlterar.Text = "Aplicar" Or btAlterarIndividual.Text = "Aplicar" Or btCriar.Text = "Aplicar") Then

                Imprimir()

            End If

            Documento_xlsx.Close(SaveChanges:=False)

        End If

        '9º encerra os processos EXCEL.EXE no gerenciador de tarefas do windows 

ExitHere:
        If Excell Is Nothing Then
            Marshal.ReleaseComObject(Documento_xlsx)
            Excell = Nothing
            Exit Sub
        Else
            Excell.Quit()
            Marshal.ReleaseComObject(Documento_xlsx)
            Marshal.ReleaseComObject(Excell)
            Excell = Nothing
        End If

ErrHandler:
        ' MsgBox(Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source, vbCritical, "Erro xx6 ")
        Resume ExitHere

    End Sub
    'OK
    Dim itfXLBook As Object
    Sub SalvarPDF()

        Dim objUDC As UDC.IUDC 'interface do programa, como se forre aberto para o usuário

        Dim itfPrinter As UDC.IUDCPrinter 'interface da impressora

        Dim itfProfile As UDC.IProfile ' interface que configura o documento


        'Dim objXLApp As Object ' EXCEL
        ' 
        Dim itfXLWorksheet As Object

        Dim itfXLPageSetup As Object

        objUDC = New UDC.APIWrapper

        itfPrinter = objUDC.Printers("Universal Document Converter")

        itfProfile = itfPrinter.Profile


        ' Use Universal Document Converter API to change settings of converterd document
        itfProfile.PageSetup.FormName = "A4"

        itfProfile.PageSetup.ResolutionX = 50

        itfProfile.PageSetup.ResolutionY = 50

        itfProfile.PageSetup.Orientation = UDC.PageOrientationID.PO_PORTRAIT

        itfProfile.FileFormat.ActualFormat = UDC.FormatID.FMT_PDF

        itfProfile.FileFormat.PDF.Multipage = UDC.MultipageModeID.MM_SINGLE

        itfProfile.Adjustments.Crop.Mode = UDC.CropModeID.CRP_AUTO
        'salvando o documento

        itfProfile.OutputLocation.Mode = UDC.LocationModeID.LM_PREDEFINED

        itfProfile.OutputLocation.FolderPath = "f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\"

        If btCriar.Text = "Aplicar" Then

            If anex = 0 Then

                itfProfile.OutputLocation.FileName = "" & lblID.Text & "-" & lblCodigox.Text & ".pdf"

            Else

                itfProfile.OutputLocation.FileName = "" & Integer.Parse(lblID.Text) + anex & "-" & lblCodigox.Text & ".pdf"

            End If

        ElseIf btAlterar.Text = "Aplicar" Then

            itfProfile.OutputLocation.FileName = "" & novoiD & "-" & lblCodigox.Text & ".pdf"

        ElseIf btAlterarIndividual.Text = "Aplicar" Then

            itfProfile.OutputLocation.FileName = "" & idIndividual & "-" & lblCodigox.Text & ".pdf"

        ElseIf btImprimirIndividual.Text = "...Imprimir" Then

            itfProfile.OutputLocation.FileName = "" & idReal & "-" & lblCodigo1.Text & ".pdf"

        End If

        itfProfile.OutputLocation.OverwriteExistingFile = True

        ' Run Microsoft Excle as COM-server
        On Error Resume Next
        'objXLApp = CreateObject("Excel.Application")

        ' Open spreadsheet from file --- abrindo o documento
        'documento_xlsx é do procedimento Imprimir() que prepara os dados para os certificados que tomo emprestado
        'se não teria que salvar e depois excluir
        itfXLBook = Documento_xlsx ' < objXLApp.Workbooks.Open("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & lblID.Text & "-" & lblCodigo.Text & ".xlsx", , ReadOnly:=True)

        ' Change active worksheet settings and print it
        itfXLWorksheet = itfXLBook.ActiveSheet
        itfXLPageSetup = itfXLWorksheet.PageSetup

        itfXLPageSetup.Orientation = 1 ' Portrait

        Call itfXLWorksheet.PrintOut(1, 1, 1, False, "Universal Document Converter")


        ' Close the spreadsheet

        itfXLBook = Nothing
        ' Call itfXLBook.Close(False)
        ' Close Microsoft Excel
        'Call objXLApp.Quit()
        'objXLApp = Nothing
    End Sub
    'OK
    Dim imprimirx As String = ""

    Sub Imprimir()
        If cbEstoque.Checked = True Then
            imprimirx = "Sim"
        End If
        If imprimirx = "" Then
            If MsgBox("Onde Imprimir?" _
                      & Chr(13) _
                      & Chr(13) _
                      & "'Estoque' = Sim" _
                      & Chr(13) _
                      & "'Qualidade' = Não", vbYesNo, "Local da Impressão") = vbYes Then
                imprimirx = "Sim"
                Documento_xlsx.PrintOutEx(From:=1, To:=1, Copies:=1, Preview:=False, ActivePrinter:="Lexmark T644 Estoque") 'imprime na impressora fisica
            Else
                imprimirx = "Não"
                Documento_xlsx.PrintOutEx(From:=1, To:=1, Copies:=1, Preview:=False, ActivePrinter:="Lexmark T644 ADM Prod") 'imprime na impressora fisica
            End If
        ElseIf imprimirx = "Sim" Then
            Documento_xlsx.PrintOutEx(From:=1, To:=1, Copies:=1, Preview:=False, ActivePrinter:="Lexmark T644 Estoque") 'imprime na impressora fisica
        ElseIf imprimirx = "Não" Then
            Documento_xlsx.PrintOutEx(From:=1, To:=1, Copies:=1, Preview:=False, ActivePrinter:="Lexmark T644 ADM Prod") 'imprime na impressora fisica
        End If

    End Sub
    'OK
    Sub Email()
        System.Threading.Thread.Sleep(5000)
        Dim AppOutlook As New Microsoft.Office.Interop.Outlook.Application
        Dim OutlookMessage As Outlook.MailItem
        OutlookMessage = AppOutlook.CreateItem(Outlook.OlItemType.olMailItem)
        'Dim Recipent As Outlook.Recipients = OutlookMessage.Recipients
        Try
            OutlookMessage.Subject = "Certificado - " & txtClientex.Text

            Dim saudacao As String
            If TimeOfDay.Hour < 12 Then
                saudacao = "Bom dia,"
            ElseIf TimeOfDay.Hour < 18 Then
                saudacao = "Boa tarde,"
            Else
                saudacao = "Boa noite,"
            End If
            OutlookMessage.Body = saudacao & "" _
                & Chr(13) _
                    & "Segue o(s) Certificado(s) em Anexo"

            If btCriar.Text = "Aplicar" Then
                VerificarRB()
                For i = 1 To verRB Step 1
                    If lblCodigo1.Text = "3007000081" Or lblCodigo1.Text = "3007000082" Or lblCodigox.Text = "3007000095" Then
                        OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & lblID.Text & "-" & lblCodigox.Text & ".xlsx")
                    Else
                        OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & lblID.Text & "-" & lblCodigox.Text & ".pdf")
                    End If
                    lblID.Text = Integer.Parse(lblID.Text) + 1
                Next
            ElseIf btAlterar.Text = "Aplicar" Then
                VerificarRB()
                For i = 1 To verRB Step 1
                    If lblCodigo1.Text = "3007000081" Or lblCodigo1.Text = "3007000082" Or lblCodigox.Text = "3007000095" Then
                        OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & novoiD & "-" & lblCodigox.Text & ".xlsx")
                    Else
                        OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & novoiD & "-" & lblCodigox.Text & ".pdf")
                    End If
                    novoiD -= 1
                    If i = 1 Then
                        Incluir_Email()
                    End If
                Next
            ElseIf btEmail.Text = "...Email" Then
                VerificarRB()
                Dim x As Integer = 0
                x = Integer.Parse(lblID.Text)
                For i = 1 To verRB Step 1
                    If lblCodigo1.Text = "3007000081" Or lblCodigo1.Text = "3007000082" Or lblCodigox.Text = "3007000095" Then
                        OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & x & "-" & lblCodigo1.Text & ".xlsx")
                    Else
                        OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & x & "-" & lblCodigo1.Text & ".pdf")
                    End If
                    x -= 1
                    If i = 1 Then
                        Incluir_Email()
                    End If
                Next
            ElseIf btAlterarIndividual.Text = "Aplicar" Then
                If lblCodigo1.Text = "3007000081" Or lblCodigo1.Text = "3007000082" Or lblCodigox.Text = "3007000095" Then
                    OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & idIndividual & "-" & lblCodigo1.Text & ".xlsx")
                Else
                    OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & idIndividual & "-" & lblCodigo1.Text & ".pdf")
                End If
                Consulta2()
                incluirEmail = Convert.ToString(ds21.Tables("tblVolume").Rows(0).Item("Email"))
            ElseIf btEmailIndividual.Text = "...Email" Then
                If lblCodigo1.Text = "3007000081" Or lblCodigo1.Text = "3007000082" Or lblCodigox.Text = "3007000095" Then
                    OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & idReal & "-" & lblCodigo1.Text & ".xlsx")
                Else
                    OutlookMessage.Attachments.Add("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & idReal & "-" & lblCodigo1.Text & ".pdf")
                End If
                Consulta2()
                incluirEmail = Convert.ToString(ds21.Tables("tblVolume").Rows(0).Item("Email"))
            End If



            If incluirEmail <> "" Or incluirEmail <> Nothing Then
                OutlookMessage.To = incluirEmail  ' Criar um grupo no outlook chamado 
                'Recipent.Add("???")

                'System.Threading.Thread.Sleep(5000)

                OutlookMessage.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
                ''
                'If (MsgBox("O E-mail está pronto para ser enviado. Deseja Enviar?" _
                '         & Chr(13) _
                '        & Chr(13) _
                '       & "'Sim' = Enviar" _
                '      & Chr(13) _
                '     & "'Não' = Alterar", vbYesNo, "Email") = vbYes) Then
                OutlookMessage.Save()
                OutlookMessage.Send()
                'MsgBox("Documento(s) enviado(s) com sucesso", , "Envio")
                'Else
                '   OutlookMessage.Display()
                '  OutlookMessage.Save()
                'End If
                ''
            Else
                MsgBox("O email não pode ser enviado. Adicione endereço(s)", , "Endereços de E-mails)")
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

    Sub Incluir_Email()
        Try
            Dim da12 As New OleDbDataAdapter
            Dim dt12 As New System.Data.DataTable
            Dim ds12 As New DataSet
            conPecasVolume.Open()
            Dim sel5 As String = "SELECT top 5 Codigo, Email FROM tblVolume where Codigo = '" & lblCodigo1.Text & "'"
            da12 = New OleDbDataAdapter(sel5, conPecasVolume)
            dt12.Clear()
            da12.Fill(dt12)
            da12.Fill(ds12, "tblVolume")
            If dt12.Rows.Count > 1 Then
                Me.DataGridView2.DataSource = ds12
                Me.DataGridView2.DataMember = "tblVolume"
            ElseIf dt12.Rows.Count = 1 Then
                incluirEmail = Nothing
                incluirEmail = ds12.Tables(name:="tblVolume").Rows(index:=0).Item(columnName:="Email").ToString()
            End If
            conPecasVolume.Close()
        Catch ex As Exception
            conPecasVolume.Close()
            MsgBox("erro wqe443 " & ex.Message)
        End Try
    End Sub
    'OK
    Private Sub btEmail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEmail.Click

        If btEmail.Text = "Email" Then
            'If MsgBox("Deseja enviar um Certificado?", vbYesNo, "Enviar Certificado") = vbYes Then
            MsgBox("Selecione um Certificado abaixo a enviar", , "Selecionar - Certificado")
            btEmail.Text = "...Email"
            GroupBox2.Enabled = False
            btCriar.Enabled = False
            btExcluir.Enabled = False
            btImprimir.Enabled = False
            btImprimirIndividual.Enabled = False
            btEmailIndividual.Enabled = False
            btAlterarIndividual.Enabled = False
            btAlterar.Enabled = False
            'Else
            'End If
        Else
        If rbSim.Checked = True Then
            If MsgBox("O Certificado a ser enviado é do Pedido: " & txtPedido.Text & " ?", vbYesNo, "Certificado Selecionado") = vbYes Then
                Email()
                LimparTudo()
            End If
        ElseIf rbNao.Checked = True Then
            MsgBox("O Certificado não pode ser enviado pois não tem nota fiscal", , "Certificado sem Nota Fiscal")
        End If
        End If

    End Sub
    'OK
    Private Sub btImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btImprimir.Click
        'se tentar imprimir e salvar em pdf usando o datagridview?

        If btImprimir.Text = "Imprimir" Then
            cbEnviarEmail.Checked = False
            cbImprimir.Checked = False
            LimparTudo()
            btImprimir.Text = "...Imprimir"
            GroupBox2.Enabled = False
            MsgBox("Selecione um item para imprimir", , "Selecione - Imprimir")
            GroupBox4.Enabled = False
            btCriar.Enabled = False
            btAlterar.Enabled = False
            btEmail.Enabled = False

        Else
            If rbSim.Checked = True Then
                VerificarRB()
                Try
                    Dim i As Byte = 0
                    For i = 1 To terRB Step 1
                        Controles()
                        If i = 1 Then
                            novoiD = novoiD - (terRB - i)
                        Else
                            novoiD += 1
                        End If

                        LansarNoExcel() ' lança no excell, salva em PDF e Imprime
                        ' adicionando quantos anexos deve-se enviar
                        idx = anex
                        val = Integer.Parse(lblID.Text)
                        arrei_anexos.SetValue(val, idx)
                        anex += 1

                    Next
                    'MsgBox("Impressão realizada com sucesso!")
                Catch ex As Exception
                    MsgBox("Erro 15 " & ex.Message)
                End Try
                btImprimir.Text = "Imprimir"
                LimparTudo()
            ElseIf rbNao.Checked = True Then
                MsgBox("O Certificado não pode ser impresso pois não tem nota fiscal", , "Certificado sem Nota Fiscal")
            End If
        End If
    End Sub
    'OK
    Dim dt19 As New System.Data.DataTable
    Sub Carregar()
        Try
            Dim da19 As New OleDbDataAdapter
            Dim ds19 As New DataSet
            conCertificado.Open()
            Dim sel_ As String = "SELECT TOP 10 * FROM tblCertificado where Pedido = '" & txtPedido.Text & "' and NotaFiscal = '" & txtNotaFiscal.Text & "' ORDER BY ID DESC "
            da19 = New OleDbDataAdapter(sel_, conCertificado)
            dt19.Clear()
            da19.Fill(dt19)
            conCertificado.Close()
            Limpar()
            QuantasLinhas()
            PreencherLinhas()
        Catch ex As Exception
            MsgBox("Erro 71 " & ex.Message)
            conCertificado.Close()
        Finally
            conCertificado.Close()
        End Try
    End Sub
    'OK
    Sub QuantasLinhas()
        Select Case dt19.Rows.Count
            Case 1
                rb1T.Checked = True
            Case 2
                rb2T.Checked = True
            Case 3
                rb3T.Checked = True
            Case 4
                rb4T.Checked = True
            Case 5
                rb5T.Checked = True
            Case 6
                rb6T.Checked = True
            Case 7
                rb7T.Checked = True
            Case 8
                rb8T.Checked = True
            Case 9
                rb9T.Checked = True
            Case 10
                rb10T.Checked = True
        End Select
    End Sub
    'OK
    Sub PreencherLinhas()
        Try
            txtOP1.Text = dt19.Rows(0)("OP")
            txtVolume1.Text = dt19.Rows(0)("Volume")
            txtQuantidade1.Text = dt19.Rows(0)("Quantidade")
            dtpDe1.Value = dt19.Rows(0)("DataFab_Inicio")
            dtpAte1.Value = dt19.Rows(0)("DataFab_Fim")
            If dt19.Rows(0)("Obs") = "" Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If

            If dt19.Rows.Count >= 2 Then
                txtOP2.Text = dt19.Rows(1)("OP")
                txtVolume2.Text = dt19.Rows(1)("Volume")
                txtQuantidade2.Text = dt19.Rows(1)("Quantidade")
                dtpDe2.Value = dt19.Rows(1)("DataFab_Inicio")
                dtpAte2.Value = dt19.Rows(1)("DataFab_Fim")
                If dt19.Rows(1)("Obs") = "" Then
                    CheckBox2.Checked = False
                Else
                    CheckBox2.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 3 Then
                txtOP3.Text = dt19.Rows(2)("OP")
                txtVolume3.Text = dt19.Rows(2)("Volume")
                txtQuantidade3.Text = dt19.Rows(2)("Quantidade")
                dtpDe3.Value = dt19.Rows(2)("DataFab_Inicio")
                dtpAte3.Value = dt19.Rows(2)("DataFab_Fim")
                If dt19.Rows(2)("Obs") = "" Then
                    CheckBox3.Checked = False
                Else
                    CheckBox3.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 4 Then
                txtOP4.Text = dt19.Rows(3)("OP")
                txtVolume4.Text = dt19.Rows(3)("Volume")
                txtQuantidade4.Text = dt19.Rows(3)("Quantidade")
                dtpDe4.Value = dt19.Rows(3)("DataFab_Inicio")
                dtpAte4.Value = dt19.Rows(3)("DataFab_Fim")
                If dt19.Rows(3)("Obs") = "" Then
                    CheckBox4.Checked = False
                Else
                    CheckBox4.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 5 Then
                txtOP5.Text = dt19.Rows(4)("OP")
                txtVolume5.Text = dt19.Rows(4)("Volume")
                txtQuantidade5.Text = dt19.Rows(4)("Quantidade")
                dtpDe5.Value = dt19.Rows(4)("DataFab_Inicio")
                dtpAte5.Value = dt19.Rows(4)("DataFab_Fim")
                If dt19.Rows(4)("Obs") = "" Then
                    CheckBox5.Checked = False
                Else
                    CheckBox5.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 6 Then
                txtOP6.Text = dt19.Rows(5)("OP")
                txtVolume6.Text = dt19.Rows(5)("Volume")
                txtQuantidade6.Text = dt19.Rows(5)("Quantidade")
                dtpDe6.Value = dt19.Rows(5)("DataFab_Inicio")
                dtpAte6.Value = dt19.Rows(5)("DataFab_Fim")
                If dt19.Rows(5)("Obs") = "" Then
                    CheckBox6.Checked = False
                Else
                    CheckBox6.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 7 Then
                txtOP7.Text = dt19.Rows(6)("OP")
                txtVolume7.Text = dt19.Rows(6)("Volume")
                txtQuantidade7.Text = dt19.Rows(6)("Quantidade")
                dtpDe7.Value = dt19.Rows(6)("DataFab_Inicio")
                dtpAte7.Value = dt19.Rows(6)("DataFab_Fim")
                If dt19.Rows(6)("Obs") = "" Then
                    CheckBox7.Checked = False
                Else
                    CheckBox7.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 8 Then
                txtOP8.Text = dt19.Rows(7)("OP")
                txtVolume8.Text = dt19.Rows(7)("Volume")
                txtQuantidade8.Text = dt19.Rows(7)("Quantidade")
                dtpDe8.Value = dt19.Rows(7)("DataFab_Inicio")
                dtpAte8.Value = dt19.Rows(7)("DataFab_Fim")
                If dt19.Rows(7)("Obs") = "" Then
                    CheckBox8.Checked = False
                Else
                    CheckBox8.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 9 Then
                txtOP9.Text = dt19.Rows(8)("OP")
                txtVolume9.Text = dt19.Rows(8)("Volume")
                txtQuantidade9.Text = dt19.Rows(8)("Quantidade")
                dtpDe9.Value = dt19.Rows(8)("DataFab_Inicio")
                dtpAte9.Value = dt19.Rows(8)("DataFab_Fim")
                If dt19.Rows(8)("Obs") = "" Then
                    CheckBox9.Checked = False
                Else
                    CheckBox9.Checked = True
                End If
            End If

            If dt19.Rows.Count >= 10 Then
                txtOP10.Text = dt19.Rows(9)("OP")
                txtVolume10.Text = dt19.Rows(9)("Volume")
                txtQuantidade10.Text = dt19.Rows(9)("Quantidade")
                dtpDe10.Value = dt19.Rows(9)("DataFab_Inicio")
                dtpAte10.Value = dt19.Rows(9)("DataFab_Fim")
                If dt19.Rows(9)("Obs") = "" Then
                    CheckBox10.Checked = False
                Else
                    CheckBox10.Checked = True
                End If
            End If

        Catch ex As Exception
            MsgBox("Erro 768 " & ex.Message)
        End Try
    End Sub
    'OK
    Sub Limpar()
        txtOP1.Clear()
        lblProduto1.Text = ""
        lblCodigo1.Text = ""
        txtCliente1.Clear()
        txtVolume1.Text = "0"
        txtQuantidade1.Text = "0"
        txtPecasPorVolume1.Text = "0"
        dtpDe1.Value = Today
        dtpAte1.Value = Today

        txtOP2.Clear()
        lblProduto2.Text = ""
        lblCodigo2.Text = ""
        txtCliente2.Clear()
        txtVolume2.Text = "0"
        txtQuantidade2.Text = "0"
        txtPecasPorVolume2.Text = "0"
        dtpDe2.Value = Today
        dtpAte2.Value = Today

        txtOP3.Clear()
        lblProduto3.Text = ""
        lblCodigo3.Text = ""
        txtCliente3.Clear()
        txtVolume3.Text = "0"
        txtQuantidade3.Text = "0"
        txtPecasPorVolume3.Text = "0"
        dtpDe3.Value = Today
        dtpAte3.Value = Today

        txtOP4.Clear()
        lblProduto4.Text = ""
        lblCodigo4.Text = ""
        txtCliente4.Clear()
        txtVolume4.Text = "0"
        txtQuantidade4.Text = "0"
        txtPecasPorVolume4.Text = "0"
        dtpDe4.Value = Today
        dtpAte4.Value = Today

        txtOP5.Clear()
        lblProduto5.Text = ""
        lblCodigo5.Text = ""
        txtCliente5.Clear()
        txtVolume5.Text = "0"
        txtQuantidade5.Text = "0"
        txtPecasPorVolume5.Text = "0"
        dtpDe5.Value = Today
        dtpAte5.Value = Today

        txtOP6.Clear()
        lblProduto6.Text = ""
        lblCodigo6.Text = ""
        txtCliente6.Clear()
        txtVolume6.Text = "0"
        txtQuantidade6.Text = "0"
        txtPecasPorVolume6.Text = "0"
        dtpDe6.Value = Today
        dtpAte6.Value = Today

        txtOP7.Clear()
        lblProduto7.Text = ""
        lblCodigo7.Text = ""
        txtCliente7.Clear()
        txtVolume7.Text = "0"
        txtQuantidade7.Text = "0"
        txtPecasPorVolume7.Text = "0"
        dtpDe7.Value = Today
        dtpAte7.Value = Today

        txtOP8.Clear()
        lblProduto8.Text = ""
        lblCodigo8.Text = ""
        txtCliente8.Clear()
        txtVolume8.Text = "0"
        txtQuantidade8.Text = "0"
        txtPecasPorVolume8.Text = "0"
        dtpDe8.Value = Today
        dtpAte8.Value = Today

        txtOP9.Clear()
        lblProduto9.Text = ""
        lblCodigo9.Text = ""
        txtCliente9.Clear()
        txtVolume9.Text = "0"
        txtQuantidade9.Text = "0"
        txtPecasPorVolume9.Text = "0"
        dtpDe9.Value = Today
        dtpAte9.Value = Today

        txtOP10.Clear()
        lblProduto10.Text = ""
        lblCodigo10.Text = ""
        txtCliente10.Clear()
        txtVolume10.Text = "0"
        txtQuantidade10.Text = "0"
        txtPecasPorVolume10.Text = "0"
        dtpDe10.Value = Today
        dtpAte10.Value = Today
    End Sub
    'OK
    Dim novoiD As Integer
    Dim idReal As Integer
    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try

            Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow

            'estáticos e iguais
            Dim ID = row.Cells(0)
            idReal = ID.Value

            Dim Pedido = row.Cells(1)
            Dim NotaFiscal = row.Cells(2)
            If Pedido.Value = txtPedido.Text And NotaFiscal.Value = txtNotaFiscal.Text Then
            Else
                Me.txtPedido.Text = Convert.ToString(Pedido.Value)
                Me.txtNotaFiscal.Text = Convert.ToString(NotaFiscal.Value)
                Carregar()

                novoiD = dt19.Rows(0)("ID")
                Me.lblID.Text = Convert.ToString(novoiD)

                Dim Produto = row.Cells(3)
                Dim Codigo = row.Cells(4)
                Dim Invoice = row.Cells(5)
                Dim Data = row.Cells(9)
                Dim Hora = row.Cells(10)
                Dim DataAlterado = row.Cells(14)
                Dim HoraAlterado = row.Cells(15)
                Dim Cliente = row.Cells(16)
                Dim Inspetor = row.Cells(17)

                'dinamicos e diferentes
                'Dim OP = row.Cells(6)
                'Dim Volume = row.Cells(7)
                'Dim Quantidade = row.Cells(8)
                'Dim DataFavI = row.Cells(11)
                'Dim DataFabF = row.Cells(12)
                'Dim Obs = row.Cells(13)ghhhhhhhhhhhhhhhhhhhhhhhhh



                If txtNotaFiscal.TextLength > 0 Then
                    rbSim.Checked = True
                    Me.txtNotaFiscal.Text = Convert.ToString(NotaFiscal.Value)
                ElseIf txtNotaFiscal.TextLength = 0 Then
                    rbNao.Checked = True
                End If
                Select Case dt19.Rows.Count
                    Case 1
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                    Case 2
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)


                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                    Case 3
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                    Case 4
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                    Case 5
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto5.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo5.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
                    Case 6
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto5.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto6.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo5.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo6.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente6.Text = Convert.ToString(Cliente.Value)
                    Case 7
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto5.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto6.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto7.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo5.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo6.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo7.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente6.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente7.Text = Convert.ToString(Cliente.Value)
                    Case 8
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto5.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto6.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto7.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto8.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo5.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo6.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo7.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo8.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente6.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente7.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente8.Text = Convert.ToString(Cliente.Value)
                    Case 9
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto5.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto6.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto7.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto8.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto9.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo5.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo6.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo7.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo8.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo9.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente6.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente7.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente8.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente9.Text = Convert.ToString(Cliente.Value)
                    Case 10
                        Me.lblProduto1.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto2.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto3.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto4.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto5.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto6.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto7.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto8.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto9.Text = Convert.ToString(Produto.Value)
                        Me.lblProduto10.Text = Convert.ToString(Produto.Value)

                        Me.lblCodigo1.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo2.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo3.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo4.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo5.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo6.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo7.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo8.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo9.Text = Convert.ToString(Codigo.Value)
                        Me.lblCodigo10.Text = Convert.ToString(Codigo.Value)

                        Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente6.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente7.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente8.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente9.Text = Convert.ToString(Cliente.Value)
                        Me.txtCliente10.Text = Convert.ToString(Cliente.Value)

                End Select

                Me.txtInvoice.Text = Convert.ToString(Invoice.Value)
                'Me.txtOP.Text = Convert.ToString(OP.Value)
                'Me.txtVolume.Text = Convert.ToString(Volume.Value)
                'Me.txtQuantidade.Text = Convert.ToString(Quantidade.Value)
                Me.lblData.Text = Convert.ToString(Data.Value)
                Me.lblHora.Text = Convert.ToString(Hora.Value)
                'Me.dtpDe.Text = Convert.ToString(DataFavI.Value)
                'Me.dtpAte.Text = Convert.ToString(DataFabF.Value)
                'Me.txtObs.Text = Convert.ToString(Obs.Value)
                Me.lblDataAlterado.Text = Convert.ToString(DataAlterado.Value)
                Me.lblHoraAlterado.Text = Convert.ToString(HoraAlterado.Value)

                Me.lblInspetor.Text = Convert.ToString(Inspetor.Value)
                Me.txtRE.Text = Convert.ToString(Inspetor.Value).Remove(4).TrimEnd()
                total()
            End If
        Catch ex As Exception
            MsgBox("Erro GT70 " & ex.Message)
        End Try
    End Sub
    'OK
    Private Sub rbT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rb1T.CheckedChanged, rb2T.CheckedChanged, rb3T.CheckedChanged, rb4T.CheckedChanged, rb5T.CheckedChanged, rb6T.CheckedChanged, rb7T.CheckedChanged, rb8T.CheckedChanged, rb9T.CheckedChanged, rb10T.CheckedChanged
        abrir()
    End Sub
    'OK
    Sub Abrir()
        VerificarRB()
        Select Case terRB
            Case 1
                Desabilita2()
                Desabilita3()
                Desabilita4()
                Desabilita5()
                Desabilita6()
                Desabilita7()
                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 2
                Abilita2()

                Desabilita3()
                Desabilita4()
                Desabilita5()
                Desabilita6()
                Desabilita7()
                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 3
                Abilita2()
                Abilita3()

                Desabilita4()
                Desabilita5()
                Desabilita6()
                Desabilita7()
                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 4
                Abilita2()
                Abilita3()
                Abilita4()

                Desabilita5()
                Desabilita6()
                Desabilita7()
                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 5
                Abilita2()
                Abilita3()
                Abilita4()
                Abilita5()

                Desabilita6()
                Desabilita7()
                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 6
                Abilita2()
                Abilita3()
                Abilita4()
                Abilita5()
                Abilita6()

                Desabilita7()
                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 7
                Abilita2()
                Abilita3()
                Abilita4()
                Abilita5()
                Abilita6()
                Abilita7()

                Desabilita8()
                Desabilita9()
                Desabilita10()
            Case 8
                Abilita2()
                Abilita3()
                Abilita4()
                Abilita5()
                Abilita6()
                Abilita7()
                Abilita8()

                Desabilita9()
                Desabilita10()
            Case 9
                Abilita2()
                Abilita3()
                Abilita4()
                Abilita5()
                Abilita6()
                Abilita7()
                Abilita8()
                Abilita9()
                Desabilita10()
            Case 10
                Abilita2()
                Abilita3()
                Abilita4()
                Abilita5()
                Abilita6()
                Abilita7()
                Abilita8()
                Abilita9()
                Abilita10()
        End Select
        total()
    End Sub
    'OK
    Sub Desabilita2()
        lbl2.Visible = False
        txtOP2.Visible = False
        lblProduto2.Visible = False
        lblCodigo2.Visible = False
        txtCliente2.Visible = False
        txtVolume2.Visible = False
        txtQuantidade2.Visible = False
        txtPecasPorVolume2.Visible = False
        dtpDe2.Visible = False
        dtpAte2.Visible = False
        CheckBox2.Visible = False
    End Sub
    Sub Abilita2()
        lbl2.Visible = True
        txtOP2.Visible = True
        lblProduto2.Visible = True
        lblCodigo2.Visible = True
        txtCliente2.Visible = True
        txtVolume2.Visible = True
        txtQuantidade2.Visible = True
        txtPecasPorVolume2.Visible = True
        dtpDe2.Visible = True
        dtpAte2.Visible = True
        CheckBox2.Visible = True
    End Sub
    Sub Desabilita3()
        lbl3.Visible = False
        txtOP3.Visible = False
        lblProduto3.Visible = False
        lblCodigo3.Visible = False
        txtCliente3.Visible = False
        txtVolume3.Visible = False
        txtQuantidade3.Visible = False
        txtPecasPorVolume3.Visible = False
        dtpDe3.Visible = False
        dtpAte3.Visible = False
        CheckBox3.Visible = False
    End Sub
    Sub Abilita3()
        lbl3.Visible = True
        txtOP3.Visible = True
        lblProduto3.Visible = True
        lblCodigo3.Visible = True
        txtCliente3.Visible = True
        txtVolume3.Visible = True
        txtQuantidade3.Visible = True
        txtPecasPorVolume3.Visible = True
        dtpDe3.Visible = True
        dtpAte3.Visible = True
        CheckBox3.Visible = True
    End Sub
    Sub Desabilita4()
        lbl4.Visible = False
        txtOP4.Visible = False
        lblProduto4.Visible = False
        lblCodigo4.Visible = False
        txtCliente4.Visible = False
        txtVolume4.Visible = False
        txtQuantidade4.Visible = False
        txtPecasPorVolume4.Visible = False
        dtpDe4.Visible = False
        dtpAte4.Visible = False
        CheckBox4.Visible = False
    End Sub
    Sub Abilita4()
        lbl4.Visible = True
        txtOP4.Visible = True
        lblProduto4.Visible = True
        lblCodigo4.Visible = True
        txtCliente4.Visible = True
        txtVolume4.Visible = True
        txtQuantidade4.Visible = True
        txtPecasPorVolume4.Visible = True
        dtpDe4.Visible = True
        dtpAte4.Visible = True
        CheckBox4.Visible = True
    End Sub
    Sub Desabilita5()
        lbl5.Visible = False
        txtOP5.Visible = False
        lblProduto5.Visible = False
        lblCodigo5.Visible = False
        txtCliente5.Visible = False
        txtVolume5.Visible = False
        txtQuantidade5.Visible = False
        txtPecasPorVolume5.Visible = False
        dtpDe5.Visible = False
        dtpAte5.Visible = False
        CheckBox5.Visible = False
    End Sub
    Sub Abilita5()
        lbl5.Visible = True
        txtOP5.Visible = True
        lblProduto5.Visible = True
        lblCodigo5.Visible = True
        txtCliente5.Visible = True
        txtVolume5.Visible = True
        txtQuantidade5.Visible = True
        txtPecasPorVolume5.Visible = True
        dtpDe5.Visible = True
        dtpAte5.Visible = True
        CheckBox5.Visible = True
    End Sub
    Sub Desabilita6()
        lbl6.Visible = False
        txtOP6.Visible = False
        lblProduto6.Visible = False
        lblCodigo6.Visible = False
        txtCliente6.Visible = False
        txtVolume6.Visible = False
        txtQuantidade6.Visible = False
        txtPecasPorVolume6.Visible = False
        dtpDe6.Visible = False
        dtpAte6.Visible = False
        CheckBox6.Visible = False
    End Sub
    Sub Abilita6()
        lbl6.Visible = True
        txtOP6.Visible = True
        lblProduto6.Visible = True
        lblCodigo6.Visible = True
        txtCliente6.Visible = True
        txtVolume6.Visible = True
        txtQuantidade6.Visible = True
        txtPecasPorVolume6.Visible = True
        dtpDe6.Visible = True
        dtpAte6.Visible = True
        CheckBox6.Visible = True
    End Sub
    Sub Desabilita7()
        lbl7.Visible = False
        txtOP7.Visible = False
        lblProduto7.Visible = False
        lblCodigo7.Visible = False
        txtCliente7.Visible = False
        txtVolume7.Visible = False
        txtQuantidade7.Visible = False
        txtPecasPorVolume7.Visible = False
        dtpDe7.Visible = False
        dtpAte7.Visible = False
        CheckBox7.Visible = False
    End Sub
    Sub Abilita7()
        lbl7.Visible = True
        txtOP7.Visible = True
        lblProduto7.Visible = True
        lblCodigo7.Visible = True
        txtCliente7.Visible = True
        txtVolume7.Visible = True
        txtQuantidade7.Visible = True
        txtPecasPorVolume7.Visible = True
        dtpDe7.Visible = True
        dtpAte7.Visible = True
        CheckBox7.Visible = True
    End Sub
    Sub Desabilita8()
        lbl8.Visible = False
        txtOP8.Visible = False
        lblProduto8.Visible = False
        lblCodigo8.Visible = False
        txtCliente8.Visible = False
        txtVolume8.Visible = False
        txtQuantidade8.Visible = False
        txtPecasPorVolume8.Visible = False
        dtpDe8.Visible = False
        dtpAte8.Visible = False
        CheckBox8.Visible = False
    End Sub
    Sub Abilita8()
        lbl8.Visible = True
        txtOP8.Visible = True
        lblProduto8.Visible = True
        lblCodigo8.Visible = True
        txtCliente8.Visible = True
        txtVolume8.Visible = True
        txtQuantidade8.Visible = True
        txtPecasPorVolume8.Visible = True
        dtpDe8.Visible = True
        dtpAte8.Visible = True
        CheckBox8.Visible = True
    End Sub
    Sub Desabilita9()
        lbl9.Visible = False
        txtOP9.Visible = False
        lblProduto9.Visible = False
        lblCodigo9.Visible = False
        txtCliente9.Visible = False
        txtVolume9.Visible = False
        txtQuantidade9.Visible = False
        txtPecasPorVolume9.Visible = False
        dtpDe9.Visible = False
        dtpAte9.Visible = False
        CheckBox9.Visible = False
    End Sub
    Sub Abilita9()
        lbl9.Visible = True
        txtOP9.Visible = True
        lblProduto9.Visible = True
        lblCodigo9.Visible = True
        txtCliente9.Visible = True
        txtVolume9.Visible = True
        txtQuantidade9.Visible = True
        txtPecasPorVolume9.Visible = True
        dtpDe9.Visible = True
        dtpAte9.Visible = True
        CheckBox9.Visible = True
    End Sub
    Sub Desabilita10()
        lbl10.Visible = False
        txtOP10.Visible = False
        lblProduto10.Visible = False
        lblCodigo10.Visible = False
        txtCliente10.Visible = False
        txtVolume10.Visible = False
        txtQuantidade10.Visible = False
        txtPecasPorVolume10.Visible = False
        dtpDe10.Visible = False
        dtpAte10.Visible = False
        CheckBox10.Visible = False
    End Sub
    Sub Abilita10()
        lbl10.Visible = True
        txtOP10.Visible = True
        lblProduto10.Visible = True
        lblCodigo10.Visible = True
        txtCliente10.Visible = True
        txtVolume10.Visible = True
        txtQuantidade10.Visible = True
        txtPecasPorVolume10.Visible = True
        dtpDe10.Visible = True
        dtpAte10.Visible = True
        CheckBox10.Visible = True
    End Sub
    'OK
    Private Sub btExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExcluir.Click
        Try
            If btExcluir.Text = "Excluir" Then
                'If MsgBox("Deseja Excluir um Certificado?", vbYesNo, "Excluir RNC") = vbYes Then
                LimparTudo()
                btExcluir.Text = "Aplicar"
                btCriar.Enabled = False
                btAlterar.Enabled = False
                btImprimir.Enabled = False
                txtPedido.Focus()
                'Else
                'End If
            Else
            conCertificado.Open()
            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet
            ds20 = New DataSet
            da20 = New OleDbDataAdapter("Delete from tblCertificado WHERE Id = " & lblID.Text & "", conCertificado)
            ds20.Clear()
            da20.Fill(ds20, "tblCertificado")
            conCertificado.Close()
            AtualizarGrid()
            LimparTudo()
                'MsgBox("Exclusão realizada com sucesso!")
            'Try
            'Kill("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & lblID.Text & "-" & lblCodigo1.Text & ".pdf") ' deleta o arquivo da pasta
            'Catch ex As Exception
            'End Try
            End If
        Catch ex As Exception
            MsgBox("Erro 84 " & ex.Message)
        End Try
    End Sub
    'OK
    Private Sub txtVolume1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume1.TextChanged, txtPecasPorVolume1.TextChanged
        If txtVolume1.TextLength = 0 Then
            txtVolume1.Text = 0
        End If
        If txtPecasPorVolume1.TextLength = 0 Then
            txtPecasPorVolume1.Text = 0
        End If
        txtQuantidade1.Text = txtVolume1.Text * txtPecasPorVolume1.Text
        total()
    End Sub
    Private Sub txtVolume2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume2.TextChanged, txtPecasPorVolume2.TextChanged
        If txtVolume2.TextLength = 0 Then
            txtVolume2.Text = 0
        End If
        If txtPecasPorVolume2.TextLength = 0 Then
            txtPecasPorVolume2.Text = 0
        End If
        txtQuantidade2.Text = txtVolume2.Text * txtPecasPorVolume2.Text
        total()
    End Sub
    Private Sub txtVolume3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume3.TextChanged, txtPecasPorVolume3.TextChanged
        If txtVolume3.TextLength = 0 Then
            txtVolume3.Text = 0
        End If
        If txtPecasPorVolume3.TextLength = 0 Then
            txtPecasPorVolume3.Text = 0
        End If
        txtQuantidade3.Text = txtVolume3.Text * txtPecasPorVolume3.Text
        total()
    End Sub
    Private Sub txtVolume4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume4.TextChanged, txtPecasPorVolume4.TextChanged
        If txtVolume4.TextLength = 0 Then
            txtVolume4.Text = 0
        End If
        If txtPecasPorVolume4.TextLength = 0 Then
            txtPecasPorVolume4.Text = 0
        End If
        txtQuantidade4.Text = txtVolume4.Text * txtPecasPorVolume4.Text
        total()
    End Sub
    Private Sub txtVolume5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume5.TextChanged, txtPecasPorVolume5.TextChanged
        If txtVolume5.TextLength = 0 Then
            txtVolume5.Text = 0
        End If
        If txtPecasPorVolume5.TextLength = 0 Then
            txtPecasPorVolume5.Text = 0
        End If
        txtQuantidade5.Text = txtVolume5.Text * txtPecasPorVolume5.Text
        total()
    End Sub
    Private Sub txtVolume6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume6.TextChanged, txtPecasPorVolume6.TextChanged
        If txtVolume6.TextLength = 0 Then
            txtVolume6.Text = 0
        End If
        If txtPecasPorVolume6.TextLength = 0 Then
            txtPecasPorVolume6.Text = 0
        End If
        txtQuantidade6.Text = txtVolume6.Text * txtPecasPorVolume6.Text
        total()
    End Sub
    Private Sub txtVolume7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume7.TextChanged, txtPecasPorVolume7.TextChanged
        If txtVolume7.TextLength = 0 Then
            txtVolume7.Text = 0
        End If
        If txtPecasPorVolume7.TextLength = 0 Then
            txtPecasPorVolume7.Text = 0
        End If
        txtQuantidade7.Text = txtVolume7.Text * txtPecasPorVolume7.Text
        total()
    End Sub
    Private Sub txtVolume8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume8.TextChanged, txtPecasPorVolume8.TextChanged
        If txtVolume8.TextLength = 0 Then
            txtVolume8.Text = 0
        End If
        If txtPecasPorVolume8.TextLength = 0 Then
            txtPecasPorVolume8.Text = 0
        End If
        txtQuantidade8.Text = txtVolume8.Text * txtPecasPorVolume8.Text
        total()
    End Sub
    Private Sub txtVolume9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume9.TextChanged, txtPecasPorVolume9.TextChanged
        If txtVolume9.TextLength = 0 Then
            txtVolume9.Text = 0
        End If
        If txtPecasPorVolume9.TextLength = 0 Then
            txtPecasPorVolume9.Text = 0
        End If
        txtQuantidade9.Text = txtVolume9.Text * txtPecasPorVolume9.Text
        total()
    End Sub
    Private Sub txtVolume10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolume10.TextChanged, txtPecasPorVolume10.TextChanged
        If txtVolume10.TextLength = 0 Then
            txtVolume10.Text = 0
        End If
        If txtPecasPorVolume10.TextLength = 0 Then
            txtPecasPorVolume10.Text = 0
        End If
        txtQuantidade10.Text = txtVolume10.Text * txtPecasPorVolume10.Text
        total()
    End Sub
    'OK
    Sub total()
        VerificarRB()
        Try
            Select Case terRB
                Case 1
                    lblTotal.Text = txtQuantidade1.Text
                Case 2
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text)
                Case 3
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text)
                Case 4
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text)
                Case 5
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text) + Integer.Parse(txtQuantidade5.Text)
                Case 6
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text) + Integer.Parse(txtQuantidade5.Text) + Integer.Parse(txtQuantidade6.Text)
                Case 7
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text) + Integer.Parse(txtQuantidade5.Text) + Integer.Parse(txtQuantidade6.Text) + Integer.Parse(txtQuantidade7.Text)
                Case 8
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text) + Integer.Parse(txtQuantidade5.Text) + Integer.Parse(txtQuantidade6.Text) + Integer.Parse(txtQuantidade7.Text) + Integer.Parse(txtQuantidade8.Text)
                Case 9
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text) + Integer.Parse(txtQuantidade5.Text) + Integer.Parse(txtQuantidade6.Text) + Integer.Parse(txtQuantidade7.Text) + Integer.Parse(txtQuantidade8.Text) + Integer.Parse(txtQuantidade9.Text)
                Case 10
                    lblTotal.Text = Integer.Parse(txtQuantidade1.Text) + Integer.Parse(txtQuantidade2.Text) + Integer.Parse(txtQuantidade3.Text) + Integer.Parse(txtQuantidade4.Text) + Integer.Parse(txtQuantidade5.Text) + Integer.Parse(txtQuantidade6.Text) + Integer.Parse(txtQuantidade7.Text) + Integer.Parse(txtQuantidade8.Text) + Integer.Parse(txtQuantidade9.Text) + Integer.Parse(txtQuantidade10.Text)
            End Select
        Catch ex As Exception
        End Try
    End Sub
    'OK
    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try

            Dim row As DataGridViewRow = Me.DataGridView2.CurrentRow

            Dim Cliente = row.Cells(2)
            Dim Email = row.Cells(4)

            Me.txtCliente1.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente2.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente3.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente4.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente5.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente6.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente7.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente8.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente9.Text = Convert.ToString(Cliente.Value)
            Me.txtCliente10.Text = Convert.ToString(Cliente.Value)

            Me.incluirEmail = Email.Value

        Catch ex As Exception
            MsgBox("Erro 70TP " & ex.Message)
        End Try
    End Sub
    'OK
    Private Sub SoNumeros(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPedido.KeyPress, txtNotaFiscal.KeyPress, txtRE.KeyPress, txtOP1.KeyPress, txtOP2.KeyPress, txtOP3.KeyPress, txtOP4.KeyPress, txtOP5.KeyPress, txtOP6.KeyPress, txtOP7.KeyPress, txtOP8.KeyPress, txtOP9.KeyPress, txtOP10.KeyPress, txtPecasPorVolume1.KeyPress, txtPecasPorVolume2.KeyPress, txtPecasPorVolume3.KeyPress, txtPecasPorVolume4.KeyPress, txtPecasPorVolume5.KeyPress, txtPecasPorVolume6.KeyPress, txtPecasPorVolume7.KeyPress, txtPecasPorVolume8.KeyPress, txtPecasPorVolume9.KeyPress, txtPecasPorVolume10.KeyPress, txtQuantidade1.KeyPress, txtQuantidade2.KeyPress, txtQuantidade3.KeyPress, txtQuantidade4.KeyPress, txtQuantidade5.KeyPress, txtQuantidade6.KeyPress, txtQuantidade7.KeyPress, txtQuantidade8.KeyPress, txtQuantidade9.KeyPress, txtQuantidade10.KeyPress, txtVolume1.KeyPress, txtVolume2.KeyPress, txtVolume3.KeyPress, txtVolume4.KeyPress, txtVolume5.KeyPress, txtVolume6.KeyPress, txtVolume7.KeyPress, txtVolume8.KeyPress, txtVolume9.KeyPress, txtVolume10.KeyPress

        Try
            Dim Keyascii As Short = CShort(Asc(e.KeyChar))
            Keyascii = CShort(Numero4(Keyascii))
            If Keyascii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            MsgBox("Erro 557gh " & ex.Message)
        End Try
    End Sub
    'OK
    Function Numero4(ByVal Keyascii As Short) As Short
        If InStr("1234567890", Chr(Keyascii)) = 0 Then
            Numero4 = 0
        Else
            Numero4 = Keyascii
        End If
        Select Case Keyascii
            Case 8 'permite backspace
                Numero4 = Keyascii
                ' Case 13
                '    Numero4 = Keyascii
                'Case 32 'permite espaço
                '   SoNumeros = Keyascii
        End Select
    End Function
    'OK
    Private Sub txtQuantidade1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtQuantidade1.TextChanged, txtQuantidade2.TextChanged, txtQuantidade3.TextChanged, txtQuantidade4.TextChanged, txtQuantidade5.TextChanged, txtQuantidade6.TextChanged, txtQuantidade7.TextChanged, txtQuantidade8.TextChanged, txtQuantidade9.TextChanged, txtQuantidade10.TextChanged
        If txtQuantidade1.TextLength = 0 Then
            txtQuantidade1.Text = 0
        End If
        If txtQuantidade2.TextLength = 0 Then
            txtQuantidade2.Text = 0
        End If
        If txtQuantidade3.TextLength = 0 Then
            txtQuantidade3.Text = 0
        End If
        If txtQuantidade4.TextLength = 0 Then
            txtQuantidade4.Text = 0
        End If
        If txtQuantidade5.TextLength = 0 Then
            txtQuantidade5.Text = 0
        End If
        If txtQuantidade6.TextLength = 0 Then
            txtQuantidade6.Text = 0
        End If
        If txtQuantidade7.TextLength = 0 Then
            txtQuantidade7.Text = 0
        End If
        If txtQuantidade8.TextLength = 0 Then
            txtQuantidade8.Text = 0
        End If
        If txtQuantidade9.TextLength = 0 Then
            txtQuantidade9.Text = 0
        End If
        If txtQuantidade10.TextLength = 0 Then
            txtQuantidade10.Text = 0
        End If

        total()
    End Sub
    'OK
    Private Sub btAlterar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterar.Click
        Try
            If btAlterar.Text = "Alterar" Then
                'If MsgBox("Deseja Alterar os Certificados de um certo Pedido?", vbYesNo, "Alterar Certificado") = vbYes Then
                Email_E_Impressao()
                LimparTudo()
                rbNao.Focus()
                btAlterar.Text = "Aplicar"
                btCriar.Enabled = False
                btExcluir.Enabled = False
                btImprimir.Enabled = False
                btEmail.Enabled = False
                lblData.Text = Today
                lblHora.Text = TimeOfDay.ToShortTimeString
                lblDataAlterado.Text = Today
                lblHoraAlterado.Text = TimeOfDay.ToShortTimeString
                GroupBox4.Enabled = False
                'Else
                ' End If
                'se botão Criar for  = Aplicar
            Else
                If txtNotaFiscal.Text = "" Then
                    cbEnviarEmail.Checked = False
                    cbImprimir.Checked = False
                End If
                interromper = "Não"
                AlterarEnviar()
                If interromper = "Não" Then
                    LimparTudo()
                End If
                AtualizarGrid()
            End If
        Catch ex As Exception
            MsgBox("Erro x84 " & ex.Message)
        End Try
    End Sub
    'OK
    Sub AlterarEnviar()
        VerPadrao()
        If interromper = "Sim" Then
        ElseIf interromper = "Não" Then
            Alterar()
        End If
    End Sub

    Sub Alterar()
        VerificarRB()
        Try
            Dim i As Byte = 0
            conCertificado.Open()
            'VerificarRB()
            For i = 1 To terRB Step 1
                Controles()

                If i = 1 Then
                    novoiD = novoiD - (terRB - i)
                Else
                    novoiD += 1
                End If
                'Try
                'Kill("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & "%" & novoiD & "-" & lblCodigo1.Text & "%" & "") ' deleta o arquivo da pasta
                'Catch ex As Exception
                'End Try
                Dim da25 As New OleDbDataAdapter
                Dim ds25 As New DataSet
                ds25 = New DataSet
                da25 = New OleDbDataAdapter("UPDATE tblCertificado SET Pedido = '" & txtPedido.Text & "', NotaFiscal = '" & txtNotaFiscal.Text & "',Produto = '" & lblProdutox.Text & "', Codigo = '" & lblCodigox.Text & "', Invoice = '" & txtInvoice.Text & "', OP = '" & txtOPx.Text & "', Volume = '" & txtVolumex.Text & "', Quantidade = '" & txtQuantidadex.Text & "', DataFab_Inicio= '" & dtpDex.Value & "', DataFab_Fim = '" & dtpAtex.Value & "', Obs = '" & Obsx.ToString() & "', DataAlteracao = '" & Today.ToShortDateString() & "', HoraAlteracao = '" & TimeOfDay.ToShortTimeString & "', Cliente = '" & txtClientex.Text & "', Inspetor= '" & lblInspetor.Text & "'  WHERE Id = " & novoiD & "", conCertificado)
                ds25.Clear()
                da25.Fill(ds25, "tblRNC")
                conCertificado.Close()

                If rbSim.Checked = True Then
                    LansarNoExcel() ' lança no excell, salva em PDF e Imprime
                    ' adicionando quantos anexos deve-se enviar
                    idx = anex
                    val = Integer.Parse(lblID.Text)
                    arrei_anexos.SetValue(val, idx)
                    anex += 1
                End If
            Next
            If cbEnviarEmail.Checked = True Then
                Email() 'Armazena os anexos para Email() e depois EneviarEmail()
            End If
            'MsgBox("Alteração registrada com sucesso!")
        Catch ex As Exception
            conCertificado.Close()
            MsgBox("Erro 15 " & ex.Message)
        End Try
    End Sub

    Sub Email_E_Impressao()
        If cbEnviarEmail.Checked = True And cbImprimir.Checked = True Then
            MsgBox("Os certificados serão:" _
                   & Chr(13) _
                   & Chr(13) _
                   & "Impressos e" _
                   & Chr(13) _
                   & Chr(13) _
                   & "Enviados por E-mail")
        ElseIf cbEnviarEmail.Checked = True And cbImprimir.Checked = False Then
            MsgBox("Os certificados serão:" _
                    & Chr(13) _
                    & Chr(13) _
                    & "Apenas enviados por E-mail")
        ElseIf cbEnviarEmail.Checked = False And cbImprimir.Checked = True Then
            MsgBox("Os certificados serão:" _
                    & Chr(13) _
                    & Chr(13) _
                    & "Apenas Impressos")
        ElseIf cbEnviarEmail.Checked = False And cbImprimir.Checked = False Then
            MsgBox("Os certificados 'NÃO' serão:" _
                    & Chr(13) _
                    & Chr(13) _
                    & "Impressos & enviados por E-mail")
        End If
    End Sub
    Dim idIndividual As Integer
    Private Sub btAlterarIndividual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAlterarIndividual.Click
        Try
            If btAlterarIndividual.Text = "Alterar" Then
                'If MsgBox("Deseja Alterar um Certificado?", vbYesNo, "Alterar Certificado") = vbYes Then
                Email_E_Impressao()
                idIndividual = InputBox(Title:="ID", Prompt:="Insira o ID do Certificado:", XPos:=615, YPos:=300)
                btAlterarIndividual.Text = "Aplicar"
                PesquisarIndividual()
                btCriar.Enabled = False
                btExcluir.Enabled = False
                btImprimir.Enabled = False
                btImprimirIndividual.Enabled = False
                btEmail.Enabled = False
                btEmailIndividual.Enabled = False
                btAlterar.Enabled = False
                rbNao.Focus()
                'Else
                'End If
            Else
            If CheckBox1.Checked = True Then
                Obs1 = InputBox(Title:="Observação", Prompt:="Insira a Observação!", XPos:=615, YPos:=300)
            End If
            If txtNotaFiscal.Text = "" Then
                cbEnviarEmail.Checked = False
                cbImprimir.Checked = False
                'Try
                'Kill("f:\RECEB.MAT.PRIMA\Banco_Dados\ProjetoCertificado\Certificados_Salvos\" & idIndividual & "-" & lblCodigo1.Text & ".pdf") ' deleta o arquivo da pasta
                'Catch ex As Exception
                'End Try
            End If
            conCertificado.Open()
            Dim da25 As New OleDbDataAdapter
            Dim ds25 As New DataSet
            ds25 = New DataSet
            da25 = New OleDbDataAdapter("UPDATE tblCertificado SET NotaFiscal = '" & txtNotaFiscal.Text & "', Produto = '" & lblProduto1.Text & "', Codigo = '" & lblCodigo1.Text & "', OP = '" & txtOP1.Text & "', Volume = '" & txtVolume1.Text & "', Quantidade = '" & txtQuantidade1.Text & "', DataFab_Inicio= '" & dtpDe1.Value & "', DataFab_Fim = '" & dtpAte1.Value & "', Obs = '" & Obs1.ToString() & "', DataAlteracao = '" & Today.ToShortDateString() & "', HoraAlteracao = '" & TimeOfDay.ToShortTimeString & "', Inspetor= '" & lblInspetor.Text & "'  WHERE Id = " & idIndividual & "", conCertificado)
            ds25.Clear()
            da25.Fill(ds25, "tblRNC")
            conCertificado.Close()
            If rbSim.Checked = True Then
                verRB = 1
                Controles()
                LansarNoExcel()
            End If
            If cbEnviarEmail.Checked = True Then
                Email()
            End If
            PesquisarIndividual2()
            LimparTudo()
                'MsgBox("Alteração registrada com sucesso!")
            AtualizarGrid()
            End If
        Catch ex As Exception
            MsgBox("Erro x84X " & ex.Message)
            conCertificado.Close()
        Finally
            GroupBox2.Enabled = True
        End Try
    End Sub
    Dim ds38 As New DataSet
    Dim dt38 As New DataTable
    Dim sel_x As String
    Sub PesquisarIndividual()
        Try
            Dim da38 As New OleDbDataAdapter
            conCertificado.Open()
            If btEmailIndividual.Text = "...Email" Or btAlterarIndividual.Text = "Aplicar" Then
                sel_x = "SELECT * FROM tblCertificado WHERE ID = " & idIndividual & ""
            Else
                sel_x = "SELECT * FROM tblCertificado WHERE ID = " & idReal & ""
            End If
            da38 = New OleDbDataAdapter(sel_x, conCertificado)
            ds38.Clear()
            da38.Fill(ds38, "tblCertificado")
            da38.Fill(dt38)
            conCertificado.Close()
            If dt38.Rows.Count() <> 0 Then
                Seguranca1()
            Else
                MsgBox("O ID não existe", , "ID Inexistente")
            End If
            If btEmailIndividual.Text = "...Email" Then
                Limpar2()
                Preencher2()
            Else
                Preencher2()
            End If
        Catch ex As Exception
            MsgBox("Erro dt38 " & ex.Message)
            conCertificado.Close()
        Finally
            conCertificado.Close()
        End Try
    End Sub
    Sub Seguranca1()
        Try
            Dim sel_p As String
            Dim dt79 As New DataTable
            Dim ds79 As New DataSet
            Dim da79 As New OleDbDataAdapter
            conCertificado.Open()
            sel_p = "SELECT Pedido, Data, Hora FROM tblCertificado WHERE Pedido = '" & dt38.Rows(0)("Pedido").ToString & "' and Data = '" & dt38.Rows(0)("Data").ToString & "' and Hora = '" & dt38.Rows(0)("Hora").ToString & "' "
            da79 = New OleDbDataAdapter(sel_p, conCertificado)
            ds79.Clear()
            da79.Fill(ds79, "tblCertificado")
            da79.Fill(dt79)
            conCertificado.Close()
            Select Case dt79.Rows.Count()
                Case 1
                Case Else
                    GroupBox2.Enabled = False
                    MsgBox("A Nota Fiscal não poderá ser alterda para o conjunto deste Pedido acima com este comando. Ultilize o alterar em grupo", , "Segurança")
            End Select
        Catch ex As Exception
            MsgBox("Erro dt79 " & ex.Message)
            conCertificado.Close()
        Finally
            conCertificado.Close()
        End Try
    End Sub

    Sub Limpar2()
        rb1T.Checked = True
        CheckBox1.Checked = False
        Obsx = ""
        lblProduto1.Text = "*"
        lblCodigo1.Text = "*"
        txtCliente1.Clear()
        txtVolume1.Text = "0"
        txtQuantidade1.Text = "0"
        txtPecasPorVolume1.Text = "0"
        dtpDe1.Value = Today.ToShortDateString()
        dtpAte1.Value = Today.ToShortDateString()
        txtOP1.Clear()
    End Sub
    Sub Preencher2()
        If btEmailIndividual.Text = "...Email" Or btAlterarIndividual.Text = "Aplicar" Then
            txtOP1.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("OP"))
            lblProduto1.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Produto"))
            lblCodigo1.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Codigo"))
            txtCliente1.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Cliente"))
            txtVolume1.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Volume"))
            txtQuantidade1.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Quantidade"))
            dtpDe1.Value = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("DataFab_Inicio"))
            dtpAte1.Value = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("DataFab_Fim"))
            Obs1 = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Obs"))
            lblInspetor.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Inspetor"))
            txtPedido.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Pedido"))
            txtNotaFiscal.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("NotaFiscal"))
            txtInvoice.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Invoice"))
            If Obs1 = Nothing Or Obs1 = "" Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
        Else
            txtOPx.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("OP"))
            lblProdutox.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Produto"))
            lblCodigox.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Codigo"))
            txtClientex.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Cliente"))
            txtVolumex.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Volume"))
            txtQuantidadex.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Quantidade"))
            dtpDex.Value = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("DataFab_Inicio"))
            dtpAtex.Value = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("DataFab_Fim"))
            Obsx = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Obs"))
            lblInspetor.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Inspetor"))
            txtPedido.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Pedido"))
            txtNotaFiscal.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("NotaFiscal"))
            txtInvoice.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Invoice"))
            ' lblData.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Data"))
            'lblHora.Text = Convert.ToString(ds38.Tables("tblCertificado").Rows(0).Item("Hora"))
            If Obsx = Nothing Or Obsx = "" Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
        End If

    End Sub
    Sub PesquisarIndividual2()
        Try
            Dim da38 As New OleDbDataAdapter
            conCertificado.Open()
            Dim sel_ As String = "SELECT * FROM tblCertificado WHERE ID = " & idIndividual & ""
            da38 = New OleDbDataAdapter(sel_, conCertificado)
            ds38.Clear()
            da38.Fill(ds38, "tblCertificado")
            conCertificado.Close()
            Me.DataGridView1.DataSource = ds38
            Me.DataGridView1.DataMember = "tblCertificado"
            FormatacaoGrid()
        Catch ex As Exception
            MsgBox("Erro dt38 " & ex.Message)
            conCertificado.Close()
        Finally
            conCertificado.Close()
        End Try
    End Sub
    Dim ds21 As New DataSet
    Dim dt21 As New DataTable
    Sub Consulta2()
        Dim da21 As New OleDbDataAdapter
        Dim dt21 As New System.Data.DataTable
        conPecasVolume.Open()
        Dim sel5 As String = "SELECT top 2 * FROM tblVolume where Codigo = '" & lblCodigo1.Text & "'"
        da21 = New OleDbDataAdapter(sel5, conPecasVolume)
        dt21.Clear()
        conPecasVolume.Close()
        da21.Fill(ds21, "tblVolume")
        da21.Fill(dt21)
    End Sub

    Private Sub btEmailIndividual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEmailIndividual.Click
        If btEmailIndividual.Text = "Email" Then
            'If MsgBox("Deseja enviar um Certificado?", vbYesNo, "Enviar Certificado") = vbYes Then
            MsgBox("Selecione um Certificado abaixo a enviar", , "Selecionar - Certificado")
            btEmailIndividual.Text = "...Email"
            GroupBox2.Enabled = False
            btCriar.Enabled = False
            btExcluir.Enabled = False
            btImprimir.Enabled = False
            btImprimirIndividual.Enabled = False
            btEmail.Enabled = False
            btAlterarIndividual.Enabled = False
            btAlterar.Enabled = False
            'Else
            'End If
        Else
        If rbSim.Checked = True Then
            If MsgBox("O Certificado a ser enviado é o: " & idReal & " ?", vbYesNo, "Certificado Selecionado") = vbYes Then
                Email()
                LimparTudo()
            End If
        ElseIf rbNao.Checked = True Then
            MsgBox("O Certificado não pode ser enviado pois não tem nota fiscal", , "Certificado sem Nota Fiscal")
        End If
        End If
    End Sub

    Private Sub btImprimirIndividual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btImprimirIndividual.Click
        Try
            If rbSim.Checked = True Then
                If btImprimirIndividual.Text = "Imprimir" Then
                    'If MsgBox("Deseja imprimir um Certificado?", vbYesNo, "Enviar Certificado") = vbYes Then
                    If MsgBox("O Certificado a ser impresso é o: " & idReal & " ?", vbYesNo, "Certificado Selecionado") = vbYes Then
                        btImprimirIndividual.Text = "...Imprimir"
                        cbImprimir.Checked = False
                        btCriar.Enabled = False
                        btExcluir.Enabled = False
                        btImprimir.Enabled = False
                        btEmailIndividual.Enabled = False
                        btEmail.Enabled = False
                        btAlterarIndividual.Enabled = False
                        btAlterar.Enabled = False


                        'Microsoft.VisualBasic.FileOpen (FileName:= ""
                        PesquisarIndividual()
                        LansarNoExcel()


                        btImprimirIndividual.Text = "Imprimir"
                        btCriar.Enabled = True
                        btExcluir.Enabled = True
                        btImprimir.Enabled = True
                        btEmailIndividual.Enabled = True
                        btEmail.Enabled = True
                        btAlterarIndividual.Enabled = True
                        btAlterar.Enabled = True
                    Else
                        MsgBox("Selecione abaixo!", , "Selecionar - Certificados")
                    End If
                    'End If
                End If
            ElseIf rbNao.Checked = True Then
                MsgBox("O Certificado não pode ser impresso pois não tem nota fiscal", , "Certificado sem Nota Fiscal")
            End If
        Catch ex As Exception
            btImprimirIndividual.Text = "Imprimir"
        Finally
            btImprimirIndividual.Text = "Imprimir"
        End Try
    End Sub

    Private Sub dtpAte1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte1.LostFocus
        VerificarRB()
        If dtpDe1.Value <= dtpAte1.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe1.Focus()
        End If
        Select Case terRB
            Case 1
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte2.LostFocus
        If dtpDe2.Value <= dtpAte2.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe2.Focus()
        End If
        Select Case terRB
            Case 2
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte3.LostFocus
        If dtpDe3.Value <= dtpAte3.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe3.Focus()
        End If
        Select Case terRB
            Case 3
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte4.LostFocus
        If dtpDe4.Value <= dtpAte4.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe4.Focus()
        End If
        Select Case terRB
            Case 4
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte5.LostFocus
        If dtpDe5.Value <= dtpAte5.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe5.Focus()
        End If
        Select Case terRB
            Case 5
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte6.LostFocus
        If dtpDe6.Value <= dtpAte6.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe6.Focus()
        End If
        Select Case terRB
            Case 6
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte7_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte7.LostFocus
        If dtpDe7.Value <= dtpAte7.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe7.Focus()
        End If
        Select Case terRB
            Case 7
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte8_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte8.LostFocus
        If dtpDe8.Value <= dtpAte8.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe8.Focus()
        End If
        Select Case terRB
            Case 8
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte9_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte9.LostFocus
        If dtpDe9.Value <= dtpAte9.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe9.Focus()
        End If
        Select Case terRB
            Case 9
                aplicar()
        End Select
    End Sub
    Private Sub dtpAte10_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpAte10.LostFocus
        If dtpDe10.Value <= dtpAte10.Value Then
        Else
            MsgBox("A data final não pode ser menor que a data inicial", , "Datas Divergentes")
            dtpDe10.Focus()
        End If
        Select Case terRB
            Case 10
                aplicar()
        End Select
    End Sub
    Sub aplicar()
        If btCriar.Text = "Aplicar" Then
            btCriar.Focus()
        ElseIf btAlterar.Text = "Aplicar" Then
            btAlterar.Focus()
        ElseIf btAlterarIndividual.Text = "Aplicar" Then
            btAlterarIndividual.Focus()
        End If
    End Sub

    Private Sub txtNotaFiscal_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNotaFiscal.LostFocus
        If txtNotaFiscal.TextLength = 0 Then
            cbEnviarEmail.Checked = False
            cbImprimir.Checked = False
            rbNao.Checked = True
        ElseIf txtNotaFiscal.TextLength <> 0 Then
            rbSim.Checked = True
        End If
    End Sub


End Class

''''  realizar um teste geral'''''''---------------
