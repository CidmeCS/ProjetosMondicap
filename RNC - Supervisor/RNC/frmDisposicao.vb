Imports System.Data.OleDb
Public Class frmDisposicao
    Dim conRNC As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Cid\Documents\Projetos\BancoDados\RNC_RNC.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim dtPrint As New DataTable
    Dim Disposicao1, Disposicao2, Disposicao3, Disposicao4, Disposicao5, Disposicao6, Disposicao7, Disposicao8, Disposicao9, Disposicao10 As String
    Dim Id1, Id2, Id3, Id4, Id5, Id6, Id7, Id8, Id9, Id10 As Integer


    Private Sub frmDisposicao_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Carregar()
    End Sub
    Sub Carregar()
        TesteAbertoRNC()
        Try
            Dim dat As New OleDbDataAdapter
            Dim dst As New DataSet

            conRNC.Open()
            Dim selt As String = "Select top 100 * from tblRNC where Status = 'Pendente' and Disposicao = 'Sem Disposição' order by ID asc "
            'Dim sel As String = "select Contador, count (*) from tblRNC group by Contador order by contador desc" 'conta quantas RNCs exitem
            dat = New OleDbDataAdapter(selt, conRNC)
            dst.Clear()
            dat.Fill(dst, "tblRNC")
            Me.DataGridView1.DataSource = dst
            Me.DataGridView1.DataMember = "tblRNC"
            FormatacaoGrid()
            lblData.Text = Today
            lblHora.Text = TimeOfDay.ToShortTimeString
            conRNC.Close()
        Catch ex As Exception
            Beep()
            MsgBox("Erro A67 " & ex.Message)
        End Try
    End Sub
    Private Sub frm_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Try

            If e.KeyChar = Convert.ToChar(13) Then
                e.Handled = True

                SendKeys.Send("{TAB}")
            End If
        Catch ex As Exception
            MsgBox("Erro 53ZD " & ex.Message)
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

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        TesteAbertoRNC()

        Dim row As DataGridViewRow = Me.DataGridView1.CurrentRow


        Dim ID = row.Cells(0)
        Dim RNC = row.Cells(1)
        Dim Origem = row.Cells(3)
        Dim Data_Abertura = row.Cells(4)
        Dim Hora = row.Cells(5)
        Dim CodProd = row.Cells(7)
        Dim Produto = row.Cells(9)
        Dim OP_Reprovado = row.Cells(10)
        Dim Maquina = row.Cells(17)
        Dim RE = row.Cells(26)
        Dim Inspetor = row.Cells(27)
        Dim Setor = row.Cells(28)
        Dim TurnoDetector = row.Cells(29)

        If lblRNC.Text = "*" Then
            'LimparDisposicao()
            Me.lblRNC.Text = RNC.Value
            AlterarCarregar()
        ElseIf RNC.Value = lblRNC.Text Then
        Else
            LimparDisposicao()
            Me.lblRNC.Text = RNC.Value
            AlterarCarregar()
        End If
        Me.lblID.Text = ID.Value
        Me.lblDetectado.Text = Origem.Value
        Me.lblData.Text = Data_Abertura.Value
        Me.lblHora.Text = Hora.Value
        Me.lblCodProduto.Text = CodProd.Value
        Me.lblProduto.Text = Produto.Value
        Me.lblOPReprovada.Text = OP_Reprovado.Value
        Me.lblMaquina.Text = Maquina.Value
        Me.lblRE.Text = RE.Value
        Me.lblInspetor.Text = Inspetor.Value
        Me.lblSetor.Text = Setor.Value
        Me.lblTurno.Text = TurnoDetector.Value
    End Sub
    Private Sub AlterarCarregar()
        Try
            conRNC.Open()
            Dim selPRINT As String = "SELECT top 10 * FROM tblRNC where RNC = " & lblRNC.Text & " order by ID asc"
            Dim daPRINT As New OleDbDataAdapter
            Dim dsPRINT As New DataSet
            'Dim dtPrint As New DataTable [estar lá na declaração de variáveis]
            daPRINT = New OleDbDataAdapter(selPRINT, conRNC)
            dsPRINT.Clear()
            dtPrint.Clear()
            daPRINT.Fill(dsPRINT, "tblRNC")
            daPRINT.Fill(dtPrint)
            conRNC.Close()

            Id1 = dsPRINT.Tables("tblRNC").Rows(0)("ID")
            lbl11.Text = dsPRINT.Tables("tblRNC").Rows(0)("Turno")
            lbl12.Text = dsPRINT.Tables("tblRNC").Rows(0)("NúmerosCaixas")
            lbl13.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_Caixas")
            lbl14.Text = dsPRINT.Tables("tblRNC").Rows(0)("QT_Reprovado")
            lbl15.Text = dsPRINT.Tables("tblRNC").Rows(0)("Cod_Defeito")
            lbl16.Text = dsPRINT.Tables("tblRNC").Rows(0)("Nao_Conformidade")
            If (dsPRINT.Tables("tblRNC").Rows(0)("Disposicao")).ToString() = "Sem Disposição" Then
                rbRT1.Checked = False
                rbRF1.Checked = False
                rbLC1.Checked = False
            ElseIf (dsPRINT.Tables("tblRNC").Rows(0)("Disposicao")).ToString() = "Retrabalhar" Then
                rbRT1.Checked = True
            ElseIf (dsPRINT.Tables("tblRNC").Rows(0)("Disposicao")).ToString() = "Liberado Condicional" Then
                rbLC1.Checked = True
            ElseIf (dsPRINT.Tables("tblRNC").Rows(0)("Disposicao")).ToString() = "Refugar" Then
                rbRF1.Checked = True
            End If

            If dtPrint.Rows.Count >= 2 Then
                Id2 = dsPRINT.Tables("tblRNC").Rows(1)("ID")
                lbl21.Text = dsPRINT.Tables("tblRNC").Rows(1)("Turno")
                lbl22.Text = dsPRINT.Tables("tblRNC").Rows(1)("NúmerosCaixas")
                lbl23.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_Caixas")
                lbl24.Text = dsPRINT.Tables("tblRNC").Rows(1)("QT_Reprovado")
                lbl25.Text = dsPRINT.Tables("tblRNC").Rows(1)("Cod_Defeito")
                lbl26.Text = dsPRINT.Tables("tblRNC").Rows(1)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(1)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT2.Checked = False
                    rbRF2.Checked = False
                    rbLC2.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(1)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT2.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(1)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC2.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(1)("Disposicao")).ToString() = "Refugar" Then
                    rbRF2.Checked = True
                End If
            End If

            If dtPrint.Rows.Count >= 3 Then
                Id3 = dsPRINT.Tables("tblRNC").Rows(2)("ID")
                lbl31.Text = dsPRINT.Tables("tblRNC").Rows(2)("Turno")
                lbl32.Text = dsPRINT.Tables("tblRNC").Rows(2)("NúmerosCaixas")
                lbl33.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_Caixas")
                lbl34.Text = dsPRINT.Tables("tblRNC").Rows(2)("QT_Reprovado")
                lbl35.Text = dsPRINT.Tables("tblRNC").Rows(2)("Cod_Defeito")
                lbl36.Text = dsPRINT.Tables("tblRNC").Rows(2)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(2)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT3.Checked = False
                    rbRF3.Checked = False
                    rbLC3.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(2)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT3.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(2)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC3.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(2)("Disposicao")).ToString() = "Refugar" Then
                    rbRF3.Checked = True
                End If
            End If
            If dtPrint.Rows.Count >= 4 Then
                Id4 = dsPRINT.Tables("tblRNC").Rows(3)("ID")
                lbl41.Text = dsPRINT.Tables("tblRNC").Rows(3)("Turno")
                lbl42.Text = dsPRINT.Tables("tblRNC").Rows(3)("NúmerosCaixas")
                lbl43.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_Caixas")
                lbl44.Text = dsPRINT.Tables("tblRNC").Rows(3)("QT_Reprovado")
                lbl45.Text = dsPRINT.Tables("tblRNC").Rows(3)("Cod_Defeito")
                lbl46.Text = dsPRINT.Tables("tblRNC").Rows(3)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(3)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT4.Checked = False
                    rbRF4.Checked = False
                    rbLC4.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(3)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT4.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(3)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC4.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(3)("Disposicao")).ToString() = "Refugar" Then
                    rbRF4.Checked = True
                End If
            End If
            If dtPrint.Rows.Count >= 5 Then
                Id5 = dsPRINT.Tables("tblRNC").Rows(4)("ID")
                lbl51.Text = dsPRINT.Tables("tblRNC").Rows(4)("Turno")
                lbl52.Text = dsPRINT.Tables("tblRNC").Rows(4)("NúmerosCaixas")
                lbl53.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_Caixas")
                lbl54.Text = dsPRINT.Tables("tblRNC").Rows(4)("QT_Reprovado")
                lbl55.Text = dsPRINT.Tables("tblRNC").Rows(4)("Cod_Defeito")
                lbl56.Text = dsPRINT.Tables("tblRNC").Rows(4)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(4)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT5.Checked = False
                    rbRF5.Checked = False
                    rbLC5.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(4)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT5.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(4)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC5.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(4)("Disposicao")).ToString() = "Refugar" Then
                    rbRF5.Checked = True
                End If
            End If
            If dtPrint.Rows.Count >= 6 Then
                Id6 = dsPRINT.Tables("tblRNC").Rows(5)("ID")
                lbl61.Text = dsPRINT.Tables("tblRNC").Rows(5)("Turno")
                lbl62.Text = dsPRINT.Tables("tblRNC").Rows(5)("NúmerosCaixas")
                lbl63.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_Caixas")
                lbl64.Text = dsPRINT.Tables("tblRNC").Rows(5)("QT_Reprovado")
                lbl65.Text = dsPRINT.Tables("tblRNC").Rows(5)("Cod_Defeito")
                lbl66.Text = dsPRINT.Tables("tblRNC").Rows(5)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(5)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT6.Checked = False
                    rbRF6.Checked = False
                    rbLC6.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(5)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT6.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(5)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC6.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(5)("Disposicao")).ToString() = "Refugar" Then
                    rbRF6.Checked = True
                End If
            End If
            If dtPrint.Rows.Count >= 7 Then
                Id7 = dsPRINT.Tables("tblRNC").Rows(6)("ID")
                lbl71.Text = dsPRINT.Tables("tblRNC").Rows(6)("Turno")
                lbl72.Text = dsPRINT.Tables("tblRNC").Rows(6)("NúmerosCaixas")
                lbl73.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_Caixas")
                lbl74.Text = dsPRINT.Tables("tblRNC").Rows(6)("QT_Reprovado")
                lbl75.Text = dsPRINT.Tables("tblRNC").Rows(6)("Cod_Defeito")
                lbl76.Text = dsPRINT.Tables("tblRNC").Rows(6)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(6)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT7.Checked = False
                    rbRF7.Checked = False
                    rbLC7.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(6)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT7.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(6)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC7.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(6)("Disposicao")).ToString() = "Refugar" Then
                    rbRF7.Checked = True
                End If
            End If

            If dtPrint.Rows.Count >= 8 Then
                Id8 = dsPRINT.Tables("tblRNC").Rows(7)("ID")
                lbl81.Text = dsPRINT.Tables("tblRNC").Rows(7)("Turno")
                lbl82.Text = dsPRINT.Tables("tblRNC").Rows(7)("NúmerosCaixas")
                lbl83.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_Caixas")
                lbl84.Text = dsPRINT.Tables("tblRNC").Rows(7)("QT_Reprovado")
                lbl85.Text = dsPRINT.Tables("tblRNC").Rows(7)("Cod_Defeito")
                lbl86.Text = dsPRINT.Tables("tblRNC").Rows(7)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(7)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT8.Checked = False
                    rbRF8.Checked = False
                    rbLC8.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(7)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT8.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(7)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC8.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(7)("Disposicao")).ToString() = "Refugar" Then
                    rbRF8.Checked = True
                End If
            End If
            If dtPrint.Rows.Count >= 9 Then
                Id9 = dsPRINT.Tables("tblRNC").Rows(8)("ID")
                lbl91.Text = dsPRINT.Tables("tblRNC").Rows(8)("Turno")
                lbl92.Text = dsPRINT.Tables("tblRNC").Rows(8)("NúmerosCaixas")
                lbl93.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_Caixas")
                lbl94.Text = dsPRINT.Tables("tblRNC").Rows(8)("QT_Reprovado")
                lbl95.Text = dsPRINT.Tables("tblRNC").Rows(8)("Cod_Defeito")
                lbl96.Text = dsPRINT.Tables("tblRNC").Rows(8)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(8)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT9.Checked = False
                    rbRF9.Checked = False
                    rbLC9.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(8)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT9.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(8)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC9.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(8)("Disposicao")).ToString() = "Refugar" Then
                    rbRF9.Checked = True
                End If
            End If
            If dtPrint.Rows.Count = 10 Then
                Id10 = dsPRINT.Tables("tblRNC").Rows(9)("ID")
                lbl101.Text = dsPRINT.Tables("tblRNC").Rows(9)("Turno")
                lbl102.Text = dsPRINT.Tables("tblRNC").Rows(9)("NúmerosCaixas")
                lbl103.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_Caixas")
                lbl104.Text = dsPRINT.Tables("tblRNC").Rows(9)("QT_Reprovado")
                lbl105.Text = dsPRINT.Tables("tblRNC").Rows(9)("Cod_Defeito")
                lbl106.Text = dsPRINT.Tables("tblRNC").Rows(9)("Nao_Conformidade")
                If (dsPRINT.Tables("tblRNC").Rows(9)("Disposicao")).ToString() = "Sem Disposição" Then
                    rbRT10.Checked = False
                    rbRF10.Checked = False
                    rbLC10.Checked = False
                ElseIf (dsPRINT.Tables("tblRNC").Rows(9)("Disposicao")).ToString() = "Retrabalhar" Then
                    rbRT10.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(9)("Disposicao")).ToString() = "Liberado Condicional" Then
                    rbLC10.Checked = True
                ElseIf (dsPRINT.Tables("tblRNC").Rows(9)("Disposicao")).ToString() = "Refugar" Then
                    rbRF10.Checked = True
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 83 " & ex.Message)
        End Try
    End Sub

    Private Sub Label9_Clck(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label44.MouseEnter, Label45.MouseEnter, Label46.MouseEnter
        frmLegenda.Show()
    End Sub
    Private Sub Label_9_Clck(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        frmLegenda.Close()
    End Sub
    Sub LimparDisposicao()

        btAlterar.Text = "Alterar"

        lblID.Text = "*"
        lblRNC.Text = "*"
        lblData.Text = "*"
        lblHora.Text = "*"
        lblOPReprovada.Text = "*"
        lblCodProduto.Text = "*"
        lblProduto.Text = "*"
        lblMaquina.Text = "*"
        lblDetectado.Text = "*"

        lblSetor.Text = "*"
        lblInspetor.Text = "*"
        lblTurno.Text = "*"
        lblRE.Text = "*"

        lbl11.Text = ""
        lbl12.Text = ""
        lbl13.Text = ""
        lbl14.Text = ""
        lbl15.Text = ""
        lbl16.Text = ""
        lbl21.Text = ""
        lbl22.Text = ""
        lbl23.Text = ""
        lbl24.Text = ""
        lbl25.Text = ""
        lbl26.Text = ""
        lbl31.Text = ""
        lbl32.Text = ""
        lbl33.Text = ""
        lbl34.Text = ""
        lbl35.Text = ""
        lbl36.Text = ""
        lbl41.Text = ""
        lbl42.Text = ""
        lbl43.Text = ""
        lbl44.Text = ""
        lbl45.Text = ""
        lbl46.Text = ""
        lbl51.Text = ""
        lbl52.Text = ""
        lbl53.Text = ""
        lbl54.Text = ""
        lbl55.Text = ""
        lbl56.Text = ""
        lbl61.Text = ""
        lbl62.Text = ""
        lbl63.Text = ""
        lbl64.Text = ""
        lbl65.Text = ""
        lbl66.Text = ""
        lbl71.Text = ""
        lbl72.Text = ""
        lbl73.Text = ""
        lbl74.Text = ""
        lbl75.Text = ""
        lbl76.Text = ""
        lbl81.Text = ""
        lbl82.Text = ""
        lbl83.Text = ""
        lbl84.Text = ""
        lbl85.Text = ""
        lbl86.Text = ""
        lbl91.Text = ""
        lbl92.Text = ""
        lbl93.Text = ""
        lbl94.Text = ""
        lbl95.Text = ""
        lbl96.Text = ""
        lbl101.Text = ""
        lbl102.Text = ""
        lbl103.Text = ""
        lbl104.Text = ""
        lbl105.Text = ""
        lbl106.Text = ""
        rbLC1.Checked = False
        rbRT1.Checked = False
        rbRF1.Checked = False
        rbLC2.Checked = False
        rbRT2.Checked = False
        rbRF2.Checked = False
        rbLC3.Checked = False
        rbRT3.Checked = False
        rbRF3.Checked = False
        rbLC4.Checked = False
        rbRT4.Checked = False
        rbRF4.Checked = False
        rbLC5.Checked = False
        rbRT5.Checked = False
        rbRF5.Checked = False
        rbLC6.Checked = False
        rbRT6.Checked = False
        rbRF6.Checked = False
        rbLC7.Checked = False
        rbRT7.Checked = False
        rbRF7.Checked = False
        rbLC8.Checked = False
        rbRT8.Checked = False
        rbRF8.Checked = False
        rbLC9.Checked = False
        rbRT9.Checked = False
        rbRF9.Checked = False
        rbLC10.Checked = False
        rbRT10.Checked = False
        rbRF10.Checked = False



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
                    Else
                    End If

                Else

                    If dtPrint.Rows.Count = 1 Then
                        Alterar1()
                    ElseIf dtPrint.Rows.Count = 2 Then
                        Alterar2()
                    ElseIf dtPrint.Rows.Count = 3 Then
                        Alterar3()
                    ElseIf dtPrint.Rows.Count = 4 Then
                        Alterar4()
                    ElseIf dtPrint.Rows.Count = 5 Then
                        Alterar5()
                    ElseIf dtPrint.Rows.Count = 6 Then
                        Alterar6()
                    ElseIf dtPrint.Rows.Count = 7 Then
                        Alterar7()
                    ElseIf dtPrint.Rows.Count = 8 Then
                        Alterar8()
                    ElseIf dtPrint.Rows.Count = 9 Then
                        Alterar9()
                    ElseIf dtPrint.Rows.Count = 10 Then
                        Alterar10()
                    End If
                    btAlterar.Text = "Alterar"
                    conRNC.Close()
                    Carregar()
                    LimparDisposicao()
                    MsgBox("Dados alterados com sucesso")
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro 72 " & ex.Message)
        End Try
    End Sub
    Sub Alterar1()
        If rbRT1.Checked = True Then
            Disposicao1 = "Retrabalhar"
            Alterar1a()
        ElseIf rbLC1.Checked = True Then
            Disposicao1 = "Liberado Condicional"
            Alterar1a()
        ElseIf rbRF1.Checked = True Then
            Disposicao1 = "Refugar"
            Alterar1a()
        End If
    End Sub
    Sub Alterar1a()
        Try
            conRNC.Open()
            Dim da20 As New OleDbDataAdapter
            Dim ds20 As New DataSet
            ds20 = New DataSet
            da20 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao1 & "' WHERE ID = " & Id1 & "", conRNC)
            ds20.Clear()
            da20.Fill(ds20, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 73 " & ex.Message)
        End Try
    End Sub
    Sub Alterar2()
        Alterar1()
        If rbRT2.Checked = True Then
            Disposicao2 = "Retrabalhar"
            alterar2a()
        ElseIf rbLC2.Checked = True Then
            Disposicao2 = "Liberado Condicional"
            alterar2a()
        ElseIf rbRF2.Checked = True Then
            Disposicao2 = "Refugar"
            alterar2a()
        End If
    End Sub
    Sub alterar2a()
        Try
            Dim da20_2 As New OleDbDataAdapter
            Dim ds20_2 As New DataSet
            ds20_2 = New DataSet
            da20_2 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao2 & "' WHERE ID = " & Id2 & "", conRNC)
            ds20_2.Clear()
            da20_2.Fill(ds20_2, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 74 " & ex.Message)
        End Try
    End Sub
    Sub Alterar3()
        Alterar2()
        If rbRT3.Checked = True Then
            Disposicao3 = "Retrabalhar"
            alterar3a()
        ElseIf rbLC3.Checked = True Then
            Disposicao3 = "Liberado Condicional"
            alterar3a()
        ElseIf rbRF3.Checked = True Then
            Disposicao3 = "Refugar"
            alterar3a()
        End If
    End Sub
    Sub alterar3a()
        Try
            Dim da20_3 As New OleDbDataAdapter
            Dim ds20_3 As New DataSet
            ds20_3 = New DataSet
            da20_3 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao3 & "' WHERE ID = " & Id3 & "", conRNC)
            ds20_3.Clear()
            da20_3.Fill(ds20_3, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 75 " & ex.Message)
        End Try
    End Sub
    Sub Alterar4()
        Alterar3()
        If rbRT4.Checked = True Then
            Disposicao4 = "Retrabalhar"
            alterar4a()
        ElseIf rbLC4.Checked = True Then
            Disposicao4 = "Liberado Condicional"
            alterar4a()
        ElseIf rbRF4.Checked = True Then
            Disposicao4 = "Refugar"
            alterar4a()
        End If
    End Sub
    Sub alterar4a()
        Try
            Dim da20_4 As New OleDbDataAdapter
            Dim ds20_4 As New DataSet
            ds20_4 = New DataSet
            da20_4 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao4 & "' WHERE ID = " & Id4 & "", conRNC)
            ds20_4.Clear()
            da20_4.Fill(ds20_4, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 76 " & ex.Message)
        End Try
    End Sub
    Sub Alterar5()
        Alterar4()
        If rbRT5.Checked = True Then
            Disposicao5 = "Retrabalhar"
            alterar5a()
        ElseIf rbLC5.Checked = True Then
            Disposicao5 = "Liberado Condicional"
            alterar5a()
        ElseIf rbRF5.Checked = True Then
            Disposicao5 = "Refugar"
            alterar5a()
        End If
    End Sub
    Sub alterar5a()
        Try
            Dim da20_5 As New OleDbDataAdapter
            Dim ds20_5 As New DataSet
            ds20_5 = New DataSet
            da20_5 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao5 & "' WHERE ID = " & Id5 & "", conRNC)
            ds20_5.Clear()
            da20_5.Fill(ds20_5, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 77 " & ex.Message)
        End Try
    End Sub
    Sub Alterar6()
        Alterar5()
        If rbRT6.Checked = True Then
            Disposicao6 = "Retrabalhar"
            alterar6a()
        ElseIf rbLC6.Checked = True Then
            Disposicao6 = "Liberado Condicional"
            alterar6a()
        ElseIf rbRF6.Checked = True Then
            Disposicao6 = "Refugar"
            alterar6a()
        End If
    End Sub
    Sub alterar6a()
        Try
            Dim da20_6 As New OleDbDataAdapter
            Dim ds20_6 As New DataSet
            ds20_6 = New DataSet
            da20_6 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao6 & "' WHERE ID = " & Id6 & "", conRNC)
            ds20_6.Clear()
            da20_6.Fill(ds20_6, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 78 " & ex.Message)
        End Try
    End Sub
    Sub Alterar7()
        Alterar6()
        If rbRT7.Checked = True Then
            Disposicao7 = "Retrabalhar"
            alterar7a()
        ElseIf rbLC7.Checked = True Then
            Disposicao7 = "Liberado Condicional"
            alterar7a()
        ElseIf rbRF7.Checked = True Then
            Disposicao7 = "Refugar"
            alterar7a()
        End If
    End Sub
    Sub alterar7a()
        Try
            Dim da20_7 As New OleDbDataAdapter
            Dim ds20_7 As New DataSet
            ds20_7 = New DataSet
            da20_7 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao7 & "' WHERE ID = " & Id7 & "", conRNC)
            ds20_7.Clear()
            da20_7.Fill(ds20_7, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 79 " & ex.Message)
        End Try
    End Sub
    Sub Alterar8()
        Alterar7()
        If rbRT8.Checked = True Then
            Disposicao8 = "Retrabalhar"
            alterar8a()
        ElseIf rbLC8.Checked = True Then
            Disposicao8 = "Liberado Condicional"
            alterar8a()
        ElseIf rbRF8.Checked = True Then
            Disposicao8 = "Refugar"
            alterar8a()
        End If
    End Sub
    Sub alterar8a()
        Try
            Dim da20_8 As New OleDbDataAdapter
            Dim ds20_8 As New DataSet
            ds20_8 = New DataSet
            da20_8 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao8 & "' WHERE ID = " & Id8 & "", conRNC)
            ds20_8.Clear()
            da20_8.Fill(ds20_8, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 80 " & ex.Message)
        End Try
    End Sub
    Sub Alterar9()
        Alterar8()
        If rbRT9.Checked = True Then
            Disposicao9 = "Retrabalhar"
            alterar9a()
        ElseIf rbLC9.Checked = True Then
            Disposicao9 = "Liberado Condicional"
            alterar9a()
        ElseIf rbRF9.Checked = True Then
            Disposicao9 = "Refugar"
            alterar9a()
        End If
    End Sub
    Sub alterar9a()
        Try
            Dim da20_9 As New OleDbDataAdapter
            Dim ds20_9 As New DataSet
            ds20_9 = New DataSet
            da20_9 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao9 & "' WHERE ID = " & Id9 & "", conRNC)
            ds20_9.Clear()
            da20_9.Fill(ds20_9, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 81 " & ex.Message)
        End Try
    End Sub
    Sub Alterar10()
        Alterar9()
        If rbRT10.Checked = True Then
            Disposicao10 = "Retrabalhar"
            alterar10a()
        ElseIf rbLC10.Checked = True Then
            Disposicao10 = "Liberado Condicional"
            alterar10a()
        ElseIf rbRF10.Checked = True Then
            Disposicao10 = "Refugar"
            alterar10a()
        End If
    End Sub
    Sub alterar10a()
        Try
            Dim da20_10 As New OleDbDataAdapter
            Dim ds20_10 As New DataSet
            ds20_10 = New DataSet
            da20_10 = New OleDbDataAdapter("UPDATE tblRNC SET  Disposicao = '" & Disposicao10 & "' WHERE ID = " & Id10 & "", conRNC)
            ds20_10.Clear()
            da20_10.Fill(ds20_10, "tblRNC")
        Catch ex As Exception
            MsgBox("Erro 82 " & ex.Message)
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
            MsgBox("Erro 71SP " & ex.Message)
        End Try
    End Sub

    Private Sub btCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancelar.Click
        LimparDisposicao()
    End Sub
End Class