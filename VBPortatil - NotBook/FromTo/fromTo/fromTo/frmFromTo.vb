Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Object
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class frmFronTo
    Dim conConsulta_OP As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb;Jet OLEDB:Database Password= projetornc;")
    Dim ds1FT As New DataSet
    Dim ds2FT As New DataSet
    Dim ds3FT As New DataSet
    Dim daFT As OleDbDataAdapter
    Dim cbFT As OleDbCommandBuilder
    Dim _connFT As String
    Dim da2FT As OleDbDataAdapter
    Dim cb2FT As OleDbCommandBuilder
    Dim AccessFT As Boolean
    Dim ExcelFT As Boolean

    Private Sub frmRNC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Today > "01/06/2015" Then
            MsgBox("Contate o Programador: Cid (15) 981797980 - cidevangelista@hotmail.com")
            Close()
        Else
            Call Teste_AbertoFT()
            Call PriMeiro_Passo()
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
        Catch e1 As Exception
            conConsulta_OP.Close()
            MessageBox.Show("Erro 3!", e1.Message)
        End Try
    End Sub
    Sub Teste_AbertoFT()
        ExcelFT = TestFT("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.xlsx")
        AccessFT = TestFT("f:\Receb.Mat.Prima\Banco_Dados\Consulta_OP.accdb")
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

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Me.Close()
    End Sub
End Class
