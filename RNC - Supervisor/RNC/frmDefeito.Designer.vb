<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDefeito
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btExcluir = New System.Windows.Forms.Button()
        Me.btAlterar = New System.Windows.Forms.Button()
        Me.btInserir = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.txtDescricaoRNC1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtCodigoRNC1 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.btCancelar = New System.Windows.Forms.Button()
        Me.btPesquisa = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btExcluir
        '
        Me.btExcluir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btExcluir.Location = New System.Drawing.Point(130, 102)
        Me.btExcluir.Name = "btExcluir"
        Me.btExcluir.Size = New System.Drawing.Size(53, 23)
        Me.btExcluir.TabIndex = 9
        Me.btExcluir.Text = "Excluir"
        Me.btExcluir.UseVisualStyleBackColor = True
        '
        'btAlterar
        '
        Me.btAlterar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btAlterar.Location = New System.Drawing.Point(71, 102)
        Me.btAlterar.Name = "btAlterar"
        Me.btAlterar.Size = New System.Drawing.Size(53, 23)
        Me.btAlterar.TabIndex = 8
        Me.btAlterar.Text = "Alterar"
        Me.btAlterar.UseVisualStyleBackColor = True
        '
        'btInserir
        '
        Me.btInserir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btInserir.Location = New System.Drawing.Point(12, 102)
        Me.btInserir.Name = "btInserir"
        Me.btInserir.Size = New System.Drawing.Size(53, 23)
        Me.btInserir.TabIndex = 7
        Me.btInserir.Text = "Inserir"
        Me.btInserir.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.Location = New System.Drawing.Point(12, 131)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(456, 693)
        Me.DataGridView1.TabIndex = 15
        Me.DataGridView1.VirtualMode = True
        '
        'txtDescricaoRNC1
        '
        Me.txtDescricaoRNC1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDescricaoRNC1.Location = New System.Drawing.Point(125, 65)
        Me.txtDescricaoRNC1.MaxLength = 60
        Me.txtDescricaoRNC1.Name = "txtDescricaoRNC1"
        Me.txtDescricaoRNC1.Size = New System.Drawing.Size(343, 20)
        Me.txtDescricaoRNC1.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label4.Location = New System.Drawing.Point(122, 39)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 13)
        Me.Label4.TabIndex = 1548
        Me.Label4.Text = "Não Conformidade"
        '
        'txtCodigoRNC1
        '
        Me.txtCodigoRNC1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodigoRNC1.Location = New System.Drawing.Point(12, 65)
        Me.txtCodigoRNC1.MaxLength = 3
        Me.txtCodigoRNC1.Name = "txtCodigoRNC1"
        Me.txtCodigoRNC1.Size = New System.Drawing.Size(68, 20)
        Me.txtCodigoRNC1.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label9.Location = New System.Drawing.Point(12, 39)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 13)
        Me.Label9.TabIndex = 1549
        Me.Label9.Text = "Código"
        '
        'btCancelar
        '
        Me.btCancelar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btCancelar.Location = New System.Drawing.Point(279, 102)
        Me.btCancelar.Name = "btCancelar"
        Me.btCancelar.Size = New System.Drawing.Size(57, 23)
        Me.btCancelar.TabIndex = 10
        Me.btCancelar.Text = "Cancelar"
        Me.btCancelar.UseVisualStyleBackColor = True
        '
        'btPesquisa
        '
        Me.btPesquisa.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btPesquisa.Location = New System.Drawing.Point(342, 102)
        Me.btPesquisa.Name = "btPesquisa"
        Me.btPesquisa.Size = New System.Drawing.Size(62, 23)
        Me.btPesquisa.TabIndex = 11
        Me.btPesquisa.Text = "Pesquisar"
        Me.btPesquisa.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(163, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(160, 20)
        Me.Label6.TabIndex = 1591
        Me.Label6.Text = "Cadastro de Defeitos"
        '
        'frmDefeito
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(480, 760)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btPesquisa)
        Me.Controls.Add(Me.btCancelar)
        Me.Controls.Add(Me.txtDescricaoRNC1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtCodigoRNC1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btExcluir)
        Me.Controls.Add(Me.btAlterar)
        Me.Controls.Add(Me.btInserir)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frmDefeito"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cadastro de Defeitos"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btExcluir As System.Windows.Forms.Button
    Friend WithEvents btAlterar As System.Windows.Forms.Button
    Friend WithEvents btInserir As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtDescricaoRNC1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCodigoRNC1 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents btCancelar As System.Windows.Forms.Button
    Friend WithEvents btPesquisa As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
