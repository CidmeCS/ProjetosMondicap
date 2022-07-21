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
        Me.Label6 = New System.Windows.Forms.Label
        Me.btPesquisa = New System.Windows.Forms.Button
        Me.btCancelar = New System.Windows.Forms.Button
        Me.txtDescricaoRNC1 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCodigoRNC1 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.btExcluir = New System.Windows.Forms.Button
        Me.btAlterar = New System.Windows.Forms.Button
        Me.btInserir = New System.Windows.Forms.Button
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(158, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(160, 20)
        Me.Label6.TabIndex = 1613
        Me.Label6.Text = "Cadastro de Defeitos"
        '
        'btPesquisa
        '
        Me.btPesquisa.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btPesquisa.Location = New System.Drawing.Point(337, 99)
        Me.btPesquisa.Name = "btPesquisa"
        Me.btPesquisa.Size = New System.Drawing.Size(62, 23)
        Me.btPesquisa.TabIndex = 1609
        Me.btPesquisa.Text = "Pesquisar"
        Me.btPesquisa.UseVisualStyleBackColor = True
        '
        'btCancelar
        '
        Me.btCancelar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btCancelar.Location = New System.Drawing.Point(274, 99)
        Me.btCancelar.Name = "btCancelar"
        Me.btCancelar.Size = New System.Drawing.Size(57, 23)
        Me.btCancelar.TabIndex = 1608
        Me.btCancelar.Text = "Cancelar"
        Me.btCancelar.UseVisualStyleBackColor = True
        '
        'txtDescricaoRNC1
        '
        Me.txtDescricaoRNC1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtDescricaoRNC1.Location = New System.Drawing.Point(120, 62)
        Me.txtDescricaoRNC1.MaxLength = 60
        Me.txtDescricaoRNC1.Name = "txtDescricaoRNC1"
        Me.txtDescricaoRNC1.Size = New System.Drawing.Size(343, 20)
        Me.txtDescricaoRNC1.TabIndex = 1604
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label4.Location = New System.Drawing.Point(117, 36)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 13)
        Me.Label4.TabIndex = 1611
        Me.Label4.Text = "Não Conformidade"
        '
        'txtCodigoRNC1
        '
        Me.txtCodigoRNC1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCodigoRNC1.Location = New System.Drawing.Point(7, 62)
        Me.txtCodigoRNC1.MaxLength = 3
        Me.txtCodigoRNC1.Name = "txtCodigoRNC1"
        Me.txtCodigoRNC1.Size = New System.Drawing.Size(68, 20)
        Me.txtCodigoRNC1.TabIndex = 1603
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label9.Location = New System.Drawing.Point(7, 36)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 13)
        Me.Label9.TabIndex = 1612
        Me.Label9.Text = "Código"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.Location = New System.Drawing.Point(7, 128)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(456, 693)
        Me.DataGridView1.TabIndex = 1610
        Me.DataGridView1.VirtualMode = True
        '
        'btExcluir
        '
        Me.btExcluir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btExcluir.Location = New System.Drawing.Point(125, 99)
        Me.btExcluir.Name = "btExcluir"
        Me.btExcluir.Size = New System.Drawing.Size(53, 23)
        Me.btExcluir.TabIndex = 1607
        Me.btExcluir.Text = "Excluir"
        Me.btExcluir.UseVisualStyleBackColor = True
        '
        'btAlterar
        '
        Me.btAlterar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btAlterar.Location = New System.Drawing.Point(66, 99)
        Me.btAlterar.Name = "btAlterar"
        Me.btAlterar.Size = New System.Drawing.Size(53, 23)
        Me.btAlterar.TabIndex = 1606
        Me.btAlterar.Text = "Alterar"
        Me.btAlterar.UseVisualStyleBackColor = True
        '
        'btInserir
        '
        Me.btInserir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btInserir.Location = New System.Drawing.Point(7, 99)
        Me.btInserir.Name = "btInserir"
        Me.btInserir.Size = New System.Drawing.Size(53, 23)
        Me.btInserir.TabIndex = 1605
        Me.btInserir.Text = "Inserir"
        Me.btInserir.UseVisualStyleBackColor = True
        '
        'frmDefeito
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(470, 826)
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
        Me.MaximumSize = New System.Drawing.Size(486, 864)
        Me.MinimumSize = New System.Drawing.Size(486, 864)
        Me.Name = "frmDefeito"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cadatro de Defeitos "
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btPesquisa As System.Windows.Forms.Button
    Friend WithEvents btCancelar As System.Windows.Forms.Button
    Friend WithEvents txtDescricaoRNC1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCodigoRNC1 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btExcluir As System.Windows.Forms.Button
    Friend WithEvents btAlterar As System.Windows.Forms.Button
    Friend WithEvents btInserir As System.Windows.Forms.Button
End Class
