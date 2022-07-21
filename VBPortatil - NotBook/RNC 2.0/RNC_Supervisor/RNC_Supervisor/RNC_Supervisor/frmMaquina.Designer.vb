<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMaquina
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
        Me.rbCelula = New System.Windows.Forms.RadioButton
        Me.rbMaquina = New System.Windows.Forms.RadioButton
        Me.lblID = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.btPesquisa = New System.Windows.Forms.Button
        Me.btCancelar = New System.Windows.Forms.Button
        Me.txtCelula = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtMaquina = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.btExcluir = New System.Windows.Forms.Button
        Me.btAlterar = New System.Windows.Forms.Button
        Me.btInserir = New System.Windows.Forms.Button
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rbCelula
        '
        Me.rbCelula.AutoSize = True
        Me.rbCelula.Location = New System.Drawing.Point(330, 111)
        Me.rbCelula.Name = "rbCelula"
        Me.rbCelula.Size = New System.Drawing.Size(54, 17)
        Me.rbCelula.TabIndex = 1636
        Me.rbCelula.TabStop = True
        Me.rbCelula.Text = "Célula"
        Me.rbCelula.UseVisualStyleBackColor = True
        '
        'rbMaquina
        '
        Me.rbMaquina.AutoSize = True
        Me.rbMaquina.Location = New System.Drawing.Point(330, 94)
        Me.rbMaquina.Name = "rbMaquina"
        Me.rbMaquina.Size = New System.Drawing.Size(66, 17)
        Me.rbMaquina.TabIndex = 1635
        Me.rbMaquina.TabStop = True
        Me.rbMaquina.Text = "Máquina"
        Me.rbMaquina.UseVisualStyleBackColor = True
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblID.Location = New System.Drawing.Point(398, 64)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(13, 13)
        Me.lblID.TabIndex = 1634
        Me.lblID.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label1.Location = New System.Drawing.Point(372, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 13)
        Me.Label1.TabIndex = 1633
        Me.Label1.Text = "ID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(158, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(220, 20)
        Me.Label6.TabIndex = 1632
        Me.Label6.Text = "Cadastro de Máquina x Célula"
        '
        'btPesquisa
        '
        Me.btPesquisa.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btPesquisa.Location = New System.Drawing.Point(401, 101)
        Me.btPesquisa.Name = "btPesquisa"
        Me.btPesquisa.Size = New System.Drawing.Size(62, 23)
        Me.btPesquisa.TabIndex = 1626
        Me.btPesquisa.Text = "Pesquisar"
        Me.btPesquisa.UseVisualStyleBackColor = True
        '
        'btCancelar
        '
        Me.btCancelar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btCancelar.Location = New System.Drawing.Point(238, 101)
        Me.btCancelar.Name = "btCancelar"
        Me.btCancelar.Size = New System.Drawing.Size(57, 23)
        Me.btCancelar.TabIndex = 1625
        Me.btCancelar.Text = "Cancelar"
        Me.btCancelar.UseVisualStyleBackColor = True
        '
        'txtCelula
        '
        Me.txtCelula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtCelula.Location = New System.Drawing.Point(178, 64)
        Me.txtCelula.MaxLength = 15
        Me.txtCelula.Name = "txtCelula"
        Me.txtCelula.Size = New System.Drawing.Size(147, 20)
        Me.txtCelula.TabIndex = 1629
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label4.Location = New System.Drawing.Point(175, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(36, 13)
        Me.Label4.TabIndex = 1630
        Me.Label4.Text = "Célula"
        '
        'txtMaquina
        '
        Me.txtMaquina.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtMaquina.Location = New System.Drawing.Point(7, 64)
        Me.txtMaquina.MaxLength = 12
        Me.txtMaquina.Name = "txtMaquina"
        Me.txtMaquina.Size = New System.Drawing.Size(147, 20)
        Me.txtMaquina.TabIndex = 1628
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label9.Location = New System.Drawing.Point(7, 38)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 13)
        Me.Label9.TabIndex = 1631
        Me.Label9.Text = "Máquina"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.Location = New System.Drawing.Point(7, 139)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(456, 680)
        Me.DataGridView1.TabIndex = 1627
        Me.DataGridView1.VirtualMode = True
        '
        'btExcluir
        '
        Me.btExcluir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btExcluir.Location = New System.Drawing.Point(125, 101)
        Me.btExcluir.Name = "btExcluir"
        Me.btExcluir.Size = New System.Drawing.Size(53, 23)
        Me.btExcluir.TabIndex = 1624
        Me.btExcluir.Text = "Excluir"
        Me.btExcluir.UseVisualStyleBackColor = True
        '
        'btAlterar
        '
        Me.btAlterar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btAlterar.Location = New System.Drawing.Point(66, 101)
        Me.btAlterar.Name = "btAlterar"
        Me.btAlterar.Size = New System.Drawing.Size(53, 23)
        Me.btAlterar.TabIndex = 1623
        Me.btAlterar.Text = "Alterar"
        Me.btAlterar.UseVisualStyleBackColor = True
        '
        'btInserir
        '
        Me.btInserir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btInserir.Location = New System.Drawing.Point(7, 101)
        Me.btInserir.Name = "btInserir"
        Me.btInserir.Size = New System.Drawing.Size(53, 23)
        Me.btInserir.TabIndex = 1622
        Me.btInserir.Text = "Inserir"
        Me.btInserir.UseVisualStyleBackColor = True
        '
        'frmMaquina
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(471, 827)
        Me.Controls.Add(Me.rbCelula)
        Me.Controls.Add(Me.rbMaquina)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btPesquisa)
        Me.Controls.Add(Me.btCancelar)
        Me.Controls.Add(Me.txtCelula)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtMaquina)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btExcluir)
        Me.Controls.Add(Me.btAlterar)
        Me.Controls.Add(Me.btInserir)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximumSize = New System.Drawing.Size(487, 865)
        Me.MinimumSize = New System.Drawing.Size(487, 865)
        Me.Name = "frmMaquina"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cadastro de Máquinas"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rbCelula As System.Windows.Forms.RadioButton
    Friend WithEvents rbMaquina As System.Windows.Forms.RadioButton
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btPesquisa As System.Windows.Forms.Button
    Friend WithEvents btCancelar As System.Windows.Forms.Button
    Friend WithEvents txtCelula As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtMaquina As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btExcluir As System.Windows.Forms.Button
    Friend WithEvents btAlterar As System.Windows.Forms.Button
    Friend WithEvents btInserir As System.Windows.Forms.Button
End Class
