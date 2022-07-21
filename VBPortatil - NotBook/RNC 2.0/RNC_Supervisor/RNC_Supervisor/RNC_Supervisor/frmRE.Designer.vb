<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRE
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
        Me.txtSetor = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblID = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.btPesquisa = New System.Windows.Forms.Button
        Me.btCancelar = New System.Windows.Forms.Button
        Me.txtInspetor = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtRE = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.btExcluir = New System.Windows.Forms.Button
        Me.btAlterar = New System.Windows.Forms.Button
        Me.btInserir = New System.Windows.Forms.Button
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtSetor
        '
        Me.txtSetor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtSetor.Location = New System.Drawing.Point(426, 62)
        Me.txtSetor.MaxLength = 15
        Me.txtSetor.Name = "txtSetor"
        Me.txtSetor.Size = New System.Drawing.Size(112, 20)
        Me.txtSetor.TabIndex = 16141
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label2.Location = New System.Drawing.Point(426, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 13)
        Me.Label2.TabIndex = 16138
        Me.Label2.Text = "Setor"
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.lblID.Location = New System.Drawing.Point(582, 62)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(13, 13)
        Me.lblID.TabIndex = 16137
        Me.lblID.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label1.Location = New System.Drawing.Point(556, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(18, 13)
        Me.Label1.TabIndex = 16136
        Me.Label1.Text = "ID"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(163, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(197, 20)
        Me.Label6.TabIndex = 16135
        Me.Label6.Text = "Cadastro de RE x Inspetor"
        '
        'btPesquisa
        '
        Me.btPesquisa.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btPesquisa.Location = New System.Drawing.Point(342, 99)
        Me.btPesquisa.Name = "btPesquisa"
        Me.btPesquisa.Size = New System.Drawing.Size(62, 23)
        Me.btPesquisa.TabIndex = 16131
        Me.btPesquisa.Text = "Pesquisar"
        Me.btPesquisa.UseVisualStyleBackColor = True
        '
        'btCancelar
        '
        Me.btCancelar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btCancelar.Location = New System.Drawing.Point(279, 99)
        Me.btCancelar.Name = "btCancelar"
        Me.btCancelar.Size = New System.Drawing.Size(57, 23)
        Me.btCancelar.TabIndex = 16130
        Me.btCancelar.Text = "Cancelar"
        Me.btCancelar.UseVisualStyleBackColor = True
        '
        'txtInspetor
        '
        Me.txtInspetor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtInspetor.Location = New System.Drawing.Point(101, 62)
        Me.txtInspetor.MaxLength = 50
        Me.txtInspetor.Name = "txtInspetor"
        Me.txtInspetor.Size = New System.Drawing.Size(303, 20)
        Me.txtInspetor.TabIndex = 16140
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label4.Location = New System.Drawing.Point(98, 36)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 16133
        Me.Label4.Text = "Inspetor"
        '
        'txtRE
        '
        Me.txtRE.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtRE.Location = New System.Drawing.Point(12, 62)
        Me.txtRE.MaxLength = 4
        Me.txtRE.Name = "txtRE"
        Me.txtRE.Size = New System.Drawing.Size(68, 20)
        Me.txtRE.TabIndex = 16139
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.Label9.Location = New System.Drawing.Point(12, 36)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(22, 13)
        Me.Label9.TabIndex = 16134
        Me.Label9.Text = "RE"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        Me.DataGridView1.ColumnHeadersHeight = 25
        Me.DataGridView1.Location = New System.Drawing.Point(12, 130)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(583, 690)
        Me.DataGridView1.TabIndex = 16132
        Me.DataGridView1.VirtualMode = True
        '
        'btExcluir
        '
        Me.btExcluir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btExcluir.Location = New System.Drawing.Point(130, 99)
        Me.btExcluir.Name = "btExcluir"
        Me.btExcluir.Size = New System.Drawing.Size(53, 23)
        Me.btExcluir.TabIndex = 16129
        Me.btExcluir.Text = "Excluir"
        Me.btExcluir.UseVisualStyleBackColor = True
        '
        'btAlterar
        '
        Me.btAlterar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btAlterar.Location = New System.Drawing.Point(71, 99)
        Me.btAlterar.Name = "btAlterar"
        Me.btAlterar.Size = New System.Drawing.Size(53, 23)
        Me.btAlterar.TabIndex = 16128
        Me.btAlterar.Text = "Alterar"
        Me.btAlterar.UseVisualStyleBackColor = True
        '
        'btInserir
        '
        Me.btInserir.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btInserir.Location = New System.Drawing.Point(12, 99)
        Me.btInserir.Name = "btInserir"
        Me.btInserir.Size = New System.Drawing.Size(53, 23)
        Me.btInserir.TabIndex = 16127
        Me.btInserir.Text = "Inserir"
        Me.btInserir.UseVisualStyleBackColor = True
        '
        'frmRE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(607, 827)
        Me.Controls.Add(Me.txtSetor)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btPesquisa)
        Me.Controls.Add(Me.btCancelar)
        Me.Controls.Add(Me.txtInspetor)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtRE)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btExcluir)
        Me.Controls.Add(Me.btAlterar)
        Me.Controls.Add(Me.btInserir)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximumSize = New System.Drawing.Size(613, 855)
        Me.MinimumSize = New System.Drawing.Size(613, 855)
        Me.Name = "frmRE"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cadatro de RE´s"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSetor As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btPesquisa As System.Windows.Forms.Button
    Friend WithEvents btCancelar As System.Windows.Forms.Button
    Friend WithEvents txtInspetor As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtRE As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btExcluir As System.Windows.Forms.Button
    Friend WithEvents btAlterar As System.Windows.Forms.Button
    Friend WithEvents btInserir As System.Windows.Forms.Button
End Class
