<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCadastro
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCadastro))
        Me.txtNome = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtBancoDados = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDiretorio = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtSenha = New System.Windows.Forms.TextBox()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.btnCancelar = New System.Windows.Forms.Button()
        Me.btnExcluir = New System.Windows.Forms.Button()
        Me.btnAlterar = New System.Windows.Forms.Button()
        Me.btnIncluir = New System.Windows.Forms.Button()
        Me.ID = New System.Windows.Forms.Label()
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtSenhaAcesso = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtNome
        '
        Me.txtNome.Location = New System.Drawing.Point(105, 28)
        Me.txtNome.Name = "txtNome"
        Me.txtNome.Size = New System.Drawing.Size(344, 20)
        Me.txtNome.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Aplicativo:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 13)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Banco de Dados:"
        '
        'txtBancoDados
        '
        Me.txtBancoDados.Location = New System.Drawing.Point(105, 48)
        Me.txtBancoDados.Name = "txtBancoDados"
        Me.txtBancoDados.Size = New System.Drawing.Size(344, 20)
        Me.txtBancoDados.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Diretório:"
        '
        'txtDiretorio
        '
        Me.txtDiretorio.Location = New System.Drawing.Point(105, 68)
        Me.txtDiretorio.Name = "txtDiretorio"
        Me.txtDiretorio.Size = New System.Drawing.Size(344, 20)
        Me.txtDiretorio.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 91)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Senha BD:"
        '
        'txtSenha
        '
        Me.txtSenha.Location = New System.Drawing.Point(105, 88)
        Me.txtSenha.Name = "txtSenha"
        Me.txtSenha.Size = New System.Drawing.Size(124, 20)
        Me.txtSenha.TabIndex = 3
        '
        'ListView1
        '
        Me.ListView1.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.Location = New System.Drawing.Point(15, 180)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(582, 193)
        Me.ListView1.TabIndex = 9
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'btnCancelar
        '
        Me.btnCancelar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCancelar.Image = CType(resources.GetObject("btnCancelar.Image"), System.Drawing.Image)
        Me.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCancelar.Location = New System.Drawing.Point(347, 146)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(75, 30)
        Me.btnCancelar.TabIndex = 8
        Me.btnCancelar.Text = "   Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = True
        '
        'btnExcluir
        '
        Me.btnExcluir.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnExcluir.Enabled = False
        Me.btnExcluir.Image = CType(resources.GetObject("btnExcluir.Image"), System.Drawing.Image)
        Me.btnExcluir.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btnExcluir.Location = New System.Drawing.Point(266, 146)
        Me.btnExcluir.Name = "btnExcluir"
        Me.btnExcluir.Size = New System.Drawing.Size(75, 30)
        Me.btnExcluir.TabIndex = 7
        Me.btnExcluir.Text = "Excluir"
        Me.btnExcluir.UseVisualStyleBackColor = True
        '
        'btnAlterar
        '
        Me.btnAlterar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnAlterar.Enabled = False
        Me.btnAlterar.Image = CType(resources.GetObject("btnAlterar.Image"), System.Drawing.Image)
        Me.btnAlterar.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btnAlterar.Location = New System.Drawing.Point(185, 146)
        Me.btnAlterar.Name = "btnAlterar"
        Me.btnAlterar.Size = New System.Drawing.Size(75, 30)
        Me.btnAlterar.TabIndex = 6
        Me.btnAlterar.Text = "Alterar"
        Me.btnAlterar.UseVisualStyleBackColor = True
        '
        'btnIncluir
        '
        Me.btnIncluir.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnIncluir.Image = CType(resources.GetObject("btnIncluir.Image"), System.Drawing.Image)
        Me.btnIncluir.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.btnIncluir.Location = New System.Drawing.Point(104, 146)
        Me.btnIncluir.Name = "btnIncluir"
        Me.btnIncluir.Size = New System.Drawing.Size(75, 30)
        Me.btnIncluir.TabIndex = 5
        Me.btnIncluir.Text = "Incluir"
        Me.btnIncluir.UseVisualStyleBackColor = True
        '
        'ID
        '
        Me.ID.AutoSize = True
        Me.ID.Location = New System.Drawing.Point(12, 11)
        Me.ID.Name = "ID"
        Me.ID.Size = New System.Drawing.Size(18, 13)
        Me.ID.TabIndex = 20
        Me.ID.Text = "ID"
        '
        'txtID
        '
        Me.txtID.Location = New System.Drawing.Point(105, 8)
        Me.txtID.Name = "txtID"
        Me.txtID.ReadOnly = True
        Me.txtID.Size = New System.Drawing.Size(124, 20)
        Me.txtID.TabIndex = 10
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 13)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Senha de acesso:"
        '
        'txtSenhaAcesso
        '
        Me.txtSenhaAcesso.Location = New System.Drawing.Point(105, 108)
        Me.txtSenhaAcesso.Name = "txtSenhaAcesso"
        Me.txtSenhaAcesso.Size = New System.Drawing.Size(124, 20)
        Me.txtSenhaAcesso.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.Color.Maroon
        Me.Label6.Location = New System.Drawing.Point(455, 51)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(101, 13)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "ex: db_nome.accdb"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Maroon
        Me.Label7.Location = New System.Drawing.Point(455, 71)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(103, 13)
        Me.Label7.TabIndex = 24
        Me.Label7.Text = "ex: C:\Caminho_bd\"
        '
        'frmCadastro
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(609, 385)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtSenhaAcesso)
        Me.Controls.Add(Me.ID)
        Me.Controls.Add(Me.txtID)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnExcluir)
        Me.Controls.Add(Me.btnAlterar)
        Me.Controls.Add(Me.btnIncluir)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSenha)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtDiretorio)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtBancoDados)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNome)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCadastro"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Cadastro"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtNome As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtBancoDados As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDiretorio As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSenha As System.Windows.Forms.TextBox
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents btnCancelar As System.Windows.Forms.Button
    Friend WithEvents btnExcluir As System.Windows.Forms.Button
    Friend WithEvents btnAlterar As System.Windows.Forms.Button
    Friend WithEvents btnIncluir As System.Windows.Forms.Button
    Friend WithEvents ID As System.Windows.Forms.Label
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSenhaAcesso As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
End Class
