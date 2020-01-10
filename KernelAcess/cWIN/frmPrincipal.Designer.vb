<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrincipal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPrincipal))
        Me.btnCompactar = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtSenha = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lbInfo = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbAplicativos = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnCompactar
        '
        Me.btnCompactar.Cursor = System.Windows.Forms.Cursors.Hand
        Me.btnCompactar.Location = New System.Drawing.Point(237, 57)
        Me.btnCompactar.Name = "btnCompactar"
        Me.btnCompactar.Size = New System.Drawing.Size(105, 45)
        Me.btnCompactar.TabIndex = 2
        Me.btnCompactar.Text = "Reparar Banco de dados"
        Me.btnCompactar.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(1, -9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(345, 37)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Compactar / Reparar - Banco de Dados"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(11, 86)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Senha:"
        '
        'txtSenha
        '
        Me.txtSenha.Location = New System.Drawing.Point(73, 82)
        Me.txtSenha.Name = "txtSenha"
        Me.txtSenha.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSenha.Size = New System.Drawing.Size(158, 20)
        Me.txtSenha.TabIndex = 1
        Me.txtSenha.Tag = "Senha"
        Me.txtSenha.UseSystemPasswordChar = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label3.Image = CType(resources.GetObject("Label3.Image"), System.Drawing.Image)
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label3.Location = New System.Drawing.Point(2, 201)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "       Fonte"
        '
        'lbInfo
        '
        Me.lbInfo.BackColor = System.Drawing.Color.Transparent
        Me.lbInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbInfo.ForeColor = System.Drawing.Color.Maroon
        Me.lbInfo.Location = New System.Drawing.Point(14, 120)
        Me.lbInfo.Name = "lbInfo"
        Me.lbInfo.Size = New System.Drawing.Size(316, 23)
        Me.lbInfo.TabIndex = 8
        Me.lbInfo.Text = "{INFORMAÇÕES}"
        Me.lbInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lbInfo.Visible = False
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label4.Image = CType(resources.GetObject("Label4.Image"), System.Drawing.Image)
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label4.Location = New System.Drawing.Point(264, 196)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(86, 22)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "   Área restrita"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbAplicativos
        '
        Me.cbAplicativos.DropDownWidth = 200
        Me.cbAplicativos.FormattingEnabled = True
        Me.cbAplicativos.Location = New System.Drawing.Point(73, 57)
        Me.cbAplicativos.Name = "cbAplicativos"
        Me.cbAplicativos.Size = New System.Drawing.Size(158, 21)
        Me.cbAplicativos.TabIndex = 0
        Me.cbAplicativos.Tag = "Aplicativo"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Location = New System.Drawing.Point(11, 61)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Aplicativo:"
        '
        'frmPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(348, 215)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.cbAplicativos)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lbInfo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtSenha)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnCompactar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPrincipal"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Kernel"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCompactar As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSenha As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lbInfo As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbAplicativos As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label

End Class
