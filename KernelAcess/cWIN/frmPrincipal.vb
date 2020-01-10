Public Class frmPrincipal
    Private hlp As New clsHelpers
    Private bll As New clsDadosBLL
    Private camposObrigatorios As String = "cbAplicativos;txtSenha"

    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtSenha.Text = ""
        bll.PreencheComboDados(Me, cbAplicativos)
        cbAplicativos.Text = ""
    End Sub

    Private Sub btnCompactar_Click(sender As Object, e As EventArgs) Handles btnCompactar.Click
        Dim senha_digitada As String
        Dim senha_acesso As String
        Dim id_ferramenta As Integer
        Dim dto As New clsDadosDTO

        'validação de campos obrigatórios
        If Not hlp.validaCamposObrigatorios(Me, camposObrigatorios) Then
            Exit Sub
        End If

        id_ferramenta = cbAplicativos.SelectedValue
        dto = bll.CarregaDadosPorId(id_ferramenta)
        senha_acesso = dto._senha_acesso.Trim
        senha_digitada = txtSenha.Text
        If senha_digitada = senha_acesso Then
            If MsgBox("Tem certeza que deseja reparar a base de dados?", vbQuestion + vbYesNo, TITULO_ALERTA) = vbYes Then
                exibeInfo()
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                compact(dto._senha_bd, dto._caminho_bd, dto._nome_bd)
                ocultaInfo()
                Cursor.Current = System.Windows.Forms.Cursors.Default
            End If
        ElseIf String.IsNullOrEmpty(senha_digitada) Then
            MsgBox("Informe a senha para executar esta rotina.", vbInformation, TITULO_ALERTA)
        Else
            MsgBox("Senha inválida.", vbExclamation, TITULO_ALERTA)
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Try
            Dim camposObrigatorios As String = "cbAplicativos"
            Dim dto As New clsDadosDTO
            Dim id_ferramenta As Integer
            Dim senha_digitada As String
            Dim senha_acesso As String
            'validação de campos obrigatórios
            If Not hlp.validaCamposObrigatorios(Me, camposObrigatorios) Then
                Exit Sub
            End If

            'captura as informações no banco
            id_ferramenta = cbAplicativos.SelectedValue
            dto = bll.CarregaDadosPorId(id_ferramenta)
            senha_acesso = dto._senha_acesso.Trim
            senha_digitada = txtSenha.Text

            If senha_digitada = senha_acesso Then
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                exibeInfo()
                System.Diagnostics.Process.Start(dto._caminho_bd)
                Cursor.Current = System.Windows.Forms.Cursors.Default
                ocultaInfo()
            ElseIf String.IsNullOrEmpty(senha_digitada) Then
                MsgBox("Informe a senha para acessar esta área.", vbInformation, TITULO_ALERTA)
                txtSenha.Focus()
            Else
                MsgBox("Senha inválida.", vbExclamation, TITULO_ALERTA)
                txtSenha.Focus()
            End If
            
        Catch ex As Exception
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            ocultaInfo()
            MsgBox(ex.Message, vbCritical, TITULO_ALERTA)
            Cursor.Current = System.Windows.Forms.Cursors.Default
        End Try

    End Sub
    Private Sub exibeInfo()
        lbInfo.Visible = True
        lbInfo.Text = "Por favor, aguarde..."
    End Sub
    Private Sub ocultaInfo()
        lbInfo.Visible = False
        'lbInfo.Text = "Por favor, aguarde..."
    End Sub
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Dim senha_digitada As String
        senha_digitada = txtSenha.Text
        If senha_digitada = SENHA_SISTEMA Then
            hlp.abrirForm(frmCadastro, True)
        ElseIf String.IsNullOrEmpty(senha_digitada) Then
            MsgBox("Informe a senha para acessar esta área.", vbInformation, TITULO_ALERTA)
        Else
            MsgBox("Senha inválida.", vbExclamation, TITULO_ALERTA)
        End If

    End Sub

    Private Sub cbAplicativos_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbAplicativos.KeyPress
        e.Handled = True
    End Sub

End Class
