Public Class frmCadastro
    Private help As New clsHelpers
    Private objDados As New clsDadosDTO
    Private bll As New clsDadosBLL
    Private hlp As New clsHelpers
    Private filtro As String = ""
    Private camposObrigatorios As String = "txtNome;txtBancoDados;txtDiretorio;txtSenha;txtSenhaAcesso"
    Private Sub frmCadastro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        hlp.LimparCampos(Me)
        bll.AtualizarListaDados()
    End Sub

    Private Sub btnIncluir_Click(sender As Object, e As EventArgs) Handles btnIncluir.Click
        If help.validaCamposObrigatorios(Me, camposObrigatorios) Then
            With objDados
                ._aplicativo = Me.txtNome.Text.Trim
                ._nome_bd = Me.txtBancoDados.Text.Trim
                ._caminho_bd = Me.txtDiretorio.Text.Trim
                ._senha_bd = Me.txtSenha.Text.Trim
                ._senha_acesso = Me.txtSenhaAcesso.Text.Trim
                .Acao = FlagAcao.Insert
            End With
            With bll 'Salvando o registro e enviando msg para usuário
                If .SalvarDados(objDados) Then
                    help.LimparCampos(Me)
                    bll.AtualizarListaDados()
                    MsgBox("Registro salvo com sucesso!", vbInformation, TITULO_ALERTA)
                End If
            End With
        End If
    End Sub

    Private Sub btnAlterar_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        If help.validaCamposObrigatorios(Me, camposObrigatorios) Then
            With objDados
                ._aplicativo = Me.txtNome.Text.Trim
                ._nome_bd = Me.txtBancoDados.Text.Trim
                ._caminho_bd = Me.txtDiretorio.Text.Trim
                ._senha_bd = Me.txtSenha.Text.Trim
                ._senha_acesso = Me.txtSenhaAcesso.Text.Trim
                .Acao = FlagAcao.Update
            End With
            With bll 'Salvando o registro e enviando msg para usuário
                If .SalvarDados(objDados) Then
                    bll.AtualizarListaDados()
                    help.LimparCampos(Me)
                    bloqueiaBotoes()
                    MsgBox("Registro alterado com sucesso!", vbInformation, TITULO_ALERTA)
                End If
            End With
        End If
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Dim id As String = Me.ListView1.SelectedItems(0).SubItems(0).Text
        If String.IsNullOrEmpty(id) Or id = 0 Then
            MsgBox("Nenhum registro foi selecionado!", MsgBoxStyle.Information, TITULO_ALERTA)
            Exit Sub 'Clique no listView vazio
        End If
        objDados = bll.CarregaDadosPorId(id)
        With objDados
            Me.txtID.Text = .ID
            Me.txtNome.Text = ._aplicativo
            Me.txtBancoDados.Text = ._nome_bd
            Me.txtDiretorio.Text = ._caminho_bd
            Me.txtSenha.Text = ._senha_bd
            Me.txtSenhaAcesso.Text = ._senha_acesso
        End With
        liberaBotoes()
    End Sub

    Private Sub btnExcluir_Click(sender As Object, e As EventArgs) Handles btnExcluir.Click
        If String.IsNullOrEmpty(Me.txtID.Text) Then
            MsgBox("Nenhum registro foi selecionado!", vbInformation, TITULO_ALERTA)
            Exit Sub
        End If
        If MsgBox("Tem certeza que deseja remover " & Me.txtNome.Text.Trim.ToUpper & " do sistema?", vbQuestion + vbYesNo, TITULO_ALERTA) = vbYes Then
            With bll
                .Deletar(Me.txtID.Text)
                .AtualizarListaDados()
            End With
            help.LimparCampos(Me)
            bloqueiaBotoes()
        End If
    End Sub

    Private Sub liberaBotoes()
        Me.btnAlterar.Enabled = True
        Me.btnExcluir.Enabled = True
        Me.btnIncluir.Enabled = False
        Me.btnCancelar.Enabled = True
    End Sub

    Private Sub bloqueiaBotoes()
        Me.btnIncluir.Enabled = True
        Me.btnAlterar.Enabled = False
        Me.btnExcluir.Enabled = False
    End Sub

    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        help.LimparCampos(Me)
        bloqueiaBotoes()
    End Sub

    Private Sub frmCadastro_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        bll.PreencheComboDados(frmPrincipal, frmPrincipal.cbAplicativos)
    End Sub
End Class