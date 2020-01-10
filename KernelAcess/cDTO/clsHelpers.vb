'Objetivo: oferecer funcionalidades de validação de dados e transformação de valores
'métodos genéricos
Imports System.Reflection
Imports System.Globalization
Imports System.Threading

Public Class clsHelpers
    'Função para validar o preenchimento de campos obrigatórios
    'argForm = nome do formulario
    'strCamposObrigatorios = lista do com o nome dos campos separados por ";"
    'tituloCampos = lista dos titulos dos campos na mesma ordem e separados por ";"
    'validaCamposObrigatorios(Me, "nomeCampo1;nomeCampo2;etc", "TituloCampo1;TituloCampo2;etc")
    Public Function validaCamposObrigatorios(ByVal argForm As Control, ByVal strCamposObrigatorios As String, Optional ByVal tituloCampos As String = "") As Boolean
        Dim nomeCampos As Object
        Dim campos As Object
        Dim valor As Object
        Dim i As Long
        Dim inicio As Long
        Dim fim As Long
        Dim ctrl As String
        'Windows.Forms.Form
        'monta os arrays
        campos = Split(strCamposObrigatorios, ";")
        nomeCampos = Split(tituloCampos, ";")
        'captura o inicio e fim do array
        inicio = LBound(campos)
        fim = UBound(campos)
        i = inicio

        'inicia a validação uma a uma
        For i = inicio To fim
            'captura o nome do tipo de campo
            ctrl = argForm.Controls(campos(i)).GetType.Name
            Select Case ctrl
                'Caso seja ComboBox
                Case "ComboBox"
                    valor = argForm.Controls(campos(i)).Text
                    If String.IsNullOrEmpty(valor) Then
                        MsgBox("Uma opção: " & argForm.Controls(campos(i)).Tag & ". Deve ser selecionada.", MsgBoxStyle.Information, TITULO_ALERTA)
                        'argForm(campos(i)).SetFocus() 'Coloca o cursor no campo
                        argForm.Controls(campos(i)).Focus()
                        argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                        validaCamposObrigatorios = False
                        Exit Function
                    Else
                        'Altera a cor de fundo para branco
                        argForm.Controls(campos(i)).BackColor = System.Drawing.Color.White
                    End If
                    'Caso seja TextBox
                Case "TextBox"
                    valor = argForm.Controls(campos(i)).Text
                    If String.IsNullOrEmpty(valor) Then
                        MsgBox("O Campo: " & argForm.Controls(campos(i)).Tag & ". Deve ser preenchido.", MsgBoxStyle.Information, TITULO_ALERTA)
                        argForm.Controls(campos(i)).Focus()
                        argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                        validaCamposObrigatorios = False
                        Exit Function
                    Else
                        'Altera a cor de fundo para branco
                        argForm.Controls(campos(i)).BackColor = System.Drawing.Color.White
                    End If
                    'Caso seja MaskeCheckBox
                Case "MaskedTextBox"
                    'tira a formatação
                    argForm.Controls(campos(i)).TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
                    'captura o valor sem a mascara
                    valor = argForm.Controls(campos(i)).Text
                    'retorna a formatação
                    argForm.Controls(campos(i)).TextMaskFormat = MaskFormat.IncludePromptAndLiterals
                    If String.IsNullOrEmpty(valor) Then
                        MsgBox("O Campo: " & argForm.Controls(campos(i)).Tag & ". Deve ser preenchido.", MsgBoxStyle.Information, TITULO_ALERTA)
                        argForm.Controls(campos(i)).Focus()
                        argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                        validaCamposObrigatorios = False
                        Exit Function
                    Else
                        'Altera a cor de fundo para branco
                        argForm.Controls(campos(i)).BackColor = System.Drawing.Color.White
                    End If
                    'Caso seja CheckBox (Normalmente esta campo é opcional)
                Case "CheckBox"
                    '    'Caso seja OptionButton
                    'Case "OptionButton"
                    '    valor = argForm.Controls(campos(i)).Text
                    '    If String.IsNullOrEmpty(valor) Then
                    '        MsgBox("Uma opção: " & argForm.Controls(campos(i)).Tag & ". Deve ser selecionada.", MsgBoxStyle.Information, TITULO_ALERTA)
                    '        argForm.Controls(campos(i)).Focus()
                    '        'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    '        validaCamposObrigatorios = False
                    '        Exit Function
                    '    End If

                    '    'Caso seja OptionGroup
                    'Case "OptionGroup"
                    '    valor = argForm.Controls(campos(i)).Text
                    '    If String.IsNullOrEmpty(valor) Then
                    '        MsgBox("Uma opção: " & argForm.Controls(campos(i)).Tag & ". Deve ser selecionada.", MsgBoxStyle.Information, TITULO_ALERTA)
                    '        argForm.Controls(campos(i)).Focus()
                    '        'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    '        validaCamposObrigatorios = False
                    '        Exit Function
                    '    End If

            End Select
        Next i
        validaCamposObrigatorios = True
    End Function
    ''Limpa objetos de um determinado formulário
    'Public Function LimparCampos(ByVal strFrmName As Form) As Boolean
    '    For Each ctl As Control In strFrmName.Controls
    '        If TypeOf (ctl) Is TextBox Then 'TextBox
    '            DirectCast(ctl, TextBox).Text = String.Empty
    '        ElseIf TypeOf (ctl) Is CheckBox Then 'CheckBox
    '            DirectCast(ctl, CheckBox).Checked = False
    '        ElseIf TypeOf (ctl) Is ComboBox Then 'ComboBox
    '            DirectCast(ctl, ComboBox).SelectedValue = 0
    '        ElseIf TypeOf (ctl) Is DataGrid Then 'Datagrid
    '            DirectCast(ctl, DataGrid).DataSource = Nothing
    '        ElseIf ctl.Controls.Count > 0 Then
    '            LimparCampos(ctl)
    '        End If
    '    Next
    '    Return True
    'End Function

    Public Sub LimparCampos(ByRef Tela As Control)
        'Caso ocorra erro, não mostrar o erro, ignorando e indo para á próxima linha
        On Error Resume Next
        'Declaramos uma variavel Campo do tipo Object
        '(Tipo Object porque iremos trabalhar com todos os campos do Form, podendo ser
        '       Label, Button, TextBox, ComboBox e outros)
        Dim Campo As Object
        'Usaremos For Each para passarmos por todos os controls do objeto atual
        For Each Campo In Tela.Controls
            'Verifica se o Campo é um GroupBox, TabPage ou Panel
            'Se for então precisa limpar os campos que estão dentro dele também...
            'Chamaremos novamente a função mas passando por referencia
            '      O GroupBox, TabPage ou Panel atual
            If TypeOf Campo Is System.Windows.Forms.GroupBox Or
                TypeOf Campo Is System.Windows.Forms.TabPage Or
                TypeOf Campo Is System.Windows.Forms.Panel Then
                LimparCampos(Campo)
            ElseIf TypeOf Campo Is System.Windows.Forms.TextBox Then
                Campo.Text = String.Empty 'Verificamos se o campo é uma TextBox se for então devemos limpar o campo
            ElseIf TypeOf Campo Is System.Windows.Forms.ComboBox Then
                'Verificamos se o campo é um ComboBox
                If Campo.DropDownStyle = ComboBoxStyle.DropDownList Then
                    Campo.SelectedIndex = -1 'Se o tipo da ComboBox for DropDownList então devemos deixar sem seleção
                    'ElseIf Campo.DropDownStyle = ComboBoxStyle.DropDown Then
                    'Campo.Text = ""
                Else
                    'Campo.Text = String.Empty
                    Campo.SelectedValue = 0
                End If
            ElseIf TypeOf Campo Is System.Windows.Forms.CheckBox Then
                Campo.Checked = False
            ElseIf TypeOf Campo Is System.Windows.Forms.DataGridView Then
                Campo.DataSource = Nothing
            ElseIf TypeOf Campo Is System.Windows.Forms.RadioButton Then
                Campo.Checked = False
            ElseIf TypeOf Campo Is System.Windows.Forms.MaskedTextBox Then
                Campo.Text = String.Empty
            End If
        Next

    End Sub


    Public Function validaCPF(ByVal argCpf As String) As Boolean
        'Função que verifica a validade de um CPF.
        Dim wSomaDosProdutos
        Dim wResto
        Dim wDigitChk1
        Dim wDigitChk2
        Dim wI
        'Inicia o valor da Soma
        wSomaDosProdutos = 0
        'Para posição I de 1 até 9
        For wI = 1 To 9
            'Soma = Soma + (valor da posição dentro do CPF x (11 - posição))
            wSomaDosProdutos = wSomaDosProdutos + Val(Mid(argCpf, wI, 1)) * (11 - wI)
        Next wI
        'Resto = Soma - ((parte inteira da divisão da Soma por 11) x 11)
        wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
        'Dígito verificador 1 = 0 (se Resto=0 ou 1 ) ou 11 - Resto (nos casos restantes)
        wDigitChk1 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)
        'Reinicia o valor da Soma
        wSomaDosProdutos = 0
        'Para posição I de 1 até 9
        For wI = 1 To 9
            'Soma = Soma + (valor da posição dentro do CPF x (12 - posição))
            wSomaDosProdutos = wSomaDosProdutos + (Val(Mid(argCpf, wI, 1)) * (12 - wI))
        Next wI
        'Soma = Soma (2 x dígito verificador 1)
        wSomaDosProdutos = wSomaDosProdutos + (2 * wDigitChk1)
        'Resto = Soma - ((parte inteira da divisão da Soma por 11) x 11)
        wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
        'Dígito verificador 2 = 0 (se Resto=0 ou 1 ) ou 11 - Resto (nos casos restantes)
        wDigitChk2 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)
        'Se o dígito da posição 10 = Dígito verificador 1 E
        'dígito da posição 11 = Dígito verificador 2 Então
        If Mid(argCpf, 10, 1) = Mid(Trim(Str(wDigitChk1)), 1, 1) And Mid(argCpf, 11, 1) = Mid(Trim(Str(wDigitChk2)), 1, 1) Then
            'CPF válido
            validaCPF = True
        Else
            'CPF inválido
            validaCPF = False
        End If
    End Function

    Public Function validaEmail(ByVal eMail As String) As Boolean
        'Função de validação do formato de um e-mail.

        Dim posicaoA As Integer
        Dim posicaoP As Integer

        'Busca posição do caracter @
        posicaoA = InStr(eMail, "@")
        'Busca a posição do ponto a partir da posição
        'do @ ou então da primeira posição
        posicaoP = InStr(posicaoA Or 1, eMail, ".")

        'Se a posição do @ for menor que 2 OU
        'a posição do ponto for menor que a posição
        'do caracter @
        If posicaoA < 2 Or posicaoP < posicaoA Then
            'Formato de e-mail inválido
            validaEmail = False
        Else
            'Formato de e-mail válido
            validaEmail = True
        End If

    End Function

    Public Function nomeProprio(ByVal argNome As String) As String
        'Função recursiva para converter a primeira letra
        'dos nomes próprios para maiúscula, mantendo os
        'aditivos em caixa baixa.
        Dim sNome As String
        Dim lEspaco As Long
        Dim lTamanho As Long
        'Pega o tamanho do nome
        lTamanho = Len(argNome)
        'Passa tudo para caixa baixa
        argNome = LCase(argNome)
        'Se o nome passado é vazio
        'acaba a função ou a recursão
        'retornando string vazia
        If lTamanho = 0 Then
            nomeProprio = ""
        Else
            'Procura a posição do primeiro espaço
            lEspaco = InStr(argNome, " ")
            'Se não tiver pega a posição da última letra
            If lEspaco = 0 Then lEspaco = lTamanho
            'Pega o primeiro nome da string
            sNome = Left(argNome, lEspaco)
            'Se não for aditivo converte a primeira letra
            If Not InStr("e da das de do dos ", sNome) > 0 Then
                sNome = UCase(Left(sNome, 1)) & LCase(Right(sNome, Len(sNome) - 1))
            End If
            'Monta o nome convertendo o restante através da recursão
            nomeProprio = sNome & nomeProprio(LCase(Trim(Right(argNome, lTamanho - lEspaco))))
        End If
    End Function

    Public Function abreviaNome(ByVal argNome As String) As String
        'Função que abrevia o penúltimo sobrenome, levando
        'em consideração os aditivos de, da, do, dos, das, e.

        'Define variáveis para controle de posição e para as
        'partes do nome que serão separadas e depois unidas
        'novamente.
        Dim ultimoEspaco As Integer, penultimoEspaco As Integer
        Dim primeiraParte As String, ultimaParte As String
        Dim parteNome As String
        Dim tamanho As Integer, i As Integer

        'Tamanho do nome passado
        'no argumento
        tamanho = Len(argNome)

        'Loop que verifica a posição do último e do penúltimo
        'espaços, utilizando apenas um loop.
        For i = tamanho To 1 Step -1
            If Mid(argNome, i, 1) = " " And ultimoEspaco <> 0 Then
                penultimoEspaco = i
                Exit For
            End If
            If Mid(argNome, i, 1) = " " And penultimoEspaco = 0 Then
                ultimoEspaco = i
            End If
        Next i

        'Caso i chegue a zero não podemos
        'abreviar o nome
        If i = 0 Then
            abreviaNome = argNome
            Exit Function
        End If

        'Separação das partes do nome em três: primeira, meio e última
        primeiraParte = Left(argNome, penultimoEspaco - 1)
        parteNome = Mid(argNome, penultimoEspaco + 1, ultimoEspaco - penultimoEspaco - 1)
        ultimaParte = Right(argNome, tamanho - ultimoEspaco)

        'Para a montagem do nome já abreviado verificamos se a parte retirada
        'não é um dos nomes de ligação: de, da ou do. Caso seja usamos o método
        'recursivo para refazer os passos.
        'Caso seja necessário basta acrescentar outros nomes de ligação para serem
        'verificados.
        If parteNome = "da" Or parteNome = "de" Or parteNome = "do" Or parteNome = "dos" Or parteNome = "das" Or parteNome = "e" Then
            abreviaNome = abreviaNome(primeiraParte & " " & parteNome) & " " & ultimaParte
        Else
            abreviaNome = primeiraParte & " " & Left(parteNome, 1) & ". " & ultimaParte
        End If
    End Function
    'Função para abrir uma caixa de seleção de arquivos
    Public Function EnderecoArqCapturar() As String
        Dim open As New OpenFileDialog()
        Try
            If open.ShowDialog = DialogResult.OK Then
                Return open.FileName.ToString
            End If
        Catch ex As Exception
            MsgBox("Não foi possível identificar o endereço do arquivo. Motivo:" & ex.Message, MsgBoxStyle.SystemModal, TITULO_ALERTA)
        End Try
        Return String.Empty
    End Function

    'para informar erros
    Public Sub InformaErro()
        Dim ErrMsg As String
        ErrMsg = "Erro nº " & Err.Number & ": " & vbNewLine & Err.Description & vbNewLine & "** O aplicativo precisa ser encerrado! **"
        If Err.Number <> 0 Then
            Select Case Err.Number
                Case 3024, 3043, 3044, 3265 ' são os códigos de erro quando não há conexão com o banco de dados
                    'Exibe a mensagem do erro
                    'Não registra log, pois o acesso a rede foi interrompido
                    MsgBox(ErrMsg, MsgBoxStyle.Exclamation, TITULO_ALERTA)
                    Application.Exit()
                    End 'interrompe imediatamente
                Case Else
                    'registra um log do erro que aconteceu ao usuário
                    'Call registraLog("Erro nº " & Err.Number & ": " & Err.Description)
                    'Exibe a mensagem do erro
                    MsgBox(ErrMsg, MsgBoxStyle.Exclamation, TITULO_ALERTA)
                    Application.Exit()
                    End 'interrompe imediatamente
            End Select
        End If
    End Sub

    'Função para preencher um combobox
    Public Sub carregaComboBox(strSQL As String, frm As Form, cb As ComboBox)
        Dim con As New clsConexao
        Dim dt As New DataTable
        dt = con.RetornaDataTable(strSQL)
        With frm
            With cb
                .DataSource = dt
                .DisplayMember = dt.Columns(1).ToString
                .ValueMember = dt.Columns(0).ToString
                .Text = Nothing
                'Se houver apenas um item no combobox este já fica selecionado
                If .Items.Count = 1 Then
                    .SelectedIndex = 0
                End If
            End With

        End With
    End Sub

    'Função para limpar as infos de um combobox
    Public Sub limpaCombobox(cb As ComboBox)
        With cb
            .DataSource = Nothing
            .DisplayMember = Nothing
            .Items.Clear()
        End With
    End Sub

    'função para abrir um formulario
    Public Sub abrirForm(frm As Form, Optional janelaRestrita As Boolean = False)
        'registrarLOG(, , GetCurrentMethodName, "Abriu: " & frm.Name.ToString)
        If janelaRestrita Then
            frm.ShowDialog()
        Else
            frm.Show()
        End If
    End Sub

    'função para fechar um formulario
    Public Sub fecharForm(frm As Form)
        'registrarLOG(, , GetCurrentMethodName, "Fechou: " & frm.Name.ToString)
        frm.Close()
    End Sub

    'função para fechar aplicativo
    Public Sub fecharAplicativo()
        'Função para fechar o aplicativo
        If MsgBox("Deseja realmente fechar o aplicativo?", vbQuestion + vbYesNo, TITULO_ALERTA) = vbYes Then
            registrarLOG(, , GetCurrentMethodName, "Fechou Aplicativo")
            Application.Exit()
        End If

    End Sub

    'função para capturar o id de rede
    Public Function capturaIdRede() As String
        capturaIdRede = Environ("USERNAME").ToString
    End Function


    'função para limitar uma quantidade minima e maxima de caracteres.
    'Utilizar no LostFocus
    'txtCartao_Leave(sender As Object, e As EventArgs) Handles txtCartao.Leave
    'hlp.validaTamanhoMinMax(txtCartao, 15, 15)
    Public Function validaTamanhoMinMax(ByVal ctl As Control, iMinLen As Integer, iMaxLen As Integer) As Boolean

        If Not String.IsNullOrEmpty(ctl.Text) Then
            'se diferente de vazio
            If Not String.IsNullOrEmpty(Replace(ctl.Text.Trim, "_", "")) Then
                'Limite Maximo
                If Len(Replace(ctl.Text.Trim, "_", "")) > iMaxLen Then
                    MsgBox("Limite máximo de " & iMaxLen & " caracteres foi excedido." & vbNewLine, vbInformation, TITULO_ALERTA)
                    ctl.Text = Left(ctl.Text.Trim, iMaxLen)
                    ctl.Focus()
                    ctl.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    Return False
                    Exit Function
                End If
                'Limite Minimo
                If Len(Replace(ctl.Text.Trim, "_", "")) < iMinLen Then
                    MsgBox("Limite mínimo de " & iMinLen & " caracteres." & vbNewLine, vbInformation, TITULO_ALERTA)
                    ctl.Text = Left(ctl.Text.Trim, iMinLen)
                    ctl.Focus()
                    ctl.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    Return False
                    Exit Function
                End If
                ctl.BackColor = System.Drawing.Color.White
            Else
                ctl.BackColor = System.Drawing.Color.White
                Return False
                Exit Function
            End If
        End If
        Return True
    End Function

    'verificar se uma data é valida.
    'Utilizar no LostFocus
    Public Function validaData(ByVal Controle As Control) As Boolean
        Dim idiomaPC As String
        Dim formato As String = ""

        'se diferente de vazio
        If Not String.IsNullOrEmpty(Replace(Replace(Controle.Text.Trim, "_", ""), "/", "").Trim) Then
            If Not IsDate(Controle.Text) Then
                'captura o idioma da maquina
                idiomaPC = CultureInfo.CurrentCulture.Name
                If idiomaPC = "pt-BR" Then
                    formato = "dia/mês/ano"
                Else
                    formato = "mês/dia/ano"
                End If
                'para o campo: " & Controle.Tag & "." & vbNewLine &
                MsgBox("Data inválida! " & vbNewLine &
                       "Possíveis motivos: " & vbNewLine &
                       " > Data inexistente." & vbNewLine &
                       " > Utilize o formato: " & idiomaPC.ToUpper & " (" & formato.ToUpper & ").", MsgBoxStyle.Information, TITULO_ALERTA)
                Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                Controle.Focus()
                Return False
            Else
                Controle.BackColor = System.Drawing.Color.White
                Return True
            End If
        Else
            Controle.BackColor = System.Drawing.Color.White
            'Controle.Text = ""
            Return False
        End If
    End Function

    'Funções para formatação de data
    Public Function DataHoraAtual() As DateTime
        Return DateTime.Now
    End Function
    Public Function DataAbreviada() As Date
        Return CDate(DateTime.Now).ToString("yyyy-MM-dd")
    End Function
    Public Function FormataHoraAbreviada(hr As DateTime) As Date
        Return CDate(hr).ToString("HH:mm:ss")
    End Function
    Public Function FormataDataAbreviada(dt As DateTime) As DateTime
        Return CDate(dt).ToString("yyyy-MM-dd")
    End Function
    'Public Function formatarData(data As Date) As Date
    '    Return FormatDateTime(data, DateFormat.ShortDate)
    'End Function
    Public Function convertDatetime(data As Object) As DateTime
        If IsDBNull(data) Then
            Return Nothing
        Else
            Return Convert.ToDateTime(data).ToString
        End If
        'Convert.ToDateTime(argData).ToString("yyyy-MM-dd")
        'Convert.ToDateTime(argData).ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    'função para ajustar um valor em decimal
    Public Function transformarMoeda(txt As String) As String

        Dim n As String = String.Empty
        Dim v As Double = 0
        Try
            'Verificando se o valor contem ',' ou '.' ou ausencia de pontuação
            If InStr(txt, ".") Or InStr(txt, ",") Then
                n = txt.Replace(",", "").Replace(".", "")
                If n.Equals("") Then n = "000"
                If n.Length > 3 And n.Substring(0, 1) = "0" Then n = n.Substring(1, n.Length - 1)
            Else 'Caso não haja pontuação apenas acrescenta 2 zeros para gerar o valor moeda
                n = txt.PadRight(txt.Length + 2, "0")

            End If
            v = Convert.ToDouble(n) / 100
            txt = String.Format("{0:N}", v)
            Return txt

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erro de conversão de valor/moeda")
            Return txt
        End Try
    End Function

    'altera o cursor para load
    Public Sub CursorPointer(bln As Boolean)
        If bln Then
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Else
            Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    Public Sub killSistema()
        'Call mdlSysOffline.colocarOffline()
    End Sub

    'AutoCloseMsgBox "Msgbox1 - Clique em OK ou aguarde 2 segundos", "Fechar MsgBox1 automaticamente", 2 
    '2 segundos
    Sub AutoCloseMsgBox(Mensagem As String, Titulo As String, Segundos As Integer)
        Dim oSHL As Object
        oSHL = CreateObject("WScript.Shell")
        oSHL.PopUp(Mensagem, Segundos, Titulo, vbOKOnly + vbInformation)
    End Sub

    Public Function versaoSistema() As String
        Return Application.ProductVersion
    End Function

    Public Sub registrarLOG(Optional ByVal erroNumero As String = "", Optional ByVal erroDescricao As String = "", Optional ByVal funcaoExecutada As Object = "", Optional ByVal acao As String = "")
        Dim logDTO As New clsLogDTO
        Dim logDAL As New clsLogDAL
        With logDTO
            .data = Now()
            .idUsuario = capturaIdRede()
            .erroDescricao = erroDescricao.ToString
            .erroNumero = erroNumero.ToString
            .funcaoExecutada = funcaoExecutada.ToString
            .versaoSis = versaoSistema()
            .idiomaPC = retornaIdiomaPC()
            .hostname = Environ("COMPUTERNAME")
            .acao = acao.ToString
        End With
        logDAL.Incluir(logDTO)
    End Sub

    'Carrega dataGrid
    Public Sub carregaDataGrid(frm As Form, dg As DataGridView, dt As DataTable)
        Try
            With frm
                With dg
                    .DataSource = dt
                End With
            End With
        Catch ex As Exception
            registrarLOG(Err.Number, Err.Description, GetCurrentMethodName)
        End Try
    End Sub

    Public Sub colarDataGridView(ByVal frm As Form, ByVal dgv As DataGridView)
        Dim dt As New DataTable
        Dim dados() As String
        Dim linhas As Integer = 0
        Dim colunasDGV As Integer = 0
        Dim colunaNome As String = ""
        Dim colunaType As Object
        'Alimentando qtde de colunas existentes do DataGridView
        With frm
            With dgv
                colunasDGV = .Columns.Count
            End With
        End With

        Try
            'Adicionar colunas conforme as existentes no DataGridView
            For i As Integer = 0 To colunasDGV - 1
                With frm
                    With dgv
                        colunaNome = .Columns(i).HeaderText
                        colunaType = .Columns(i).HeaderCell.ValueType
                    End With
                End With
                dt.Columns.Add(colunaNome, colunaType)
            Next

            'Rodando a área de transferência e incluindo em uma nova linha do dataTable
            For Each line As String In Clipboard.GetText.Split(vbNewLine)
                dados = line.Trim.Split(vbTab)


                If dados.Length = colunasDGV Then 'Evitando colocar a última linha do clipboard que é em branco
                    dt.Rows.Add() 'Adicionando nova linha
                    For i As Integer = 0 To colunasDGV - 1 'For das Colunas
                        dt.Rows(linhas).Item(i) = dados(i)
                    Next
                End If
                linhas = linhas + 1 'Próxima linha do DataGridView

            Next

            With frm
                With dgv
                    .DataSource = dt
                End With
            End With

        Catch ex As Exception
            registrarLOG(Err.Number, Err.Description, GetCurrentMethodName)
            MsgBox(Err.Number & " - " & Err.Description, MsgBoxStyle.Information, TITULO_ALERTA)
        End Try

    End Sub

    Public Sub carregaBarraProgresso(ByVal frm As Form, ByVal nomeBarraProgresso As ProgressBar, Optional maximo As Integer = 0, Optional limpeza As Boolean = True, Optional saltoProgresso As Boolean = False)
        With frm
            With nomeBarraProgresso
                If Not saltoProgresso Then
                    .Maximum = maximo
                    .Minimum = 0
                    .Step = 1
                    .Value = 0
                    .Visible = IIf(limpeza, False, True)
                Else
                    .PerformStep()
                    Application.DoEvents()
                End If
            End With
        End With
    End Sub

    'Função para executar um delay variando entre 1 e 5 segundos dependendo da qualidade da rede
    Public Sub DelayMacro()
        Dim numero As Integer
        Dim iCount As Long = 0
        Dim time1, time2
        'Inicializa o gerador de números aleatórios.
        Randomize()
        numero = CInt(Int((3 * Rnd()) + 1)) 'Generate random value between 1 and 3.
        'Debug.Print(Now & " > " & numero) 'Mensagem de hora inicial (teste)
        time1 = Now
        time2 = Now.AddSeconds(numero) 'TimeValue("0:00:0" & numero)
        Do Until time1 >= time2
            'Application.DoEvents
            time1 = Now()
        Loop
        'Debug.Print(Now & " > " & numero)  'Mensagem de hora final (teste)
        Exit Sub
    End Sub

    'A classe StackFrame devolve a pilha de execuções, 
    'algo como a janela Call Stack no Visual Studio. O item 0 é esta própria função, 
    'então precisamos pegar o item 1, quem chamou essa função, e retornar o nome do método. 
    'Assim podemos obter em run-time o nome do método em execução e utilizar esse recurso em padrões de trace para nossos sistemas.
    Public Function GetCurrentMethodName() As String
        Dim stack As New System.Diagnostics.StackFrame(1)
        Return stack.GetMethod().Name
    End Function

    'função para retornar caminho/nome do arquivo onde devemos "salvar Como"
    Public Function SavarComo(Optional ByVal nomeArquivo As String = "") As String
        Dim saveFileDialog1 As New SaveFileDialog()
        With saveFileDialog1
            .Filter = "txt files (*.txt)|*.txt|csv files (*.csv)|*.csv|All files (*.*)|*.*"
            .Title = "Salvar arquivo em..."
            '.InitialDirectory = nomeArquivo
            .FileName = nomeArquivo
            '.ShowDialog()
            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                If .FileName <> "" Then
                    Return .FileName
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End With
    End Function


    'retorna um caminho pessoal no c:
    Public Function retornaDirPessoal() As String
        retornaDirPessoal = "\\" & Environ("COMPUTERNAME") & "\c$\Users\" & Environ("USERNAME") & "\"
    End Function
    'abrir um determinado Arquivo
    Public Sub abrirArquivo(arquivo As String)
        System.Diagnostics.Process.Start(arquivo)
    End Sub
    'Função que verifica se um determinado arquivo esta aberto
    Public Function IsFileOpen(ByVal filename As String) As Boolean

        Dim filenum As Integer, errnum As Integer
        On Error Resume Next   ' Turn error checking off.
        filenum = FreeFile()   ' Get a free file number.
        FileOpen(filenum, filename, OpenMode.Random, OpenAccess.ReadWrite)
        FileClose(filenum)  'close the file.
        errnum = Err.Number 'Save the error number that occurred.
        On Error GoTo 0        'Turn error checking back on.
        ' Check to see which error occurred.
        Select Case errnum
            ' No error occurred.
            ' File is NOT already open by another user.
            Case 0
                Return False
                ' Error number for "Permission Denied."
                ' File is already opened by another user.
            Case 70, 55, 75
                Return True
                ' Another error occurred.
            Case Else
                Error errnum
        End Select

    End Function

    'para copiar para um determinado local
    Public Sub CopiaArquivo(ByVal origem As String, ByVal destino As String, ByVal arquivo As String)
        Dim novoNome As String
        Try
            origem = Replace(origem, arquivo, "") & arquivo
            'atribui um novo nome unico para o arquivo
            novoNome = capturaIdRede() & " " & Format(Now, "ddMMyyyy HHmmss") & "." & PegarExtensao(arquivo)
            destino = destino & novoNome

            'novoNome = capturaIdRede() & " " & Format(Now, "ddMMyyyy HHmmss") & "." & PegarExtensao(arquivo)
            'se o arquivo não existir
            If Len(Dir(destino)) = 0 Then
                'se não existir nenhum arquivo, pode copiar
                Microsoft.VisualBasic.FileCopy(origem, destino)
                'caso existir apaga e depois copia
            Else
                'verifica se o arquivo ja esta em uso
                If IsFileOpen(destino) Then
                    MsgBox("Arquivo já em uso!", vbInformation, TITULO_ALERTA)
                    Exit Sub
                End If
                'apaga o arquivo antigo
                Microsoft.VisualBasic.Kill(destino)
                'copia o novo arquivo
                Microsoft.VisualBasic.FileCopy(origem, destino)
            End If
        Catch ex As Exception
            MsgBox(ex.Message & "Erro nº: " & Err.Number, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Function desacentua(ByVal argTexto As String) As String
        'Função que retira acentos de qualquer texto.
        Dim strAcento As String
        Dim strNormal As String
        Dim strLetra As String
        Dim strNovoTexto As String = ""
        Dim intPosicao As Integer
        Dim i As Integer

        'Informa as duas sequências de caracteres, com e sem acento
        strAcento = "ÃÁÀÂÄÉÈÊËÍÌÎÏÕÓÒÔÖÚÙÛÜÝÇÑãáàâäéèêëíìîïõóòôöúùûüýçñ'"
        strNormal = "AAAAAEEEEIIIIOOOOOUUUUYCNaaaaaeeeeiiiiooooouuuuycn_"

        'Retira os espaços antes e após
        argTexto = Trim(argTexto)
        'Para i de 1 até o tamanho do texto
        For i = 1 To Len(argTexto)
            'Retira a letra da posição atual
            strLetra = Mid(argTexto, i, 1)
            'Busca a posição da letra na sequência com acento
            intPosicao = InStr(1, strAcento, strLetra)
            'Se a posição for maior que zero
            If intPosicao > 0 Then
                'Retira a letra na mesma posição na
                'sequência sem acentos.
                strLetra = Mid(strNormal, intPosicao, 1)
            End If
            'Remonta o novo texto, sem acento
            strNovoTexto = strNovoTexto & strLetra
        Next
        'Devolve o resultado
        desacentua = strNovoTexto
    End Function

    'Limpa objetos de um determinado formulário
    Public Sub CapturaNomeCamposForm(strFrmName As Form)

        Dim txt As New clsEscreveArquivoTxt
        Dim ctrl As Control
        Dim a As String = ""
        Dim arquivo As String
        Dim Nome As String
        Nome = "campos"
        For Each ctrl In strFrmName.Controls
            If TypeOf ctrl Is ComboBox Then
                a = a & ctrl.Name & vbNewLine
            End If
            If TypeOf ctrl Is TextBox Then
                a = a & ctrl.Name & vbNewLine
            End If
            If TypeOf ctrl Is CheckBox Then
                a = a & ctrl.Name & vbNewLine
            End If
            If TypeOf ctrl Is MaskedTextBox Then
                a = a & ctrl.Name & vbNewLine
            End If
            If TypeOf ctrl Is RadioButton Then
                a = a & ctrl.Name & vbNewLine
            End If
        Next
        arquivo = retornaDirPessoal() & Trim(Nome) & ".txt"
        'cria arquivo txt
        With txt
            .CriaArquivo(arquivo)
            .EscreveLn(a)
            .FechaStrm()
        End With
        'abre o arquivo
        abrirArquivo(arquivo)
    End Sub

    Public Function retornaIdiomaPC() As String
        ''Dim culture As CultureInfo = CultureInfo.CurrentCulture
        'Dim a = Thread.CurrentThread.CurrentCulture.Name
        'Dim b = culture.Name
        Return CultureInfo.CurrentCulture.Name.ToUpper.Trim
        Application.DoEvents()
    End Function

    'Metodo para desabilitar o botão "X" fechar
    'Disable the button on the current form:
    'RemoveXButton(Me.Handle())
    Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
    Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&
    Public Function RemoveXButton(ByVal iHWND As Integer) As Integer
        Dim iSysMenu As Integer
        iSysMenu = GetSystemMenu(iHWND, False)
        Return RemoveMenu(iSysMenu, SC_CLOSE, MF_BYCOMMAND)
    End Function

    Public Function somenteNumero(ctrl As Control) As Boolean
        If Not IsNumeric(ctrl.Text.Trim) And Not String.IsNullOrEmpty(ctrl.Text) Then
            MsgBox("Número do " & ctrl.Tag & " inválido.", MsgBoxStyle.Information, TITULO_ALERTA)
            ctrl.Focus()
            ctrl.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
            Return False
            Exit Function
        Else
            ctrl.BackColor = System.Drawing.Color.White
            Return True
        End If
    End Function

    Public Function TotalLinhasArquivoTxt(caminhoArquivo As String) As Long
        'Regular Expression e contar as quebras de linhas:
        Dim re As New System.Text.RegularExpressions.Regex("\r\n")
        Dim sr As New System.IO.StreamReader(caminhoArquivo)
        Dim txt As String = sr.ReadToEnd()
        Dim qtdLinhas As Long = re.Matches(txt).Count + 1
        sr.Close()
        'No final eu somo +1 para contar a última linha, que não tem a quebra de linha que é contada acima.
        Return qtdLinhas
    End Function

    'pegar extensao do arquivo
    Public Function PegarExtensao(arquivo As String) As String
        Dim i As Integer
        Dim j As Integer
        i = InStrRev(arquivo, ".")
        j = InStrRev(arquivo, "\")
        If j = 0 Then j = InStrRev(arquivo, ":")
        'End If
        If j < i Or i > 0 Then
            PegarExtensao = Right(arquivo, (Len(arquivo) - i))
        Else
            Return ""
        End If
    End Function

    'Procura uma determinada palavra em um texto e retorna verdadeiro caso encontre
    Public Function procurarPalavra(texto As String, texto_a_procurar As String) As Boolean
        Dim resultado As Long
        resultado = InStr(texto, texto_a_procurar)
        If resultado > 0 Then
            procurarPalavra = True
        Else : procurarPalavra = False
        End If
    End Function

    'Função para retornar vazio para campos textbox com data
    Public Function RetornaDataTextBox(argValor As DateTime) As String
        Dim idiomaPC As String
        Dim formato As String = ""
        Dim dataVazia As DateTime = Nothing
        If CDate(argValor).ToString("yyyy-MM-dd HH:mm:ss") = CDate(dataVazia).ToString("yyyy-MM-dd HH:mm:ss") Then
            Return ""
        Else

            'captura o idioma da maquina
            idiomaPC = CultureInfo.CurrentCulture.Name
            If idiomaPC = "pt-BR" Then
                formato = "dd/MM/yyyy HH:mm:ss" 'dia/mês/ano
            Else
                formato = "MM/dd/yyyy HH:mm:ss" 'mês/dia/ano"
            End If

            'Dim sFormat As System.Globalization.DateTimeFormatInfo = New System.Globalization.DateTimeFormatInfo()
            'sFormat.LongDatePattern = formato ' "yyyy-MM-DD HH:mm:ss" ' ShortDatePattern
            'Return Format(Convert.ToDateTime(argValor.ToString, sFormat), vbLongDate) 'MM/dd/yyyy HH:mm:ss
            Return Format(argValor, formato)
            'Return Convert.ToDateTime(argValor.ToString, sFormat)

        End If
    End Function

    'Chamada
    'hlp.CarregaComboBoxManualmente("FAB;INQ;Tudo", Me, Me.cFilas)
    'função
    'Carregamento de Combobox de forma manual
    Public Sub CarregaComboBoxManualmente(ByVal strItens As String, ByVal frm As Form, ByVal cb As ComboBox)
        Dim itens As Object
        itens = Split(strItens, ";")
        'limpando o combobox para evitar duplo carregamento
        With frm
            With cb
                .Items.Clear()
            End With
        End With
        'Carregando itens
        For i = LBound(itens) To UBound(itens)
            With frm
                With cb
                    .Items.Add(itens(i))
                End With
            End With
        Next
    End Sub

    Public Sub killProcesso()
        'captura o processo do aplicativo
        Dim proc As Process = Process.GetCurrentProcess
        'captura o nome do processo deste aplicativo
        Dim processo As String = proc.ProcessName.ToString

        'percorrendo todos os processos abertos
        For Each prog As Process In Process.GetProcesses
            'fecha o processo deste aplicativo
            If prog.ProcessName = processo Then
                prog.Kill()
            End If
        Next
    End Sub

End Class