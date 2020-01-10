'Solução - Framework
'Projeto Windows Forms - Camada de Apresentação - CamadaWin
'Projeto Class Library - Camada de Negócios - CamadaBLL
'Projeto Class Library - Camda de Acesso aos Dados - CamadaDAL
'Projeto Class Library - Camda de Transferência de dados - CamadaDTO - Considerada uma camada auxiliar onde iremos declarar as nossas classes.

'Dependências:
'A camada de apresentação - projeto CamadaWin - deverá possuir uma dependência para o projeto CamadaBLL que é a nossa camada de negócios;
'A camada de negócios - projeto CamadaBLL - deverá possuir uma dependência para o projeto CamadaDAL que é nossa camada de acesso aos dados;
'A camada de transferência de dados - DTO - Data Transfer Object - que deverá ser visto pelos demais projetos;
'Assim teremos a seguinte hierarquia: CamadaWin (CamadaDTO) => CamadaBLL (CamadaDTO) => CamadaDAL => CamadaDTO
Imports System.Data
Imports System.Data.OleDb
Imports ADODB

Public Class clsConexao

    Private conexao As New ADODB.Connection 'conexao usando ADO
    Private cmd As New ADODB.Command 'Command
    Private isTran As Boolean
    Private password_bd As String = BD_PWD
    Private diretorio_bd As String = BD_PATH
    Private banco_dados As String = BD_NOME

    Public Function GetStringConexao() As String
        Dim strconexao As String
        'String de conexao com banco de dados MSACCESS com senha
        strconexao = "Provider=Microsoft.ACE.OLEDB.12.0;"
        'strconexao = "Provider=Microsoft.JET.OLEDB.4.0;"
        strconexao = strconexao & "Data Source=" & diretorio_bd & banco_dados & ";"
        strconexao = strconexao & "Jet OLEDB:Database Password= " & password_bd & ";"
        strconexao = strconexao & "Persist Security Info=False"
        Return strconexao
    End Function

    'Metodo para efetuar uma conexão
    'Optional ByVal senha As Boolean = False
    Public Function Conectar() As Boolean
        Dim bln As Boolean
        Try
            If conexao.State = ConnectionState.Closed Then
                With conexao
                    .ConnectionString = GetStringConexao()
                    .Mode = ConnectModeEnum.adModeReadWrite 'modo de conexao leitura e escrita
                    .Open()

                End With
            End If
            bln = True

        Catch ex As Exception 'ex As OleDbException

            If MsgBox("Erro ao efetuar a conexão com a base de dados : " & ex.Message, vbCritical + MsgBoxStyle.RetryCancel, TITULO_ALERTA) = MsgBoxResult.Retry Then
                Conectar()
            Else
                conexao = Nothing
                bln = False
                Application.Exit()
            End If

            'MsgBox("Erro ao efetuar a conexão com a base de dados : " & ex.Message & Err.Number, vbCritical, TITULO_ALERTA)
            'Application.Exit()
        End Try
        Return bln
    End Function

    ' Procedimento para desconectar do banco de dados.
    Public Sub Desconectar()
        Try
            If Not conexao Is Nothing Then
                If Not conexao.State = ConnectionState.Closed Then
                    conexao.Close()
                End If
            End If
        Catch ex As Exception
            MsgBox("Erro ao desconectar com a base de dados : " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    ' Procedimento para testar conexão com o banco de dados.
    Public Sub Testaconexao()
        Try
            Conectar()
            MsgBox("Conexão realizada com sucesso!!!")
        Catch ex As Exception
            MsgBox("Não foi possível conectar o banco de dados." & ex.Message, vbCritical, TITULO_ALERTA)
            conexao = Nothing
        End Try
        Desconectar()
    End Sub

    'Executa um comando SQL e retorna um boleano
    Public Function ExecutaQuery(ByVal strSql As String) As Boolean
        Try
            'verifica se a conexao esta fechada
            If conexao.State = ConnectionState.Closed Then
                Conectar()
            End If
            'executa a consulta
            With conexao
                .Execute(strSql)
            End With
            'retorna verdadeiro
            ExecutaQuery = True
            Desconectar()
        Catch ex As Exception
            Desconectar()
            MsgBox("Erro ao efetuar a conexão com a base de dados em ExecutaQuery: " & vbNewLine & _
                   "(" & Err.Number & " " & ex.Message & ")" & vbNewLine & _
                   "O aplicativo precisa ser encerrado.", vbCritical, TITULO_ALERTA)
            Application.Exit()
            Return False
        End Try
    End Function

    Public Function RetornaDataTable(ByVal strSQL As String) As DataTable
        Dim objDA As New OleDbDataAdapter
        Dim objDT As New DataTable
        Dim rsObjt As ADODB.Recordset
        Try
            If conexao.State = ConnectionState.Closed Then
                Conectar()
            End If
            rsObjt = RetornaRs(strSQL)
            objDT = RecordSetToDataTable(rsObjt)
            Desconectar()
        Catch ex As Exception
            Desconectar()
            MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
            Application.Exit()
        End Try
        RetornaDataTable = objDT
    End Function

    'Retorna um recordset
    Public Function RetornaRs(ByVal strSQL As String) As ADODB.Recordset
        Dim ADORecordset As New ADODB.Recordset
        If conexao.State = ConnectionState.Closed Then
            Conectar()
        End If
        With ADORecordset
            .CursorLocation = CursorLocationEnum.adUseClient
        End With
        ADORecordset.Open(strSQL, conexao, CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        RetornaRs = ADORecordset
        ADORecordset = Nothing
    End Function
    'Procedimento para retornar um Objeto do tipo DataTable através de um recordset
    Public Function RecordSetToDataTable(ByVal objRS As ADODB.Recordset) As DataTable
        Dim objDA As New OleDbDataAdapter()
        Dim objDT As New DataTable()
        objDA.Fill(objDT, objRS)
        Desconectar()
        Return objDT
    End Function

    'Inicia uma transação;
    Public Sub BeginTransaction()
        Try
            If conexao.State = ConnectionState.Closed Then
                Conectar()
            End If
            conexao.BeginTransaction()

        Catch ex As Exception
            MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    'Faz um commit na transação;
    Public Sub CommitTransaction()
        Try
            conexao.CommitTrans()
            conexao.Close()
        Catch ex As Exception
            MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    'Cancela a transação
    Public Sub RollBackTransaction()
        Try
            conexao.RollbackTrans()
            conexao.Close()
        Catch ex As Exception
            MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Function logicoSql(ByVal argValor As Boolean) As String
        'Função que troca os valores lógicos Verdadeiro/Falso
        'para True/False para utilização em consultas SQL

        'Se o valor for verdadeiro
        If argValor Then
            'Troca por True
            logicoSql = "True"
        Else
            'Senão troca por False
            logicoSql = "False"
        End If

    End Function

    Public Function pontoVirgula(ByVal varValor As Object) As String
        'Função que troca a vírgula de um valor decimal por
        'um ponto para utilização em consultas SQL

        Dim strValor As String
        Dim strInteiro As String
        Dim strDecimal As String
        Dim intPosicao As Integer

        'Converte o valor em string
        strValor = CStr(varValor)

        'Busca a posição da vírgula
        intPosicao = InStr(strValor, ",")

        'Se há uma vírgula em alguma posição
        If intPosicao > 0 Then
            'Retira a parte inteira
            strInteiro = Left(strValor, intPosicao - 1)
            'Retira a parte decimal
            strDecimal = Right(strValor, Len(strValor) - intPosicao)
            'Junta os dois novamente incluindo
            'agora o ponto no lugar da vírgula
            pontoVirgula = strInteiro & "." & strDecimal
        Else
            'Senão devolve o mesmo valor
            pontoVirgula = strValor
        End If

    End Function

    Public Function HoraSql(ByVal argData As DateTime) As String
        'Função que formata uma data para o modo SQL
        'com a cerquilha: #YYYY-MM-DD HH:MM:SS#
        'sempre retorna uma string
        Dim strDataCompleta As String
        'Remonta no formato adequado (Padrão banco de dados)
        strDataCompleta = CDate(argData).ToString("HH:mm:ss")
        HoraSql = "#" & strDataCompleta & "#"
    End Function

    Public Function dataSql(ByVal argData As DateTime) As String
        'Função que formata uma data para o modo SQL
        'com a cerquilha: #YYYY-MM-DD HH:MM:SS#
        'sempre retorna uma string
        Dim strDataCompleta As String
        'Remonta no formato adequado (Padrão banco de dados)
        strDataCompleta = CDate(argData).ToString("yyyy-MM-dd HH:mm:ss")
        dataSql = "#" & strDataCompleta & "#"
    End Function

    Public Function dataSqlAbreviada(ByVal argData As DateTime) As String
        'Função que formata uma data para o modo SQL
        'com a cerquilha: #YYYY-MM-DD HH:MM:SS#
        'sempre retorna uma string
        Dim strDataCompleta As String
        'Remonta no formato adequado (Padrão banco de dados)
        strDataCompleta = CDate(argData).ToString("yyyy-MM-dd")
        dataSqlAbreviada = "#" & strDataCompleta & "#"
    End Function


    Public Function valorSql(ByVal argValor As Object) As String
        'Função que formata valores para utilização
        'em consultas SQL
        valorSql = Nothing

        If argValor = Nothing Then
            valorSql = "Null"
        End If
        'Seleciona o tipo de valor informado
        Select Case VarType(argValor)
            'Caso seja vazio ou nulo apenas
            'devolve a string Null
            Case vbEmpty, vbNull
                valorSql = "Null"
                'Caso seja inteiro ou longo apenas
                'converte em string
            Case vbInteger, vbLong
                valorSql = CStr(argValor)
                'Caso seja simples, duplo, decimal ou moeda
                'substitui a vírgula por ponto
            Case vbSingle, vbDouble, vbDecimal, vbCurrency
                valorSql = pontoVirgula(argValor)
                'Caso seja data chama a função dataSql()
            Case vbDate
                'verifica se esta vazio e retorna Null
                'Or argValor = "00:00:00" Or argValor = "12:00:00 AM"
                Dim dataVazia As DateTime = Nothing
                If CDate(argValor).ToString("yyyy-MM-dd HH:mm:ss") = CDate(dataVazia).ToString("yyyy-MM-dd HH:mm:ss") Then

                    valorSql = "Null"
                Else
                    valorSql = dataSql(argValor)
                End If
                'Caso seja string acrescenta aspas simples
            Case vbString
                If String.IsNullOrEmpty(argValor) Or argValor = "" Then
                    'devolve a string Null
                    valorSql = "Null"
                Else
                    'acrescenta aspas simples para valores diferentes de vazio
                    valorSql = "'" & argValor & "'"
                End If
                'Caso seja lógico chama a função logicoSql()
            Case vbBoolean
                valorSql = logicoSql(argValor)
        End Select
        Return valorSql
    End Function
    'Função para retornar um valor vazio ao invés de nulo.
    'para utilização nas classes DTO
    'Para setar campo data como null/nothing:
    'campoDeData = objCon.retornaVazioParaValorNulo(drRow("data_inicial_viagem"), Nothing)
    Public Function retornaVazioParaValorNulo(ByVal valor As Object, Optional ByVal valorRetorno As Object = "") As Object
        'verificamos se a variavel esta vazia ou nulla e retornamos vazio e/ou nothing nos casos de data vazia
        If String.IsNullOrEmpty(If(IsDBNull(valor), valorRetorno, valor)) Then
            Return valorRetorno
        ElseIf IsDBNull(valor) Then 'novo
            Return valorRetorno
        Else
            Return valor
        End If
    End Function

    Public Sub salvaDataTable(tabela As String, dt As DataTable)
        Try
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
            Dim qtdcolunas As Integer = 0
            Dim i As Integer = 0 'contador de colunas
            Dim irow As Integer = 0 'contador de linhas
            Dim sql As String = ""
            Dim sqlValues As String = ""
            Dim sqlInto As String = ""
            Dim sqlFinal As String = ""
            Dim column As DataColumn
            Dim cont As Long = 0 'contador
            Dim totalRegistros As Long = dt.Rows.Count
            Dim linhas As Integer
            Dim colunas As Integer
            'percorre toda a datatable e incluir na tabela correspondente
            If dt.Rows.Count > 0 Then 'verifica se existem registros

                Dim hlp As New clsHelpers
                'hlp.abrirForm(frmBarraProgresso, True)
                frmBarraProgresso.Show()
                hlp.carregaBarraProgresso(frmBarraProgresso, frmBarraProgresso.ProgressBar1, totalRegistros, False, False)

                qtdcolunas = dt.Columns.Count 'calcular a qtd de colunas
                'cria sql
                sql = "Insert into [" & tabela & "] "
                sql = sql & "("
                'percorremos as colunas
                For Each column In dt.Columns
                    sqlInto = sqlInto & "[" & column.ToString & "]"
                    If i < (qtdcolunas - 1) Then
                        sqlInto = sqlInto & "," 'Incluimos a virgula em todos menos na ultima coluna
                    Else
                        sqlInto = sqlInto & ") " 'fechamos o parentese
                    End If
                    i = i + 1
                Next
                sql = sql & sqlInto & "Values("
                'percorremos as linhas
                For linhas = 0 To totalRegistros - 1 'total de registros do datatable
                    For colunas = 0 To qtdcolunas - 1 'total de colunas
                        sqlValues = sqlValues & valorSql(dt.Rows(linhas).Item(colunas))
                        If colunas < qtdcolunas - 1 Then
                            sqlValues = sqlValues & ", " 'Incluimos a virgula em todos menos na ultima coluna
                        Else
                            sqlValues = sqlValues & ") " 'fechamos o parentese
                        End If
                        'efetua o insert da linha completa
                        If colunas = qtdcolunas - 1 Then
                            cont = cont + 1 'contador geral
                            'linhas = linhas + 1 'pula para a proxima linha
                            sqlFinal = sql & sqlValues
                            ExecutaQuery(sqlFinal)
                            Application.DoEvents()
                            hlp.carregaBarraProgresso(frmBarraProgresso, frmBarraProgresso.ProgressBar1, , False, True)
                            'colunas = 0
                            sqlValues = ""
                            sqlFinal = ""
                        End If
                    Next
                Next
            End If
            frmBarraProgresso.Hide()
            Cursor.Current = System.Windows.Forms.Cursors.Default
        Catch ex As Exception
            MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub
End Class

