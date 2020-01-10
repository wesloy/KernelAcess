Public Class clsDadosDAL
    Private sql As String
    Private objCon As New clsConexao
    Private dt As DataTable ' OleDb.OleDbDataReader
    Private hlp As New clsHelpers

    Public Function DeletaDadosPorID(ByVal ID_Dados As Integer) As Boolean
        sql = "Delete * from sysDados Where id  = " & objCon.valorSql(ID_Dados)
        DeletaDadosPorID = objCon.ExecutaQuery(sql)
    End Function
    'Captura todo o objeto, com opção de filtrar o dado desejado
    Public Function GetDados(Optional ByVal filtro As String = "") As DataTable
        sql = "Select * from sysDados "
        GetDados = objCon.RetornaDataTable(sql)
    End Function

    'Captura a área por ID e retorna um Objeto
    Public Function GetDadosPorID(ByVal ID_Dados As Integer) As clsDadosDTO
        Dim objDados As New clsDadosDTO
        sql = "Select * from sysDados where id = " & objCon.valorSql(ID_Dados)
        dt = objCon.RetornaDataTable(sql)
        If dt.Rows.Count > 0 Then 'verifica se existem registros
            For Each drRow As DataRow In dt.Rows 'efetua o loop até o fim
                With objDados
                    .ID = objCon.retornaVazioParaValorNulo(drRow("id"))
                    ._aplicativo = objCon.retornaVazioParaValorNulo(drRow("aplicativo"))
                    ._senha_bd = objCon.retornaVazioParaValorNulo(drRow("senha_bd"))
                    ._nome_bd = objCon.retornaVazioParaValorNulo(drRow("nome_bd"))
                    ._caminho_bd = objCon.retornaVazioParaValorNulo(drRow("caminho_bd"))
                    ._senha_acesso = objCon.retornaVazioParaValorNulo(drRow("senha_acesso"))
                    .Acao = FlagAcao.NoAction
                End With
            Next drRow
        End If
        Return objDados
    End Function

    Public Function Incluir(ByVal objDados As clsDadosDTO) As Boolean
        sql = "Insert into sysDados "
        sql = sql & "(aplicativo, "
        sql = sql & "senha_bd, "
        sql = sql & "nome_bd, "
        sql = sql & "caminho_bd, "
        sql = sql & "senha_acesso) "
        sql = sql & "values ( "
        sql = sql & objCon.valorSql(objDados._aplicativo.Trim) & ", "
        sql = sql & objCon.valorSql(objDados._senha_bd.Trim) & ", "
        sql = sql & objCon.valorSql(objDados._nome_bd.Trim) & ", "
        sql = sql & objCon.valorSql(objDados._caminho_bd.Trim) & ", "
        sql = sql & objCon.valorSql(objDados._senha_acesso.Trim) & " )"
        Incluir = objCon.ExecutaQuery(sql)
        Return Incluir
    End Function

    Public Function Atualizar(ByVal objDados As clsDadosDTO) As Boolean
        sql = "Update sysDados "
        sql = sql & "Set aplicativo = " & objCon.valorSql(objDados._aplicativo.Trim) & ", "
        sql = sql & "senha_bd = " & objCon.valorSql(objDados._senha_bd.Trim) & ", "
        sql = sql & "nome_bd = " & objCon.valorSql(objDados._nome_bd.Trim) & ", "
        sql = sql & "caminho_bd = " & objCon.valorSql(objDados._caminho_bd) & ", "
        sql = sql & "senha_acesso = " & objCon.valorSql(objDados._senha_acesso) & " "
        sql = sql & "Where id = " & objDados.ID & " "
        Atualizar = objCon.ExecutaQuery(sql)
        Return Atualizar
    End Function

    Public Function Salvar(ByVal objDados As clsDadosDTO) As Boolean
        If objDados.Acao = FlagAcao.Insert Then
            Salvar = Me.Incluir(objDados)
        ElseIf objDados.Acao = FlagAcao.Update Then
            Salvar = Me.Atualizar(objDados)
        End If
        Return Salvar
    End Function

    Public Sub GetComboboxDados(frm As Form, cb As ComboBox)
        hlp.carregaComboBox("Select * from sysDados order by aplicativo asc ", frm, cb)
    End Sub

End Class

