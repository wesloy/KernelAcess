'Classe DAL
Public Class clsLogDAL
    Private dto As New clsLogDTO
    Private con As New clsConexao
    Private sql As String

    Public Function Incluir(ByVal log As clsLogDTO) As Boolean
        sql = "Insert into sysLOG ("
        sql = sql & "[data], "
        sql = sql & "[funcaoExecutada], "
        sql = sql & "[erroNumero], "
        sql = sql & "[erroDescricao], "
        sql = sql & "[idUsuario], "
        sql = sql & "[versaoSis], "
        sql = sql & "[idioma], "
        sql = sql & "[hostname], "
        sql = sql & "[acao]) "
        sql = sql & "values ( "
        sql = sql & con.valorSql(log.data) & ", "
        sql = sql & con.valorSql(log.funcaoExecutada) & ", "
        sql = sql & con.valorSql(log.erroNumero) & ", "
        sql = sql & con.valorSql(log.erroDescricao) & ", "
        sql = sql & con.valorSql(log.idUsuario) & ", "
        sql = sql & con.valorSql(log.versaoSis) & ", "
        sql = sql & con.valorSql(log.idiomaPC) & ", "
        sql = sql & con.valorSql(log.hostname) & ", "
        sql = sql & con.valorSql(log.acao) & ") "
        Incluir = con.ExecutaQuery(sql)
    End Function
End Class

