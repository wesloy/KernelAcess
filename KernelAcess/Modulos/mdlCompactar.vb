Imports System.IO

Module mdlCompactar
    Private hlp As New clsHelpers
    Public Sub compact(bd_senha As String, bd_path As String, bd_nome As String)
        Dim password As String = bd_senha
        'path da base de datos a ser reparada/compactada
        Dim caminho_bd_original As String = bd_path & bd_nome
        'arquivo temporario não deve existir
        Dim extensao As String = hlp.PegarExtensao(caminho_bd_original)
        Dim caminho_bd_Temp As String = BD_PATH & "BDTemporario__." & extensao

        Try
            Dim objDbEngine As New Microsoft.Office.Interop.Access.Dao.DBEngine()
            'apaga o arquivo temp se existir
            If Dir(caminho_bd_Temp) <> "" Then Kill(caminho_bd_Temp)
            'formata a senha no formato ";pwd=PASSWORD" se a mesma existir
            If password <> "" Then password = ";pwd=" & bd_senha
            'compacta a base criando um novo banco de dados
            objDbEngine.CompactDatabase(caminho_bd_original, caminho_bd_Temp, Nothing, Nothing, password)
            If (New FileInfo(caminho_bd_Temp)).Length = (New FileInfo(caminho_bd_original)).Length Then
                'não foi compactada e tem o mesmo tamanho de antes
                File.Delete(caminho_bd_Temp)
                MsgBox("Banco de dados não foi reparado.")
            Else
                'Sucesso na compactação
                'copia de segurança do arquivo original, IMPORTANTE!!!
                Dim cria_backup As String
                cria_backup = bd_path & "Backup_NaoCompactado_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & "_" & bd_nome
                File.Copy(caminho_bd_original, cria_backup)
                'apagar db original
                File.Delete(caminho_bd_original)
                'mover o db compactada
                File.Move(caminho_bd_Temp, caminho_bd_original)
                hlp.registrarLOG(, , hlp.GetCurrentMethodName(), bd_nome)
                MsgBox("Banco de dados reparado com sucesso!", vbInformation, TITULO_ALERTA)
            End If
        Catch ex As Exception
            Dim mensagem As String = ex.Message & Err.Number
            If Err.Number = 3704 Then
                mensagem = "Você tentou abrir um banco de dados que já está sendo utilizado por outro usuário. Por favor, tente novamente quando o banco de dados não estiver sendo utilizado." '& vbNewLine & vbNewLine & "A FERRAMENTA NÃO PODE ESTAR SENDO UTILIZADA!"
            Else
                mensagem = ex.Message
            End If
            MsgBox(mensagem, vbExclamation, TITULO_ALERTA)
        End Try

    End Sub
End Module
