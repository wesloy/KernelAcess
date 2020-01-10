Imports System.IO

'Forma de uso:
'Dim teste As New clsEscreveArquivoTxt
'    With teste
'        .CriaArquivo(caminhoNomeArquivo)
'        .EscreveLn("A")
'        .EscreveLn("B")
'        .FechaStrm()
'    End With

Public Class clsEscreveArquivoTxt
    Private strm As StreamWriter

    'cria uma instância de StreamWriter para escrever no
    'arquivo desejado
    Public Function CriaArquivo(ByVal caminhoNomeArquivo As String) As StreamWriter
        strm = New StreamWriter(caminhoNomeArquivo)
        Return strm
    End Function

    Public Sub EscreveLn(ByVal linha As String)
        strm.WriteLine(linha)
    End Sub

    Public Sub SaltandoLinhas(ByVal quantidadeSaltos As Integer)
        For i As Integer = 1 To quantidadeSaltos
            strm.WriteLine("")
        Next i
    End Sub

    Public Sub FechaStrm()
        strm.Close() 'fecha o objeto strm
    End Sub


End Class
