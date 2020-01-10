Module mdlConstantes
    'OFICIAL - R - S250
    'Path de configurações do sistema
    Public Const BD_PWD = "cicpkdx0"
    Public Const BD_NOME = "db_kernel.accdb"
    Public Const BD_PATH = "\\D5174s006\bd\S250_Apl_Fraude\03_Database\01_Amex\Kernel\vbnet\base\"
    Public Const PATH_ICONS = "\\d5174s006\D5174\Compartilhado\Entre_secoes\S250_Apl_Fraude\03_Database\01_Amex\Kernel\imagens\"

    ' ''TESTE ALGAR
    ' ''Path de configurações do sistema
    'Public Const BD_PWD = "cicpkdx0"
    'Public Const BD_NOME = "db_kernel.accdb"
    'Public Const BD_PATH = "C:\Users\rafaelsan\Documents\Projetos\BancoDados\"
    'Public Const PATH_ICONS = "C:\Users\rafaelsan\Documents\Projetos\Everest\Framework\Imagens\"

    'Constantes opcionais
    Public Const TITULO_ALERTA = "Alerta do Sistema"
    Public Const FormatoDataUniversal = "yyyy-MM-dd"
    Public Const FormatoDataHoraUniversal = "yyyy-MM-dd HH:mm:ss"
    Public Const CREDITOS = "Desenvolvido por: Rafael Alvarenga"
    Public Const SENHA_SISTEMA = "cicpkdx0"

    Public Enum FlagAcao
        Insert = 1
        Update = 2
        Delete = 3
        NoAction = 0
    End Enum

    Public Function imglist() As ImageList
        'cria um imagelist se necessario
        Dim imageListSmall As New ImageList
        With imageListSmall
            '.ImageSize = New Size(16, 16) ' (the default is 16 x 16).
            .Images.Add(1, Image.FromFile(PATH_ICONS & "01.ico"))
            .Images.Add(2, Image.FromFile(PATH_ICONS & "02.ico"))
            .Images.Add(3, Image.FromFile(PATH_ICONS & "03.ico"))
            .Images.Add(4, Image.FromFile(PATH_ICONS & "04.ico"))
            .Images.Add(5, Image.FromFile(PATH_ICONS & "05.ico"))
            .Images.Add(6, Image.FromFile(PATH_ICONS & "06.ico"))
            .Images.Add(7, Image.FromFile(PATH_ICONS & "07.ico"))
            .Images.Add(8, Image.FromFile(PATH_ICONS & "08.ico"))
            .Images.Add(9, Image.FromFile(PATH_ICONS & "09.ico"))
            .Images.Add(10, Image.FromFile(PATH_ICONS & "10.ico"))
            .Images.Add(11, Image.FromFile(PATH_ICONS & "11.ico"))
            .Images.Add(12, Image.FromFile(PATH_ICONS & "12.ico"))
            .Images.Add(13, Image.FromFile(PATH_ICONS & "13.ico"))
            .Images.Add(14, Image.FromFile(PATH_ICONS & "14.ico"))
        End With
        'fim da criacao do imagelist
        Return imageListSmall
    End Function

    'Armazena as sessoes de login
    Public sessaoNomeUsuario As String
    Public sessaoIdUsuario As String
    Public sessaoIdModulo As Integer
    Public sessaoModulo As String
    Public sessaoIdArea As Integer
    Public sessaoArea As String
End Module
