Public Class clsDadosBLL

    Private dt As DataTable 'OleDb.OleDbDataReader
    Private dto As New clsDadosDTO
    Private dal As New clsDadosDAL

    Public Function AtualizarListaDados() As Boolean
        frmCadastro.ListView1.Clear()
        dt = dal.GetDados()
        'Modelando as colunas
        With frmCadastro.ListView1
            .View = View.Details
            .LabelEdit = False
            .CheckBoxes = False
            .SmallImageList = imglist() 'Utilizando um modulo publico
            .GridLines = True
            .FullRowSelect = True
            .HideSelection = False
            .MultiSelect = False
            .Columns.Add("ID", 50, HorizontalAlignment.Center)
            .Columns.Add("Aplicativo", 150, HorizontalAlignment.Left)
            .Columns.Add("Banco de dados", 250, HorizontalAlignment.Left)
        End With

        'POPULANDO
        If dt.Rows.Count > 0 Then 'verifica se existem registros
            For Each drRow As DataRow In dt.Rows
                Dim item As New ListViewItem()
                item.Text = drRow("id")
                item.SubItems.Add(drRow("aplicativo").ToString)
                item.SubItems.Add(drRow("nome_bd").ToString)
                frmCadastro.ListView1.Items.Add(item)
            Next drRow
        End If
        Return True
    End Function

    Public Sub PreencheComboDados(frm As Form, cb As ComboBox)
        dal.GetComboboxDados(frm, cb)
    End Sub

    Public Function SalvarDados(objDados As clsDadosDTO) As Boolean
        Return dal.Salvar(objDados)
    End Function

    Public Function CarregaDadosPorId(ByVal id As Integer) As clsDadosDTO
        CarregaDadosPorId = dal.GetDadosPorID(id)
    End Function

    Public Function Deletar(ByVal id As Integer) As Boolean
        Deletar = dal.DeletaDadosPorID(id)
    End Function

End Class
