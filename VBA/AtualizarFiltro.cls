VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
' Atualiza o filtro da tabela dinâmica a partir do valor de uma célula
    ' Definindo as variáveis
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    Dim xStr As String
    
    ' Em caso de erro, exibir mensagem
    On Error GoTo ErrMsg
    If Intersect(Target, Range("B3:D4")) Is Nothing Then Exit Sub
    
    ' Desativa atualização de tela durante a busca
    Application.ScreenUpdating = False
    
    ' Atribuindo valores às variáveis
    Set xPTable = Worksheets("Interface").PivotTables("Tabela dinâmica1")
    Set xPFile = xPTable.PivotFields("Serviço")
    xStr = Target.Text
    
    ' Limpando todos os filtros da tabela dinâmica
    xPFile.ClearAllFilters
    
    ' Atribuindo novo filtro
    xPFile.CurrentPage = xStr
    
    ' Retomando atualização de tela
    Application.ScreenUpdating = True
    Range("B3").Select
Exit Sub

' Mensagem de erro caso a busca não se complete
ErrMsg:
    MsgBox ("Serviço não existente. O filtro será resetado.")
    Range("B3").Select
    
End Sub
