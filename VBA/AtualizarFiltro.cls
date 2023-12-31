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
' Atualiza o filtro da tabela din�mica a partir do valor de uma c�lula
    ' Definindo as vari�veis
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    Dim xStr As String
    
    ' Em caso de erro, exibir mensagem
    On Error GoTo ErrMsg
    If Intersect(Target, Range("B3:D4")) Is Nothing Then Exit Sub
    
    ' Desativa atualiza��o de tela durante a busca
    Application.ScreenUpdating = False
    
    ' Atribuindo valores �s vari�veis
    Set xPTable = Worksheets("Interface").PivotTables("Tabela din�mica1")
    Set xPFile = xPTable.PivotFields("Servi�o")
    xStr = Target.Text
    
    ' Limpando todos os filtros da tabela din�mica
    xPFile.ClearAllFilters
    
    ' Atribuindo novo filtro
    xPFile.CurrentPage = xStr
    
    ' Retomando atualiza��o de tela
    Application.ScreenUpdating = True
    Range("B3").Select
Exit Sub

' Mensagem de erro caso a busca n�o se complete
ErrMsg:
    MsgBox ("Servi�o n�o existente. O filtro ser� resetado.")
    Range("B3").Select
    
End Sub
