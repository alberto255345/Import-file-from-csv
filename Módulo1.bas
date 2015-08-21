Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub lixo()
'
' lixo Macro
'

'

Dim oApp As Object 'Objeto application
Dim Arquivo As Variant
Dim NovaPasta As Variant
Dim Caminho As String
Dim strDate As String
Dim test As String
Dim nume As Integer
Dim cont As Integer
Dim val As String
Dim NomedoArquivo As String

Sheets.Add After:=Sheets(Sheets.Count)

'Arquivo = Application.GetOpenFilename("Excel Files (*.xls),*.xls," & "Excel Files (*.xlsx),*.xlsx," & "Excel Files (*.xlsm),*.xlsm," & "Add-in Files (*.xla),*.xla", , "Select Excel file or add-in for XL Start-up")


Arquivo = Application.GetOpenFilename("Arquivo CSV (*.csv), *.csv," & "Arquivo ZIP (*.zip), *.zip")
'Arquivo = "C:\Users\oi344916\Downloads\TarefasAbertas_13082015_0705.zip"
test = Arquivo

nume = Len(test)
cont = 2
Do While cont <> nume
If Left(Right(test, cont), 1) = "\" Then

Exit Do

Else
cont = cont + 1
End If
Loop

Caminho = Left(test, Len(test) - cont + 1)
NomedoArquivo = Left(Right(test, cont - 1), cont - 5)
'Caminho = Left(test, InStr(test, "\"))

If Arquivo = False Then
Application.DisplayAlerts = False
ActiveSheet.Delete
Application.DisplayAlerts = True
    End
  End If
  
If Right(Arquivo, 4) = ".zip" Then
   
   NovaPasta = Caminho & NomedoArquivo
   'Cria um nova pasta
   On Error GoTo ErrHandler:
   MkDir NovaPasta
   
ErrHandler:
If Err.Number <> 0 Then

If FileExists(NovaPasta & "\*") = True Then
Kill NovaPasta & "\*"
End If
RmDir NovaPasta & "\"

MkDir NovaPasta
End If
   
   
   'Extrai os arquivos para a pasta criada
   Set oApp = CreateObject("Shell.Application")
oApp.Namespace(NovaPasta).CopyHere oApp.Namespace(Arquivo).items
Arquivo = NovaPasta & "\" & NomedoArquivo & ".csv"

  End If

With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & Arquivo, Destination:=Range("A1"))
        .Name = "Base"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
Kill NovaPasta & "\*"
RmDir NovaPasta & "\"


    ActiveSheet.Name = "Lixo"
    Application.AutoRecover.Enabled = False
    Sheets("Macro").Select

End Sub
Sub extrair()
'
' extrair Macro
'

'
    
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Sheets("Base")
    On Error GoTo 0
    If ws Is Nothing Then
        'MsgBox "Não existe uma Planilha com esse nome!", vbCritical
    Sheets.Add
    ActiveSheet.Name = "Base"
    Else
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    Sheets.Add
    ActiveSheet.Name = "Base"
    End If
    
    Sheets("Base").Select
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Range("A1").Select
    
    Sheets("Lixo").Select
    
    If Not ActiveSheet.AutoFilterMode Then
    ActiveSheet.Range("A1").AutoFilter
    End If
    ActiveSheet.Range("$A$1:$CB$1048575").AutoFilter Field:=68, Criteria1:="CE"
    
    ''''''''''''''''' seleção geral ''''''''''''''''''''''
    'fimr = 1048576
    contR = 1048576
    contC = 1
    
    Do While Cells(contR, 68).Value <> "CE"
    contR = contR - 1
    'fimr = fimr - 1
    Loop
    
    'contR = 1048576 - fimr
    
    valor = Cells(1, contC).Value
    
    Do While valor <> ""
    contC = contC + 1
    valor = Cells(1, contC).Value
    Loop
    contC = contC - 1
    Range(Cells(1, 1), Cells(contR, contC)).Select
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Call Plan1.AutoFilter.Range.Copy
    'Call Plan3.Paste
     
    'Cells.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    
    Selection.Copy
    Sheets("Base").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    
    Application.CutCopyMode = False
    Sheets("Lixo").Select
    Range("A1").Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
    Sheets("Macro").Select
    
End Sub
Sub editar()
'
' editar Macro
'

'
    Sheets("Base").Select
    Cells.Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    Cells.Replace What:="INSTALAÇÃO VELOX (INS KIT)", Replacement:= _
        "INSTALAÇÃO VELOX", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    Cells.Replace What:="INS MUDEND - MUDANÇA ENDEREÇO VOZ", Replacement:= _
        "MUDEND VOZ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    Cells.Replace What:="INS MUD AREA - MUDANÇA ÁREA VOZ", Replacement:= _
        "MUDEND VOZ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    Cells.Replace What:="INS MUD LOC - MUDANÇA LOCALIDADE VOZ", Replacement:= _
        "MUDEND VOZ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
        
    Cells.Replace What:="INS MUDANÇA SUBNUM VOZ", Replacement:= _
        "MUDEND VOZ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False

    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AB1").Select
    ActiveCell.FormulaR1C1 = "PERÍODO"
    Range("AB2").Select
    
    contDat = 2
    verdat = Left(Cells(contDat, 26).Value, 10)
    conti = Cells(contDat, 1).Value
    Do While conti <> ""
    
    If verdat = Format(Date, "yyyy-mm-dd") Then
    
    verdat = Left(Right(Cells(contDat, 26).Value, 10), 2)
    
    If verdat < 12 Then
    Cells(contDat, 28).Value = "Manhã"
    ElseIf verdat < 18 Then
    Cells(contDat, 28).Value = "Tarde"
    Else
    Cells(contDat, 28).Value = "Noite"
    End If
    
    contDat = contDat + 1
    verdat = Left(Cells(contDat, 26).Value, 10)
    
    Else
    
    Cells(contDat, 28).Value = "Outra Data"
    contDat = contDat + 1
    verdat = Left(Cells(contDat, 26).Value, 10)
    
    End If
    
    conti = Cells(contDat, 1).Value
    
    Loop
    
    Sheets("Macro").Select
    
End Sub

Sub limp()
Sheets("Base").Select
    Cells.Select
    Range("A1").Activate
    Selection.ClearContents
    Range("A1").Select
    Sheets("Macro").Select
End Sub

Sub dinam()

Dim sht As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim pvt2 As PivotTable
Dim pvt3 As PivotTable
Dim pvt4 As PivotTable
Dim StartPvt As String
Dim SrcData As String
Dim pf4 As PivotTable
Dim pf5 As PivotTable
Dim pf6 As PivotTable
Dim i As Long
Dim ws As Worksheet
Dim sPlanilha As String
Dim tbl As ListObject
Dim DataRange As String
Dim Destination As Range
Dim Destination2 As Range
Dim Destination3 As Range

sPlanilha = "Dina"
    
    
    
    On Error Resume Next
    Set ws = Sheets(sPlanilha)
    On Error GoTo 0
    If ws Is Nothing Then
        'MsgBox "Não existe uma Planilha com esse nome!", vbCritical
    Else
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If


'Determine the data range you want to pivot
  SrcData = Worksheets("Base").Name & "!" & Range("A1:CC1048575").Address(ReferenceStyle:=xlR1C1)
  
'Create a new worksheet
  Set sht = Sheets.Add
  ActiveSheet.Name = "Dina"

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="Tabela dinâmica1")


'Set pvt = PivotTables("PivotTable1")
    
  'Add item to the Report Filter
    pvt.PivotFields("PRONTOPARAEXECUCAO").Orientation = xlPageField
      
  'Add item to the Column Labels
    pvt.PivotFields("PERÍODO").Orientation = xlColumnField
    
  'Add item to the Row Labels
    pvt.PivotFields("ATIVIDADE").Orientation = xlRowField
    pvt.PivotFields("NRBA").Orientation = xlDataField
    
    
    
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PERÍODO").PivotItems(5).Visible = False
     ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PRONTOPARAEXECUCAO").CurrentPage = "Sim"
     'ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PRONTOPARAEXECUCAO").PivotItems(3).Visible = False
     'ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PRONTOPARAEXECUCAO").PivotItems(1).Visible = False
     'ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PRONTOPARAEXECUCAO").PivotItems(2).Visible = True
     ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PERÍODO").PivotItems("Manhã").Position = 1
     ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PERÍODO").PivotItems("Tarde").Position = 2
     ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PERÍODO").PivotItems("Noite").Position = 3
     ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("PERÍODO").PivotItems("Outra Data").Position = 4
     
     Range("A1:B1").Interior.Color = 9944516
     
     If ActiveWindow.DisplayGridlines = True Then
   ActiveWindow.DisplayGridlines = False
   
Else
   
End If


pvt.TableStyle2 = "PivotStyleMedium2"

Range("A1:f4").Font.Bold = True
    
    
    conrow = 1
    vernum = Cells(conrow, 1).Value
    Do While vernum <> "Total Geral"
    
    conrow = conrow + 1
    vernum = Cells(conrow, 1).Value
    
    Loop
    
    Range(Cells(5, 2), Cells(conrow, 4)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    Range("B3").Select
    ActiveSheet.PivotTables("Tabela dinâmica1").CompactLayoutColumnHeader = "Período"
    
'''''''''''''''''''' dina 2 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
 'DataRange = Worksheets("Base").Name & "!" & Range("A1:CC1048575").Address(ReferenceStyle:=xlR1C1)

Set Destination = Worksheets("Dina").Cells(1, 8)

  
'Set pvtCache2 = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=DataRange)
    
Set pvt2 = pvtCache.CreatePivotTable(TableDestination:=Destination, TableName:="Tabela dinâmica2")
    
     
Set pf4 = ActiveSheet.PivotTables("Tabela dinâmica2")
pf4.PivotFields("PRONTOPARAEXECUCAO").Orientation = xlPageField
pf4.PivotFields("PERÍODO").Orientation = xlPageField
pf4.PivotFields("ATIVIDADE").Orientation = xlColumnField
pf4.PivotFields("GRA").Orientation = xlRowField
pf4.PivotFields("SETOR").Orientation = xlRowField
pf4.PivotFields("NRBA").Orientation = xlDataField
pf4.PivotFields("PERÍODO").Position = 2
pf4.PivotFields("PRONTOPARAEXECUCAO").Position = 1
pf4.TableStyle2 = "PivotStyleMedium2"
pf4.PivotFields("ATIVIDADE").ClearAllFilters
pf4.PivotFields("ATIVIDADE").EnableMultiplePageItems = True
pf4.PivotFields("PERÍODO").EnableMultiplePageItems = True
pf4.PivotFields("PRONTOPARAEXECUCAO").EnableMultiplePageItems = True
For i = 2 To pf4.PivotFields("ATIVIDADE").PivotItems.Count
    pf4.PivotFields("ATIVIDADE").PivotItems(i).Visible = False
Next
    pf4.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VELOX").Visible = True
    pf4.PivotFields("ATIVIDADE").PivotItems("REPARO VELOX").Visible = True
    pf4.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VOZ").Visible = True
    pf4.PivotFields("ATIVIDADE").PivotItems("MUDEND VOZ").Visible = True
    pf4.PivotFields("ATIVIDADE").PivotItems("REPARO VOZ").Visible = True
    pf4.PivotFields("ATIVIDADE").PivotItems(1).Visible = False
    
pf4.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VELOX").Position = 1
pf4.PivotFields("ATIVIDADE").PivotItems("REPARO VELOX").Position = 2
pf4.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VOZ").Position = 3
pf4.PivotFields("ATIVIDADE").PivotItems("MUDEND VOZ").Position = 4
pf4.PivotFields("ATIVIDADE").PivotItems("REPARO VOZ").Position = 5

ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("PRONTOPARAEXECUCAO").CurrentPage = "Sim"

With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("PERÍODO")
        .PivotItems("Outra Data").Visible = False
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
        On Error Resume Next
    End With
    
    With ActiveSheet.PivotTables("Tabela dinâmica2").PivotFields("GRA")
        .PivotItems("Tratamento CE").Visible = False
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
        On Error Resume Next
    End With

Range("H1:I2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
Range("H1:N5").Select
    Selection.Font.Bold = True


'''''''''''''''''''''''''''''''''''''' dina 3'''''''''''''''''''''''''''''''''

 'DataRange = Worksheets("Base").Name & "!" & Range("A1:CC1048575").Address(ReferenceStyle:=xlR1C1)
'Worksheets("Base").Range ("A1:CC1048575")
Set Destination2 = Worksheets("Dina").Cells(1, 16)

  
'Set pvtCache3 = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=DataRange)
    
Set pvt3 = pvtCache.CreatePivotTable(TableDestination:=Destination2, TableName:="Tabela dinâmica3")
    
     
Set pf5 = ActiveSheet.PivotTables("Tabela dinâmica3")
pf5.PivotFields("PRONTOPARAEXECUCAO").Orientation = xlPageField
pf5.PivotFields("PERÍODO").Orientation = xlColumnField
pf5.PivotFields("ATIVIDADE").Orientation = xlColumnField
pf5.PivotFields("GRA").Orientation = xlRowField
pf5.PivotFields("SETOR").Orientation = xlRowField
pf5.PivotFields("NRBA").Orientation = xlDataField
pf5.PivotFields("ATIVIDADE").Position = 2
pf5.PivotFields("PERÍODO").Position = 1
pf5.TableStyle2 = "PivotStyleMedium15"
pf5.PivotFields("ATIVIDADE").ClearAllFilters
pf5.PivotFields("ATIVIDADE").EnableMultiplePageItems = True
pf5.PivotFields("PERÍODO").EnableMultiplePageItems = True
pf5.PivotFields("PRONTOPARAEXECUCAO").CurrentPage = "Sim"
'pf5.PivotFields("PRONTOPARAEXECUCAO").EnableMultiplePageItems = True
For i = 2 To pf5.PivotFields("ATIVIDADE").PivotItems.Count
    pf5.PivotFields("ATIVIDADE").PivotItems(i).Visible = False
Next
    'pf5.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VELOX").Visible = True
    'pf5.PivotFields("ATIVIDADE").PivotItems("REPARO VELOX").Visible = True
    pf5.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VOZ").Visible = True
    pf5.PivotFields("ATIVIDADE").PivotItems("MUDEND VOZ").Visible = True
    pf5.PivotFields("ATIVIDADE").PivotItems("REPARO VOZ").Visible = True
    pf5.PivotFields("ATIVIDADE").PivotItems(1).Visible = False
    
    
    
'pf5.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VELOX").Position = 1
'pf5.PivotFields("ATIVIDADE").PivotItems("REPARO VELOX").Position = 2
pf5.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VOZ").Position = 1
pf5.PivotFields("ATIVIDADE").PivotItems("MUDEND VOZ").Position = 2
pf5.PivotFields("ATIVIDADE").PivotItems("REPARO VOZ").Position = 3

ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("PRONTOPARAEXECUCAO").CurrentPage = "Sim"

With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("PERÍODO")
        .PivotItems("Outra Data").Visible = False
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
        On Error Resume Next
    End With
    
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("GRA")
        .PivotItems("Tratamento CE").Visible = False
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
        On Error Resume Next
    End With
    
    With ActiveSheet.PivotTables("Tabela dinâmica3").PivotFields("PERÍODO")
        .PivotItems("Outra Data").Visible = True
    End With
pf5.PivotFields("PERÍODO").PivotItems("Manhã").Position = 1
pf5.PivotFields("PERÍODO").PivotItems("Tarde").Position = 2
pf5.PivotFields("PERÍODO").PivotItems("Noite").Position = 3
pf5.PivotFields("PERÍODO").PivotItems("Outra Data").Position = 4


'''''''''''''''''''''''''' dina 4 '''''''''''''''''''''''''''''''''''''''''''''''''

 'DataRange = Worksheets("Base").Name & "!" & Range("A1:CC1048575").Address(ReferenceStyle:=xlR1C1)
'Worksheets("Base").Range ("A1:CC1048575")
Set Destination3 = Worksheets("Dina").Cells(1, 35)

  
'Set pvtCache3 = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=DataRange)
    
Set pvt4 = pvtCache.CreatePivotTable(TableDestination:=Destination3, TableName:="Tabela dinâmica4")
    
     
Set pf6 = ActiveSheet.PivotTables("Tabela dinâmica4")
pf6.PivotFields("PRONTOPARAEXECUCAO").Orientation = xlPageField
pf6.PivotFields("PERÍODO").Orientation = xlColumnField
pf6.PivotFields("ATIVIDADE").Orientation = xlColumnField
pf6.PivotFields("GRA").Orientation = xlRowField
pf6.PivotFields("SETOR").Orientation = xlRowField
pf6.PivotFields("NRBA").Orientation = xlDataField
pf6.PivotFields("ATIVIDADE").Position = 2
pf6.PivotFields("PERÍODO").Position = 1
pf6.TableStyle2 = "PivotStyleMedium15"
pf6.PivotFields("ATIVIDADE").ClearAllFilters
pf6.PivotFields("ATIVIDADE").EnableMultiplePageItems = True
pf6.PivotFields("PERÍODO").EnableMultiplePageItems = True
pf6.PivotFields("PRONTOPARAEXECUCAO").CurrentPage = "Sim"
'pf6.PivotFields("PRONTOPARAEXECUCAO").EnableMultiplePageItems = True
For i = 2 To pf6.PivotFields("ATIVIDADE").PivotItems.Count
    pf6.PivotFields("ATIVIDADE").PivotItems(i).Visible = False
Next
    pf6.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VELOX").Visible = True
    pf6.PivotFields("ATIVIDADE").PivotItems("REPARO VELOX").Visible = True
    'pf6.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VOZ").Visible = True
    'pf6.PivotFields("ATIVIDADE").PivotItems("MUDEND VOZ").Visible = True
    'pf6.PivotFields("ATIVIDADE").PivotItems("REPARO VOZ").Visible = True
    pf6.PivotFields("ATIVIDADE").PivotItems(1).Visible = False
    
    
    
pf6.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VELOX").Position = 1
pf6.PivotFields("ATIVIDADE").PivotItems("REPARO VELOX").Position = 2
'pf6.PivotFields("ATIVIDADE").PivotItems("INSTALAÇÃO VOZ").Position = 1
'pf6.PivotFields("ATIVIDADE").PivotItems("MUDEND VOZ").Position = 2
'pf6.PivotFields("ATIVIDADE").PivotItems("REPARO VOZ").Position = 3

ActiveSheet.PivotTables("Tabela dinâmica4").PivotFields("PRONTOPARAEXECUCAO").CurrentPage = "Sim"

With ActiveSheet.PivotTables("Tabela dinâmica4").PivotFields("PERÍODO")
        .PivotItems("Outra Data").Visible = False
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
        On Error Resume Next
    End With
    
    With ActiveSheet.PivotTables("Tabela dinâmica4").PivotFields("GRA")
        .PivotItems("Tratamento CE").Visible = False
        On Error Resume Next
        .PivotItems("(blank)").Visible = False
        On Error Resume Next
    End With
    
    With ActiveSheet.PivotTables("Tabela dinâmica4").PivotFields("PERÍODO")
        .PivotItems("Outra Data").Visible = True
    End With
pf6.PivotFields("PERÍODO").PivotItems("Manhã").Position = 1
pf6.PivotFields("PERÍODO").PivotItems("Tarde").Position = 2
pf6.PivotFields("PERÍODO").PivotItems("Noite").Position = 3
pf6.PivotFields("PERÍODO").PivotItems("Outra Data").Position = 4
    
    
 Cells.Select
    Range("AJ3").Activate
    With Selection.Font
        .Name = "Calibri"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = -1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.EntireColumn.AutoFit
    Columns("AH:AH").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("P:P").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("H:H").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("P1:AT5").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("O11").Select
    
    Range("P1:AT5").Select
    Selection.Font.Bold = True
    
    Range("A1").Select

End Sub



Sub rapidinho()

Call lixo
Call extrair
Call editar
Call dinam


End Sub
