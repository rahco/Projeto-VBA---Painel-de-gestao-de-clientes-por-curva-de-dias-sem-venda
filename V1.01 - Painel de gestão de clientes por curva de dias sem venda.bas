Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False
    
    Call Datas_atividades
    Call Base_tratada_2
    
    If Sheets("BASE TRATADA (2)").Cells(3, 115).Value > 0 Then
        Sheets("BASE TRATADA (2)").Select
        Range("DK5").Select
        MsgBox ("Ajustar Área x HC")
    Else
        Call Base_geral
        
    Sheets("MACROS").Select
    Range("B7").Select

    End If

    Application.ScreenUpdating = True

End Sub

Sub Datas_atividades()

    Application.ScreenUpdating = False
    
    Sheets("DATAS ATIVIDADES").Select
    Range("D5").Select
    Selection.Copy
    Range("D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D5").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True

End Sub

Sub Base_tratada_2()
Attribute Base_tratada_2.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE TRATADA (2)").Range("C3").Value)
    final = Abs(Worksheets("BASE TRATADA (2)").Range("B3").Value)
 
    Do While atual > final
        Sheets("BASE TRATADA (2)").Select
        Range("B5").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B5").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE TRATADA (2)").Range("C3").Value)
        final = Abs(Worksheets("BASE TRATADA (2)").Range("B3").Value)
    Loop

    Sheets("BASE TRATADA (2)").Select
    Range("B5").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C3").Value > 0 Then
        linhaf = linhai - Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C3").Value < 0 Then
        linhaf = linhai + Range("C3").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B5").Select

    Sheets("BASE TRATADA").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BASE TRATADA (2)").Select
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False
    Range("CN5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("CN6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B5").Select

    Application.ScreenUpdating = True

End Sub

Sub Base_geral()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE GERAL").Range("C1").Value)
    final = Abs(Worksheets("BASE GERAL").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE GERAL").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE GERAL").Range("C1").Value)
        final = Abs(Worksheets("BASE GERAL").Range("B1").Value)
    Loop

    Sheets("BASE GERAL").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select

    Sheets("BASE TRATADA (2)").Select
    Range("DR4").Select
    ActiveSheet.Range("$B$4:$DS$100000").AutoFilter Field:=122, Criteria1:="=1", _
        Operator:=xlAnd
    Range("CY5:DR5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE GERAL").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BASE TRATADA (2)").Select
    Range("DS4").Select
    ActiveSheet.Range("$B$4:$DS$100000").AutoFilter Field:=122
    Range("B5").Select
    Sheets("BASE GERAL").Select
    Range("B4").Select
    Range("E3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("E3:E100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("D3:D100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("F3:F100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("G3:G100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("N3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("N3:N100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("O3").Select
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("O3:O100000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE GERAL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B3").Select
    Selection.End(xlToRight).Select
    Range("S3").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Columns.AutoFit
    Range("B4").Select

    ActiveWorkbook.RefreshAll

    Application.ScreenUpdating = True

End Sub

Sub Arquivo_de_envio()

    Application.ScreenUpdating = False
    
    ActiveWorkbook.Save

    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C10").Value & " - Gestão de POS sem venda - Dados até dia " & Worksheets("MACROS").Range("C11").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

    Sheets("QUADRO GERENCIAL").Select
    ActiveWindow.DisplayHeadings = True
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B5").Select
    Application.CutCopyMode = False
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE GERAL").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    ActiveWindow.DisplayHeadings = False
    Rows("1:1").Select
    Range("B1").Activate
    Selection.ClearContents
    Range("B4").Select
    Sheets(Array("DATAS ATIVIDADES", "BASE TRATADA", "BASE TRATADA (2)", "TD", _
        "TABELA DE ENVIO")).Select
    Sheets("TD").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=-2
    Sheets(Array("MACROS", "DADOS GERAIS", "DATAS ATIVIDADES", "BASE TRATADA", _
        "BASE TRATADA (2)", "TD", "TABELA DE ENVIO")).Select
    Sheets("DATAS ATIVIDADES").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("QUADRO GERENCIAL").Select
    ActiveWorkbook.Save
    
    Application.ScreenUpdating = True

End Sub
