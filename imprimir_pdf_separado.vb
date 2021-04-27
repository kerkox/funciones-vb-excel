Sub ImprimirPdfSeparadosV2()
'
' Por Paul Cortes
' Declaracion de Variables
Dim paginasDocumento As Integer
Dim totalPaginas As Integer
Dim pagActual As Integer
Dim carpeta As String
Dim nombreDocs As String
Dim maximoDocs As Integer
maximoDocs = 0
' Asignacion de valores a variables
totalPaginas = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
paginasDocumento = InputBox("¿Cuántas páginas tiene cada documento?", "Número de páginas", 1)
carpeta = InputBox("Copie aquí la dirección de la carpeta destino. Por ejemplo: ", "Carpeta destino", "c:\temp\")
maximoDocs = InputBox("Cuantos documentos a generar?", "Cantidad Documentos", 0)

If (maximoDocs > 0) Then
    totalPaginas = paginasDocumento * maximoDocs
End If


pagActual = 1
Set miRango = ActiveDocument.Content
Do While pagActual <= totalPaginas
    miRango.Find.Execute FindText:="(==)*(==)", MatchWildcards:=True
    If (miRango.Find.Found = True) Then
        miRango.Bold = True
        nombreDocs = miRango
        miRango.Delete
                nombreDocs = Replace(nombreDocs, "==", "")
       'MsgBox (nombreDocs)
    End If
    ActiveDocument.ExportAsFixedFormat OutputFileName:= _
        carpeta & "\" & nombreDocs & ".pdf", ExportFormat:= _
        wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportFromTo, From:=pagActual, To:=pagActual + paginasDocumento - 1, Item:= _
        wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    ChangeFileOpenDirectory carpeta
pagActual = pagActual + paginasDocumento
Loop
MsgBox ("Generación Terminada.")
End Sub

