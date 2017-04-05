Attribute VB_Name = "ModIdioma"
Public lstTextos(30) As String

Public Sub CambiarIdioma(nLenguaje As Integer)
Select Case nLenguaje
    Case 0: Call IdiomaEspanol 'Cambiar a idioma español
    Case 1: Call IdiomaIngles 'Cambiar a idioma español
End Select

FrmFrameWork.CmbDesde_Click
End Sub

Private Sub IdiomaEspanol()
Dim auxInt As Integer
With FrmFrameWork
    .MnuArchivo.Caption = "Archivo"
    .MnuIdioma.Caption = "Idioma"
    .MnuSalir.Caption = "Salir"
    .MnuAyuda.Caption = "Ayuda"
    .MnuAcercade.Caption = "Acerca de ..."
    
    .FrameNombres.Caption = "Obtner lista de nombres"
    .Label5.Caption = "Obtener desde"
    .Label12.Caption = "Nodo"
    .Label22.Caption = "Propiedad"
    .Label1.Caption = "Extensión:"
    .Label2.Caption = "Ruta"
    auxInt = .CmbDesde.ListIndex
    .CmbDesde.Clear
    .CmbDesde.AddItem "Directorio"
    .CmbDesde.AddItem "Texto plano(.txt)"
    .CmbDesde.AddItem "XML (.xml)"
    .CmbDesde.ListIndex = auxInt
    
    .FrameArchivo.Caption = "Directorio de archivos a renombrar"
    .Label4.Caption = "Extensión:"
    .Label3.Caption = "Ruta"
    
    .Frame3.Caption = "Parametros"
    .lblAlgoritmo.Caption = "Algoritmo"
    .Label10.Caption = "Similitud Min"
    .Label14.Caption = "Repetir SNAP al calcular"
    .Label21.Caption = "Permitir SNAP repetidas al renombrar"
    .LblCheckTurbo.Caption = "Acelerar busqueda en listas ordenadas"
    
    .FrameResultados.Caption = "Resultados"
    .LblTitulo(0).Caption = "Usar"
    .LblTitulo(1).Caption = "Nombre original"
    .LblTitulo(2).Caption = "Archivos a renombrar"
    
    '.fraFiltro.Caption = "Filtros"
    '.LblFiltro(0).Caption = "Todos"
    '.LblFiltro(1).Caption = "Con error"
    '.LblFiltro(2).Caption = "Sin match"
    '.LblFiltro(3).Caption = "Repetidos"
    '.LblFiltro(4).Caption = "Sin usar"
    .LblFiltros.Caption = "Filtro"
    .BtnLstFiltros(0).Caption = "Todos"
    .BtnLstFiltros(1).Caption = "Con error"
    .BtnLstFiltros(2).Caption = "Sin match"
    .BtnLstFiltros(3).Caption = "Repetidos"
    .BtnLstFiltros(4).Caption = "Sin usar"
    .BtnLstFiltros(5).Caption = "Marcados para usar"
    .BtnFiltros.Caption = .BtnLstFiltros(.Filtro).Caption
    
    .LblOrdenar.Caption = "Ordenar por:"
    .BtnLstOrdenar(0).Caption = "Ninguno"
    .BtnLstOrdenar(1).Caption = "Nombre"
    .BtnLstOrdenar(2).Caption = "Archivo"
    .BtnLstOrdenar(3).Caption = "%"
    .BtnOrdenar.Caption = .BtnLstOrdenar(.Ordenar).Caption
    .BtnBuscar.Caption = "Buscar"
    
    .BtnCalcular.Caption = "Calcular"
    .BtnRenombrar.Caption = "Acciones"
    
    .BtnCancelarLstArchivos.Caption = "Cancelar"
    
    .Label17.Caption = "Acciones"
    .LblRom.Caption = "Nombre"
    .Label19.Caption = "Archivos"
    .Label18.Caption = "a usar"
    .Label24.Caption = "a usar"
    .Label25.Caption = "sin coincidencia"
    .Label28.Caption = "sin coincidencia"
    .Label26.Caption = "totales"
    .Label29.Caption = "totales"
    .Label20.Caption = "total ARCHIVOS a ser renombrados más de una vez"
    .FrmAcciones.Caption = "Acciones"
    .BtnRenombrarSNAP.Caption = "Renombrar archivos"
    .BtnMoverSNAP.Caption = "Mover archivos"
    .Check4.Caption = "Mover también las ROM"
    .Check5.Caption = "Mantener copia de los archivos"
    .frmExportarTexto.Caption = "Exportar a archivo de texto"
    .BtnExportarNombresUsados.Caption = "Nombres usados"
    .BtnExportarNombresNOUsados.Caption = "Nombres no usados"
    .BtnExportarNombres.Caption = "Todos los nombres"
    .BtnExportarArchivosUsados.Caption = "Archivos usados"
    .BtnExportarArchivosNOusados.Caption = "Archivos no usados"
    .BtnExportarArchivos.Caption = "Todos los archivos"
    .BtnTerminar.Caption = "Cerrar"
    
    lstTextos(0) = "La ruta para obtener los nombres no existe"
    lstTextos(1) = "El archivo para obtener los nombres no existe"
    lstTextos(2) = "El campo NODO no puede estar vacio para archivos XML"
    lstTextos(3) = "La ruta SNAP no existe"
    lstTextos(4) = "La similitud mínima debe ser un número entre 0 y 100"
    lstTextos(5) = "La ruta asginada no es valida"
    lstTextos(6) = "No se encontraron archivos a renombrar"
    lstTextos(7) = "No encontrado"
    lstTextos(8) = "No se encontraron nombres validos"
    lstTextos(9) = "Se ha terminado de exportar"
    lstTextos(10) = "puede encontrar la lista en: "
    lstTextos(11) = "No se puede renombrar archivos si hay coincidencias repetidas" & vbCrLf & _
    "Marque la opción de mover en vez de renombrar" & vbCrLf & _
    "o resuelva los coflictos de coincidencia"
    lstTextos(12) = "Se ha completado el proceso de renombrado"
    lstTextos(13) = "Completado"
    lstTextos(14) = "Se ha completado el proceso de renombrado" & vbCrLf & vbCrLf & "ERROR" & vbCrLf & _
    "Se detectaron " & ErrorCount & " errores durante el proceso" & vbCrLf & _
    "Revise el log.txt para más información"
    lstTextos(15) = "La actual lista de nombres ha cambiado y no refleja" & vbCrLf & _
"Los arcivos reales" & vbCrLf & _
"¿Desea recalcular la lista?"
    lstTextos(16) = "Recalcular"
    lstTextos(17) = "ingrese nombre del archivo"
    lstTextos(18) = "Exportar lista"
    lstTextos(19) = ""
    lstTextos(20) = "Extensión:"
    lstTextos(21) = "Los nombres tienen extensión"
    
    lstTextos(22) = "Todos"
    lstTextos(23) = "Sólo con error"
    lstTextos(24) = "Sin match"
    lstTextos(25) = "Repetidos"
    lstTextos(26) = "Sin usar"
    lstTextos(27) = "Marcados para usar"
    lstTextos(28) = "Ingrese texto a buscar en Nombres."
    lstTextos(29) = "Buscar"
    lstTextos(30) = "No se han encontrado elementos que coincidan con los criterios de la busqueda."
End With
End Sub

Private Sub IdiomaIngles()
With FrmFrameWork
    .MnuArchivo.Caption = "File"
    .MnuIdioma.Caption = "Language"
    .MnuSalir.Caption = "Exit"
    .MnuAyuda.Caption = "Help"
    .MnuAcercade.Caption = "About ..."
    
    .FrameNombres.Caption = "Name list"
    .Label5.Caption = "Get from"
    .Label12.Caption = "Node"
    .Label22.Caption = "Attribute"
    .Label1.Caption = "File Extension:"
    .Label2.Caption = "Path"
    auxInt = .CmbDesde.ListIndex
    .CmbDesde.Clear
    .CmbDesde.AddItem "Folder"
    .CmbDesde.AddItem "Plain text(.txt)"
    .CmbDesde.AddItem "XML (.xml)"
    .CmbDesde.ListIndex = auxInt
    
    .FrameArchivo.Caption = "Folder to rename"
    .Label4.Caption = "File Extension:"
    .Label3.Caption = "Path"
    
    .Frame3.Caption = "Settings"
    .lblAlgoritmo.Caption = "Algorithm"
    .Label10.Caption = "Min Similarity"
    .Label14.Caption = "Allow repeated files when calculating"
    .Label21.Caption = "Allow repeated files"
    .LblCheckTurbo.Caption = "Speed up search"
    
    .FrameResultados.Caption = "Results"
    .LblTitulo(0).Caption = "Check"
    .LblTitulo(1).Caption = "Names"
    .LblTitulo(2).Caption = "Files"
    
    '.fraFiltro.Caption = "Filters"
    
    .LblFiltros.Caption = "Filter"
    .BtnLstFiltros(0).Caption = "All"
    .BtnLstFiltros(1).Caption = "With error"
    .BtnLstFiltros(2).Caption = "Miss Match"
    .BtnLstFiltros(3).Caption = "Repeated"
    .BtnLstFiltros(4).Caption = "No selection"
    .BtnLstFiltros(5).Caption = "With check mark"
    .BtnFiltros.Caption = .BtnLstFiltros(.Filtro).Caption
    
    .LblOrdenar.Caption = "Sort for:"
    .BtnLstOrdenar(0).Caption = "None"
    .BtnLstOrdenar(1).Caption = "Name"
    .BtnLstOrdenar(2).Caption = "File"
    .BtnLstOrdenar(3).Caption = "%"
    .BtnOrdenar.Caption = .BtnLstOrdenar(.Ordenar).Caption
    .BtnBuscar.Caption = "Search"
    
    .BtnCalcular.Caption = "Compare"
    .BtnRenombrar.Caption = "Rename"
    
    .BtnCancelarLstArchivos.Caption = "Cancel"
    
    .Label17.Caption = "Actions"
    .LblRom.Caption = "Names"
    .Label19.Caption = "Files"
    .Label18.Caption = "to be used"
    .Label24.Caption = "to be used"
    .Label25.Caption = "miss match"
    .Label28.Caption = "miss match"
    .Label26.Caption = "total"
    .Label29.Caption = "total"
    .Label20.Caption = "files to be renamed more than one time"
    .FrmAcciones.Caption = "Actions"
    .BtnRenombrarSNAP.Caption = "Rename files"
    .BtnMoverSNAP.Caption = "Move files"
    .Check4.Caption = "Move ROM too"
    .Check5.Caption = "Keep a copy"
    .frmExportarTexto.Caption = "Export lists to text file"
    .BtnExportarNombresUsados.Caption = "Used name"
    .BtnExportarNombresNOUsados.Caption = "not used name"
    .BtnExportarNombres.Caption = "all name"
    .BtnExportarArchivosUsados.Caption = "used files"
    .BtnExportarArchivosNOusados.Caption = "not used files"
    .BtnExportarArchivos.Caption = "all files"
    .BtnTerminar.Caption = "Close"
    
    lstTextos(0) = "The path for names list is invalid"
    lstTextos(1) = "The file for names list is invalid"
    lstTextos(2) = "The node field cannot be empty when select generated names list from XML"
    lstTextos(3) = "The path for files list is invalid"
    lstTextos(4) = "The minimum similarity must have a number between 0 and 100"
    lstTextos(5) = "Invalid path"
    lstTextos(6) = "Not found files to rename"
    lstTextos(7) = "Not found"
    lstTextos(8) = "Not found valid names list"
    lstTextos(9) = "Export finished"
    lstTextos(10) = "The export file can be found in: "
    lstTextos(11) = "Can not rename files if there are repeated matches" & vbCrLf & _
    "Use move option instead " & vbCrLf & _
    "or solve repeat files."
    lstTextos(12) = "Rename files process finished"
    lstTextos(13) = "Finished"
    lstTextos(14) = "Rename files process finished" & vbCrLf & vbCrLf & "ERROR" & vbCrLf & _
    "Total error: " & ErrorCount & " " & vbCrLf & _
    "Check log.txt for more information"
    lstTextos(15) = "Want update names and files lists?"
    lstTextos(16) = "Recalculate"
    lstTextos(17) = "Select name for export file"
    lstTextos(18) = "Export"
    lstTextos(19) = ""
    lstTextos(20) = "Extension:"
    lstTextos(21) = "Search filename extension"
    
    lstTextos(22) = "All"
    lstTextos(23) = "With error"
    lstTextos(24) = "Miss Match"
    lstTextos(25) = "Repeated"
    lstTextos(26) = "No selection"
    lstTextos(27) = "With check mark"
    lstTextos(28) = "Find What:"
    lstTextos(29) = "Find"
    lstTextos(30) = "Can't find the text."
End With
End Sub
