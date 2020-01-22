Attribute VB_Name = "RepairnConvert"
Function Padre() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Selecciona la carpeta de origen, guacho"
        .AllowMultiSelect = False
        .InitialFileName = Application.ActiveWorkbook.Path
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    Padre = sItem
    Set fldr = Nothing
End Function
Function Mijo() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Selecciona la carpeta de destino, paparrin"
        .AllowMultiSelect = False
        .InitialFileName = Application.ActiveWorkbook.Path
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    Mijo = sItem
    Set fldr = Nothing
End Function

Sub NonRecursiveDrill_Convert()
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  Application.Calculation = xlCalculationManual
Dim fso, oFolder, oSubfolder, oFile, queue, papis, nombres As Collection
Dim Directorio, Nombre, Destino As String
Dim i As Integer
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set queue = New Collection
 Set papis = New Collection
 Set nombres = New Collection
 queue.Add fso.GetFolder(Padre) 'Uso una function para tener un input variable
    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        For Each oSubfolder In oFolder.SubFolders
            queue.Add oSubfolder 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.Files  'levanta todos los archivos sin diferenciar formatos
        Directorio = oFile.Path
        Nombre = Left(oFile.Name, InStr(oFile.Name, "xlsx") - 1) & " - ESTUDIANTES" 'el "ESTUDIANTES" tiene q ser variable independiente (meter userform con combobox por programa)
        papis.Add Directorio 'armo mi coleccion con todas las direcciones de los archivitos
        nombres.Add Nombre  'armo mi coleccion con los nombres, en la misma posicion relativa de las direcciones
        Next oFile
    Loop
  'On Error GoTo errorhandler
  Destino = Mijo & "\"
  For i = 1 To papis.Count
  Workbooks.Open Filename:=papis(i), UpdateLinks:=0, CorruptLoad:=2
  ActiveSheet.Cells.Range("K4").NumberFormat = "General" 'para prevenir error de validacion fecha
  ActiveWorkbook.Sheets("ESTUDIANTES").Activate 'la solapa "Estudiantes", siempre q este bien nombrada y demas... una pija
  ActiveWorkbook.SaveAs Filename:=Destino & nombres(i) & ".csv", FileFormat:=xlCSV, CreateBackup:=False, Local:=True
  ActiveWorkbook.Close False
  Next i 'las iteaciones reflejan la cantidad de archivos
  Application.DisplayAlerts = True
  Application.EnableEvents = True
  Application.DisplayStatusBar = True
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Exit Sub
errorhandler:
  Debug.Print papis(1)
  MsgBox "Error en la planilla " & papis(i) & ", iteracion" & i
End Sub
