VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "AVANZAR 2018"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11025
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function loadrow() As Double
On Error GoTo lastrow
ref = ComboBox100.Value & ComboBox101.Value & ComboBox102.Value & ComboBox103.Value & TextBox2.Value
loadrow = Application.WorksheetFunction.Match(ref, Worksheets("DATOS CARGADOS").Range("EK:EK"), 0)
Exit Function
lastrow:
loadrow = Worksheets("DATOS CARGADOS").Range("C50000").End(xlUp).Row + 1
End Function
Function loadrow_UE() As Double
On Error GoTo lastrow
ref = ComboBox100.Value & ComboBox101.Value & ComboBox102.Value & ComboBox103.Value
loadrow_UE = Application.WorksheetFunction.Match(ref, Worksheets("UNIDADES DE EVALUACION").Range("P:P"), 0)
Exit Function
lastrow:
loadrow_UE = Worksheets("UNIDADES DE EVALUACION").Range("B1000").End(xlUp).Row + 1
End Function
Private Sub ComboBox101_Change()
If ComboBox101.Value = "" Then
 TextBox8.Value = ""
 TextBox9.Value = ""
 TextBox15.Value = ""
End If
If ComboBox101.Value <> "" And ComboBox102.Value <> "" And ComboBox103.Value <> "" Then
 On Error GoTo Err1
 m = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value & ComboBox103.Value, Worksheets("REF").Range("O:Q"), 3, False)
 TextBox9.Value = m
 Else
End If
If ComboBox101.Value <> "" And ComboBox102.Value <> "" Then
On Error GoTo Err
n = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value, Worksheets("REF").Range("K:P"), 2, False)
TextBox8.Value = n
Exit Sub
Else
End If
Exit Sub
Err:
TextBox8.Value = ""
TextBox9.Value = ""
MsgBox "Distrito y Sede inconsistentes."
Err1:
TextBox9.Value = ""
TextBox8.Value = ""
MsgBox "Distrito, Sede y Unidad de Evaluación inconsistentes."
End Sub
Private Sub ComboBox102_AfterUpdate()
If ComboBox102.Value = "" Then
 TextBox8.Value = ""
 TextBox9.Value = ""
 TextBox15.Value = ""
End If
If ComboBox101.Value <> "" And ComboBox102.Value <> "" And ComboBox103.Value <> "" Then
 On Error GoTo Err1
 m = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value & ComboBox103.Value, Worksheets("REF").Range("O:Q"), 3, False)
 TextBox9.Value = m
 Else
End If
If ComboBox101.Value <> "" And ComboBox102.Value <> "" Then
On Error GoTo Err
n = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value, Worksheets("REF").Range("K:P"), 2, False)
TextBox8.Value = n
q = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value, Worksheets("REF").Range("K:T"), 10, False)
TextBox15.Value = q
Else
End If
Exit Sub
Err:
TextBox8.Value = ""
TextBox9.Value = ""
TextBox15.Value = ""
ComboBox102.Value = ""
MsgBox "Distrito y Sede inconsistentes."
Exit Sub
Err1:
TextBox8.Value = ""
TextBox9.Value = ""
TextBox15.Value = ""
ComboBox102.Value = ""
ComboBox103.Value = ""
MsgBox "Distrito, Sede y Unidad de Evaluación inconsistentes."
End Sub
Private Sub ComboBox103_Change()
If ComboBox103.Value = "" Then
 TextBox9.Value = ""
End If
If ComboBox101.Value <> "" And ComboBox102.Value <> "" And ComboBox103.Value <> "" Then
On Error GoTo Err1
m = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value & ComboBox103.Value, Worksheets("REF").Range("O:Q"), 3, False)
TextBox9.Value = m
TextBox14.Value = Application.WorksheetFunction.VLookup(ComboBox101.Value & ComboBox102.Value & ComboBox103.Value, Worksheets("REF").Range("O:R"), 4, False)
Exit Sub
Else
End If
Exit Sub
Err1:
TextBox9.Value = ""
TextBox14.Value = ""
ComboBox103.Value = ""
MsgBox "Distrito, Sede y Unidad de Evaluación inconsistentes."
End Sub
Private Sub ComboBox104_AfterUpdate()
If ComboBox104.ListIndex < 0 Then
   ComboBox104.Value = ""
   ComboBox104.SetFocus
   MsgBox "'CUE' erroneo.", , "Error"
   Exit Sub
 End If
End Sub
Private Sub ComboBox65_AfterUpdate()
 If ComboBox65.Value = "a" Then
  ComboBox66.Enabled = True
  ComboBox67.Enabled = True
  ComboBox68.Enabled = True
 ElseIf ComboBox65.Value = "b" Then
  ComboBox66.Enabled = False
  ComboBox67.Enabled = False
  ComboBox68.Enabled = False
 Else
  ComboBox65.Value = ""
  ComboBox65.SetFocus
  MsgBox "Elegí entre las opciones."
 End If
End Sub
Private Sub ComboBox67_AfterUpdate()
If ComboBox67.Value = "a" Then
 ComboBox68.Enabled = True
ElseIf ComboBox67.Value = "b" Then
 ComboBox68.Enabled = False
Else
 ComboBox68.Enabled = True
 ComboBox68.Value = ""
 ComboBox68.SetFocus
 MsgBox "Elegí entre las opciones."
End If
End Sub
Private Sub ComboBox86_AfterUpdate()
If ComboBox86.Value = "a" Then
 ComboBox87.Enabled = False
Else
 ComboBox87.Enabled = True
End If
End Sub
Private Sub ComboBox87_AfterUpdate()
If ComboBox87.Value = "b" Then
 ComboBox88.Enabled = False
 ComboBox89.Enabled = False
Else
 ComboBox88.Enabled = True
 ComboBox89.Enabled = True
End If
End Sub
Private Sub ComboBox90_AfterUpdate()
If ComboBox90.Value = "a" Then
 ComboBox91.Enabled = False
Else
 ComboBox91.Enabled = True
End If
End Sub
Private Sub CommandButton2_Click() 'limpiar
Dim ctrl As Control
Dim nombre As String
'For Each ctrl In UserForm1.Controls
 'If TypeName(ctrl) = "ComboBox" Then
  'ctrl.Value = ""
 'End If
'Next ctrl
 For n = 1 To 38
  Controls("ComboBox" & n).Value = ""
 Next n
 TextBox3.Value = ""
End Sub
Private Sub CommandButton3_Click() 'actualizar, disabled
Sheets("DATOS CARGADOS").Unprotect
If TextBox2.Value = "" Then
MsgBox "Pone un 'N° de Alumno' para ver lo que tiene cargado."
Else
Range("B1").Value = TextBox2.Value
Range("A1").FormulaR1C1 = "=+MATCH(RC[1],C[2],0)"
If IsNumeric(Range("A1").Value) Then
 rw = Range("A1").Value
Else
 MsgBox "El 'N° de Alumno' consultado no tiene datos cargados."
 Range("B1").Value = ""
 Range("A1").Value = ""
 Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
 Exit Sub
End If
Range("B1").Value = ""
Range("A1").Value = ""
TextBox1.Value = Cells(rw, 2).Value
TextBox2.Value = Cells(rw, 3).Value
TextBox3.Value = Cells(rw, 4).Value
For n = 1 To 38
Controls("ComboBox" & n).Value = Cells(rw, n + 4).Value
Next n
End If
Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Private Sub CommandButton4_Click() 'verificar y cargar
Sheets("DATOS CARGADOS").Unprotect
Dim ctrl As Control
Dim nombre As String
Dim i, r As Integer
i = 0
For Each ctrl In UserForm1.Controls
 If TypeName(ctrl) = "ComboBox" Then
 If Right(ctrl.Name, Len(ctrl.Name) - 8) <= 58 And Right(ctrl.Name, Len(ctrl.Name) - 8) >= 39 Then
  If ctrl.ListIndex < 0 And ctrl.Value <> "" Then
   ctrl.Value = ""
   'nombre = ctrl.Name
   ctrl.SetFocus
   MsgBox ("Cargaste mal.")
   Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
   Exit Sub
  End If
  If ctrl.Value = "" And ctrl.Enabled = True Then
  i = i + 1
  End If
  End If
 End If
Next ctrl
If TextBox4.Value <> "" Then
 If i > 0 Then
  response = MsgBox("Dejaste sin cargar " & i & " campos.¿Continuar?", vbYesNo, "Campos Vacíos")
 Else
  response = MsgBox("Cargaste todos los campos.¿Continuar?", vbYesNo, "Carga")
 End If
End If
If response = vbNo Then
    Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub
End If
If TextBox2.Value <> "" And TextBox4.Value <> "" Then
 rw = loadrow
 Worksheets("DATOS CARGADOS").Cells(rw, 2).Value = ComboBox104.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 3).Value = TextBox2.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 44).Value = TextBox4.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 135).Value = ComboBox100.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 136).Value = ComboBox101.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 141).Value = ComboBox100.Value & ComboBox101.Value & ComboBox102.Value & ComboBox103.Value & TextBox2.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 137).Value = TextBox15.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 138).Value = ComboBox102.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 139).Value = ComboBox103.Value
 For n = 39 To 58
 Worksheets("DATOS CARGADOS").Cells(loadrow, n + 6).Value = Controls("ComboBox" & n).Value
 Next n
 Else
 MsgBox "'N° de Alumno' quedó en blanco."
 Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
 TextBox2.SetFocus
 Exit Sub
End If
For n = 39 To 58
 Controls("ComboBox" & n).Value = ""
Next n
TextBox2.Value = ""
TextBox4.Value = ""
TextBox2.SetFocus
Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Private Sub CommandButton5_Click()
For n = 39 To 58
 Controls("ComboBox" & n).Value = ""
Next n
TextBox4.Value = ""
End Sub

Private Sub CommandButton7_Click()
 response = MsgBox("¿Seguro?", vbYesNo, "Ojo")
 If response = vbNo Then
    Exit Sub
 End If
 TextBox5.Value = ""
 TextBox6.Value = ""
 TextBox7.Value = ""
 ComboBox59.Value = ""
 ComboBox60.Value = ""
 ComboBox61.Value = ""
 ComboBox62.Value = ""
 For n = 63 To 99 'comboboxes en secuencia
 Controls("ComboBox" & n).Value = ""
 Next n
 For Each ctrl In UserForm1.Controls
  If TypeName(ctrl) = "ComboBox" Then
   If Right(ctrl.Name, Len(ctrl.Name) - 8) <= 99 And Right(ctrl.Name, Len(ctrl.Name) - 8) >= 63 Then
    ctrl.Enabled = True
   End If
  End If
 Next ctrl
 For Each ctrl In UserForm1.Controls
    If TypeName(ctrl) = "CheckBox" Then
     ctrl.Value = False
    End If
   Next ctrl
End Sub
Private Sub CommandButton8_Click()
Sheets("DATOS CARGADOS").Unprotect
Dim ctrl As Control
Dim nombre As String
Dim i, r As Integer
i = 0
For Each ctrl In UserForm1.Controls
 If TypeName(ctrl) = "ComboBox" Then
  If Right(ctrl.Name, Len(ctrl.Name) - 8) >= 59 And Right(ctrl.Name, Len(ctrl.Name) - 8) <= 99 Then
  If ctrl.ListIndex < 0 And ctrl.Value <> "" Then
   ctrl.Value = ""
   'nombre = ctrl.Name
   ctrl.SetFocus
   MsgBox ("Cargaste mal.")
   Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
   Exit Sub
  End If
  If ctrl.Value = "" And ctrl.Enabled = True Then
  i = i + 1
  End If
  End If
 End If
Next ctrl
If TextBox5.Value = "" Then i = i + 1
If TextBox6.Value = "" Then i = i + 1
If TextBox7.Value = "" Then i = i + 1
 If i > 0 Then
  response = MsgBox("Dejaste sin cargar " & i & " campos.¿Continuar?", vbYesNo, "Campos Vacíos")
 Else
  response = MsgBox("Cargaste todos los campos.¿Continuar?", vbYesNo, "Carga")
 End If
If response = vbNo Then
    Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub
End If
If TextBox2.Value <> "" Then
 rw = loadrow
 Worksheets("DATOS CARGADOS").Cells(rw, 2).Value = ComboBox104.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 3).Value = TextBox2.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 135).Value = ComboBox100.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 136).Value = ComboBox101.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 141).Value = ComboBox100.Value & ComboBox101.Value & ComboBox102.Value & ComboBox103.Value & TextBox2.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 137).Value = TextBox15.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 138).Value = ComboBox102.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 139).Value = ComboBox103.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 67).Value = TextBox5.Value 'textboxes
 Worksheets("DATOS CARGADOS").Cells(rw, 69).Value = TextBox6.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 96).Value = TextBox7.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 66).Value = ComboBox59.Value 'comboboxes fuera de secuencia
 Worksheets("DATOS CARGADOS").Cells(rw, 68).Value = ComboBox60.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 79).Value = ComboBox61.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 95).Value = ComboBox62.Value
 For n = 63 To 99 'comboboxes en secuencia
 Worksheets("DATOS CARGADOS").Cells(rw, n + 34) = Controls("ComboBox" & n).Value
 Next n
 C = 0 'aux
 d = 0
   For Each ctrl In UserForm1.Controls
    If TypeName(ctrl) = "CheckBox" Then
     If Len(ctrl.Name) = 9 Then
      Worksheets("DATOS CARGADOS").Cells(rw, 70 + C).Value = ctrl.Value
      C = C + 1
     ElseIf Len(ctrl.Name) = 10 Then
      Worksheets("DATOS CARGADOS").Cells(rw, 80 + d).Value = ctrl.Value
      d = d + 1
     Else
     End If
    End If
   Next ctrl
Else
MsgBox "'N° de Alumno' quedó en blanco."
Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
TextBox2.SetFocus
Exit Sub
End If
'clean
 TextBox5.Value = ""
 TextBox6.Value = ""
 TextBox7.Value = ""
 ComboBox59.Value = ""
 ComboBox60.Value = ""
 ComboBox61.Value = ""
 ComboBox62.Value = ""
 For n = 63 To 99 'comboboxes en secuencia
 Controls("ComboBox" & n).Value = ""
 Next n
 For Each ctrl In UserForm1.Controls
  If TypeName(ctrl) = "ComboBox" Then
   If Right(ctrl.Name, Len(ctrl.Name) - 8) <= 99 And Right(ctrl.Name, Len(ctrl.Name) - 8) >= 63 Then
    ctrl.Enabled = True
   End If
  End If
 Next ctrl
 For Each ctrl In UserForm1.Controls
    If TypeName(ctrl) = "CheckBox" Then
     ctrl.Value = False
    End If
   Next ctrl
TextBox2.Value = ""
TextBox2.SetFocus
Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Private Sub CommandButton9_Click() 'datos extra
 If ComboBox100.ListIndex < 0 Then
   ComboBox100.Value = ""
   ComboBox100.SetFocus
   MsgBox "Elegí una región", , "Error"
   Exit Sub
 End If
 If ComboBox101.ListIndex < 0 Then
   ComboBox101.Value = ""
   ComboBox101.SetFocus
   MsgBox "Elegí un distrito", , "Error"
   Exit Sub
 End If
 If ComboBox102.ListIndex < 0 Then
   ComboBox102.Value = ""
   ComboBox102.SetFocus
   MsgBox "Elegí una sede de la lista.", , "Error"
   Exit Sub
 End If
 If ComboBox103.ListIndex < 0 Then
   ComboBox103.Value = ""
   ComboBox103.SetFocus
   MsgBox "Elegí una 'UNIDAD DE EVALUACIÓN' de la lista.", , "Error"
   Exit Sub
 End If
 If ComboBox104.ListIndex < 0 Then
   ComboBox104.Value = ""
   ComboBox104.SetFocus
   MsgBox "Elegí 'CUE' de la lista.", , "Error"
   Exit Sub
 End If
Sheets("UNIDADES DE EVALUACION").Unprotect
rw = loadrow_UE
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 2).Value = ComboBox100.Value 'REGION
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 3).Value = ComboBox101.Value  'DISTRITO
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 4).Value = ComboBox104.Value  'CUE
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 5).Value = TextBox15.Value
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 6) = ComboBox102.Value 'SEDE
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 7) = ComboBox103.Value 'UE
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 14) = TextBox14.Value 'APLICADOR
TextBox13.Value = Worksheets("UNIDADES DE EVALUACION").Cells(rw, 9).Value 'ESTUDIANTES CARGADOS
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 16) = ComboBox100.Value & ComboBox101.Value & ComboBox102.Value & ComboBox103.Value 'ref
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 11) = TextBox10.Value 'blancos
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 12) = TextBox11.Value
Worksheets("UNIDADES DE EVALUACION").Cells(rw, 13) = TextBox12.Value
Sheets("UNIDADES DE EVALUACION").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Private Sub MultiPage1_Click(ByVal Index As Long)
'If ComboBox100.Value <> "" And ComboBox101.Value <> "" And ComboBox102.Value <> "" And ComboBox103.Value <> "" Then
If TextBox8.Value <> "" And TextBox9.Value <> "" And ComboBox104.Value <> "" Then
MultiPage1.Pages(1).Enabled = True
MultiPage1.Pages(2).Enabled = True
MultiPage1.Pages(3).Enabled = True
TextBox2.Enabled = True
Else
MultiPage1.Pages(1).Enabled = False
MultiPage1.Pages(2).Enabled = False
MultiPage1.Pages(3).Enabled = False
TextBox2.Enabled = False
End If
End Sub
Private Sub TextBox2_AfterUpdate()
TextBox2.Value = Trim(TextBox2.Value)
End Sub
Private Sub TextBox4_Change()
For n = 39 To 58
Controls("ComboBox" & n).Enabled = True
Next n
 If TextBox4.Value = 1 Then
  ComboBox54.Enabled = False
 ElseIf TextBox4.Value = 3 Then
  ComboBox53.Enabled = False
 ElseIf TextBox4.Value = 2 Or TextBox4.Value = 4 Then
  Exit Sub
 ElseIf TextBox4.Value = "" Then
  For n = 39 To 58
   Controls("ComboBox" & n).Enabled = False
  Next n
 Else
  TextBox4.Value = ""
  MsgBox "Ingresa un valor entre '1','2','3' y '4'.", , "Pancho"
 End If
End Sub
Private Sub TextBox5_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
If IsDate(Trim(TextBox5.Text)) = False Then
TextBox5.Value = ""
MsgBox "Pone bien la fecha."
TextBox5.SetFocus
End If
TextBox5.Text = Format(TextBox5.Text, "dd mmmm yyyy")
End Sub
Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 If IsNumeric(TextBox6.Value) = False And TextBox6.Value <> "" Then
        TextBox6.Value = ""
        MsgBox "Ingresá un valor numérico"
        Cancel = True
    End If
End Sub
Private Sub TextBox7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 If IsNumeric(TextBox7.Value) = False And TextBox7.Value <> "" Then
        TextBox7.Value = ""
        MsgBox "Ingresá un valor numérico"
        Cancel = True
    End If
End Sub
Private Sub UserForm_Initialize()
Dim cLoc As Range
Dim ws As Worksheet
Dim ctrl As Control
Set wsref = Worksheets("LookupLists")
For Each cLoc In wsref.Range("A2:A8")
  For Each ctrl In UserForm1.Controls
   If TypeName(ctrl) = "ComboBox" Then
    If Right(ctrl.Name, Len(ctrl.Name) - 8) <= 58 Then
     With ctrl
      .AddItem cLoc.Value
     End With
    End If
   End If
  Next ctrl
Next cLoc
 For Each ctrl In UserForm1.Controls
  If TypeName(ctrl) = "ComboBox" Then
   If Right(ctrl.Name, Len(ctrl.Name) - 8) <= 58 Then
    ctrl.Enabled = False
   End If
  End If
 Next ctrl
'verdadero/falsos en matematica
ComboBox8.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox9.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox10.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox11.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox14.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox15.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox16.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox17.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox27.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox28.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox29.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox30.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox33.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox34.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox35.List = Worksheets("LookupLists").Range("O2:O6").Value
ComboBox36.List = Worksheets("LookupLists").Range("O2:O6").Value
'Cuestionario del estudiante, poblado de comboboxes <=====8
ComboBox59.AddItem "Masculino"
ComboBox59.AddItem "Femenino"
ComboBox59.AddItem "Otro"
ComboBox60.AddItem "En provincia de Buenos Aires"
ComboBox60.AddItem "En otra provincia"
ComboBox60.AddItem "En otro pais"
ComboBox61.AddItem "1"
ComboBox61.AddItem "2"
ComboBox61.AddItem "3"
ComboBox61.AddItem "4"
ComboBox61.AddItem "5"
ComboBox61.AddItem "6"
ComboBox61.AddItem "7 o más"
ComboBox61.AddItem "No sé"
ComboBox62.AddItem "a"
ComboBox62.AddItem "b"
ComboBox62.AddItem "c"
ComboBox63.List = Worksheets("LookupLists").Range("B2:B11").Value
ComboBox64.List = Worksheets("LookupLists").Range("B2:B11").Value
ComboBox65.AddItem "a"
ComboBox65.AddItem "b"
ComboBox66.AddItem "a"
ComboBox66.AddItem "b"
ComboBox66.AddItem "c"
ComboBox67.AddItem "a"
ComboBox67.AddItem "b"
ComboBox68.List = Worksheets("LookupLists").Range("B2:B9").Value
ComboBox69.List = Worksheets("LookupLists").Range("C2:C5").Value
ComboBox70.List = Worksheets("LookupLists").Range("C2:C5").Value
ComboBox71.List = Worksheets("LookupLists").Range("C2:C5").Value
ComboBox72.List = Worksheets("LookupLists").Range("B2:B5").Value
ComboBox73.List = Worksheets("LookupLists").Range("B2:B5").Value
ComboBox74.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox75.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox76.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox77.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox78.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox79.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox80.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox81.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox82.List = Worksheets("LookupLists").Range("D2:D7").Value
ComboBox83.List = Worksheets("LookupLists").Range("B2:B9").Value
ComboBox84.List = Worksheets("LookupLists").Range("B2:B9").Value
ComboBox85.List = Worksheets("LookupLists").Range("B2:B3").Value
ComboBox86.List = Worksheets("LookupLists").Range("B2:B4").Value
ComboBox87.List = Worksheets("LookupLists").Range("B2:B3").Value
ComboBox88.List = Worksheets("LookupLists").Range("B2:B5").Value
ComboBox89.List = Worksheets("LookupLists").Range("B2:B3").Value
ComboBox90.List = Worksheets("LookupLists").Range("B2:B3").Value
ComboBox91.List = Worksheets("LookupLists").Range("B2:B3").Value
ComboBox92.List = Worksheets("LookupLists").Range("B2:B5").Value
ComboBox93.List = Worksheets("LookupLists").Range("B2:B3").Value
ComboBox94.List = Worksheets("LookupLists").Range("E2:E4").Value
ComboBox95.List = Worksheets("LookupLists").Range("E2:E4").Value
ComboBox96.List = Worksheets("LookupLists").Range("E2:E4").Value
ComboBox97.List = Worksheets("LookupLists").Range("E2:E4").Value
ComboBox98.List = Worksheets("LookupLists").Range("E2:E4").Value
ComboBox99.List = Worksheets("LookupLists").Range("B2:B6").Value
ComboBox100.List = Worksheets("LookupLists").Range("F2:F26").Value
ComboBox101.List = Worksheets("LookupLists").Range("G2:G82").Value
ComboBox102.List = Worksheets("LookupLists").Range("I2:I176").Value
ComboBox103.List = Worksheets("LookupLists").Range("M2:M10").Value
ComboBox104.List = Worksheets("Lookuplists").Range("K2:K98").Value
MultiPage1.Pages(1).Enabled = False
MultiPage1.Pages(2).Enabled = False
MultiPage1.Pages(3).Enabled = False
TextBox2.Enabled = False
End Sub
Private Sub TextBox3_Change()
For n = 1 To 38
Controls("ComboBox" & n).Enabled = True
Next n
If TextBox3.Value = 1 Then
 ComboBox7.Enabled = True 'la pregunta 7
 ComboBox8.Enabled = False
 ComboBox9.Enabled = False
 ComboBox10.Enabled = False
 ComboBox11.Enabled = False
 ComboBox13.Enabled = True 'la pregunta 9
 ComboBox14.Enabled = False
 ComboBox15.Enabled = False
 ComboBox16.Enabled = False
 ComboBox17.Enabled = False
 ComboBox19.Enabled = True 'la pregunta 10.b
 ComboBox26.Enabled = False 'la pregunta 17
 ComboBox27.Enabled = True
 ComboBox28.Enabled = True
 ComboBox29.Enabled = True
 ComboBox30.Enabled = True
 ComboBox32.Enabled = True 'la pregunta 19
 ComboBox33.Enabled = False
 ComboBox34.Enabled = False
 ComboBox35.Enabled = False
 ComboBox36.Enabled = False
 ComboBox38.Enabled = False 'la pregunta 20.b
ElseIf TextBox3.Value = 2 Then
 ComboBox7.Enabled = False 'la pregunta 7
 ComboBox8.Enabled = True
 ComboBox9.Enabled = True
 ComboBox10.Enabled = True
 ComboBox11.Enabled = True
 ComboBox13.Enabled = True 'la pregunta 9
 ComboBox14.Enabled = False
 ComboBox15.Enabled = False
 ComboBox16.Enabled = False
 ComboBox17.Enabled = False
 ComboBox19.Enabled = False 'la pregunta 10.b
 ComboBox26.Enabled = True 'la pregunta 17
 ComboBox27.Enabled = False
 ComboBox28.Enabled = False
 ComboBox29.Enabled = False
 ComboBox30.Enabled = False
 ComboBox32.Enabled = True 'la pregunta 19
 ComboBox33.Enabled = False
 ComboBox34.Enabled = False
 ComboBox35.Enabled = False
 ComboBox36.Enabled = False
 ComboBox38.Enabled = True 'la pregunta 20.b
ElseIf TextBox3.Value = 3 Then
 ComboBox7.Enabled = True 'la pregunta 7
 ComboBox8.Enabled = False
 ComboBox9.Enabled = False
 ComboBox10.Enabled = False
 ComboBox11.Enabled = False
 ComboBox13.Enabled = True 'la pregunta 9
 ComboBox14.Enabled = False
 ComboBox15.Enabled = False
 ComboBox16.Enabled = False
 ComboBox17.Enabled = False
 ComboBox19.Enabled = True 'la pregunta 10.b
 ComboBox26.Enabled = True 'la pregunta 17
 ComboBox27.Enabled = False
 ComboBox28.Enabled = False
 ComboBox29.Enabled = False
 ComboBox30.Enabled = False
 ComboBox32.Enabled = False 'la pregunta 19
 ComboBox33.Enabled = True
 ComboBox34.Enabled = True
 ComboBox35.Enabled = True
 ComboBox36.Enabled = True
 ComboBox38.Enabled = False 'la pregunta 20.b
ElseIf TextBox3.Value = 4 Then
 ComboBox7.Enabled = True 'la pregunta 7
 ComboBox8.Enabled = False
 ComboBox9.Enabled = False
 ComboBox10.Enabled = False
 ComboBox11.Enabled = False
 ComboBox13.Enabled = False 'la pregunta 9
 ComboBox14.Enabled = True
 ComboBox15.Enabled = True
 ComboBox16.Enabled = True
 ComboBox17.Enabled = True
 ComboBox19.Enabled = False 'la pregunta 10.b
 ComboBox26.Enabled = True 'la pregunta 17
 ComboBox27.Enabled = False
 ComboBox28.Enabled = False
 ComboBox29.Enabled = False
 ComboBox30.Enabled = False
 ComboBox32.Enabled = True 'la pregunta 19
 ComboBox33.Enabled = False
 ComboBox34.Enabled = False
 ComboBox35.Enabled = False
 ComboBox36.Enabled = False
 ComboBox38.Enabled = True 'la pregunta 20.b
ElseIf TextBox3.Value = "" Then
 For n = 1 To 38
  Controls("ComboBox" & n).Enabled = False
 Next n
Exit Sub
Else
TextBox3.Value = ""
MsgBox "Ingresa un valor entre '1','2','3' y '4'.", , "Pancho"
End If
End Sub
Private Sub CommandButton1_Click() 'verificar y cargar
Sheets("DATOS CARGADOS").Unprotect
Dim ctrl As Control
Dim nombre As String
Dim i, r As Integer
i = 0
For Each ctrl In UserForm1.Controls
 If TypeName(ctrl) = "ComboBox" Then
  If Right(ctrl.Name, Len(ctrl.Name) - 8) <= 38 Then
  If ctrl.ListIndex < 0 And ctrl.Value <> "" Then
   ctrl.Value = ""
   ctrl.SetFocus
   MsgBox ("Cargaste mal.")
   Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
   Exit Sub
  End If
  If ctrl.Value = "" And ctrl.Enabled = True Then
  i = i + 1
  End If
  End If
 End If
Next ctrl
If TextBox2.Value <> "" And TextBox3.Value <> "" Then
 If i > 0 Then
  response = MsgBox("Dejaste sin cargar " & i & " campos.¿Continuar?", vbYesNo, "Campos Vacíos")
 Else
  response = MsgBox("Cargaste todos los campos.¿Continuar?", vbYesNo, "Carga")
 End If
 If response = vbNo Then
    Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub
 End If
 rw = loadrow
 Worksheets("DATOS CARGADOS").Cells(rw, 2).Value = ComboBox104.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 3).Value = TextBox2.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 4).Value = TextBox3.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 135).Value = ComboBox100.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 136).Value = ComboBox101.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 141).Value = ComboBox100.Value & ComboBox101.Value & ComboBox102.Value & ComboBox103.Value & TextBox2.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 137).Value = TextBox15.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 138).Value = ComboBox102.Value
 Worksheets("DATOS CARGADOS").Cells(rw, 139).Value = ComboBox103.Value
 For n = 1 To 38
 Worksheets("DATOS CARGADOS").Cells(rw, n + 4).Value = Controls("ComboBox" & n).Value
 Next n
 Else
 MsgBox "'N° de Alumno' quedó en blanco."
 Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
 TextBox2.SetFocus
 Exit Sub
End If
 'clean
 For n = 1 To 38
  Controls("ComboBox" & n).Value = ""
 Next n
 TextBox2.Value = "" 'borrar linea para pato
 TextBox3.Value = ""
 TextBox2.SetFocus
Sheets("DATOS CARGADOS").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
