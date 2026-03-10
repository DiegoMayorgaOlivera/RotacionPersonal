VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CONTROL_CARNETS 
   Caption         =   "CONTROL DE CARNETS Y CORDONES"
   ClientHeight    =   8820.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17760
   OleObjectBlob   =   "CONTROL_CARNETS.frx":0000
End
Attribute VB_Name = "CONTROL_CARNETS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'====================================================================================================================================
' Formulario para gestionar el control de carnets y cordones de empleados.
' Permite registrar nuevos movimientos relacionados con carnets y cordones, así como consultar los registros existentes.
'====================================================================================================================================

Dim Borrador As Boolean
Option Explicit

'=====================================
' RUTA BASE DE LAS FOTOS
'=====================================
Public Function ObtenerRutaFotos() As String

    ObtenerRutaFotos = ThisWorkbook.Path & "\FOTOS\"
    
End Function


Private Sub UBICACION_ESPECIFICA_NUEVO_ETIQUETA_Click()

End Sub

Private Sub UserForm_Initialize()

    Ocultar
        
        '==========================================================================
        'BUSQUEDA DE MOTIVO DE CARNET
                    
        Dim Motivos, Mot As Range
        
        ' Referenciar la tabla
            Set Motivos = Hoja24.ListObjects("MOTIVO_CARNET").DataBodyRange
            
            Me.MOTIVO.Clear
                    
            ' Agregar cada valor, omitiendo vacios
            For Each Mot In Motivos
                If Trim(Mot.Value) <> "" Then  ' Omite celdas vacias
                    Me.MOTIVO.AddItem Mot.Value
                End If
            Next Mot
        '==========================================================================
        
End Sub


Private Sub UserForm_Activate()

    Me.Top = 0
    Me.REGISTRO.BackColor = RGB(250, 232, 250)
    
End Sub


Private Sub MOTIVO_AfterUpdate()

        MotivoSeleccionado
    
End Sub

Private Sub AGREGAR_NUEVO_REGISTRO_Click()

    If Me.MOTIVO.Value = "CORRECCION" And Me.CODIGO_CARNET_CORDON.Value = "" Then
        MsgBox "Ingrese el codigo del Registro a Corregir", vbDefaultButton1 + vbExclamation, "Dato Incompleto"
        'Me.AGREGAR_NUEVO_REGISTRO.Visible = False
    Else
        MotivoSeleccionado
    End If
    
End Sub

Private Sub MotivoSeleccionado()

Dim Mot As String

Mot = Me.MOTIVO.Value


Select Case True
    
    Case Mot = ""
        Call UserForm_Initialize
        Call UserForm_Activate
        Me.AGREGAR_NUEVO_REGISTRO.Visible = True
        Exit Sub
    
    Case Mot = "TRASLADO Y NOMBRAMIENTO"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = True
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Me.CARGO_NUEVO_ETIQUETA.Visible = True
        Me.CARGO_NUEVO.Visible = True
        Me.UBICACION_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_NUEVO_ETIQUETA.Top = 60
        Me.UBICACION_NUEVO.Visible = True
        Me.UBICACION_NUEVO.Top = 84
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 60
        Me.UBICACION_GENERAL_NUEVO.Visible = True
        Me.UBICACION_GENERAL_NUEVO.Top = 84
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Top = 114
        Me.UBICACION_ESPECIFICA_NUEVO.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO.Top = 138
        Exit Sub
        
    Case Mot = "TRASLADO"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = True
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Me.CARGO_NUEVO_ETIQUETA.Visible = False
        Me.CARGO_NUEVO.Visible = False
        Me.UBICACION_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_NUEVO_ETIQUETA.Top = 6
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 6
        Me.UBICACION_NUEVO.Visible = True
        Me.UBICACION_NUEVO.Top = 30
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 6
        Me.UBICACION_GENERAL_NUEVO.Visible = True
        Me.UBICACION_GENERAL_NUEVO.Top = 30
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Top = 60
        Me.UBICACION_ESPECIFICA_NUEVO.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO.Top = 84
        Exit Sub
    
    Case Mot = "NOMBRAMIENTO"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = True
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Me.CARGO_NUEVO_ETIQUETA.Visible = True
        Me.CARGO_NUEVO.Visible = True
        Me.UBICACION_NUEVO_ETIQUETA.Visible = False
        Me.UBICACION_NUEVO.Visible = False
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = False
        Me.UBICACION_GENERAL_NUEVO.Visible = False
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = False
        Me.UBICACION_ESPECIFICA_NUEVO.Visible = False
        Exit Sub
    
    Case Mot = "NOMBRAMIENTO DEFINITIVO"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Exit Sub
    
    Case Mot = "PERIODO DE PRUEBA"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Exit Sub
        
    Case Mot = "PASANTIA"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = False
        Me.EMP.Visible = False
        Exit Sub
     
    Case Mot = "DETERIORO"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Exit Sub
    
        
    Case Mot = "PERDIDA"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Exit Sub
    
    Case Mot = "VISITA INSTITUCIONAL"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = False
        Me.EMP.Visible = False
        Exit Sub
    
    Case Mot = "ESCUELA ADUANERA"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = False
        Me.EMP.Visible = False
        Exit Sub
    
    Case Mot = "AGENTE ADUANERO"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = False
        Me.EMP.Visible = False
        Exit Sub
    
    Case Mot = "PRUEBA DE MAQUINA"
        Mostrar
        Me.AGREGAR_NUEVO_REGISTRO.Visible = False
        Me.REGISTRO.Value = 0
        Me.REGISTRO.Pages(1).Visible = False
        Me.EMP_ETIQUETA.Visible = True
        Me.EMP.Visible = True
        Exit Sub
        
End Select
End Sub


Private Sub Mostrar()

    'Mostrados Todo el Formulario
    Me.CODIGO_CARNET_CORDON.Visible = True
    Me.height = 470
    Me.width = 900
    
    'Mostrar Arriba
    Me.FECHA_ETIQUETA.Visible = True
    Me.FECHA.Visible = True
    Me.NOMBRES_ETIQUETA.Visible = True
    Me.NOMBRES.Visible = True
        
    'Ocultar Multipagina
    Me.REGISTRO.Visible = True
    
    'Ocultar Abajo
    Me.JUSTIFICACION_ETIQUETA.Visible = True
    Me.JUSTIFICACION.Visible = True
    Me.OBSERVACIONES_ETIQUETA.Visible = True
    Me.OBSERVACIONES.Visible = True
    Me.RESPONSABLE_IMPRESION_ETIQUETA.Visible = True
    Me.RESPONSABLE_IMPRESION.Visible = True
    Me.REGISTRAR.Visible = False
    
    'Ocultar Lateral
    
    MostrarSinFoto
    Me.FOTO_TRABAJADOR.Visible = True
    Me.SIN_FOTO.Visible = True
    Me.ACTUALIZAR_FOTO.Visible = True
    Me.AGREGAR_FOTO.Visible = True
    Me.ELIMINAR_FOTO.Visible = True
    Me.CARPETA_FOTO.Visible = True
    Me.COMBO.Visible = True
    Me.CANTIDAD_ETIQUETA.Visible = True
    Me.CANTIDAD.Visible = True
    
    
End Sub


Private Sub Ocultar()
    
    'Mostrados al Iniciar Formulario
    Me.CODIGO_CARNET_CORDON.Visible = True
    Me.CODIGO_CARNET_CORDON.Value = ""
    Me.AGREGAR_NUEVO_REGISTRO.Visible = True
    Me.height = 100
    Me.width = 530
    
    'Ocultados Arriba
    Me.FECHA_ETIQUETA.Visible = False
    Me.FECHA.Visible = False
    Me.FECHA.Value = ""
    Me.EMP_ETIQUETA.Visible = False
    Me.EMP.Visible = False
    Me.EMP.Value = ""
    Me.NOMBRES_ETIQUETA.Visible = False
    Me.NOMBRES.Visible = False
    Me.NOMBRES.Value = ""
    
    'Ocultar Multipagina
    Me.REGISTRO.Visible = False
    
    'Ocultar Abajo
    Me.JUSTIFICACION_ETIQUETA.Visible = False
    Me.JUSTIFICACION.Visible = False
    Me.JUSTIFICACION.Value = ""
    Me.OBSERVACIONES_ETIQUETA.Visible = False
    Me.OBSERVACIONES.Visible = False
    Me.OBSERVACIONES.Value = ""
    Me.RESPONSABLE_IMPRESION_ETIQUETA.Visible = False
    Me.RESPONSABLE_IMPRESION.Visible = False
    Me.RESPONSABLE_IMPRESION.Clear
    Me.RESPONSABLE_IMPRESION.Value = ""
    Me.REGISTRAR.Visible = False
    
    'Ocultar Lateral
    
    MostrarSinFoto
    Me.FOTO_TRABAJADOR.Visible = False
    Me.SIN_FOTO.Visible = False
    Me.ACTUALIZAR_FOTO.Visible = False
    Me.AGREGAR_FOTO.Visible = False
    Me.ELIMINAR_FOTO.Visible = False
    Me.CARPETA_FOTO.Visible = False
    Me.COMBO.Visible = False
    Me.CARNET.Value = False
    Me.CORDON.Value = False
    Me.AMBOS.Value = False
    Me.CANTIDAD_ETIQUETA.Visible = False
    Me.CANTIDAD.Visible = False
    Me.CANTIDAD.Value = ""

End Sub

Private Sub UBICACION_Change()


Dim Depe, Ubi As Range
    If Me.UBICACION.Value = Empty Then
    Me.UBICACION_GENERAL.Clear
    Else
    'Asignar a cada Dependencia su area Especifica
    Set Depe = Hoja24.ListObjects(Me.UBICACION.Value).DataBodyRange
    Me.UBICACION_GENERAL.Clear
        
        'Agregar Cada area Especifica al Listado
        For Each Ubi In Depe
        Me.UBICACION_GENERAL.AddItem Ubi.Value
        Next Ubi
    End If
    
End Sub


Private Sub UBICACION_NUEVO_Change()

Dim Depe, Ubi As Range
    If Me.UBICACION_NUEVO.Value = Empty Then
    Me.UBICACION_GENERAL_NUEVO.Clear
    Else
    'Asignar a cada Dependencia su area Especifica
    Set Depe = Hoja24.ListObjects(Me.UBICACION_NUEVO.Value).DataBodyRange
    Me.UBICACION_GENERAL_NUEVO.Clear
        
        'Agregar Cada area Especifica al Listado
        For Each Ubi In Depe
        Me.UBICACION_GENERAL_NUEVO.AddItem Ubi.Value
        Next Ubi
    End If
    
End Sub


Private Sub NOMBRES_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub


Private Sub CARGO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub

Private Sub UBICACION_ESPECIFICA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub

Private Sub CARGO_NUEVO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub

Private Sub UBICACION_ESPECIFICA_NUEVO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub

Private Sub CARNET_Click()

If Me.CARNET.Value = True Then
Me.REGISTRAR.Visible = True
End If

End Sub
Private Sub CORDON_Click()

If Me.CORDON.Value = True Then
Me.REGISTRAR.Visible = True
End If

End Sub
Private Sub AMBOS_Click()

If Me.AMBOS.Value = True Then
Me.REGISTRAR.Visible = True
End If

End Sub

Private Sub REGISTRAR_Click()

   
    Call UserForm_Initialize

End Sub



'=====================================
'CARGAR FOTO AUTOMATICAMENTE
'=====================================
Public Sub CargarFoto(nombreCompleto As String)

    Dim rutaBase As String
    Dim archivo As String
    Dim extensiones As Variant
    Dim i As Integer
    
    rutaBase = ObtenerRutaFotos()
    
    extensiones = Array(".jpg", ".jpeg")
    
    Me.FOTO_TRABAJADOR.Picture = Nothing
    
    For i = LBound(extensiones) To UBound(extensiones)
        
        archivo = rutaBase & nombreCompleto & extensiones(i)
        
        If Dir(archivo) <> "" Then
            Me.FOTO_TRABAJADOR.Picture = LoadPicture(archivo)
            Me.FOTO_TRABAJADOR.BackColor = RGB(20, 71, 255)
            Me.SIN_FOTO.Visible = False
            Me.FOTO_TRABAJADOR.ControlTipText = ""
            Exit Sub
        End If
        
    Next i
    
    MostrarSinFoto

End Sub


'=====================================
'VOLVER A CARGAR LA FOTO POR SI HUBO ALGUNA ACTUALIZACION
'=====================================

Private Sub ACTUALIZAR_FOTO_Click()
    Call CargarFoto(Me.NOMBRES.Value)
End Sub


'=====================================
'MOSTRAR MENSAJE CUANDO NO HAYA FOTO CARGADA
'=====================================
Private Sub MostrarSinFoto()

    On Error Resume Next
    Me.FOTO_TRABAJADOR.Picture = LoadPicture("")
    On Error GoTo 0
    
    Me.FOTO_TRABAJADOR.PictureAlignment = fmPictureAlignmentCenter
    
    'Opcional: mostrar mensaje visual
    Me.FOTO_TRABAJADOR.Picture = LoadPicture("")
    Me.FOTO_TRABAJADOR.ControlTipText = "Ingrese cOdigo de trabajador"
    Me.FOTO_TRABAJADOR.BackColor = RGB(240, 240, 240)
    Me.SIN_FOTO.Visible = True

End Sub


'=====================================
'AGREGAR FOTO
'=====================================

Private Sub AGREGAR_FOTO_Click()

    Dim fd As fileDialog
    Dim rutaSeleccionada As String
    Dim rutaDestino, rutaEliminar As String
    Dim nombreTrabajador, nombreRuta1, nombreRuta2, rutaCompleta As String
    Dim rutaBase As String
    Dim ext As String
    Dim Sobreescribir As VbMsgBoxResult
    
    'Validar que exista la carpeta de las Fotos
    rutaBase = ObtenerRutaFotos()
    
    If Dir(rutaBase, vbDirectory) = "" Then
        MsgBox "La Carpeta No Existe.", vbCritical
        Exit Sub
    End If
    
    'Validar que exista ese trabajador
    nombreTrabajador = Trim(Me.NOMBRES.Value)
    
    If nombreTrabajador = "" Then
        MsgBox "No hay Nombre de Trabajador.", vbExclamation
        Exit Sub
    End If
    
    'Validar si existe una foto con el nombre
    
    nombreRuta1 = rutaBase & nombreTrabajador & ".jpg"
    nombreRuta2 = rutaBase & nombreTrabajador & ".jpeg"
    
    'Validar si hay un .jpg O .jpeg
    If Dir(nombreRuta1) <> "" Or Dir(nombreRuta2) <> "" Then
        Sobreescribir = MsgBox( _
            "Ya Existe una Foto .JPG Asociada al Trabajador:" & vbNewLine & _
            nombreTrabajador & ext & vbNewLine & vbNewLine & _
            "¿Aun asi, Desea Remplazarla?", vbQuestion + vbYesNo + vbDefaultButton2, _
            "Foto existente")
        
        If Sobreescribir = vbNo Then
            MsgBox "OperaciOn cancelada.", vbInformation
            Exit Sub
        End If
        
        
    End If
    
    'Seleccionar una foto desde el Explorador
            Set fd = Application.fileDialog(msoFileDialogFilePicker)
            
            With fd
                .Title = "Seleccione la Foto del Trabajador"
                .InitialFileName = rutaBase
                .Filters.Clear
                .Filters.Add "Imagenes", "*.jpg; *.jpeg"
                
                If .Show = -1 Then
                    rutaSeleccionada = .SelectedItems(1)
                Else
                    Exit Sub
                End If
            End With
    
    
    
    'Asignar a rutaDestino el nombre que tendra la foto a Asignar
    ext = Mid(rutaSeleccionada, InStrRev(rutaSeleccionada, "."))
    rutaDestino = rutaBase & nombreTrabajador & ext
    
    
    'Eliminar fotos existentes
    If nombreRuta1 = rutaSeleccionada Or nombreRuta2 = rutaSeleccionada Then
    MsgBox "Se SeleccionO la Misma Foto", vbExclamation
    Exit Sub
    
    Else
        If Dir(nombreRuta1) <> "" Then Kill nombreRuta1
        If Dir(nombreRuta2) <> "" Then Kill nombreRuta2
    
        FileCopy rutaSeleccionada, rutaDestino
    
    End If

    'CargarFoto nombreTrabajador
    CargarFoto (nombreTrabajador)
    MsgBox "Foto Actualizada Correctamente.", vbInformation

End Sub


'=====================================
'ELIMINAR FOTO
'=====================================
Private Sub ELIMINAR_FOTO_Click()
    
    Dim rutaFotos As String
    ' Obtenemos la ruta directamente (funciOn que ya tienes)
    rutaFotos = ObtenerRutaFotos()
    
    Dim nombre As String: nombre = Trim(Me.NOMBRES.Value)
    
    If nombre = "" Then Exit Sub
        If Me.FOTO_TRABAJADOR.Picture Is Nothing Then
        MsgBox "El Trabajador No Tiene Foto para Eliminar.", vbInformation
        Exit Sub
        End If
    
    If MsgBox("¿Esta seguro que desea eliminar la foto de " & nombre & "?", _
              vbQuestion + vbYesNo, "Confirmar") = vbNo Then Exit Sub
    
    ' --- EliminaciOn ---------------------------------------
    Dim seElimino As Boolean: seElimino = True
    
    On Error Resume Next
    If Dir(rutaFotos & nombre & ".jpg") <> "" Then Kill rutaFotos & nombre & ".jpg": seElimino = True
    If Dir(rutaFotos & nombre & ".jpeg") <> "" Then Kill rutaFotos & nombre & ".jpeg": seElimino = True
    On Error GoTo 0
    
    If seElimino Then
        MsgBox "Foto Eliminada Correctamente.", vbInformation
    Else
        MsgBox "No se Encontro Ninguna Foto para Eliminar.", vbInformation
    End If
    MostrarSinFoto
    
End Sub

'=====================================
'ABRIR LA CARPETA DONDE SE ENCUENTRAN LAS FOTOS
'=====================================

Private Sub CARPETA_FOTO_Click()
    
    Dim rutaFotos As String
    
    ' Obtenemos la ruta directamente (funciOn que ya tienes)
    rutaFotos = ObtenerRutaFotos()
    
    ' Verificamos que la carpeta exista
    If Dir(rutaFotos, vbDirectory) = "" Then
        MsgBox "La Carpeta No Existe." & vbCrLf & _
               "Ruta buscada: " & rutaFotos, vbCritical, "Error"
        Exit Sub
    End If
    
    ' Abrimos la carpeta en el Explorador de Windows
    Shell "explorer.exe """ & rutaFotos & """", vbNormalFocus
    
End Sub

