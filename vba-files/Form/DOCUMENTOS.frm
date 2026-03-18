VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DOCUMENTOS 
   Caption         =   "IMPRIMIR FICHA DE CONTRATACION"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17355
   OleObjectBlob   =   "DOCUMENTOS.frx":0000
End
Attribute VB_Name = "DOCUMENTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'====================================================================================================================================
' Formulario para gestionar las bajas de empleados.
' Permite registrar nuevas bajas y consultar las ya registradas.
'====================================================================================================================================

Option Explicit

Private colOpts As Collection
'=====================================
' RUTA BASE DE LAS FOTOS
'=====================================
Public Function ObtenerRutaFotos() As String
    ObtenerRutaFotos = ThisWorkbook.Path & "\FOTOS\"
End Function


Private Sub ATRAS_Click()

Unload Me
MENU_PRINCIPAL.Show
Load MENU_PRINCIPAL

End Sub

Private Sub UserForm_Activate()
  Me.EMP.Value = ""
  Me.Top = 0
  Me.Left = 0
  Call Reset

End Sub


' =========================
' NORMALIZAR TEXTO PARA TAG
' =========================
Private Function NormalizarTag(txt As String) As String
    Dim t As String
    t = UCase(txt)

    t = Replace(t, "'Á", "A")
    t = Replace(t, "'É", "E")
    t = Replace(t, "'Í", "I")
    t = Replace(t, "'Ó", "O")
    t = Replace(t, "'Ú", "U")
    t = Replace(t, "'Ñ", "N")

    t = Replace(t, "-", "")
    t = Replace(t, "/", "")
    t = Replace(t, ".", "")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    t = Replace(t, " ", "_")

    NormalizarTag = t
End Function

' =========================
' ASIGNAR TAGS A CONTROLES
' =========================
Public Sub AsignarTagsDocumentos()

    Dim pg As MSForms.page
    Dim fra As MSForms.Frame
    Dim lbl As MSForms.Label
    Dim ctrl As Control
    Dim baseTag As String

    For Each pg In Me.Multipagina_Documentos.Pages
        For Each ctrl In pg.Controls

            If TypeOf ctrl Is MSForms.Frame Then
                Set fra = ctrl
                Set lbl = Nothing

                ' Buscar Label del documento
                Dim c As Control
                For Each c In fra.Controls
                    If TypeOf c Is MSForms.Label Then
                        Set lbl = c
                        Exit For
                    End If
                Next c

                If Not lbl Is Nothing Then
                    baseTag = NormalizarTag(lbl.Caption)

                    ' Tag del Frame = documento
                    fra.Tag = baseTag

                    ' Tags de los OptionButton
                    For Each c In fra.Controls
                        If TypeOf c Is MSForms.OptionButton Then
                            Select Case c.Caption
                                Case "C"
                                    c.Tag = baseTag & "|C"
                                Case "NC"
                                    c.Tag = baseTag & "|NC"
                                Case "NA"
                                    c.Tag = baseTag & "|NA"
                            End Select
                        End If
                    Next c
                End If
            End If

        Next ctrl
    Next pg

End Sub

' ===============================
' RESETEO DEL FORMULARIO
' ===============================
Private Sub Reset()
    
    Dim pg As MSForms.page
    Dim ctrl As Control
    Dim subCtrl As Control
    Dim obj As clsOption

    Set colOpts = New Collection

    For Each pg In Multipagina_Documentos.Pages

        For Each ctrl In pg.Controls

            If TypeOf ctrl Is MSForms.Frame Then

                ' Reset visual
                For Each subCtrl In ctrl.Controls
                    If TypeOf subCtrl Is MSForms.Label Then
                        subCtrl.ForeColor = vbBlack
                    End If

                    If TypeOf subCtrl Is MSForms.OptionButton Then
                        subCtrl.Value = False
                        subCtrl.ForeColor = vbBlack

                        Set obj = New clsOption
                        Set obj.Opt = subCtrl
                        colOpts.Add obj
                    End If
                Next subCtrl

            End If

        Next ctrl

    Next pg

    Me.NOMBRES.Value = ""
    Me.UBICACION_GENERAL.Value = ""
    Me.UBICACION_ESPECIFICA.Value = ""
    Me.CARGO.Value = ""
    Me.SALARIO.Value = ""
    Me.FECHA.Value = ""
    Me.RESPONSABLE.Value = ""
    Me.Multipagina_Documentos.Value = 0
    Me.DOCUMENTOS_EXTRAS.Visible = False
    Me.DOCUMENTOS_EXTRAS.Value = ""
    Me.RECLUTADOR.Value = ""
    Me.EMP.Visible = True
    Me.EMP_ETIQUETA.Visible = True
    Me.NOMBRES.Visible = False
    Me.NOMBRES_ETIQUETA.Visible = False
    Me.UBICACION_GENERAL.Visible = False
    Me.UBICACION_GENERAL_ETIQUETA.Visible = False
    Me.UBICACION_ESPECIFICA.Visible = False
    Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = False
    Me.CARGO.Visible = False
    Me.CARGO_ETIQUETA.Visible = False
    Me.SALARIO.Visible = False
    Me.SALARIO_ETIQUETA.Visible = False
    Me.FECHA.Visible = False
    Me.FECHA_ETIQUETA.Visible = False
    Me.RESPONSABLE.Visible = False
    Me.RESPONSABLE_ETIQUETA.Visible = False
    Me.Multipagina_Documentos.Visible = True
    Me.DOCUMENTOS_EXTRAS.Visible = False
    Me.RECLUTADOR.Visible = False
    Me.RECLUTADOR_ETIQUETA.Visible = False
    Me.REGISTRAR.Visible = False
    Me.IMPRIMIR.Visible = False
    
    AsignarTagsDocumentos
    
    MostrarSinFoto
    
    Me.FOTO_TRABAJADOR.Visible = False
    Me.ACTUALIZAR_FOTO.Visible = False
    Me.AGREGAR_FOTO.Visible = False
    Me.ELIMINAR_FOTO.Visible = False
    Me.CARPETA_FOTO.Visible = False
    Me.SIN_FOTO.Visible = False
    Me.ATRAS.Visible = True
    Me.ATRAS.Top = 20
    Me.ATRAS.Left = 460
    
    Me.ScrollTop = 0
    Me.height = 100
    Me.width = 510
    Me.ScrollBars = fmScrollBarsNone
    Me.Top = 0
    Me.Left = 0
    
    Me.EMP.SetFocus
    
    
End Sub

' ===============================
' BUSCAR EMP
' ===============================
Private Sub EMP_EXIT(ByVal Cancel As MSForms.ReturnBoolean)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow
    Dim empVal As String
    Dim encontrado As Boolean

    Set ws = Sheets("ALTAS")
    Set tbl = ws.ListObjects("ALTAS")

    empVal = EMP.Value
    
    If empVal = "" Then
        Call Reset
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Call Reset
    Call Mostrar
    Application.ScreenUpdating = True
                '=====================================
                'Cargar datos ya existentes de DOCUMENTOS
                '=====================================
                    
                For Each r In tbl.ListRows
                    If r.Range(tbl.ListColumns("No. EMP").INDEX).Value = empVal Then
                        encontrado = True
                        CargarFila r
                        Exit For
                    End If
                Next r
                  
                '=====================================
                'Cargar datos ya existentes del TRABAJADOR
                '=====================================
                
                Dim Ultimo, f As Integer
                Dim SALARIO As Variant
            
                empVal = Me.EMP.Value
                
                Ultimo = Hoja4.Range("G" & rows.Count).End(xlUp).row
                                
                For f = 5 To Ultimo
                
                    If empVal = Hoja4.Cells(f, 7).Value Then
            
                        Me.NOMBRES.Value = Hoja4.Cells(f, 8).Value
                        Me.UBICACION_GENERAL.Value = Hoja4.Cells(f, 12).Value
                        Me.UBICACION_ESPECIFICA.Value = Hoja4.Cells(f, 13).Value
                        Me.CARGO.Value = Hoja4.Cells(f, 14).Value
                        
                            SALARIO = Hoja4.Cells(f, 18).Value
                            Me.SALARIO.Value = "C$ " & Format(SALARIO, "#,##0.00")
                            
                        Me.FECHA.Value = Hoja4.Cells(f, 19).Value
                            
                            
                        'BUSQUEDA DE RECLUTADORES Y RESPONSABLES
                                    
                            Dim Responsables As Range
                            Dim Resp As Range
                            'revisar duplicado - Dim tbl As ListObject
                            
                            If Me.EMP.Value = Empty Then
                                Me.RESPONSABLE.Clear
                            Else
                                ' Referenciar la tabla
                                Set tbl = Hoja24.ListObjects("RESPONSABLE")
                                
                                ' Obtener SOLO la columna "Nombres y Apellidos"
                                Set Responsables = tbl.ListColumns("Nombres y Apellidos").DataBodyRange
                                
                                Me.RESPONSABLE.Clear
                                Me.RECLUTADOR.Clear
                                
                                ' Agregar cada valor, omitiendo vacios
                                For Each Resp In Responsables
                                    If Trim(Resp.Value) <> "" Then  ' Omite celdas vacias
                                        Me.RECLUTADOR.AddItem Resp.Value
                                        Me.RESPONSABLE.AddItem Resp.Value
                                    End If
                                Next Resp
                            End If
                            
                            
                        Me.RECLUTADOR.Text = Hoja4.Cells(f, "AM").Value
                        Me.RESPONSABLE.Text = Hoja4.Cells(f, "AO").Value
                        Me.DOCUMENTOS_EXTRAS.Text = Hoja4.Cells(f, "CC").Value
                        Call CargarFoto(Me.NOMBRES.Value)
                    End If
                        
                Next f
                
                If Not encontrado Then

                    Call Reset
                    Me.EMP.Value = ""
                    MsgBox "Empleado no encontrado", vbExclamation
                    Cancel = True
                    Exit Sub
                End If
        ActualizarBarraProgreso
        
        Me.ATRAS.Top = 480
        Me.ATRAS.Left = 800
        
        
    End Sub


' ===============================
' CARGAR DATOS EN FORM
' ===============================
Private Sub CargarFila(r As ListRow)

    Dim pg As MSForms.page
    Dim ctrl As Control
    Dim subCtrl As Control
    Dim lblCaption As String
    Dim colIndex As Long
    Dim val As String
    
    For Each pg In Multipagina_Documentos.Pages
        For Each ctrl In pg.Controls

            If TypeOf ctrl Is MSForms.Frame Then

                lblCaption = ""

                ' Obtener caption del Label
                For Each subCtrl In ctrl.Controls
                    If TypeOf subCtrl Is MSForms.Label Then
                        lblCaption = subCtrl.Caption
                        Exit For
                    End If
                Next subCtrl

                If lblCaption = "" Then GoTo siguiente

                On Error Resume Next
                colIndex = r.Parent.ListColumns(lblCaption).INDEX
                On Error GoTo 0

                If colIndex = 0 Then GoTo siguiente

                val = r.Range(colIndex).Value

                If val = "" Then
                    ' label rojo
                    For Each subCtrl In ctrl.Controls
                        If TypeOf subCtrl Is MSForms.Label Then
                            subCtrl.ForeColor = vbRed
                        End If
                    Next subCtrl
                Else
                    ' marcar opciOn
                    For Each subCtrl In ctrl.Controls
                        If TypeOf subCtrl Is MSForms.OptionButton Then
                            If subCtrl.Caption = val Then
                                subCtrl.Value = True
                                subCtrl.ForeColor = vbBlue
                            End If
                        End If
                    Next subCtrl
                End If

            End If
siguiente:
        Next ctrl
    Next pg
    
ActualizarBarraProgreso

End Sub

' ===============================
' LIMPIAR FORM
' ===============================
Private Sub LimpiarTodo()

    Dim pg As MSForms.page
    Dim ctrl As Control
    Dim subCtrl As Control

    For Each pg In Multipagina_Documentos.Pages
        For Each ctrl In pg.Controls

            If TypeOf ctrl Is MSForms.Frame Then

                For Each subCtrl In ctrl.Controls

                    If TypeOf subCtrl Is MSForms.OptionButton Then
                        subCtrl.Value = False
                        subCtrl.ForeColor = vbBlack
                    End If

                    If TypeOf subCtrl Is MSForms.Label Then
                        subCtrl.ForeColor = vbBlack
                    End If

                Next subCtrl

            End If

        Next ctrl
    Next pg

End Sub

' ===============================
' REGISTRAR
' ===============================
Private Sub REGISTRAR_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim r As ListRow
    Dim empVal As String
    Dim pg As MSForms.page
    Dim ctrl As Control
    Dim subCtrl As Control
    Dim lblCaption As String

    Set ws = Sheets("ALTAS")
    Set tbl = ws.ListObjects("ALTAS")

    empVal = EMP.Value

    For Each r In tbl.ListRows

        If r.Range(tbl.ListColumns("No. EMP").INDEX).Value = empVal Then

            For Each pg In Multipagina_Documentos.Pages
                For Each ctrl In pg.Controls

                    If TypeOf ctrl Is MSForms.Frame Then

                        lblCaption = ""

                        For Each subCtrl In ctrl.Controls
                            If TypeOf subCtrl Is MSForms.Label Then
                                lblCaption = subCtrl.Caption
                                Exit For
                            End If
                        Next subCtrl

                        If lblCaption = "" Then GoTo siguiente2

                        For Each subCtrl In ctrl.Controls
                            If TypeOf subCtrl Is MSForms.OptionButton Then
                                If subCtrl.Value = True Then
                                    r.Range(tbl.ListColumns(lblCaption).INDEX).Value = subCtrl.Caption
                                End If
                            End If
                        Next subCtrl

                    End If
siguiente2:
                Next ctrl
            Next pg
            
                    ' ===============================
                    ' REGISTRAR DOCUMENTOS EXTRAS Y RESPONSABLE
                    ' ===============================
                    Dim Ultimo, f As Integer
                                        
                    empVal = Me.EMP.Value
                    
                    Ultimo = Hoja4.Range("G" & rows.Count).End(xlUp).row
                    
                    
                        For f = 5 To Ultimo
                            If empVal = Hoja4.Cells(f, "G").Value Then
    
                                Hoja4.Cells(f, "AO").Value = Me.RESPONSABLE.Value
                                Hoja4.Cells(f, "CC").Value = Me.DOCUMENTOS_EXTRAS.Value
                            
                            End If
                        Next f
                        
            Call Reset
            MsgBox "Guardado correctamente"
            UserForm_Activate
            Exit Sub

        End If

    Next r

    MsgBox "Empleado no encontrado", vbExclamation
    

End Sub

' ===============================
' COMPROBAR SI EN OTROS DOCUMENTOS ESTA MARCADA LA OPCION DE CUMPLE
' ===============================
Private Sub ControlOtros()

    If optOtros_C.Value = True Then
        
        DOCUMENTOS_EXTRAS.Visible = True
        
    Else
        
        DOCUMENTOS_EXTRAS.Value = ""
        DOCUMENTOS_EXTRAS.Visible = False
        
    End If

End Sub

' ===============================
' EJECUTAR POR CADA OPCION LA FUNCION DE COMPROBAR
' ===============================

Private Sub optOtros_C_Click()
    ControlOtros
End Sub

Private Sub optOtros_NC_Click()
    ControlOtros
End Sub

Private Sub optOtros_NA_Click()
    ControlOtros
End Sub



' ===============================
' VALIDAR ANTES DE IMPRIMIR
' ===============================
Private Sub IMPRIMIR_Click()

    Dim pg As Object
    Dim fra As Object
    Dim subCtrl As Object
    
    Dim seleccionado As Boolean
    Dim listaFaltantes As String
    Dim contador As Long
    
    contador = 0
    listaFaltantes = ""

    'Recorrer paginas del MultiPage
    For Each pg In Me.Multipagina_Documentos.Pages
        
        For Each fra In pg.Controls
            
            'Solo Frames de documentos
            If TypeName(fra) = "Frame" _
               And fra.Tag <> "PROGRESO" _
               And fra.Tag <> "" Then
                
                seleccionado = False
                
                'Buscar OptionButtons seleccionados
                For Each subCtrl In fra.Controls
                    
                    If TypeName(subCtrl) = "OptionButton" Then
                        If subCtrl.Value = True Then
                            seleccionado = True
                            Exit For
                        End If
                    End If
                    
                Next subCtrl
                
                'Si no hay ninguno seleccionado
                If seleccionado = False Then
                    
                    'Buscar el Label del documento
                    For Each subCtrl In fra.Controls
                        
                        If TypeName(subCtrl) = "Label" Then
                            
                            contador = contador + 1
                            
                            listaFaltantes = listaFaltantes & _
                                contador & ". " & subCtrl.Caption & vbCrLf
                            
                            Exit For
                            
                        End If
                        
                    Next subCtrl
                    
                End If
                
            End If
            
        Next fra
        
    Next pg

    'Mostrar resultado
    If contador > 0 Then
        
        MsgBox "Faltan los siguientes documentos:" & vbCrLf & vbCrLf & _
               listaFaltantes, vbExclamation, "VALIDACION"
    
    Else
        If Me.RESPONSABLE.Value = "" Then
        MsgBox "Falta ingresar un RESPONSABLE", vbExclamation
        Else
            
            If Me.DOCUMENTOS_EXTRAS.Visible = True And Me.DOCUMENTOS_EXTRAS.Value = "" Then
                        MsgBox "Falta Ingresar DOCUMENTO EXTRA", vbExclamation
                        
            Else
                
                MsgBox "Todos los documentos estan completos." & vbCrLf & _
                       "Procediendo a imprimir...", vbInformation
                
                Hoja2.Range("C7").Value = Me.EMP.Value
                
                Call REGISTRAR_Click
                Call UserForm_Activate
            End If
        
        End If
        
    End If

End Sub



Private Sub Mostrar()
    
    Me.EMP.Visible = True
    Me.EMP_ETIQUETA.Visible = True
    Me.NOMBRES.Visible = True
    Me.NOMBRES_ETIQUETA.Visible = True
    Me.UBICACION_GENERAL.Visible = True
    Me.UBICACION_GENERAL_ETIQUETA.Visible = True
    Me.UBICACION_ESPECIFICA.Visible = True
    Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = True
    Me.CARGO.Visible = True
    Me.CARGO_ETIQUETA.Visible = True
    Me.SALARIO.Visible = True
    Me.SALARIO_ETIQUETA.Visible = True
    Me.FECHA.Visible = True
    Me.FECHA_ETIQUETA.Visible = True
    Me.RESPONSABLE.Visible = True
    Me.RESPONSABLE_ETIQUETA.Visible = True
    Me.Multipagina_Documentos.Visible = True
    Me.DOCUMENTOS_EXTRAS.Visible = False
    Me.RECLUTADOR.Visible = True
    Me.RECLUTADOR_ETIQUETA.Visible = True
    Me.REGISTRAR.Visible = True
    Me.IMPRIMIR.Visible = True
    Me.height = 555
    Me.width = 880
    Me.Top = 0
    Me.ScrollBars = fmScrollBarsVertical
    Me.FOTO_TRABAJADOR.Visible = True
    Me.ACTUALIZAR_FOTO.Visible = True
    Me.AGREGAR_FOTO.Visible = True
    Me.ELIMINAR_FOTO.Visible = True
    Me.CARPETA_FOTO.Visible = True
    Me.SIN_FOTO.Visible = True
    
    
End Sub

' =========================
' MASIVO
' =========================
Private Sub MASIVO_Click()

    Dim pg As Object
    Dim ctrl As Object
    Dim fra As Object
    Dim c As Object

    For Each pg In Multipagina_Documentos.Pages
        
        For Each ctrl In pg.Controls
            
            If TypeName(ctrl) = "Frame" Then
                
                Set fra = ctrl
                
                For Each c In fra.Controls
                    If TypeName(c) = "OptionButton" Then
                        If UCase(c.Caption) = "CUMPLE" Then
                            c.Value = True
                            c.ForeColor = vbBlue
                            Exit For
                        End If
                    End If
                Next c
                
            End If
            
        Next ctrl
        
    Next pg

    ActualizarBarraProgreso

End Sub

' =========================
' PROGRESO
' =========================
Sub ActualizarBarraProgreso()

    Dim total As Long
    Dim completos As Long
    Dim pg As Object
    Dim fra As Object
    Dim ctrl As Object
    
    total = 0
    completos = 0

    For Each pg In Me.Multipagina_Documentos.Pages
        
        For Each fra In pg.Controls
            
            If TypeName(fra) = "Frame" _
               And fra.Tag <> "" _
               And fra.Tag <> "PROGRESO" Then
                
                total = total + 1
                
                For Each ctrl In fra.Controls
                    If TypeName(ctrl) = "OptionButton" Then
                        If ctrl.Value = True Then
                            completos = completos + 1
                            Exit For
                        End If
                    End If
                Next ctrl
                
            End If
            
        Next fra
        
    Next pg

    If total = 0 Then Exit Sub

    Dim porcentaje As Double
    porcentaje = completos / total

    'Barra visual
    Me.lblBarra.width = porcentaje * Me.fraProgreso.width
    Me.lblProgresoTexto.Caption = completos & "/" & total & " Documentos (" & Format(porcentaje, "0%") & ")"


End Sub


Private Sub RESPONSABLE_AfterUpdate()

Dim tbl As ListObject
    Dim c As Range
    Dim Resp As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("RESPONSABLE")
    
    Resp = Trim(Me.RESPONSABLE.Value)
    
    
    'Si esta vacio, limpiar y salir.
    If Resp = "" Then
        Me.IMPRIMIR.Enabled = False
        Me.REGISTRAR.Enabled = False
        Exit Sub
    End If
    
    'Buscar responsable en columna 1 (Col A = UbicaciOn)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Resp, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        MsgBox "Responsable no Existe", vbExclamation
        Me.REGISTRAR.Enabled = False
        Me.IMPRIMIR.Enabled = False
        Me.RESPONSABLE.Value = ""
        Me.RESPONSABLE.SetFocus
        Exit Sub
        
    End If
    
    Me.IMPRIMIR.Enabled = True
    Me.REGISTRAR.Enabled = True
    
End Sub


Private Sub DOCUMENTOS_EXTRAS_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub



'=====================================
'CARGAR FOTO AUTOMATICAMENTE
'=====================================
Public Sub CargarFoto(nombreCompleto As String)

    Dim rutaBase As String
    Dim archivo As String
    Dim extensiones As Variant
    Dim i As Integer
    
    
    If nombreCompleto = "" Then
        Exit Sub
    End If
    
    rutaBase = ObtenerRutaFotos()
    
    extensiones = Array(".jpg", ".jpeg")
    
    Me.FOTO_TRABAJADOR.Picture = Nothing
    
    For i = LBound(extensiones) To UBound(extensiones)
        
        archivo = rutaBase & nombreCompleto & extensiones(i)
        
        If Dir(archivo) <> "" Then
            Me.FOTO_TRABAJADOR.Visible = True
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
    Me.FOTO_TRABAJADOR.Visible = True
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
        MsgBox "No se EncontrO Ninguna Foto para Eliminar.", vbInformation
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

