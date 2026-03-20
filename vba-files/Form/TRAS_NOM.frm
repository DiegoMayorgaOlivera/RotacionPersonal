VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TRAS_NOM 
   Caption         =   "TRASLADOS Y NOMBRAMIENTOS"
   ClientHeight    =   9450.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17760
   OleObjectBlob   =   "TRAS_NOM.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "TRAS_NOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'====================================================================================================================================
' Formulario para gestionar los traslados y nombramientos de empleados.
' Permite registrar nuevos movimientos y consultar los ya registrados.
'====================================================================================================================================

Dim Borrador As Boolean
Option Explicit
Private IgnorarValidacionEMP As Boolean
Private IgnorarValidacionCOD As Boolean

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


Private Sub MOTIVO_EXIT(ByVal Cancel As MSForms.ReturnBoolean)

Dim tbl As ListObject
    Dim c As Range
    Dim OpMot As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("MOTIVO")
    
    OpMot = Trim(Me.MOTIVO.Value)
    
    'Si esta vacio ? limpiar y salir
    If OpMot = "" Then
        UserForm_Initialize
        UserForm_Activate
        Me.EMP.Value = ""
        Exit Sub

    End If
    
    'Buscar empleado en columna 1 (Col A = No. EMP)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=OpMot, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        LimpiarCampos
        MsgBox "Motivo Incorrecto", vbExclamation
        Me.MOTIVO.Value = ""
        Me.EMP.Value = ""
        UserForm_Initialize
        UserForm_Activate
        Cancel = True
        Exit Sub
        
    End If
    
    Me.EMP_ETIQUETA.Visible = True
    Me.EMP.Visible = True
    Me.EMP.Value = ""
    Me.EMP.SetFocus

End Sub



Private Sub UserForm_Initialize()

Me.ScrollTop = 0
Me.Top = 0
Me.ScrollBars = fmScrollBarsNone
Me.KeepScrollBarsVisible = fmScrollBarsNone

IgnorarValidacionEMP = False 'validacion de EMP
IgnorarValidacionCOD = False 'validacion de COD

Hoja1.rows("1:50").EntireRow.Hidden = False
Hoja25.rows("1:2000").EntireRow.Hidden = False
Hoja1.Range("M6:M9").Value = ""
Hoja1.Range("M12:M17").Value = ""

Me.COD.Value = ""
Me.COD.Locked = False
Me.COD.tabstop = True
Me.COD.MousePointer = fmMousePointerDefault
Me.COD.tabindex = 50

Me.MOTIVO.Value = ""
Me.MOTIVO.Locked = False
Me.MOTIVO.tabstop = True
Me.MOTIVO.MousePointer = fmMousePointerDefault
Me.NOMBRES_ETIQUETA.Visible = False
Me.NOMBRES.Visible = False
Me.NOMBRES.Value = ""
Me.EMP_ETIQUETA.Visible = True
Me.EMP.Visible = True
Me.EMP.Value = ""
Me.EMP.Locked = False
Me.EMP.tabstop = True
Me.EMP.MousePointer = fmMousePointerDefault
Me.CEDULA_ETIQUETA.Visible = False
Me.CEDULA.Visible = False
Me.CEDULA.Value = ""
Me.EDAD_ETIQUETA.Visible = False
Me.EDAD.Visible = False
Me.SEXO_ETIQUETA.Visible = False
Me.SEXO.Visible = False
Me.FEMENINO.Value = False
Me.MASCULINO.Value = False
Me.CARGO_ETIQUETA.Visible = False
Me.CARGO.Visible = False
Me.CARGO.Value = ""
Me.CARGO_OCUPACIONAL_ETIQUETA.Visible = False
Me.CARGO_OCUPACIONAL.Visible = False
Me.CARGO_OCUPACIONAL.Value = ""
Me.CLASIFICACION_CARGO_ETIQUETA.Visible = False
Me.CLASIFICACION_CARGO.Visible = False
Me.CLASIFICACION_CARGO.Clear
Me.UBICACION_ETIQUETA.Visible = False
Me.UBICACION.Visible = False
Me.UBICACION.Value = ""
Me.UBICACION_GENERAL_ETIQUETA.Visible = False
Me.UBICACION_GENERAL.Visible = False
Me.UBICACION_GENERAL.Value = ""
Me.UBICACION_GENERAL.Clear
Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = False
Me.UBICACION_ESPECIFICA.Visible = False
Me.UBICACION_ESPECIFICA.Value = ""
Me.UBICACION_GENERAL_ETIQUETA.Visible = False
Me.UBICACION_GENERAL.Visible = False
Me.UBICACION_GENERAL.Value = ""
Me.CARGO_NUEVO_ETIQUETA.Visible = False
Me.CARGO_NUEVO.Visible = False
Me.CARGO_NUEVO.Value = ""
Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = False
Me.CARGO_OCUPACIONAL_NUEVO.Visible = False
Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = False
Me.CLASIFICACION_CARGO_NUEVO.Visible = False
Me.CLASIFICACION_CARGO_NUEVO.Clear
Me.UBICACION_NUEVO_ETIQUETA.Visible = False
Me.UBICACION_NUEVO.Visible = False
Me.UBICACION_NUEVO.Value = ""
Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = False
Me.UBICACION_GENERAL_NUEVO.Visible = False
Me.UBICACION_GENERAL_NUEVO.Value = ""
Me.UBICACION_GENERAL_NUEVO.Clear
Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = False
Me.UBICACION_ESPECIFICA_NUEVO.Visible = False
Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
Me.RESPONSABLE_ETIQUETA.Visible = False
Me.RESPONSABLE.Visible = False
Me.REGISTRAR.Visible = False
Me.FECHA_ETIQUETA.Visible = False
Me.FECHA.Visible = False
Me.FECHA.Value = ""
Me.MEMO_ETIQUETA.Visible = False
Me.MEMO.Visible = False
Me.MEMO.Value = ""
Me.DOMICILIO_ETIQUETA.Visible = False
Me.DOMICILIO.Visible = False
Me.URBANO.Value = False
Me.RURAL.Value = False
Me.OBSERVACIONES_ETIQUETA.Visible = False
Me.OBSERVACIONES.Visible = False
Me.OBSERVACIONES.Value = ""
Me.ATRAS.Visible = True
Me.DEPENDIENTES_ETIQUETA.Visible = False
Me.DEPENDIENTES.Visible = False
Me.SI.Value = False
Me.NO.Value = False



Me.ATRAS.Top = 20
Me.ATRAS.Left = 445


Me.FOTO_TRABAJADOR.Visible = False
Me.SIN_FOTO.Visible = False
Me.ACTUALIZAR_FOTO.Visible = False
Me.AGREGAR_FOTO.Visible = False
Me.ELIMINAR_FOTO.Visible = False
Me.CARPETA_FOTO.Visible = False
Me.IMPRIMIR.Visible = False
Me.EDITAR.Visible = False
Me.COMPARATIVA.Visible = False
    
    'BUSQUEDA DE TIPOS

            Dim Motivos, Mot As Range
                
                ' Referenciar la tabla
                Set Motivos = Hoja24.ListObjects("MOTIVO").DataBodyRange
                
                Me.MOTIVO.Clear
                
                ' Agregar cada valor, omitiendo vacios
                For Each Mot In Motivos
                    If Trim(Mot.Value) <> "" Then  ' Omite celdas vacias
                        Me.MOTIVO.AddItem Mot.Value
                    End If
                Next Mot
    
    
End Sub

Private Sub UserForm_Activate()
Me.Top = 0
Me.Left = 0
Me.height = 120
Me.width = 500
Me.ScrollBars = fmScrollBarsNone
Me.KeepScrollBarsVisible = fmScrollBarsNone
Me.COMPARATIVA.BackColor = RGB(222, 242, 252)
DesbloquearTodo

End Sub

Private Sub COD_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IgnorarValidacionCOD Then Exit Sub

    Dim tblMov As ListObject
    Dim c As Range
    Dim codigo, MOTIVO As String
    
    'Referencia a la tabla
    Set tblMov = Hoja5.ListObjects("MOVIMIENTOS")
    
    codigo = Trim(Me.COD.Value)
    
    'Si esta vacio ? limpiar y salir
    If codigo = "" Then
        UserForm_Initialize
        UserForm_Activate
        Me.MOTIVO.Value = ""
        
            
        Exit Sub
    End If
    'Buscar Movimiento
    Set c = tblMov.ListColumns(2).DataBodyRange.Find( _
                What:=codigo, _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    
    'Si NO existe
    If c Is Nothing Then
        
        MsgBox "El Movimiento de Empleado no Existe", vbExclamation
        UserForm_Initialize
        UserForm_Activate
        Me.MOTIVO.Value = ""
        
        Me.COD.Value = ""
        Me.COD.SetFocus
        Cancel = True
        Exit Sub
    End If
 
    IgnorarValidacionEMP = True
        
    MOTIVO = c.Offset(0, 3).Value
    
    Me.COMPARATIVA.Visible = True
    Me.COMPARATIVA.Value = 0
    Hoja1.Range("M2").Value = Me.COD.Value
        
        
        Select Case True
        Case MOTIVO = "TN"
            Me.MOTIVO.Value = "TRASLADO Y NOMBRAMIENTO"
            Me.CARGO_NUEVO_ETIQUETA.Visible = True
            Me.CARGO_NUEVO.Visible = True
            Me.CARGO_NUEVO.Value = ""
            Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = True
            Me.CARGO_OCUPACIONAL_NUEVO.Visible = True
            Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
            Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = True
            Me.CLASIFICACION_CARGO_NUEVO.Visible = True
            Me.CLASIFICACION_CARGO_NUEVO.Clear
            Me.CLASIFICACION_CARGO_NUEVO.Value = ""
            Me.UBICACION_NUEVO_ETIQUETA.Visible = True
            Me.UBICACION_NUEVO_ETIQUETA.Left = 6
            Me.UBICACION_NUEVO_ETIQUETA.Top = 114
            Me.UBICACION_NUEVO.Visible = True
            Me.UBICACION_NUEVO.Left = 6
            Me.UBICACION_NUEVO.Top = 138
            Me.UBICACION_NUEVO.Value = ""
            Me.UBICACION_NUEVO.Clear
            Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = True
            Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Left = 258
            Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 114
            Me.UBICACION_GENERAL_NUEVO.Visible = True
            Me.UBICACION_GENERAL_NUEVO.Left = 258
            Me.UBICACION_GENERAL_NUEVO.Top = 138
            Me.UBICACION_GENERAL_NUEVO.Value = ""
            Me.UBICACION_GENERAL_NUEVO.Clear
            Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = True
            Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Left = 6
            Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Top = 168
            Me.UBICACION_ESPECIFICA_NUEVO.Visible = True
            Me.UBICACION_ESPECIFICA_NUEVO.Left = 6
            Me.UBICACION_ESPECIFICA_NUEVO.Top = 192
            Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
            Hoja1.rows("24:29").EntireRow.Hidden = False
            Hoja1.Range("F23").Value = "Datos de la Ubicacion Nueva y Cargo Laboral Nuevo"
            
        Case MOTIVO = "T"
            Me.MOTIVO.Value = "TRASLADO"
            Me.CARGO_NUEVO_ETIQUETA.Visible = False
            Me.CARGO_NUEVO.Visible = False
            Me.CARGO_NUEVO.Value = ""
            Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = False
            Me.CARGO_OCUPACIONAL_NUEVO.Visible = False
            Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
            Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = False
            Me.CLASIFICACION_CARGO_NUEVO.Visible = False
            Me.CLASIFICACION_CARGO_NUEVO.Clear
            Me.CLASIFICACION_CARGO_NUEVO.Value = ""
            Me.UBICACION_NUEVO_ETIQUETA.Visible = True
            Me.UBICACION_NUEVO_ETIQUETA.Top = 6
            Me.UBICACION_NUEVO.Visible = True
            Me.UBICACION_NUEVO.Top = 30
            Me.UBICACION_NUEVO.Value = ""
            Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = True
            Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 6
            Me.UBICACION_GENERAL_NUEVO.Visible = True
            Me.UBICACION_GENERAL_NUEVO.Top = 30
            Me.UBICACION_GENERAL_NUEVO.Value = ""
            Me.UBICACION_GENERAL_NUEVO.Clear
            Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = True
            Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Top = 60
            Me.UBICACION_ESPECIFICA_NUEVO.Visible = True
            Me.UBICACION_ESPECIFICA_NUEVO.Top = 84
            Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
            Hoja1.rows("24:29").EntireRow.Hidden = False
            Hoja1.rows("28:29").EntireRow.Hidden = True
            Hoja1.Range("F23").Value = "Datos de la Ubicacion Nueva"

        Case MOTIVO = "N"
            Me.MOTIVO.Value = "NOMBRAMIENTO"
            Me.CARGO_NUEVO_ETIQUETA.Visible = True
            Me.CARGO_NUEVO.Visible = True
            Me.CARGO_NUEVO.Value = ""
            Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = True
            Me.CLASIFICACION_CARGO_NUEVO.Visible = True
            Me.CLASIFICACION_CARGO_NUEVO.Clear
            Me.CLASIFICACION_CARGO_NUEVO.Value = ""
            Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = True
            Me.CARGO_OCUPACIONAL_NUEVO.Visible = True
            Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
            Me.UBICACION_NUEVO_ETIQUETA.Visible = False
            Me.UBICACION_NUEVO.Visible = False
            Me.UBICACION_NUEVO.Value = ""
            Me.UBICACION_NUEVO.Clear
            Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = False
            Me.UBICACION_GENERAL_NUEVO.Visible = False
            Me.UBICACION_GENERAL_NUEVO.Value = ""
            Me.UBICACION_GENERAL_NUEVO.Clear
            Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = False
            Me.UBICACION_ESPECIFICA_NUEVO.Visible = False
            Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
            Hoja1.rows("24:29").EntireRow.Hidden = False
            Hoja1.rows("24:27").EntireRow.Hidden = True
            Hoja1.Range("F23").Value = "Datos del Cargo Laboral Nuevo"
            
        End Select
        Me.COMPARATIVA.Value = 0
        
        LimpiarCampos
        Me.height = 510
        Me.width = 900
        Me.NOMBRES_ETIQUETA.Visible = True
        Me.NOMBRES.Visible = True
        Me.CEDULA_ETIQUETA.Visible = True
        Me.CEDULA.Visible = True
        Me.EDAD_ETIQUETA.Visible = True
        Me.EDAD.Visible = True
        Me.SEXO_ETIQUETA.Visible = True
        Me.SEXO.Visible = True
        Me.CARGO_ETIQUETA.Visible = True
        Me.CARGO.Visible = True
        Me.CLASIFICACION_CARGO_ETIQUETA.Visible = True
        Me.CLASIFICACION_CARGO.Visible = True
        Me.CARGO_OCUPACIONAL_ETIQUETA.Visible = True
        Me.CARGO_OCUPACIONAL.Visible = True
        Me.UBICACION_ETIQUETA.Visible = True
        Me.UBICACION.Visible = True
        Me.UBICACION_GENERAL_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL.Visible = True
        Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA.Visible = True
        Me.FOTO_TRABAJADOR.Visible = True
        Me.SIN_FOTO.Visible = True
        Me.ACTUALIZAR_FOTO.Visible = True
        Me.AGREGAR_FOTO.Visible = True
        Me.ELIMINAR_FOTO.Visible = True
        Me.CARPETA_FOTO.Visible = True
        Me.RESPONSABLE_ETIQUETA.Visible = True
        Me.RESPONSABLE.Visible = True
        Me.REGISTRAR.Visible = False
        Me.EDITAR.Visible = True
        Me.FECHA_ETIQUETA.Visible = True
        Me.FECHA.Visible = True
        Me.MEMO_ETIQUETA.Visible = True
        Me.MEMO.Visible = True
        Me.DOMICILIO_ETIQUETA.Visible = True
        Me.DOMICILIO.Visible = True
        Me.SI.Value = False
        Me.NO.Value = False
        Me.DEPENDIENTES_ETIQUETA.Visible = True
        Me.DEPENDIENTES.Visible = True
        Me.URBANO.Value = False
        Me.RURAL.Value = False
        Me.OBSERVACIONES_ETIQUETA.Visible = True
        Me.OBSERVACIONES.Visible = True
        
        
        'DATOS GENERALES
        Me.EMP.Value = c.Offset(0, 4).Value
            Hoja1.Range("M4").Value = Me.EMP.Value
        Me.NOMBRES.Value = c.Offset(0, 5).Value
        Me.CEDULA.Value = c.Offset(0, 6).Value
        Me.EDAD.Value = c.Offset(0, 7).Value
        SeleccionarGenero c.Offset(0, 21).Value
        SeleccionarDomicilio c.Offset(0, 23).Value
        Me.FECHA.Value = c.Offset(0, 20).Value
        Me.MEMO.Value = c.Offset(0, 22).Value
        
        'DATOS ACTUALES
        Me.CARGO.Value = c.Offset(0, 8).Value
            Hoja1.Range("M6").Value = Me.CARGO.Value
        Me.CARGO_OCUPACIONAL.Value = c.Offset(0, 9).Value
            Hoja1.Range("M7").Value = Me.CARGO_OCUPACIONAL.Value
        Me.CLASIFICACION_CARGO.Value = c.Offset(0, 10).Value
        Me.UBICACION.Value = c.Offset(0, 11).Value
        Me.UBICACION_GENERAL.Value = c.Offset(0, 12).Value
            Hoja1.Range("L9").Value = Me.UBICACION_GENERAL.Value
        Me.UBICACION_ESPECIFICA.Value = c.Offset(0, 13).Value
            Hoja1.Range("L8").Value = Me.UBICACION_ESPECIFICA.Value


        'DATOS NUEVOS
        Me.CARGO_NUEVO.Value = c.Offset(0, 14).Value
        Me.CARGO_OCUPACIONAL_NUEVO.Value = c.Offset(0, 15).Value
        Me.CLASIFICACION_CARGO_NUEVO.Value = c.Offset(0, 16).Value

        If Me.CLASIFICACION_CARGO_NUEVO.Value = "DIRECTIVO" Then
            Hoja1.Range("M12").Value = "DIRECTIVO"
        Else
            Hoja1.Range("M12").Value = ""
        End If

        Me.UBICACION_NUEVO.Value = c.Offset(0, 17).Value
        Me.UBICACION_GENERAL_NUEVO.Value = c.Offset(0, 18).Value
            Hoja1.Range("M16").Value = Me.UBICACION_GENERAL_NUEVO.Value
        Me.UBICACION_ESPECIFICA_NUEVO.Value = c.Offset(0, 19).Value
            Hoja1.Range("M17").Value = Me.UBICACION_ESPECIFICA_NUEVO.Value

        
        'DATOS ABAJO
        Me.DEPENDIENTES.Visible = True
        Me.DEPENDIENTES_ETIQUETA.Visible = True
        SeleccionarDependientes c.Offset(0, 24).Value
        
        If c.Offset(0, 25).Value = "Temporal" Then
            Me.Temporal.Value = True
            Hoja1.Range("M3").Value = "TEMPORAL"
            Me.Definitivo.Value = False
        Else
            Me.Temporal.Value = False
        End If
        
        If c.Offset(0, 25).Value = "Definitivo" Then
            Me.Definitivo.Value = True
            Hoja1.Range("M3").Value = "DEFINITIVO"
            Me.Temporal.Value = False
        Else
            Me.Definitivo.Value = False
        End If

        Me.OBSERVACIONES.Value = c.Offset(0, 27).Value
        Me.RESPONSABLE.Value = c.Offset(0, 26).Value
        Me.REGISTRAR.Visible = False
        Me.IMPRIMIR.Visible = True


        'Preparar datos para el Memo
        Hoja1.Range("M38").Value = 2745 'Cambiar cuando haya una Persona responsable definida

            If c.Offset(0, 25).Value = "DEFINITIVO" Then
                Hoja1.Range("M3").Value = ""
                Me.Temporal.Value = False
            Else
                Hoja1.Range("M3").Value = "TEMPORAL"
                Me.Temporal.Value = True
            End If
                
        If Hoja1.Range("M36").Value = 0 Then

            Me.EMP_ENVIA_MEMO.Value = Hoja1.Range("M38").Value
        Else
            Me.EMP_ENVIA_MEMO.Value = Hoja1.Range("M36").Value
        End If
        
        Me.NOMBRE_ENVIA_MEMO.Value = UCase(Hoja1.Range("M39").Value)
        Me.CARGO_ENVIA_MEMO.Value = UCase(Hoja1.Range("M40").Value)
        Me.UBICACION_ENVIA_MEMO.Value = UCase(Hoja1.Range("M41").Value)
        Me.VER_ENVIAR.Value = True


        Me.COMPARATIVA.Value = 0

        CargarFoto (Me.NOMBRES.Value)
        
        BloquearTodo

        Me.ATRAS.Top = 440
        Me.ATRAS.Left = 845
        Me.IMPRIMIR.SetFocus
        
End Sub


Private Sub EMP_EXIT(ByVal Cancel As MSForms.ReturnBoolean)


    If IgnorarValidacionEMP Then Exit Sub 'validacion de COD

    Dim tbl As ListObject
    Dim c As Range
    Dim empVal, Mot, BAJA As String
    
    'Referencia a la tabla
    Set tbl = Hoja25.ListObjects("NOMINA")
    
    Mot = Me.MOTIVO.Value
    empVal = Trim(Me.EMP.Value)
    
    If Mot = "" Then
        MsgBox "Seleccione un Motivo de Movimiento", vbExclamation
        Exit Sub
    End If

    'Si esta vacio ? limpiar y salir
    If empVal = "" Then
        UserForm_Initialize
        UserForm_Activate
        LimpiarCampos
        Me.MOTIVO.Value = Mot
        IgnorarValidacionCOD = False
        Cancel = False

        Exit Sub
    End If
    
    'Buscar empleado en columna 1 (Col A = No. EMP)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=empVal, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        UserForm_Initialize
        UserForm_Activate
        LimpiarCampos
        Me.MOTIVO.Value = Mot
        IgnorarValidacionCOD = False
        MsgBox "Empleado No Existe en la Nomina", vbExclamation
        Me.EMP.Value = ""
        Cancel = True
        Exit Sub

    End If
    
    
    BAJA = c.Offset(0, 17).Value
    
    If BAJA <> "" Then
        
        UserForm_Initialize
        UserForm_Activate
        LimpiarCampos
        Me.MOTIVO.Value = Mot
        MsgBox "Empleado se Encuentra en Estado Inactivo (BAJA)", vbExclamation
        Me.EMP.Value = ""
        Cancel = True
        Exit Sub
    End If
    
    '===========================
    'Mostrar campos de los datos
    '===========================
    
    IgnorarValidacionCOD = True
    DesbloquearTodo
    Me.COD.Value = ""
    Me.COD.Locked = True
    Me.COD.tabstop = False
    Me.COD.MousePointer = fmMousePointerNoDrop

    Me.COMPARATIVA.Visible = True
    Me.COMPARATIVA.Value = 0
    

    Select Case True
    
    Case Mot = "TRASLADO Y NOMBRAMIENTO"
    
        Me.height = 510
        Me.width = 900
        Me.NOMBRES_ETIQUETA.Visible = True
        Me.NOMBRES.Visible = True
        Me.NOMBRES.Value = ""
        Me.CEDULA_ETIQUETA.Visible = True
        Me.CEDULA.Visible = True
        Me.CEDULA.Value = ""
        Me.EDAD_ETIQUETA.Visible = True
        Me.EDAD.Visible = True
        Me.EDAD.Value = ""
        Me.SEXO_ETIQUETA.Visible = True
        Me.SEXO.Visible = True
        Me.FEMENINO.Value = False
        Me.MASCULINO.Value = False
        Me.FEMENINO.ForeColor = &H0&
        Me.MASCULINO.ForeColor = &H0&
        Me.CARGO_ETIQUETA.Visible = True
        Me.CARGO.Visible = True
        Me.CARGO.Value = ""
        Me.CLASIFICACION_CARGO_ETIQUETA.Visible = True
        Me.CLASIFICACION_CARGO.Visible = True
        Me.CLASIFICACION_CARGO.Value = ""
        Me.CLASIFICACION_CARGO.Clear
        Me.CARGO_OCUPACIONAL_ETIQUETA.Visible = True
        Me.CARGO_OCUPACIONAL.Visible = True
        Me.CARGO_OCUPACIONAL.Value = ""
        Me.UBICACION_ETIQUETA.Visible = True
        Me.UBICACION.Visible = True
        Me.UBICACION.Value = ""
        Me.UBICACION_GENERAL_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL.Visible = True
        Me.UBICACION_GENERAL.Value = ""
        Me.UBICACION_GENERAL.Clear
        Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA.Visible = True
        Me.UBICACION_ESPECIFICA.Value = ""
        
        Me.CARGO_NUEVO_ETIQUETA.Visible = True
        Me.CARGO_NUEVO.Visible = True
        Me.CARGO_NUEVO.Value = ""
        Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = True
        Me.CARGO_OCUPACIONAL_NUEVO.Visible = True
        Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
        Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = True
        Me.CLASIFICACION_CARGO_NUEVO.Visible = True
        Me.CLASIFICACION_CARGO_NUEVO.Clear
        Me.CLASIFICACION_CARGO_NUEVO.Value = ""
        Me.UBICACION_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_NUEVO_ETIQUETA.Left = 6
        Me.UBICACION_NUEVO_ETIQUETA.Top = 114
        Me.UBICACION_NUEVO.Visible = True
        Me.UBICACION_NUEVO.Left = 6
        Me.UBICACION_NUEVO.Top = 138
        Me.UBICACION_NUEVO.Value = ""
        Me.UBICACION_NUEVO.Clear
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Left = 258
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 114
        Me.UBICACION_GENERAL_NUEVO.Visible = True
        Me.UBICACION_GENERAL_NUEVO.Left = 258
        Me.UBICACION_GENERAL_NUEVO.Top = 138
        Me.UBICACION_GENERAL_NUEVO.Value = ""
        Me.UBICACION_GENERAL_NUEVO.Clear
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Left = 6
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Top = 168
        Me.UBICACION_ESPECIFICA_NUEVO.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO.Left = 6
        Me.UBICACION_ESPECIFICA_NUEVO.Top = 192
        Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
        
        Me.RESPONSABLE_ETIQUETA.Visible = True
        Me.RESPONSABLE.Visible = True
        Me.RESPONSABLE.Value = ""
        Me.REGISTRAR.Visible = False
        Me.FECHA_ETIQUETA.Visible = True
        Me.FECHA.Visible = True
        Me.FECHA.Value = ""
        Me.MEMO_ETIQUETA.Visible = True
        Me.MEMO.Visible = True
        Me.MEMO.Value = ""
        Me.DOMICILIO_ETIQUETA.Visible = True
        Me.DOMICILIO.Visible = True
        Me.URBANO.Value = False
        Me.RURAL.Value = False
        Me.OBSERVACIONES_ETIQUETA.Visible = True
        Me.OBSERVACIONES.Visible = True
        Me.OBSERVACIONES.Value = ""
        Me.EMP.SetFocus

        Hoja1.rows("24:29").EntireRow.Hidden = False
        Hoja1.Range("F23").Value = "Datos de la Ubicacion Nueva y Cargo Laboral Nuevo"
    
    Case Mot = "TRASLADO"
    
        Me.height = 510
        Me.width = 900
        Me.NOMBRES_ETIQUETA.Visible = True
        Me.NOMBRES.Visible = True
        Me.NOMBRES.Value = ""
        Me.CEDULA_ETIQUETA.Visible = True
        Me.CEDULA.Visible = True
        Me.CEDULA.Value = ""
        Me.EDAD_ETIQUETA.Visible = True
        Me.EDAD.Visible = True
        Me.EDAD.Value = ""
        Me.SEXO_ETIQUETA.Visible = True
        Me.SEXO.Visible = True
        Me.FEMENINO.Value = False
        Me.MASCULINO.Value = False
        Me.FEMENINO.ForeColor = &H0&
        Me.MASCULINO.ForeColor = &H0&
        Me.CARGO_ETIQUETA.Visible = True
        Me.CARGO.Visible = True
        Me.CARGO.Value = ""
        Me.CLASIFICACION_CARGO_ETIQUETA.Visible = True
        Me.CLASIFICACION_CARGO.Visible = True
        Me.CLASIFICACION_CARGO.Value = ""
        Me.CLASIFICACION_CARGO.Clear
        Me.CARGO_OCUPACIONAL_ETIQUETA.Visible = True
        Me.CARGO_OCUPACIONAL.Visible = True
        Me.CARGO_OCUPACIONAL.Value = ""
        Me.UBICACION_ETIQUETA.Visible = True
        Me.UBICACION.Visible = True
        Me.UBICACION.Value = ""
        Me.UBICACION_GENERAL_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL.Visible = True
        Me.UBICACION_GENERAL.Value = ""
        Me.UBICACION_GENERAL.Clear
        Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA.Visible = True
        Me.UBICACION_ESPECIFICA.Value = ""
        
        Me.CARGO_NUEVO_ETIQUETA.Visible = False
        Me.CARGO_NUEVO.Visible = False
        Me.CARGO_NUEVO.Value = ""
        Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = False
        Me.CARGO_OCUPACIONAL_NUEVO.Visible = False
        Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
        Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = False
        Me.CLASIFICACION_CARGO_NUEVO.Visible = False
        Me.CLASIFICACION_CARGO_NUEVO.Clear
        Me.CLASIFICACION_CARGO_NUEVO.Value = ""
        Me.UBICACION_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_NUEVO_ETIQUETA.Top = 6
        Me.UBICACION_NUEVO.Visible = True
        Me.UBICACION_NUEVO.Top = 30
        Me.UBICACION_NUEVO.Value = ""
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Top = 6
        Me.UBICACION_GENERAL_NUEVO.Visible = True
        Me.UBICACION_GENERAL_NUEVO.Top = 30
        Me.UBICACION_GENERAL_NUEVO.Value = ""
        Me.UBICACION_GENERAL_NUEVO.Clear
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Top = 60
        Me.UBICACION_ESPECIFICA_NUEVO.Visible = True
        Me.UBICACION_ESPECIFICA_NUEVO.Top = 84
        Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
        
        Me.RESPONSABLE_ETIQUETA.Visible = True
        Me.RESPONSABLE.Visible = True
        Me.RESPONSABLE.Value = ""
        Me.REGISTRAR.Visible = False
        Me.FECHA_ETIQUETA.Visible = True
        Me.FECHA.Visible = True
        Me.FECHA.Value = ""
        Me.MEMO_ETIQUETA.Visible = True
        Me.MEMO.Visible = True
        Me.MEMO.Value = ""
        Me.DOMICILIO_ETIQUETA.Visible = True
        Me.DOMICILIO.Visible = True
        Me.URBANO.Value = False
        Me.RURAL.Value = False
        Me.OBSERVACIONES_ETIQUETA.Visible = True
        Me.OBSERVACIONES.Visible = True
        Me.OBSERVACIONES.Value = ""
        Me.EMP.SetFocus

        Hoja1.rows("24:29").EntireRow.Hidden = False
        Hoja1.rows("28:29").EntireRow.Hidden = True
        Hoja1.Range("F23").Value = "Datos de la Ubicacion Nueva"
        
    Case Mot = "NOMBRAMIENTO"
    
        Me.height = 510
        Me.width = 900
        Me.NOMBRES_ETIQUETA.Visible = True
        Me.NOMBRES.Visible = True
        Me.NOMBRES.Value = ""
        Me.CEDULA_ETIQUETA.Visible = True
        Me.CEDULA.Visible = True
        Me.CEDULA.Value = ""
        Me.EDAD_ETIQUETA.Visible = True
        Me.EDAD.Visible = True
        Me.EDAD.Value = ""
        Me.SEXO_ETIQUETA.Visible = True
        Me.SEXO.Visible = True
        Me.FEMENINO.Value = False
        Me.MASCULINO.Value = False
        Me.FEMENINO.ForeColor = &H0&
        Me.MASCULINO.ForeColor = &H0&
        Me.CARGO_ETIQUETA.Visible = True
        Me.CARGO.Visible = True
        Me.CARGO.Value = ""
        Me.CLASIFICACION_CARGO_ETIQUETA.Visible = True
        Me.CLASIFICACION_CARGO.Visible = True
        Me.CLASIFICACION_CARGO.Value = ""
        Me.CLASIFICACION_CARGO.Clear
        Me.CARGO_OCUPACIONAL_ETIQUETA.Visible = True
        Me.CARGO_OCUPACIONAL.Visible = True
        Me.CARGO_OCUPACIONAL.Value = ""
        Me.UBICACION_ETIQUETA.Visible = True
        Me.UBICACION.Visible = True
        Me.UBICACION.Value = ""
        Me.UBICACION_GENERAL_ETIQUETA.Visible = True
        Me.UBICACION_GENERAL.Visible = True
        Me.UBICACION_GENERAL.Value = ""
        Me.UBICACION_GENERAL.Clear
        Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = True
        Me.UBICACION_ESPECIFICA.Visible = True
        Me.UBICACION_ESPECIFICA.Value = ""
        
        Me.CARGO_NUEVO_ETIQUETA.Visible = True
        Me.CARGO_NUEVO.Visible = True
        Me.CARGO_NUEVO.Value = ""
        Me.CLASIFICACION_CARGO_NUEVO_ETIQUETA.Visible = True
        Me.CLASIFICACION_CARGO_NUEVO.Visible = True
        Me.CLASIFICACION_CARGO_NUEVO.Clear
        Me.CLASIFICACION_CARGO_NUEVO.Value = ""
        Me.CARGO_OCUPACIONAL_NUEVO_ETIQUETA.Visible = True
        Me.CARGO_OCUPACIONAL_NUEVO.Visible = True
        Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
        Me.UBICACION_NUEVO_ETIQUETA.Visible = False
        Me.UBICACION_NUEVO.Visible = False
        Me.UBICACION_NUEVO.Value = ""
        Me.UBICACION_NUEVO.Clear
        Me.UBICACION_GENERAL_NUEVO_ETIQUETA.Visible = False
        Me.UBICACION_GENERAL_NUEVO.Visible = False
        Me.UBICACION_GENERAL_NUEVO.Value = ""
        Me.UBICACION_GENERAL_NUEVO.Clear
        Me.UBICACION_ESPECIFICA_NUEVO_ETIQUETA.Visible = False
        Me.UBICACION_ESPECIFICA_NUEVO.Visible = False
        Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
        
        Me.RESPONSABLE_ETIQUETA.Visible = True
        Me.RESPONSABLE.Visible = True
        Me.RESPONSABLE.Value = ""
        Me.REGISTRAR.Visible = False
        Me.FECHA_ETIQUETA.Visible = True
        Me.FECHA.Visible = True
        Me.FECHA.Value = ""
        Me.MEMO_ETIQUETA.Visible = True
        Me.MEMO.Visible = True
        Me.MEMO.Value = ""
        Me.DOMICILIO_ETIQUETA.Visible = True
        Me.DOMICILIO.Visible = True
        Me.URBANO.Value = False
        Me.RURAL.Value = False
        Me.OBSERVACIONES_ETIQUETA.Visible = True
        Me.OBSERVACIONES.Visible = True
        Me.OBSERVACIONES.Value = ""
        Me.EMP.SetFocus
        Hoja1.rows("24:29").EntireRow.Hidden = False
        Hoja1.rows("24:27").EntireRow.Hidden = True
        Hoja1.Range("F23").Value = "Datos del Cargo Laboral Nuevo"


    End Select

        
'===========================
    'CARGAR DATOS
    '===========================
    LimpiarCampos
    Me.FOTO_TRABAJADOR.Visible = True
    Me.SIN_FOTO.Visible = True
    Me.ACTUALIZAR_FOTO.Visible = True
    Me.AGREGAR_FOTO.Visible = True
    Me.ELIMINAR_FOTO.Visible = True
    Me.CARPETA_FOTO.Visible = True
    
    Me.COMPARATIVA.Value = 0
        Hoja1.Range("M4").Value = Me.EMP.Value
    Me.NOMBRES.Value = c.Offset(0, 1).Value     'Col B
    Me.CEDULA.Value = c.Offset(0, 12).Value     'Col M
    Me.EDAD.Value = c.Offset(0, 15).Value       'Col P
    
    'Genero (Col O)
    SeleccionarGenero c.Offset(0, 14).Value
    
    Me.CARGO.Value = c.Offset(0, 8).Value                'Col I
        Hoja1.Range("M6").Value = Me.CARGO.Value
    Me.CARGO_OCUPACIONAL.Value = c.Offset(0, 7).Value  'Col H
        Hoja1.Range("M7").Value = Me.CARGO_OCUPACIONAL.Value
    Me.CLASIFICACION_CARGO.Value = c.Offset(0, 9).Value  'Col J

            If Me.CLASIFICACION_CARGO.Value = Empty Then
                Dim Clas, Carg As Range

                'Agregar los elementos de la tabla Clasificacion_Cargo de la Hoja24 a la variable Clas
                Set Clas = Hoja24.ListObjects("CLASIFICACION_CARGO").DataBodyRange
                Me.CLASIFICACION_CARGO.Clear
                
                'Agregar Cada Clasificacion de Cargos al Listado
                For Each Carg In Clas
                Me.CLASIFICACION_CARGO.AddItem Carg.Value
                Next Carg
            End If

    Me.UBICACION.Value = c.Offset(0, 3).Value            'Col D
    Me.UBICACION_GENERAL.Value = c.Offset(0, 5).Value    'Col F
        Hoja1.Range("M9").Value = c.Offset(0, 5).Value    'para el Memo
    Me.UBICACION_ESPECIFICA.Value = c.Offset(0, 6).Value 'Col G
        Hoja1.Range("M8").Value = Me.UBICACION_ESPECIFICA.Value    'para el Memo


        'Preparar datos para el Memo
        Hoja1.Range("M38").Value = 2745
        Hoja1.Range("M3").Value = ""
        
        
        If Hoja1.Range("M36").Value = 0 Then

            Me.EMP_ENVIA_MEMO.Value = Hoja1.Range("M38").Value
        Else
            Me.EMP_ENVIA_MEMO.Value = Hoja1.Range("M36").Value
        End If
        
        Me.NOMBRE_ENVIA_MEMO.Value = UCase(Hoja1.Range("M39").Value)
        Me.CARGO_ENVIA_MEMO.Value = UCase(Hoja1.Range("M40").Value)
        Me.UBICACION_ENVIA_MEMO.Value = UCase(Hoja1.Range("M41").Value)
        Me.DEPENDIENTES_ETIQUETA.Visible = True
        Me.DEPENDIENTES.Visible = True
        Me.VER_ENVIAR.Value = True
        Me.Temporal.Value = False
        Me.Definitivo.Value = False

        Dim Depe, Ubi As Range
        If Me.EMP.Value = Empty Then
        Me.UBICACION_NUEVO.Clear
        
        Else
        'Asignar a cada Dependencia su area ESPECIFICA
        Set Depe = Hoja24.ListObjects("UBICACION").DataBodyRange
        Me.UBICACION_NUEVO.Clear
        
                'Agregar Cada area ESPECIFICA al Listado
                For Each Ubi In Depe
                Me.UBICACION_NUEVO.AddItem Ubi.Value
                Next Ubi
        End If
        
        Dim Responsables, Res As Range
        If Me.EMP.Value = Empty Then
        Me.RESPONSABLE.Clear
        
        Else
        'Asignar la Tabla Responsable
        Set Responsables = Hoja24.ListObjects("RESPONSABLE").DataBodyRange
        Me.RESPONSABLE.Clear
        
                'Agregar Cada Responsable al Listado
                For Each Res In Responsables
                Me.RESPONSABLE.AddItem Res.Value
                Next Res
        End If


        
    ' Buscar el último COD en la tabla MOVIMIENTOS y asignar el siguiente
    Dim tblMov As ListObject
    Set tblMov = Hoja5.ListObjects("MOVIMIENTOS")
    Dim maxNum As Long
    maxNum = 0
    Dim r As ListRow
    For Each r In tblMov.ListRows
        Dim codStr As String
        codStr = Trim(r.Range(2).Value) ' Columna 2 es COD
        If Left(codStr, 1) = "M" Then
            Dim num As Long
            On Error Resume Next
            num = CLng(Mid(codStr, 2))
            On Error GoTo 0
            If num > maxNum Then maxNum = num
        End If
    Next r
    Me.COD.Value = "M" & (maxNum + 1)
    
    Me.COD.Locked = True
    Me.COD.tabstop = False
    Me.COD.MousePointer = fmMousePointerNoDrop
    
    BloquearActuales
    BloquearGenerales
    DesbloquearNuevos
    DesbloquearAbajo
' ============================================
' EVENTO PARA CARGAR LAS FOTOS DEL TRABAJADOR
' ============================================


    Call CargarFoto(Me.NOMBRES.Value)
    Me.ATRAS.Top = 440
    Me.ATRAS.Left = 845
    
End Sub


Sub EMP_ENVIA_MEMO_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim tbl As ListObject
    Dim c As Range
    Dim empVal, Mot, BAJA As String
    
    'Referencia a la tabla
    Set tbl = Hoja25.ListObjects("NOMINA")
    
    empVal = Trim(Me.EMP_ENVIA_MEMO.Value)
    
    'Si esta vacio ? limpiar y salir
    If empVal = "" Then
        Me.NOMBRE_ENVIA_MEMO.Value = ""
        Me.CARGO_ENVIA_MEMO.Value = ""
        Me.UBICACION_ENVIA_MEMO.Value = ""
        Cancel = False
        
        Exit Sub
    End If
    
    'Buscar empleado en columna 1 (Col A = No. EMP)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=empVal, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        MsgBox "Empleado No Existe en la Nomina", vbExclamation
        Me.EMP_ENVIA_MEMO.Value = ""
        Cancel = True
        Exit Sub

    End If
    Hoja1.Range("M38").Value = Me.EMP_ENVIA_MEMO.Value

    Call LetrasMemo(Hoja1.Range("M39"), Hoja1.Range("M37"))

    Me.NOMBRE_ENVIA_MEMO.Value = c.Offset(0, 1).Value     'Col B
    Me.CARGO_ENVIA_MEMO.Value = c.Offset(0, 8).Value
    Me.UBICACION_ENVIA_MEMO.Value = c.Offset(0, 4).Value

End Sub



Sub Temporal_Change()

    If Me.Temporal.Value = False Then

        Me.Temporal.ForeColor = RGB(0, 0, 0)
    Else
        Me.Temporal.ForeColor = RGB(0, 0, 255)
        
    End If
    
End Sub



Private Sub Temporal_Click()
    
        If Me.Temporal.Value = True Then
        
        Me.Definitivo.Value = False
        Me.Definitivo.ForeColor = RGB(0, 0, 0)
        
        Hoja1.Range("M3").Value = "Temporal"
        Me.Temporal.ForeColor = RGB(0, 0, 255)
                
    Else
        
        Hoja1.Range("M3").Value = ""
        Me.Temporal.ForeColor = RGB(0, 0, 0)
    End If

End Sub
Private Sub Definitivo_Click()

    If Me.Definitivo.Value = True Then
    
        Me.Temporal.Value = False
        Me.Temporal.ForeColor = RGB(0, 0, 0)
        
        Hoja1.Range("M3").Value = "Definitivo"
        Me.Definitivo.ForeColor = RGB(0, 0, 255)
        
   Else

        Hoja1.Range("M3").Value = ""
        Me.Definitivo.ForeColor = RGB(0, 0, 0)

    End If

End Sub


Sub Ver_Enviar_Change()

    If Me.VER_ENVIAR.Value = False Then

        Me.VER_ENVIAR.ForeColor = RGB(0, 0, 0)
    Else
        Me.VER_ENVIAR.ForeColor = RGB(0, 0, 255)
        
    End If
    
End Sub


Sub Ver_Enviar_Click()

    If Me.VER_ENVIAR.Value = True Then
        Me.CARGO_ENVIA_MEMO.Enabled = True
        Hoja1.Range("M42").Value = "Ver"
        Hoja1.rows("13").EntireRow.Hidden = False
    Else
        Me.CARGO_ENVIA_MEMO.Enabled = False
        Hoja1.Range("M42").Value = "No Ver"
        Hoja1.rows("13").EntireRow.Hidden = True
    End If
    
End Sub

Private Sub BloquearTodo()

    '===========================
    ' Bloquear todos los campos
    '===========================
    BloquearGenerales
    BloquearActuales
    BloquearNuevos
    BloquearAbajo
    BloquearDatosMemo

End Sub

Sub BloquearGenerales()

        'Datos generales
    Me.EMP.Locked = True
    Me.MOTIVO.Locked = True
    Me.NOMBRES.Locked = True
    Me.CEDULA.Locked = True
    Me.EDAD.Locked = True
    Me.FEMENINO.Locked = True
    Me.MASCULINO.Locked = True



    'Datos generales
    Me.EMP.tabstop = False
    Me.MOTIVO.tabstop = False
    Me.NOMBRES.tabstop = False
    Me.CEDULA.tabstop = False
    Me.EDAD.tabstop = False
    Me.FEMENINO.tabstop = False
    Me.MASCULINO.tabstop = False



    'Datos generales
    Me.EMP.MousePointer = fmMousePointerNoDrop
    Me.MOTIVO.MousePointer = fmMousePointerNoDrop
    Me.NOMBRES.MousePointer = fmMousePointerNoDrop
    Me.CEDULA.MousePointer = fmMousePointerNoDrop
    Me.EDAD.MousePointer = fmMousePointerNoDrop
    Me.FEMENINO.MousePointer = fmMousePointerNoDrop
    Me.MASCULINO.MousePointer = fmMousePointerNoDrop

End Sub


Sub BloquearActuales()

    'datos actuales
    Me.CARGO.Locked = True
    Me.CARGO_OCUPACIONAL.Locked = True
    Me.UBICACION.Locked = True
    Me.UBICACION_GENERAL.Locked = True
    Me.UBICACION_ESPECIFICA.Locked = True

    'datos actuales
    Me.CARGO.tabstop = False
    Me.CARGO_OCUPACIONAL.tabstop = False
    Me.UBICACION.tabstop = False
    Me.UBICACION_GENERAL.tabstop = False
    Me.UBICACION_ESPECIFICA.tabstop = False

    'datos actuales
    Me.CARGO.MousePointer = fmMousePointerNoDrop
    Me.CARGO_OCUPACIONAL.MousePointer = fmMousePointerNoDrop
    Me.UBICACION.MousePointer = fmMousePointerNoDrop
    Me.UBICACION_GENERAL.MousePointer = fmMousePointerNoDrop
    Me.UBICACION_ESPECIFICA.MousePointer = fmMousePointerNoDrop
    
    If Me.CLASIFICACION_CARGO.Value = "" Then
        Me.CLASIFICACION_CARGO.Locked = False
        Me.CLASIFICACION_CARGO.tabstop = True
        Me.CLASIFICACION_CARGO.MousePointer = fmMousePointerDefault
    Else
        Me.CLASIFICACION_CARGO.Locked = True
        Me.CLASIFICACION_CARGO.tabstop = False
        Me.CLASIFICACION_CARGO.MousePointer = fmMousePointerNoDrop
    End If

End Sub

Sub BloquearNuevos()

    'datos nuevos
    Me.CARGO_NUEVO.Locked = True
    Me.CARGO_OCUPACIONAL_NUEVO.Locked = True
    Me.CLASIFICACION_CARGO_NUEVO.Locked = True
    Me.UBICACION_NUEVO.Locked = True
    Me.UBICACION_GENERAL_NUEVO.Locked = True
    Me.UBICACION_ESPECIFICA_NUEVO.Locked = True
    
    'datos nuevos
    Me.CARGO_NUEVO.tabstop = False
    Me.CARGO_OCUPACIONAL_NUEVO.tabstop = False
    Me.CLASIFICACION_CARGO_NUEVO.tabstop = False
    Me.UBICACION_NUEVO.tabstop = False
    Me.UBICACION_GENERAL_NUEVO.tabstop = False
    Me.UBICACION_ESPECIFICA_NUEVO.tabstop = False

    'datos nuevos
    Me.CARGO_NUEVO.MousePointer = fmMousePointerNoDrop
    Me.CARGO_OCUPACIONAL_NUEVO.MousePointer = fmMousePointerNoDrop
    Me.CLASIFICACION_CARGO_NUEVO.MousePointer = fmMousePointerNoDrop
    Me.UBICACION_NUEVO.MousePointer = fmMousePointerNoDrop
    Me.UBICACION_GENERAL_NUEVO.MousePointer = fmMousePointerNoDrop
    Me.UBICACION_ESPECIFICA_NUEVO.MousePointer = fmMousePointerNoDrop

End Sub


Sub BloquearAbajo()
    
    'datos abajo
    Me.OBSERVACIONES.Locked = True
    Me.RESPONSABLE.Locked = True
    Me.URBANO.Locked = True
    Me.RURAL.Locked = True

    'datos abajo
    Me.OBSERVACIONES.tabstop = False
    Me.RESPONSABLE.tabstop = False
    Me.URBANO.tabstop = False
    Me.RURAL.tabstop = False

    'datos abajo
    Me.OBSERVACIONES.MousePointer = fmMousePointerNoDrop
    Me.RESPONSABLE.MousePointer = fmMousePointerNoDrop
    Me.URBANO.MousePointer = fmMousePointerNoDrop
    Me.RURAL.MousePointer = fmMousePointerNoDrop

End Sub

Sub BloquearDatosMemo()

    'Bloquear datos del memo
    Me.EMP_ENVIA_MEMO.Locked = True
    Me.FECHA.Locked = True
    Me.MEMO.Locked = True
    Me.Temporal.Locked = True
    Me.Definitivo.Locked = True

    Me.EMP_ENVIA_MEMO.tabstop = False
    Me.FECHA.tabstop = False
    Me.MEMO.tabstop = False
    Me.Temporal.tabstop = False
    Me.Definitivo.tabstop = False

    Me.EMP_ENVIA_MEMO.MousePointer = fmMousePointerNoDrop
    Me.FECHA.MousePointer = fmMousePointerNoDrop
    Me.MEMO.MousePointer = fmMousePointerNoDrop
    Me.Temporal.MousePointer = fmMousePointerNoDrop
    Me.Definitivo.MousePointer = fmMousePointerNoDrop

End Sub

Private Sub DesbloquearTodo()

    '===========================
    ' Desbloquear todos los campos
    '===========================
    DesbloquearGenerales
    DesbloquearActuales
    DesbloquearNuevos
    DesbloquearAbajo
    DesbloquearDatosMemo

End Sub


Sub DesbloquearGenerales()
        
    'Datos generales
    Me.EMP.Locked = False
    Me.MOTIVO.Locked = False
    Me.NOMBRES.Locked = False
    Me.CEDULA.Locked = False
    Me.EDAD.Locked = False
    Me.FEMENINO.Locked = False
    Me.MASCULINO.Locked = False

    'Datos generales
    Me.EMP.tabstop = True
    Me.MOTIVO.tabstop = True
    Me.NOMBRES.tabstop = True
    Me.CEDULA.tabstop = True
    Me.EDAD.tabstop = True
    Me.FEMENINO.tabstop = True
    Me.MASCULINO.tabstop = True


    'Datos generales
    Me.EMP.MousePointer = fmMousePointerDefault
    Me.MOTIVO.MousePointer = fmMousePointerDefault
    Me.NOMBRES.MousePointer = fmMousePointerDefault
    Me.CEDULA.MousePointer = fmMousePointerDefault
    Me.EDAD.MousePointer = fmMousePointerDefault
    Me.FEMENINO.MousePointer = fmMousePointerDefault
    Me.MASCULINO.MousePointer = fmMousePointerDefault


End Sub

Sub DesbloquearActuales()
    
       'datos actuales
    Me.CARGO.Locked = False
    Me.CARGO_OCUPACIONAL.Locked = False
    Me.CLASIFICACION_CARGO.Locked = False
    Me.UBICACION.Locked = False
    Me.UBICACION_GENERAL.Locked = False
    Me.UBICACION_ESPECIFICA.Locked = False

    'datos actuales
    Me.CARGO.tabstop = True
    Me.CARGO_OCUPACIONAL.tabstop = True
    Me.CLASIFICACION_CARGO.tabstop = True
    Me.UBICACION.tabstop = True
    Me.UBICACION_GENERAL.tabstop = True
    Me.UBICACION_ESPECIFICA.tabstop = True

    'datos actuales
    Me.CARGO.MousePointer = fmMousePointerDefault
    Me.CARGO_OCUPACIONAL.MousePointer = fmMousePointerDefault
    Me.CLASIFICACION_CARGO.MousePointer = fmMousePointerDefault
    Me.UBICACION.MousePointer = fmMousePointerDefault
    Me.UBICACION_GENERAL.MousePointer = fmMousePointerDefault
    Me.UBICACION_ESPECIFICA.MousePointer = fmMousePointerDefault

End Sub

Sub DesbloquearNuevos()
    
    'datos nuevos
    Me.CARGO_NUEVO.Locked = False
    Me.CARGO_OCUPACIONAL_NUEVO.Locked = False
    Me.CLASIFICACION_CARGO_NUEVO.Locked = False
    Me.UBICACION_NUEVO.Locked = False
    Me.UBICACION_GENERAL_NUEVO.Locked = False
    Me.UBICACION_ESPECIFICA_NUEVO.Locked = False

    'datos nuevos
    Me.CARGO_NUEVO.tabstop = True
    Me.CARGO_OCUPACIONAL_NUEVO.tabstop = True
    Me.CLASIFICACION_CARGO_NUEVO.tabstop = True
    Me.UBICACION_NUEVO.tabstop = True
    Me.UBICACION_GENERAL_NUEVO.tabstop = True
    Me.UBICACION_ESPECIFICA_NUEVO.tabstop = True

    'datos nuevos
    Me.CARGO_NUEVO.MousePointer = fmMousePointerDefault
    Me.CARGO_OCUPACIONAL_NUEVO.MousePointer = fmMousePointerDefault
    Me.CLASIFICACION_CARGO_NUEVO.MousePointer = fmMousePointerDefault
    Me.UBICACION_NUEVO.MousePointer = fmMousePointerDefault
    Me.UBICACION_GENERAL_NUEVO.MousePointer = fmMousePointerDefault
    Me.UBICACION_ESPECIFICA_NUEVO.MousePointer = fmMousePointerDefault

End Sub

Sub DesbloquearAbajo()
        
    'datos abajo
    Me.OBSERVACIONES.Locked = False
    Me.RESPONSABLE.Locked = False
    Me.URBANO.Locked = False
    Me.RURAL.Locked = False
   

    'datos abajo
    Me.OBSERVACIONES.tabstop = True
    Me.RESPONSABLE.tabstop = True
    Me.URBANO.tabstop = True
    Me.RURAL.tabstop = True
    
    'datos abajo
    Me.OBSERVACIONES.MousePointer = fmMousePointerDefault
    Me.RESPONSABLE.MousePointer = fmMousePointerDefault
    Me.URBANO.MousePointer = fmMousePointerDefault
    Me.RURAL.MousePointer = fmMousePointerDefault

End Sub

Sub DesbloquearDatosMemo()

    'Desbloquear datos del memo
    Me.EMP_ENVIA_MEMO.Locked = False
    Me.FECHA.Locked = False
    Me.MEMO.Locked = False
    Me.Temporal.Locked = False
    Me.Definitivo.Locked = False

    Me.EMP_ENVIA_MEMO.tabstop = True
    Me.FECHA.tabstop = True
    Me.MEMO.tabstop = True
    Me.Temporal.tabstop = True
    Me.Definitivo.tabstop = True

    Me.EMP_ENVIA_MEMO.MousePointer = fmMousePointerDefault
    Me.FECHA.MousePointer = fmMousePointerDefault
    Me.MEMO.MousePointer = fmMousePointerDefault
    Me.Temporal.MousePointer = fmMousePointerDefault
    Me.Definitivo.MousePointer = fmMousePointerDefault
    
End Sub


Private Sub CARGO_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim Clas, Carg As Range

    'Verificar si el campo de Cargo esta vacio

    If TRIM(Me.CARGO.Value) = "" Then
        Me.CLASIFICACION_CARGO.Clear
        Cancel = False
        Exit Sub

    'Agregar los elementos de la tabla Clasificacion_Cargo de la Hoja24 a la variable Clas
    Else
        Set Clas = Hoja24.ListObjects("CLASIFICACION_CARGO").DataBodyRange
        Me.CLASIFICACION_CARGO.Clear
        
        'Agregar Cada Clasificacion de Cargos al Listado
        For Each Carg In Clas
        Me.CLASIFICACION_CARGO.AddItem Carg.Value
        Next Carg
    End If

End Sub

Private Sub Cargo_Nuevo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
Dim Clas, Carg As Range
                'Verificar si el campo de Cargo_Nuevo esta vacio

    If TRIM(Me.CARGO_NUEVO.Value) = "" Then
        Me.CLASIFICACION_CARGO_NUEVO.Clear
        Cancel = False
        Exit Sub
    
    'Agregar los elementos de la tabla Clasificacion_Cargo de la Hoja24 a la variable Clas
    Else
        Set Clas = Hoja24.ListObjects("CLASIFICACION_CARGO").DataBodyRange
        Me.CLASIFICACION_CARGO_NUEVO.Clear
        
        'Agregar Cada Clasificacion de Cargos al Listado
        For Each Carg In Clas
        Me.CLASIFICACION_CARGO_NUEVO.AddItem Carg.Value
        Next Carg
    End If

    Hoja1.Range("M13").Value = Me.CARGO_NUEVO.Value
End Sub

Private Sub Cargo_Ocupacional_Nuevo_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Hoja1.Range("M14").Value = Me.CARGO_OCUPACIONAL_NUEVO.Value
    
End Sub

Private Sub CLASIFICACION_CARGO_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.CARGO.Value = "" Then
        Me.CLASIFICACION_CARGO.Value = ""
        Cancel = False
        Exit Sub

    End If
   
   
    Dim tbl As ListObject
    Dim c As Range
    Dim Clas As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("CLASIFICACION_CARGO")
    
    Clas = Trim(Me.CLASIFICACION_CARGO.Value)
    
    'Si esta vacio ? limpiar y salir
    If Clas = "" Then
        Me.CLASIFICACION_CARGO.Value = ""
        Cancel = False
        Exit Sub
    End If
    
    'Buscar ubicacion en columna 1 (Col A = Ubicacion)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Clas, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        Me.CLASIFICACION_CARGO.Value = ""
        MsgBox "La Clasificacion de Cargo no Existe", vbExclamation
        Cancel = True
        Me.CLASIFICACION_CARGO.SetFocus
        Me.CLASIFICACION_CARGO.Dropdown
        Exit Sub
        
    End If

    Hoja1.Range("M13").Value = Me.CLASIFICACION_CARGO.Value

End Sub


Private Sub CLASIFICACION_CARGO_NUEVO_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Me.CARGO_NUEVO.Value = "" Then
        Me.CLASIFICACION_CARGO_NUEVO.Value = ""
        Cancel = False
        Exit Sub

    End If


    Dim tbl As ListObject
    Dim c As Range
    Dim Clas As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("CLASIFICACION_CARGO")
    
    Clas = Trim(Me.CLASIFICACION_CARGO_NUEVO.Value)
    
    'Si esta vacio ? limpiar y salir
    If Clas = "" Then
        Me.CLASIFICACION_CARGO_NUEVO.Value = ""
        Cancel = False
        Exit Sub
    End If
    
    'Buscar ubicacion en columna 1 (Col A = Ubicacion)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Clas, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        Me.CLASIFICACION_CARGO_NUEVO.Value = ""
        MsgBox "La Clasificacion de Cargo Nuevo no Existe", vbExclamation
        Cancel = True
        Me.CLASIFICACION_CARGO_NUEVO.SetFocus
        Me.CLASIFICACION_CARGO_NUEVO.Dropdown
        Exit Sub
        
    End If

        If Me.CLASIFICACION_CARGO_NUEVO.Value = "DIRECTIVO" Then
            Hoja1.Range("M12").Value = "DIRECTIVO"
        Else
            Hoja1.Range("M12").Value = ""
        End If

End Sub


Private Sub UBICACION_exit(ByVal Cancel As MSForms.ReturnBoolean)

    'Si ingresa una ubicacion que no existe, limpiar el campo de Ubicacion
        Dim tbl As ListObject
        Dim c As Range
        Dim Ubi As Variant
        'Referencia a la tabla
        Set tbl = Hoja24.ListObjects("UBICACION")
        Ubi = Trim(Me.UBICACION.Value)
        'Si esta vacio ? limpiar y salir
        If Ubi = "" Then
            Me.UBICACION_GENERAL.Clear
            Cancel = False
            Exit Sub
        End If
        'Buscar ubicacion en columna 1 (Col A = Ubicacion)
        Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                    What:=Ubi, _
                    LookAt:=xlWhole, _
                    MatchCase:=False)
        'Si NO existe
        If c Is Nothing Then
            Me.UBICACION_GENERAL.Clear
            MsgBox "La Ubicacion no Existe", vbExclamation
            Me.UBICACION.Value = ""
            Cancel = True
            Me.UBICACION.SetFocus
            Me.UBICACION.Dropdown
            Exit Sub
        End If

    'Asignar a cada Dependencia su area ESPECIFICA
    Dim Dependencia, Dep As Range
        Set Dependencia = Hoja24.ListObjects(Me.UBICACION.Value).DataBodyRange
        
        'Solo rellenar si UBICACION_GENERAL está vacío
        If Me.UBICACION_GENERAL.Value = "" Then
            Me.UBICACION_GENERAL.Clear
            Me.UBICACION_GENERAL.Value = ""
            
            'Agregar Cada area ESPECIFICA al Listado
            For Each Dep In Dependencia
            Me.UBICACION_GENERAL.AddItem Dep.Value
            Next Dep
        End If
End Sub

Private Sub UBICACION_NUEVO_exit(ByVal Cancel As MSForms.ReturnBoolean)

    Dim tbl As ListObject
    Dim c As Range
    Dim Ubi As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("UBICACION")
    
    Ubi = Trim(Me.UBICACION_NUEVO.Value)
    
    'Si esta vacio ? limpiar y salir
    If Ubi = "" Then
        Me.UBICACION_GENERAL_NUEVO.Clear
        Cancel = False
        Exit Sub
    End If
    
    'Buscar ubicacion en columna 1 (Col A = Ubicacion)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Ubi, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        Me.UBICACION_GENERAL_NUEVO.Clear
        MsgBox "La Ubicacion no Existe", vbExclamation
        Me.UBICACION_NUEVO.Value = ""
        Cancel = True
        Me.UBICACION_NUEVO.SetFocus
        Me.UBICACION_NUEVO.Dropdown
        Exit Sub
        
    End If
    
    'Asignar a cada Dependencia su area ESPECIFICA
    Dim Dependencia, Dep As Range
        Set Dependencia = Hoja24.ListObjects(Me.UBICACION_NUEVO.Value).DataBodyRange
        
        'Solo rellenar si UBICACION_GENERAL_NUEVO está vacío
        If Me.UBICACION_GENERAL_NUEVO.Value = "" Then
            Me.UBICACION_GENERAL_NUEVO.Clear
            Me.UBICACION_GENERAL_NUEVO.Value = ""
            
            'Agregar Cada area ESPECIFICA al Listado
            For Each Dep In Dependencia
                Me.UBICACION_GENERAL_NUEVO.AddItem Dep.Value
            Next Dep
        End If
    
    Hoja1.Range("M15").Value = Me.UBICACION_NUEVO.Value
End Sub


Private Sub UBICACION_GENERAL_EXIT(ByVal Cancel As MSForms.ReturnBoolean)

Dim tbl As ListObject
    Dim c As Range
    Dim UbiGen As String
    
    If TRIM(Me.UBICACION.Value) = "" Then
        Me.UBICACION_GENERAL.Clear
        Me.UBICACION_GENERAL.Value = ""
        Cancel = False
        Exit Sub
        
    End If

    'Referencia a la tabla
    
    Set tbl = Hoja24.ListObjects(Me.UBICACION.Value)
    
    UbiGen = Trim(Me.UBICACION_GENERAL.Value)
    
    'Si esta vacio ? limpiar y salir
    If UbiGen = "" Then
        Cancel = False
        Exit Sub
    End If
    
    'Buscar ubicacion en columna 1 (Col A = Ubicacion)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=UbiGen, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        Me.UBICACION_GENERAL.Value = ""
        MsgBox "La Ubicacion General no Existe", vbExclamation
        Cancel = True
        Me.UBICACION_GENERAL.SetFocus
        Me.UBICACION_GENERAL.Dropdown
        Exit Sub
        
    End If

    Me.UBICACION_GENERAL.Value = Hoja1.Range("M9").Value

End Sub

Private Sub UBICACION_GENERAL_NUEVO_EXIT(ByVal Cancel As MSForms.ReturnBoolean)

Dim tbl As ListObject
    Dim c As Range
    Dim UbiGen As String
    
    If TRIM(Me.UBICACION_NUEVO.Value) = "" Then
        Me.UBICACION_GENERAL_NUEVO.Clear
        Me.UBICACION_GENERAL_NUEVO.Value = ""
        Cancel = False
        Exit Sub
        
    End If

    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects(Me.UBICACION_NUEVO.Value)
    
    UbiGen = Trim(Me.UBICACION_GENERAL_NUEVO.Value)
    
    'Si esta vacio ? limpiar y salir
    If UbiGen = "" Then
        Cancel = False
        Exit Sub
    End If
    
    'Buscar ubicacion en columna 1 (Col A = Ubicacion)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=UbiGen, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        Me.UBICACION_GENERAL_NUEVO.Value = ""
        MsgBox "La Ubicacion General no Existe", vbExclamation
        Cancel = True
        Me.UBICACION_GENERAL_NUEVO.SetFocus
        Me.UBICACION_GENERAL_NUEVO.Dropdown
        Exit Sub
        
    End If

        Hoja1.Range("M16").Value = Me.UBICACION_GENERAL_NUEVO.Value
    
End Sub

Private Sub Ubicacion_Especifica_Nuevo_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Hoja1.Range("M17").Value = Me.UBICACION_ESPECIFICA_NUEVO.Value
    
End Sub

Private Sub URBANO_Click()

    Me.URBANO.Value = True
    Me.URBANO.ForeColor = &HFF0000
    Me.RURAL.Value = False
    Me.RURAL.ForeColor = &H0&
    
End Sub

Private Sub RURAL_Click()
        
    Me.RURAL.Value = True
    Me.RURAL.ForeColor = &HFF0000
    Me.URBANO.Value = False
    Me.URBANO.ForeColor = &H0&

End Sub

Private Sub Si_Click()

    Me.SI.Value = True
    Me.SI.ForeColor = &HFF0000
    Me.NO.Value = False
    Me.NO.ForeColor = &H0&
    
End Sub

Private Sub NO_Click()

    Me.NO.Value = True
    Me.NO.ForeColor = &HFF0000
    Me.SI.Value = False
    Me.SI.ForeColor = &H0&
    
End Sub

'=========================
'Seleccion de Genero
'=========================
Private Sub SeleccionarGenero(valor As String)

    'Reset
    Me.FEMENINO.Value = False
    Me.MASCULINO.Value = False
    
    Me.FEMENINO.ForeColor = &H0&
    Me.MASCULINO.ForeColor = &H0&
    
    Select Case UCase(valor)
        
        Case "F"
            Me.FEMENINO.Value = True
            Me.FEMENINO.ForeColor = &HFF0000
            
        Case "M"
            Me.MASCULINO.Value = True
            Me.MASCULINO.ForeColor = &HFF0000
            
    End Select

End Sub

'=========================
'Seleccion de Dependientes
'=========================
Private Sub SeleccionarDependientes(valor As String)

    'Reset
    Me.SI.Value = False
    Me.NO.Value = False
    
    Me.SI.ForeColor = &H0&
    Me.NO.ForeColor = &H0&
    
    Select Case UCase(valor)
        
        Case "SI"
            Me.SI.Value = True
            Me.SI.ForeColor = &HFF0000
            
        Case "NO"
            Me.NO.Value = True
            Me.NO.ForeColor = &HFF0000
            
    End Select

End Sub

'=========================
'Seleccion de Domicilio
'=========================
Private Sub SeleccionarDomicilio(valor As String)

    'Reset
    Me.URBANO.Value = False
    Me.RURAL.Value = False
    
    Me.URBANO.ForeColor = &H0&
    Me.RURAL.ForeColor = &H0&
    
    Select Case UCase(valor)
        
        Case "URBANO"
            Me.URBANO.Value = True
            Me.URBANO.ForeColor = &HFF0000
            
        Case "RURAL"
            Me.RURAL.Value = True
            Me.RURAL.ForeColor = &HFF0000
            
    End Select

End Sub

'=========================
'Limpiar campos
'=========================



Private Sub LimpiarCampos()
    

    Me.NOMBRES.Value = ""
    Me.CEDULA.Value = ""
    Me.EDAD.Value = ""
    
    Me.FEMENINO.Value = False
    Me.MASCULINO.Value = False
    Me.FEMENINO.ForeColor = &H0&
    Me.MASCULINO.ForeColor = &H0&
    
    Me.CARGO.Value = ""
    Me.CARGO_OCUPACIONAL.Value = ""
    Me.CLASIFICACION_CARGO.Clear
    Me.UBICACION.Clear
    Me.UBICACION.Value = ""
    Me.UBICACION_GENERAL.Clear
    Me.UBICACION_ESPECIFICA.Value = ""

    Me.CARGO_NUEVO.Value = ""
    Me.CARGO_OCUPACIONAL_NUEVO.Value = ""
    Me.CLASIFICACION_CARGO_NUEVO.Clear
    Me.UBICACION_NUEVO.Clear
    Me.UBICACION_GENERAL_NUEVO.Clear
    Me.UBICACION_ESPECIFICA_NUEVO.Value = ""
    MostrarSinFoto
    
    Me.FECHA.Value = ""
    Me.MEMO.Value = ""
    Me.URBANO.Value = False
    Me.RURAL.Value = False
    Me.URBANO.ForeColor = &H0&
    Me.RURAL.ForeColor = &H0&
    
    Me.OBSERVACIONES.Value = ""
    
    Me.RESPONSABLE.Clear
    Me.REGISTRAR.Visible = False
    
End Sub



Private Sub RESPONSABLE_Change()
    If Me.RESPONSABLE.Value = "" Then
        Me.REGISTRAR.Enabled = False
        Me.IMPRIMIR.Enabled = False
        Me.IMPRIMIR.Visible = False

    Else
        Me.REGISTRAR.Visible = True
        Me.REGISTRAR.Enabled = True
        Me.IMPRIMIR.Enabled = True
        Me.IMPRIMIR.Visible = True
        
    End If

End Sub

Private Sub EMP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, /, backspace, delete, enter, tab

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub EMP_ENVIA_MEMO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, /, backspace, delete, enter, tab

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub RESPONSABLE_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim tbl As ListObject
    Dim c As Range
    Dim Resp As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("RESPONSABLE")
    
    Resp = Trim(Me.RESPONSABLE.Value)
    
    'Si esta vacio ? limpiar y salir
    If Resp = "" Then
        Me.REGISTRAR.Visible = False
        Me.REGISTRAR.Enabled = False
        Me.IMPRIMIR.Enabled = False
        Me.IMPRIMIR.Visible = False
        Cancel = False
        Exit Sub
    End If
    
    'Buscar responsable en columna 1 (Col A = Ubicacion)
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
        Cancel = True
        Me.RESPONSABLE.SetFocus
        Me.RESPONSABLE.Dropdown
        Exit Sub
        
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
Private Sub CARGO_OCUPACIONAL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
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
Private Sub CARGO_OCUPACIONAL_NUEVO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub

Private Sub UBICACION_ESPECIFICA_NUEVO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub

Private Sub OBSERVACIONES_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If

End Sub

Private Sub MEMO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, /, backspace, delete, enter, tab

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub MEMO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Me.MEMO.Value = "" Then
        Cancel = False
        Exit Sub
    End If

    If Not IsNumeric(Me.MEMO.Value) Then
        MsgBox "El campo MEMO solo acepta numeros.", vbExclamation, "Error"
        Cancel = True   '  impide que el foco salga del control
        Exit Sub
    End If

    Dim tbl As ListObject
    Dim c, m As Range
    Dim Clas As String
    
    'Referencia a la tabla
    Set tbl = Hoja5.ListObjects("MOVIMIENTOS")
    
    MEMO = Trim(Me.MEMO.Value)
    COD = Trim(Me.COD.Value)

    'Si esta vacio limpiar y salir
    If MEMO = "" Then
        Me.MEMO.Value = ""
        Cancel = False
        Exit Sub
    End If
    
    If Len(MEMO) = 1 Then
        MEMO = "000" & MEMO
    ElseIf Len(MEMO) = 2 Then
        MEMO = "00" & MEMO
    ElseIf Len(MEMO) = 3 Then
        MEMO = "0" & MEMO
    ElseIf Len(MEMO) > 4 Then
        MsgBox "El numero de MEMO no puede tener mas de 4 digitos.", vbExclamation, "Error"
        Me.MEMO.Value = ""
        Cancel = True   '  impide que el foco salga del control
        Exit Sub
    End If


        'Buscar ubicacion en columna COD de la tabla MOVIMIENTOS
    Set c = tbl.ListColumns(2).DataBodyRange.Find( _
                What:=COD, _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                MatchCase:=False)

    'Si no existe, es único (dato válido)
    If c Is Nothing Then
        
        'Buscar ubicacion en columna Memo de la tabla MOVIMIENTOS
        Set m = tbl.ListColumns(24).DataBodyRange.Find( _
                    What:=MEMO, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    MatchCase:=False)

        'Si no existe, es único (dato válido)
        If m Is Nothing Then
            ' No hay duplicado, se puede continuar
            Cancel = False
            Exit Sub

        ' Si existe, duplicado encontrado
        Else
            MsgBox "El numero de MEMO: " & MEMO & " ya existe." & vbCrLf & vbCrLf & "Por favor, ingrese un numero diferente.", vbExclamation, "Error de Duplicado"
            Me.MEMO.Value = ""
            Cancel = True   ' impide que el foco salga del control
            Me.MEMO.SetFocus
            Exit Sub
        End If
    End If

End Sub

Private Sub FECHA_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Not EsFechaValidaYFormatear(Me.FECHA) Then
        Cancel = True   '  impide que el foco salga del control
        Exit Sub
    End If
    
 End Sub

Private Sub FECHA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, /, backspace, delete, enter, tab

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub FECHA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 8 Then
       Borrador = True
    Else
       Borrador = False
    End If
    
End Sub

Private Sub FECHA_Change()
        If Borrador = False Then
        
            If Len(Me.FECHA.Value) > 10 Then
                
                Me.FECHA.Value = Mid(Me.FECHA.Value, 1, 10)
                MsgBox "Fecha de Registro Incorrecta"
            
            Else
                
                If Len(Me.FECHA.Value) = 2 Then
                Me.FECHA.Value = Me.FECHA.Value & "/"
                End If
                
                If Len(Me.FECHA.Value) = 5 Then
                Me.FECHA.Value = Me.FECHA.Value & "/"
                End If
                            
            End If
        End If
End Sub


' =============================================================
' Funcion para validar fechas
' =============================================================
Private Function EsFechaValidaYFormatear(ctrl As MSForms.TextBox, _
                                        Optional formatoSalida As String = "dd/mm/yyyy") As Boolean
    
    Dim texto As String
    Dim fechaTemp As Date
    
    texto = Trim(ctrl.Text)
    
    ' Permitir campo vacio (si es opcional en tu formulario)
    If texto = "" Then
        EsFechaValidaYFormatear = True
        Exit Function
    End If
    
' Intentamos convertir a fecha valida segun configuracion regional
    If Not IsDate(texto) Then
        MsgBox "La fecha ingresada NO es valida.", vbExclamation, "Error"
        
        ctrl.SetFocus
        ctrl.SelStart = 0
        ctrl.SelLength = Len(ctrl.Text)
        EsFechaValidaYFormatear = False
        Exit Function
    End If
    
    ' Si llego aqui ? es una fecha valida
    fechaTemp = CDate(texto)   ' o DateValue(texto)
    
    ' Formateamos al formato deseado (recomendado dd/mm/yyyy para Nicaragua)
    ctrl.Text = Format(fechaTemp, formatoSalida)
    
    EsFechaValidaYFormatear = True
    
End Function


' ============================================
' VALIDACION PARA TRASLADO Y NOMBRAMIENTO
' ============================================
Private Function ValidarTrasladoYNombramiento() As String
    Dim faltantes As String
    Dim listaCampos As String
    Dim num As Integer

    num = 0
    ' Verificar campos obligatorios
    If Trim(Me.CLASIFICACION_CARGO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CLASIFICACION CARGO ACTUAL" & vbCrLf
    End If
    
    If Trim(Me.CARGO_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CARGO FUNCIONAL NUEVO" & vbCrLf
    End If
    
    If Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CARGO OCUPACIONAL NUEVO" & vbCrLf
    End If
    
    If Trim(Me.CLASIFICACION_CARGO_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CLASIFICACION DE CARGO NUEVO" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". UBICACION NUEVA" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_GENERAL_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". UBICACION GENERAL NUEVA" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_ESPECIFICA_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". UBICACION ESPECIFICA NUEVA" & vbCrLf
    End If

    If Trim(Me.EMP_ENVIA_MEMO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". DATOS DEL EMPLEADO QUE ENVIA MEMO" & vbCrLf
    End If
    
    If Trim(Me.FECHA.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". FECHA" & vbCrLf
    End If

    
    If Trim(Me.MEMO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". MEMO" & vbCrLf
    End If
    
    ' Validar al menos una opcion URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        num = num + 1
        faltantes = faltantes & num & ". URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If
    
    'Validar al menos una opcion de Dependientes SI/NO
    If Me.SI.Value = False And Me.NO.Value = False Then
        num = num + 1
        faltantes = faltantes & num & ". DEPENDIENTES SI/NO (debe seleccionar uno)" & vbCrLf
    End If

    ValidarTrasladoYNombramiento = faltantes

End Function

' ============================================
' VALIDACION PARA TRASLADO
' ============================================
Private Function ValidarTraslado() As String
    Dim faltantes As String
    Dim num As Integer

    num = 0

    ' Verificar campos obligatorios
    If Trim(Me.CLASIFICACION_CARGO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CLASIFICACION CARGO ACTUAL" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". UBICACION NUEVA" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_GENERAL_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". UBICACION GENERAL NUEVA" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_ESPECIFICA_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". UBICACION ESPECIFICA NUEVA" & vbCrLf
    End If

    If Trim(Me.EMP_ENVIA_MEMO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". DATOS DEL EMPLEADO QUE ENVIA MEMO" & vbCrLf
    End If
    
    If Trim(Me.FECHA.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". FECHA" & vbCrLf
    End If
    
    If Trim(Me.MEMO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". MEMO" & vbCrLf
    End If
    
    ' Validar al menos una opcion URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        num = num + 1
        faltantes = faltantes & num & ". URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If

    'Validar al menos una opcion de Dependientes SI/NO
    If Me.SI.Value = False And Me.NO.Value = False Then
        num = num + 1
        faltantes = faltantes & num & ". DEPENDIENTES SI/NO (debe seleccionar uno)" & vbCrLf
    End If
    
    ValidarTraslado = faltantes

End Function

' ============================================
' VALIDACION PARA NOMBRAMIENTO
' ============================================
Private Function ValidarNombramiento() As String
    Dim faltantes As String
    Dim num As Integer

    num = 0
' Verificar campos obligatorios
    If Trim(Me.CLASIFICACION_CARGO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CLASIFICACION CARGO ACTUAL" & vbCrLf
    End If
    
    If Trim(Me.CARGO_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CARGO FUNCIONAL NUEVO" & vbCrLf
    End If
    
    If Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CARGO OCUPACIONAL NUEVO" & vbCrLf
    End If
    
    If Trim(Me.CLASIFICACION_CARGO_NUEVO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". CLASIFICACION DE CARGO NUEVO" & vbCrLf
    End If

    If Trim(Me.EMP_ENVIA_MEMO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". DATOS DEL EMPLEADO QUE ENVIA MEMO" & vbCrLf
    End If
    
    If Trim(Me.FECHA.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". FECHA" & vbCrLf
    End If
    
    If Trim(Me.MEMO.Value) = "" Then
        num = num + 1
        faltantes = faltantes & num & ". MEMO" & vbCrLf
    End If
    
    ' Validar al menos una opcion URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        num = num + 1
        faltantes = faltantes & num & ". URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If

    'Validar al menos una opcion de Dependientes SI/NO
    If Me.SI.Value = False And Me.NO.Value = False Then
        num = num + 1
        faltantes = faltantes & num & ". DEPENDIENTES SI/NO (debe seleccionar uno)" & vbCrLf
    End If
    
    ValidarNombramiento = faltantes

End Function

' ============================================
' FUNCION PARA ENFOCAR EL PRIMER CAMPO FALTANTE
' ============================================
Private Sub EnfocarPrimerCampoFaltante(MOTIVO As String)
    ' Esta funcion puede enfocar el primer campo que falta
    ' segun el motivo, para facilitar la correccion
    
    If Me.CLASIFICACION_CARGO.Value = "" Then
        Me.COMPARATIVA.Value = 0
        Me.CLASIFICACION_CARGO.SetFocus
        Exit Sub
    End If

    Select Case MOTIVO
        Case "TRASLADO Y NOMBRAMIENTO"
            Me.COMPARATIVA.Value = 1
            If Trim(Me.CARGO_NUEVO.Value) = "" Then
                Me.CARGO_NUEVO.SetFocus
            ElseIf Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
                Me.CARGO_OCUPACIONAL_NUEVO.SetFocus
            ElseIf Trim(Me.CLASIFICACION_CARGO_NUEVO.Value) = "" Then
                Me.CLASIFICACION_CARGO_NUEVO.SetFocus
                Me.CLASIFICACION_CARGO_NUEVO.Dropdown
            ElseIf Trim(Me.UBICACION_NUEVO.Value) = "" Then
                Me.UBICACION_NUEVO.SetFocus
                Me.UBICACION_NUEVO.Dropdown
            ElseIf Trim(Me.UBICACION_GENERAL_NUEVO.Value) = "" Then
                Me.UBICACION_GENERAL_NUEVO.SetFocus
                Me.UBICACION_GENERAL_NUEVO.Dropdown
            ElseIf Trim(Me.UBICACION_ESPECIFICA_NUEVO.Value) = "" Then
                Me.UBICACION_ESPECIFICA_NUEVO.SetFocus
            ElseIf Me.EMP_ENVIA_MEMO.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.EMP_ENVIA_MEMO.SetFocus
            ElseIf Me.FECHA.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.FECHA.SetFocus
            ElseIf Me.MEMO.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.MEMO.SetFocus
            ElseIf Me.SI.Value = False And Me.NO.Value = False Then
                Me.SI.SetFocus
            ElseIf Me.URBANO.Value = False And Me.RURAL.Value = False Then
                Me.URBANO.SetFocus
            End If

        Case "TRASLADO"
            Me.COMPARATIVA.Value = 1
            If Trim(Me.UBICACION_NUEVO.Value) = "" Then
                Me.UBICACION_NUEVO.SetFocus
                Me.UBICACION_NUEVO.Dropdown
            ElseIf Trim(Me.UBICACION_GENERAL_NUEVO.Value) = "" Then
                Me.UBICACION_GENERAL_NUEVO.SetFocus
                Me.UBICACION_GENERAL_NUEVO.Dropdown
            ElseIf Trim(Me.UBICACION_ESPECIFICA_NUEVO.Value) = "" Then
                Me.UBICACION_ESPECIFICA_NUEVO.SetFocus
            ElseIf Me.EMP_ENVIA_MEMO.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.EMP_ENVIA_MEMO.SetFocus
            ElseIf Me.FECHA.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.FECHA.SetFocus
            ElseIf Me.MEMO.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.MEMO.SetFocus
            ElseIf Me.SI.Value = False And Me.NO.Value = False Then
                Me.SI.SetFocus
            ElseIf Me.URBANO.Value = False And Me.RURAL.Value = False Then
                Me.URBANO.SetFocus
            End If

        Case "NOMBRAMIENTO"
            Me.COMPARATIVA.Value = 1
            If Trim(Me.CARGO_NUEVO.Value) = "" Then
                Me.CARGO_NUEVO.SetFocus
            ElseIf Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
                Me.CARGO_OCUPACIONAL_NUEVO.SetFocus
            ElseIf Trim(Me.CLASIFICACION_CARGO_NUEVO.Value) = "" Then
                Me.CLASIFICACION_CARGO_NUEVO.SetFocus
                Me.CLASIFICACION_CARGO_NUEVO.Dropdown
            ElseIf Me.EMP_ENVIA_MEMO.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.EMP_ENVIA_MEMO.SetFocus
            ElseIf Me.FECHA.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.FECHA.SetFocus
            ElseIf Me.MEMO.Value = "" Then
                Me.COMPARATIVA.Value = 2
                Me.MEMO.SetFocus
            ElseIf Me.SI.Value = False And Me.NO.Value = False Then
                Me.SI.SetFocus
            ElseIf Me.URBANO.Value = False And Me.RURAL.Value = False Then
                Me.URBANO.SetFocus
            End If

    End Select

End Sub

Private Sub EDITAR_Click()

    ' Permite editar UN registro ingresado previamente, cargando los datos en el formulario.
        
    DesbloquearNuevos
    DesbloquearDatosMemo
    DesbloquearAbajo
    Me.REGISTRAR.Visible = True
    Me.IMPIRMIR.Visible = True
    Me.EDITAR.Visible = False

End Sub

Private Sub REGISTRAR_Click()

    Dim MOTIVO As String
    Dim camposFaltantes As String
    Dim totalCampos As Integer
    Dim msgTitulo As String
    
    ' Obtener motivo
    MOTIVO = Trim(Me.MOTIVO.Value)
    
    ' Validar segun el motivo
    Select Case MOTIVO
        Case "TRASLADO Y NOMBRAMIENTO"
            camposFaltantes = ValidarTrasladoYNombramiento()
            
        Case "TRASLADO"
            camposFaltantes = ValidarTraslado()
            
        Case "NOMBRAMIENTO"
            camposFaltantes = ValidarNombramiento()
            
        Case Else
            MsgBox "Debe seleccionar un MOTIVO valido", vbExclamation, "Error"
            Me.MOTIVO.SetFocus
            Exit Sub
    End Select
    
    ' Si hay campos faltantes, mostrar mensaje con formato
    If camposFaltantes <> "" Then
        ' Contar lineas para el mensaje
        totalCampos = UBound(Split(camposFaltantes, vbCrLf))
        MsgBox "CAMPOS OBLIGATORIOS FALTANTES: " & totalCampos & vbCrLf & vbCrLf & _
               camposFaltantes & vbCrLf & _
               "--------------------------------------" & vbCrLf & _
               "Complete todos los campos mencionados", _
               vbExclamation, "Validacion de datos"
        
        ' Opcional: Enfocar el primer campo faltante
        EnfocarPrimerCampoFaltante MOTIVO
    Else
        ' Todo correcto
        If MsgBox("Esta seguro de guardar el registro?", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
            
            GuardarRegistro
            
            If Me.Temporal.Value = False Then
                MsgBox "El " & MOTIVO & " fue guardado con Exito." & vbCrLf & _
                "Y los Datos de "& Me.NOMBRES.Value &" fueron editados en NOMINA", _
                vbInformation, "Exito"
            Else
               MsgBox "El " & MOTIVO & " fue guardado con Exito." & vbCrLf & _
               "NO se actualizaron los datos de "& Me.NOMBRES.Value &" en NOMINA, por ser TEMPORAL.", _
               vbInformation, "Exito"
            End If

        End If
    End If

        Call UserForm_Initialize
        Call UserForm_Activates

End Sub


Private Sub GuardarRegistro()
Dim CeldaMotivo, Mot As String

Hoja5.Select

CeldaMotivo = Hoja5.ListObjects("MOVIMIENTOS").HeaderRowRange.Find("MOTIVO").Address(False, False)
    
     Select Case True
        Case Me.MOTIVO.Value = "TRASLADO Y NOMBRAMIENTO"
            Mot = "TN"
        Case Me.MOTIVO.Value = "TRASLADO"
            Mot = "T"
        Case Me.MOTIVO.Value = "NOMBRAMIENTO"
            Mot = "N"
    End Select
                
    If Range(CeldaMotivo).Offset(1, 0).Value = "" Then
    
        'Registrar cuando no hay ningun registro
        Range(CeldaMotivo).Offset(1, 0).Value = Mot
               
    Else
      
        'Registrar cuando ya hay uno o mas registros en la tabla
    
        Range(CeldaMotivo).End(xlDown).Offset(1, 0).Value = Mot
        
    End If
    
    'Registrar datos del Traslado y Nombramiento
    
    Range(CeldaMotivo).End(xlDown).Offset(0, 1).Value = Me.EMP.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 2).Value = Me.NOMBRES.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 3).Value = Me.CEDULA.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 4).Value = Me.EDAD.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 5).Value = Me.CARGO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 6).Value = Me.CARGO_OCUPACIONAL.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 7).Value = Me.CLASIFICACION_CARGO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 8).Value = Me.UBICACION.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 9).Value = Me.UBICACION_GENERAL.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 10).Value = Me.UBICACION_ESPECIFICA.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 11).Value = Me.CARGO_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 12).Value = Me.CARGO_OCUPACIONAL_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 13).Value = Me.CLASIFICACION_CARGO_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 14).Value = Me.UBICACION_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 15).Value = Me.UBICACION_GENERAL_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 16).Value = Me.UBICACION_ESPECIFICA_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 17).Value = Me.FECHA.Value
        If Me.FEMENINO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 18).Value = "F"
        ElseIf Me.MASCULINO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 18).Value = "M"
        End If
    Range(CeldaMotivo).End(xlDown).Offset(0, 19).Value = Me.MEMO.Value
        If Me.URBANO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 20).Value = "URBANO"
        ElseIf Me.RURAL.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 20).Value = "RURAL"
        End If

        If Me.SI.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 21).Value = "SI"
        ElseIf Me.NO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 21).Value = "NO"
        End If

        If Me.Temporal.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 22).Value = "TEMPORAL"
        ElseIf Me.Definitivo.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 22).Value = "DEFINITIVO"
        End If
    

    Range(CeldaMotivo).End(xlDown).Offset(0, 23).Value = Me.RESPONSABLE.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 24).Value = Me.OBSERVACIONES.Value


    'Si está marcado Temporal omitir este paso de actualizar en NOMINA
    If Me.Temporal.Value = True Then
        Call LimpiarCampos
        ActiveWorkbook.Save
        Exit Sub
    End If

    'Si es distinto de Temporal entonces si se actualizan los datos en NOMINA
    'Actualizar los datos del trabajador en NOMINA
    
    Dim Ultimo, f, empVal As Integer
                        
    empVal = Me.EMP.Value
    
    Ultimo = Hoja25.Range("A" & rows.Count).End(xlUp).row
    
    For f = 2 To Ultimo
        If empVal = Hoja25.Cells(f, "A").Value Then
            
            Hoja25.Cells(f, "D").Value = Me.UBICACION_NUEVO.Value
            Hoja25.Cells(f, "F").Value = Me.UBICACION_GENERAL_NUEVO.Value
            Hoja25.Cells(f, "G").Value = Me.UBICACION_ESPECIFICA_NUEVO.Value
            Hoja25.Cells(f, "I").Value = Me.CARGO_NUEVO.Value
            Hoja25.Cells(f, "H").Value = Me.CARGO_OCUPACIONAL_NUEVO.Value
            Hoja25.Cells(f, "J").Value = Me.CLASIFICACION_CARGO_NUEVO.Value
            
            '========================================================================================
            'Nivel de Aduana y Dependencia Completa
            
            Dim Ult, d As Integer
            Dim DepVal As String
            DepVal = Me.UBICACION_GENERAL_NUEVO.Value
            
            Ult = Hoja24.Range("O" & rows.Count).End(xlUp).row
            
            For d = 2 To Ultimo
                If DepVal = Hoja24.Cells(d, "O").Value Then
                
                Hoja25.Cells(f, "C").Value = Hoja24.Cells(d, "P").Value 'Nivel de Aduana
                Hoja25.Cells(f, "E").Value = Hoja24.Cells(d, "N").Value 'Dependencia Completa
                End If
            Next d
            '========================================================================================
        End If
    Next f
    
    
Call LimpiarCampos
ActiveWorkbook.Save



Call UserForm_Initialize
Call UserForm_Activate

End Sub


Private Sub IMPRIMIR_Click()

    Dim MOTIVO As String
    Dim camposFaltantes As String
    Dim totalCampos As Integer
    
    ' Obtener motivo
    MOTIVO = Trim(Me.MOTIVO.Value)
    
    ' Validar segun el motivo
    Select Case MOTIVO
        Case "TRASLADO Y NOMBRAMIENTO"
            camposFaltantes = ValidarTrasladoYNombramiento()
            
        Case "TRASLADO"
            camposFaltantes = ValidarTraslado()
            
        Case "NOMBRAMIENTO"
            camposFaltantes = ValidarNombramiento()
            
        Case Else
            MsgBox "Debe seleccionar un MOTIVO valido", vbExclamation, "Error"
            Me.MOTIVO.SetFocus
            Exit Sub
    End Select
    
    ' Si hay campos faltantes, mostrar mensaje y salir
    If camposFaltantes <> "" Then
        totalCampos = UBound(Split(camposFaltantes, vbCrLf))
        MsgBox "CAMPOS OBLIGATORIOS FALTANTES: " & totalCampos & vbCrLf & vbCrLf & _
               camposFaltantes & vbCrLf & _
               "--------------------------------------" & vbCrLf & _
               "Complete todos los campos mencionados", _
               vbExclamation, "Validacion de datos"
        EnfocarPrimerCampoFaltante MOTIVO
        Exit Sub
    End If
    
    ' Todo correcto, confirmar guardar
    If MsgBox("Esta seguro de guardar el registro e imprimir?", vbQuestion + vbYesNo, "Confirmar") = vbNo Then
        Exit Sub
    End If
    
    
    ' Pedir copias
    Dim respuesta As String
    Dim numCopias As Long

    Do
        respuesta = Trim(InputBox("Ingrese el número de copias a imprimir (1, 2 o 3):", "Cantidad de Impresiones", "1"))
        If respuesta = "" Then
            Exit Sub
        End If
        If IsNumeric(respuesta) Then
            numCopias = CLng(respuesta)
        Else
            numCopias = 0
        End If
        If numCopias < 1 Or numCopias > 3 Then
            MsgBox "Por favor ingrese 1, 2 o 3 copias.", vbExclamation, "Valor inválido"
        End If
    Loop While numCopias < 1 Or numCopias > 3

    ' Imprimir
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("MEMOS")
    On Error GoTo 0

    If sh Is Nothing Then
        MsgBox "Hoja 'MEMOS' no encontrada", vbInformation, "Aviso"
    End If

    sh.PrintOut From:=2, To:=2, Copies:=numCopias
    
        If numCopias = 1 Then
            
            If Me.Temporal.Value = False Then
                MsgBox "El " & MOTIVO & " fue guardado con Exito." & vbCrLf & _
                "Se imprimio" & numCopias & " hoja." & vbCrLf & _
                "Y los Datos de "& Me.NOMBRES.Value &" fueron editados en NOMINA", _
                vbInformation, "Exito"
            Else
                MsgBox "El " & MOTIVO & " fue guardado con Exito." & vbCrLf & _
                "Se imprimio" & numCopias & " hoja." & vbCrLf & _
                "NO se actualizaron los datos de "& Me.NOMBRES.Value &" en NOMINA, por ser TEMPORAL.", _
               vbInformation, "Exito"
            End If

        Else
            
            If Me.Temporal.Value = False Then
                MsgBox "El " & MOTIVO & " fue guardado con Exito." & vbCrLf & _
                "Se imprimieron" & numCopias & " hojas." & vbCrLf & _
                "Y los Datos de "& Me.NOMBRES.Value &" fueron editados en NOMINA", _
                vbInformation, "Exito"
            Else
                MsgBox "El " & MOTIVO & " fue guardado con Exito." & vbCrLf & _
                "Se imprimieron" & numCopias & " hojas." & vbCrLf & _
                "NO se actualizaron los datos de "& Me.NOMBRES.Value &" en NOMINA, por ser TEMPORAL.", _
               vbInformation, "Exito"
            End If
            
        End If

    ' Guardar
    GuardarRegistro

End Sub


' ========================================================================================
' FORMULARIO CON GESTION DE FOTOS DE TRABAJADORES
' ========================================================================================


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
    Me.FOTO_TRABAJADOR.ControlTipText = "Ingrese codigo de trabajador"
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
            "Aun asi, Desea Remplazarla?", vbQuestion + vbYesNo + vbDefaultButton2, _
            "Foto existente")
        
        If Sobreescribir = vbNo Then
            MsgBox "Operacion cancelada.", vbInformation
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
    MsgBox "Se Selecciono la Misma Foto", vbExclamation
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
    ' Obtenemos la ruta directamente (funcion que ya tienes)
    rutaFotos = ObtenerRutaFotos()
    
    Dim nombre As String: nombre = Trim(Me.NOMBRES.Value)
    
    If nombre = "" Then Exit Sub
        If Me.FOTO_TRABAJADOR.Picture Is Nothing Then
        MsgBox "El Trabajador No Tiene Foto para Eliminar.", vbInformation
        Exit Sub
        End If
    
    If MsgBox("Esta seguro que desea eliminar la foto de " & nombre & "?", _
              vbQuestion + vbYesNo, "Confirmar") = vbNo Then Exit Sub
    
    ' --- Eliminacion ---------------------------------------
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
    
    ' Obtenemos la ruta directamente (funcion que ya tienes)
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

