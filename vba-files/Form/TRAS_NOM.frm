VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TRAS_NOM 
   Caption         =   "TRASLADOS Y NOMBRAMIENTOS"
   ClientHeight    =   9615.001
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

Private Sub UserForm_Initialize()

Me.ScrollTop = 0
Me.Top = 0
Me.ScrollBars = fmScrollBarsNone
Me.KeepScrollBarsVisible = fmScrollBarsNone

IgnorarValidacionEMP = False 'validacion de EMP
IgnorarValidacionCOD = False 'validacion de COD

me.cod.value = ""
me.cod.locked = false
me.cod.tabstop = true
me.cod.mousepointer = fmMousePointerDefault

me.motivo.value = ""
me.motivo.locked = false
me.motivo.tabstop = true
me.motivo.mousepointer = fmMousePointerDefault
Me.NOMBRES_ETIQUETA.Visible = False
Me.NOMBRES.Visible = False
Me.NOMBRES.Value = ""
Me.EMP_ETIQUETA.Visible = True
Me.EMP.Visible = True
Me.EMP.Value = ""
me.emp.locked = false
me.emp.tabstop = true
me.emp.mousepointer = fmMousePointerDefault
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

Me.FOTO_TRABAJADOR.Visible = False
Me.SIN_FOTO.Visible = False
Me.ACTUALIZAR_FOTO.Visible = False
Me.AGREGAR_FOTO.Visible = False
Me.ELIMINAR_FOTO.Visible = False
Me.CARPETA_FOTO.Visible = False
Me.IMPRIMIR.Visible = False

    
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
        Me.FECHA_ETIQUETA.Visible = True
        Me.FECHA.Visible = True
        Me.MEMO_ETIQUETA.Visible = True
        Me.MEMO.Visible = True
        Me.DOMICILIO_ETIQUETA.Visible = True
        Me.DOMICILIO.Visible = True
        Me.URBANO.Value = False
        Me.RURAL.Value = False
        Me.OBSERVACIONES_ETIQUETA.Visible = True
        Me.OBSERVACIONES.Visible = True
        
        
        'DATOS GENERALES
        Me.EMP.Value = c.Offset(0, 4).Value
        Me.NOMBRES.Value = c.Offset(0, 5).Value
        Me.CEDULA.Value = c.Offset(0, 6).Value
        Me.EDAD.Value = c.Offset(0, 7).Value
        SeleccionarGenero c.Offset(0, 21).Value
        SeleccionarDomicilio c.Offset(0, 23).Value
        Me.FECHA.Value = c.Offset(0, 20).Value
        Me.MEMO.Value = c.Offset(0, 22).Value
        
        'DATOS ACTUALES
        Me.CARGO.Value = c.Offset(0, 8).Value
        Me.CARGO_OCUPACIONAL.Value = c.Offset(0, 9).Value
        Me.CLASIFICACION_CARGO.Value = c.Offset(0, 10).Value
        Me.UBICACION.Value = c.Offset(0, 11).Value
        Me.UBICACION_GENERAL.Value = c.Offset(0, 12).Value
        Me.UBICACION_ESPECIFICA.Value = c.Offset(0, 13).Value
        
        
        'DATOS NUEVOS
        Me.CARGO_NUEVO.Value = c.Offset(0, 14).Value
        Me.CARGO_OCUPACIONAL_NUEVO.Value = c.Offset(0, 15).Value
        Me.CLASIFICACION_CARGO_NUEVO.Value = c.Offset(0, 16).Value
        Me.UBICACION_NUEVO.Value = c.Offset(0, 17).Value
        Me.UBICACION_GENERAL_NUEVO.Value = c.Offset(0, 18).Value
        Me.UBICACION_ESPECIFICA_NUEVO.Value = c.Offset(0, 19).Value
        
        'DATOS ABAJO
        
        Me.OBSERVACIONES.Value = c.Offset(0, 25).Value
        Me.RESPONSABLE.Value = c.Offset(0, 24).Value
        Me.REGISTRAR.Visible = False
        Me.IMPRIMIR.Visible = True

        me.comparativa.value = 0

        Call CargarFoto(Me.NOMBRES.Value)
        
        Call BloquearTodo

        Me.IMPRIMIR.SetFocus
        
End Sub


Private Sub BloquearTodo()

        '===========================
        ' Bloquear todos los campos
        '===========================
        bloquearGenerales
        bloquearActuales
        bloquearNuevos
        bloquearAbajo

end sub

end Sub BloquearGenerales()

            'Datos generales
        me.emp.locked = true
        me.motivo.locked = true
        me.nombres.locked = true
        me.cedula.locked = true
        me.edad.locked = true
        me.femenino.locked = true
        me.masculino.locked = true
        me.urbano.locked = true
        me.rural.locked = true
        me.memo.locked = true
        me.fecha.locked = true


        'Datos generales
        me.emp.tabstop = false
        me.motivo.tabstop = false
        me.nombres.tabstop = false
        me.cedula.tabstop = false
        me.edad.tabstop = false
        me.femenino.tabstop = false
        me.masculino.tabstop = false
        me.urbano.tabstop = false
        me.rural.tabstop = false
        me.memo.tabstop = false
        me.fecha.tabstop = false

        'Datos generales
        me.emp.mousepointer = fmMousePointerNoDrop
        me.motivo.mousepointer = fmMousePointerNoDrop
        me.nombres.mousepointer = fmMousePointerNoDrop
        me.cedula.mousepointer = fmMousePointerNoDrop
        me.edad.mousepointer = fmMousePointerNoDrop
        me.femenino.mousepointer = fmMousePointerNoDrop
        me.masculino.mousepointer = fmMousePointerNoDrop
        me.urbano.mousepointer = fmMousePointerNoDrop
        me.rural.mousepointer = fmMousePointerNoDrop
        me.memo.mousepointer = fmMousePointerNoDrop
        me.fecha.mousepointer = fmMousePointerNoDrop
        
End Sub

Sub BloquearActuales()

        'datos actuales
        me.cargo.locked = true
        me.cargo_ocupacional.locked = true
        me.clasificacion_cargo.locked = true
        me.ubicacion.locked = true
        me.ubicacion_general.locked = true
        me.ubicacion_especifica.locked = true

        'datos actuales
        me.cargo.tabstop = false
        me.cargo_ocupacional.tabstop = false
        me.clasificacion_cargo.tabstop = false
        me.ubicacion.tabstop = false
        me.ubicacion_general.tabstop = false
        me.ubicacion_especifica.tabstop = false

        'datos actuales
        me.cargo.mousepointer = fmMousePointerNoDrop
        me.cargo_ocupacional.mousepointer = fmMousePointerNoDrop
        me.clasificacion_cargo.mousepointer = fmMousePointerNoDrop
        me.ubicacion.mousepointer = fmMousePointerNoDrop
        me.ubicacion_general.mousepointer = fmMousePointerNoDrop
        me.ubicacion_especifica.mousepointer = fmMousePointerNoDrop
        

End Sub

Sub BloquearNuevos()

        'datos nuevos
        me.cargo_nuevo.locked = true
        me.cargo_ocupacional_nuevo.locked = true
        me.clasificacion_cargo_nuevo.locked = true
        me.ubicacion_nuevo.locked = true
        me.ubicacion_general_nuevo.locked = true
        me.ubicacion_especifica_nuevo.locked = true
        
        'datos nuevos
        me.cargo_nuevo.tabstop = false
        me.cargo_ocupacional_nuevo.tabstop = false
        me.clasificacion_cargo_nuevo.tabstop = false
        me.ubicacion_nuevo.tabstop = false
        me.ubicacion_general_nuevo.tabstop = false
        me.ubicacion_especifica_nuevo.tabstop = false

        'datos nuevos
        me.cargo_nuevo.mousepointer = fmMousePointerNoDrop
        me.cargo_ocupacional_nuevo.mousepointer = fmMousePointerNoDrop
        me.clasificacion_cargo_nuevo.mousepointer = fmMousePointerNoDrop
        me.ubicacion_nuevo.mousepointer = fmMousePointerNoDrop
        me.ubicacion_general_nuevo.mousepointer = fmMousePointerNoDrop
        me.ubicacion_especifica_nuevo.mousepointer = fmMousePointerNoDrop
End Sub


Sub BloquearAbajo()
    
        'datos abajo
        me.observaciones.locked = true
        me.responsable.locked = true

        'datos abajo
        me.observaciones.tabstop = false
        me.responsable.tabstop = false

        'datos abajo
        me.observaciones.mousepointer = fmMousePointerNoDrop
        me.responsable.mousepointer = fmMousePointerNoDrop
End Sub    



    Private Sub DesbloquearTodo()

    '===========================
    ' Desbloquear todos los campos
    '===========================
    desbloquearGenerales
    desbloquearActuales
    desbloquearNuevos
    desbloquearAbajo
        

    end sub


    Sub DesbloquearGenerales()
        
    

    'Datos generales
    me.emp.locked = false
    me.motivo.locked = false
    me.nombres.locked = false
    me.cedula.locked = false
    me.edad.locked = false
    me.femenino.locked = false
    me.masculino.locked = false
    me.urbano.locked = false
    me.rural.locked = false
    me.memo.locked = false
    me.fecha.locked = false

    'Datos generales
    me.emp.tabstop = true
    me.motivo.tabstop = true
    me.nombres.tabstop = true
    me.cedula.tabstop = true
    me.edad.tabstop = true
    me.femenino.tabstop = true
    me.masculino.tabstop = true
    me.urbano.tabstop = true
    me.rural.tabstop = true
    me.memo.tabstop = true
    me.fecha.tabstop = true

    'Datos generales
    me.emp.mousepointer = fmMousePointerDefault
    me.motivo.mousepointer = fmMousePointerDefault
    me.nombres.mousepointer = fmMousePointerDefault
    me.cedula.mousepointer = fmMousePointerDefault
    me.edad.mousepointer = fmMousePointerDefault
    me.femenino.mousepointer = fmMousePointerDefault
    me.masculino.mousepointer = fmMousePointerDefault
    me.urbano.mousepointer = fmMousePointerDefault
    me.rural.mousepointer = fmMousePointerDefault
    me.memo.mousepointer = fmMousePointerDefault
    me.fecha.mousepointer = fmMousePointerDefault

    End Sub

    Sub DesbloquearActuales()
        
   

    'datos actuales
    me.cargo.locked = false
    me.cargo_ocupacional.locked = false
    me.clasificacion_cargo.locked = false
    me.ubicacion.locked = false
    me.ubicacion_general.locked = false
    me.ubicacion_especifica.locked = false

    'datos actuales
    me.cargo.tabstop = true
    me.cargo_ocupacional.tabstop = true
    me.clasificacion_cargo.tabstop = true
    me.ubicacion.tabstop = true
    me.ubicacion_general.tabstop = true
    me.ubicacion_especifica.tabstop = true

    'datos actuales
    me.cargo.mousepointer = fmMousePointerDefault
    me.cargo_ocupacional.mousepointer = fmMousePointerDefault
    me.clasificacion_cargo.mousepointer = fmMousePointerDefault
    me.ubicacion.mousepointer = fmMousePointerDefault
    me.ubicacion_general.mousepointer = fmMousePointerDefault
    me.ubicacion_especifica.mousepointer = fmMousePointerDefault

    End Sub

    Sub DesbloquearNuevos()
        


    'datos nuevos
    me.cargo_nuevo.locked = false
    me.cargo_ocupacional_nuevo.locked = false
    me.clasificacion_cargo_nuevo.locked = false
    me.ubicacion_nuevo.locked = false
    me.ubicacion_general_nuevo.locked = false
    me.ubicacion_especifica_nuevo.locked = false

    'datos nuevos
    me.cargo_nuevo.tabstop = true
    me.cargo_ocupacional_nuevo.tabstop = true
    me.clasificacion_cargo_nuevo.tabstop = true
    me.ubicacion_nuevo.tabstop = true
    me.ubicacion_general_nuevo.tabstop = true
    me.ubicacion_especifica_nuevo.tabstop = true

    'datos nuevos
    me.cargo_nuevo.mousepointer = fmMousePointerDefault
    me.cargo_ocupacional_nuevo.mousepointer = fmMousePointerDefault
    me.clasificacion_cargo_nuevo.mousepointer = fmMousePointerDefault
    me.ubicacion_nuevo.mousepointer = fmMousePointerDefault
    me.ubicacion_general_nuevo.mousepointer = fmMousePointerDefault
    me.ubicacion_especifica_nuevo.mousepointer = fmMousePointerDefault

    End Sub

    Sub DesbloquearAbajo()
        


    'datos abajo
    me.observaciones.locked = false
    me.responsable.locked = false
   

    'datos abajo
    me.observaciones.tabstop = true
    me.responsable.tabstop = true


    'datos abajo
    me.observaciones.mousepointer = fmMousePointerDefault
    me.responsable.mousepointer = fmMousePointerDefault

    End Sub

Private Sub CARGO_Change()
Dim Clas, Carg As Range
                'Verificar si el campo de Cargo_Nuevo esta vacio
                If Me.CARGO.Value = Empty Then
                Me.CLASIFICACION_CARGO.Clear
                
                'Agregar los elementos de la tabla Clasificacion_Cargo de la Hoja24 a la variable Clas
                Else
                Set Clas = Hoja24.ListObjects("CLASIFICACION_CARGO").DataBodyRange
                Me.CLASIFICACION_CARGO.Clear
                
                'Agregar Cada ClasificaciOn de Cargos al Listado
                For Each Carg In Clas
                Me.CLASIFICACION_CARGO.AddItem Carg.Value
                Next Carg
                End If
End Sub


Private Sub CLASIFICACION_CARGO_AfterUpdate()

Dim tbl As ListObject
    Dim c As Range
    Dim Clas As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("CLASIFICACION_CARGO")
    
    Clas = Trim(Me.CLASIFICACION_CARGO.Value)
    
    'Si esta vacio ? limpiar y salir
    If Clas = "" Then
        Exit Sub
    End If
    
    'Buscar ubicaciOn en columna 1 (Col A = UbicaciOn)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Clas, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        MsgBox "La Clasificacion de Cargo no Existe", vbExclamation
        Me.CLASIFICACION_CARGO.SetFocus
        Exit Sub
        
    End If

End Sub
Private Sub CLASIFICACION_CARGO_NUEVO_AfterUpdate()

Dim tbl As ListObject
    Dim c As Range
    Dim Clas As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("CLASIFICACION_CARGO")
    
    Clas = Trim(Me.CLASIFICACION_CARGO_NUEVO.Value)
    
    'Si esta vacio ? limpiar y salir
    If Clas = "" Then
        Exit Sub
    End If
    
    'Buscar ubicaciOn en columna 1 (Col A = UbicaciOn)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Clas, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        MsgBox "La Clasificacion de Cargo Nuevo no Existe", vbExclamation
        Me.CLASIFICACION_CARGO_NUEVO.SetFocus
        Exit Sub
        
    End If

End Sub

Private Sub RESPONSABLE_Change()
    If Me.RESPONSABLE.Value = "" Then
        'Me.REGISTRAR.Visible = False
        Me.REGISTRAR.Enabled = False
    Else
        Me.REGISTRAR.Visible = True
        Me.REGISTRAR.Enabled = True
        
    End If

End Sub

Private Sub RESPONSABLE_AfterUpdate()

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
        Me.RESPONSABLE.Value = ""
        Me.RESPONSABLE.SetFocus
        Exit Sub
        
    End If
        
End Sub

Private Sub UBICACION_Change()
Dim Depe, Ubi As Range
        If Me.UBICACION.Value = Empty Then
        Me.UBICACION_GENERAL.Clear
        
        Else
        'Asignar a cada Dependencia su area ESPECIFICA
        Set Depe = Hoja24.ListObjects(Me.UBICACION.Value).DataBodyRange
        Me.UBICACION_GENERAL.Clear
        
                'Agregar Cada area ESPECIFICA al Listado
                For Each Ubi In Depe
                Me.UBICACION_GENERAL.AddItem Ubi.Value
                Next Ubi
        End If
    
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

Private Sub UBICACION_NUEVO_AfterUpdate()
    
    Dim tbl As ListObject
    Dim c As Range
    Dim Ubi As String
    
    'Referencia a la tabla
    Set tbl = Hoja24.ListObjects("UBICACION")
    
    Ubi = Trim(Me.UBICACION_NUEVO.Value)
    
    'Si esta vacio ? limpiar y salir
    If Ubi = "" Then
        Me.UBICACION_GENERAL_NUEVO.Clear
        Exit Sub
    End If
    
    'Buscar ubicaciOn en columna 1 (Col A = UbicaciOn)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Ubi, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        
        Me.UBICACION_GENERAL_NUEVO.Clear
        MsgBox "La Ubicacion no Existe", vbExclamation
        Me.UBICACION_NUEVO.SetFocus
        Exit Sub
        
    End If
    
    'Asignar a cada Dependencia su area ESPECIFICA
    Dim Dependencia, Dep As Range
        Set Dependencia = Hoja24.ListObjects(Me.UBICACION_NUEVO.Value).DataBodyRange
        Me.UBICACION_GENERAL_NUEVO.Clear
        
        'Agregar Cada area ESPECIFICA al Listado
        For Each Dep In Dependencia
            Me.UBICACION_GENERAL_NUEVO.AddItem Dep.Value
        Next Dep
    
End Sub


Private Sub EMP_AfterUpdate()
    If IgnorarValidacionEMP Then Exit Sub 'validacion de COD

    Dim tbl As ListObject
    Dim c As Range
    Dim empVal, Mot, Baja As String
    
    'Referencia a la tabla
    Set tbl = Hoja25.ListObjects("NOMINA")
    
    Mot = Me.MOTIVO.Value
    empVal = Trim(Me.EMP.Value)
    
    'Si esta vacio ? limpiar y salir
    If empVal = "" Then
        UserForm_Initialize
        UserForm_Activate
        LimpiarCampos
        Me.MOTIVO.Value = Mot
        IgnorarValidacionCOD = False
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
        Me.EMP.SetFocus
        Exit Sub

    End If
    
    
    Baja = c.Offset(0, 17).Value
    
    If Baja <> "" Then
        
        UserForm_Initialize
        UserForm_Activate
        LimpiarCampos
        Me.MOTIVO.Value = Mot
        MsgBox "Empleado se Encuentra en Estado Inactivo (BAJA)", vbExclamation
        Me.EMP.Value = ""
        Me.EMP.SetFocus
        Exit Sub
    End If
    
    '===========================
    'Mostrar campos de los datos
    '===========================
    
    IgnorarValidacionCOD = True
    desbloquear
    me.cod.value = ""
    me.cod.locked = true
    me.cod.tabstop = false
    me.cod.mousepointer = fmMousePointerNodrop

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
        
    End Select
        Me.COMPARATIVA.Value = 0
    
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
    Me.NOMBRES.Value = c.Offset(0, 1).Value     'Col B
    Me.CEDULA.Value = c.Offset(0, 12).Value     'Col M
    Me.EDAD.Value = c.Offset(0, 15).Value       'Col P
    
    'Genero (Col O)
    SeleccionarGenero c.Offset(0, 14).Value
    
    Me.CARGO_OCUPACIONAL.Value = c.Offset(0, 7).Value  'Col H
    Me.CARGO.Value = c.Offset(0, 8).Value                'Col I
    Me.CLASIFICACION_CARGO.Value = c.Offset(0, 9).Value  'Col J
    Me.UBICACION.Value = c.Offset(0, 3).Value            'Col D
    Me.UBICACION_GENERAL.Value = c.Offset(0, 5).Value    'Col F
    Me.UBICACION_ESPECIFICA.Value = c.Offset(0, 6).Value 'Col G
    
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
        
        
' ============================================
' ValidaciOn para la clasificaciOn del cargo actual
' ============================================
    If Me.CLASIFICACION_CARGO.Value = "" Then
        Me.CLASIFICACION_CARGO.MousePointer = fmMousePointerDefault
        Me.CLASIFICACION_CARGO.Locked = False
    Else
        Me.CLASIFICACION_CARGO.MousePointer = fmMousePointerNoDrop
        Me.CLASIFICACION_CARGO.Locked = True
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
    
    me.cod.locked = true
    me.cod.tabstop = false
    me.cod.mousepointer = fmMousePointerNoDrop

' ============================================
' EVENTO PARA CARGAR LAS FOTOS DEL TRABAJADOR
' ============================================
   
    Call CargarFoto(Me.NOMBRES.Value)


    
End Sub

'=========================
'SelecciOn de Genero
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
'SelecciOn de Domicilio
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


Private Sub EMP_Exit(ByVal Cancel As MSForms.ReturnBoolean)


' ValidaciOn adicional al perder el foco

    If Not EMP.Text Like "*[!0-9]*" Then
        ' Texto valido (contiene solo numeros y /)
    Else
        MsgBox "Solo se permiten numeros", vbExclamation
        Cancel = True
        EMP.SetFocus
    End If

End Sub

Private Sub MOTIVO_AfterUpdate()

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
        Exit Sub
        
    End If
    
    Me.EMP_ETIQUETA.Visible = True
    Me.EMP.Visible = True
    Me.EMP.Value = ""
    Me.EMP.SetFocus

End Sub
Private Sub COD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
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


Private Sub CARGO_NUEVO_Change()

Dim Clas, Carg As Range
                'Verificar si el campo de Cargo_Nuevo esta vacio
                If Me.CARGO_NUEVO.Value = Empty Then
                Me.CLASIFICACION_CARGO_NUEVO.Clear
                
                'Agregar los elementos de la tabla Clasificacion_Cargo de la Hoja24 a la variable Clas
                Else
                Set Clas = Hoja24.ListObjects("CLASIFICACION_CARGO").DataBodyRange
                Me.CLASIFICACION_CARGO_NUEVO.Clear
                
                'Agregar Cada ClasificaciOn de Cargos al Listado
                For Each Carg In Clas
                Me.CLASIFICACION_CARGO_NUEVO.AddItem Carg.Value
                Next Carg
                End If
                
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
        If MsgBox("¿Esta seguro de guardar el registro?", vbQuestion + vbYesNo, "Confirmar") = vbYes Then
            GuardarRegistro
        End If
    End If
End Sub

' ============================================
' VALIDACION PARA TRASLADO Y NOMBRAMIENTO
' ============================================
Private Function ValidarTrasladoYNombramiento() As String
    Dim faltantes As String
    Dim listaCampos As String
    
    ' Verificar campos obligatorios
    If Trim(Me.CLASIFICACION_CARGO.Value) = "" Then
        faltantes = faltantes & "+ CLASIFICACION CARGO ACTUAL" & vbCrLf
    End If
    
    If Trim(Me.CARGO_NUEVO.Value) = "" Then
        faltantes = faltantes & "+ CARGO FUNCIONAL NUEVO" & vbCrLf
    End If
    
    If Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
        faltantes = faltantes & "+ CARGO OCUPACIONAL NUEVO" & vbCrLf
    End If
    
    If Me.CLASIFICACION_CARGO_NUEVO.ListIndex = -1 Then
        faltantes = faltantes & "+ CLASIFICACION DE CARGO NUEVO" & vbCrLf
    End If
    
    If Me.UBICACION_NUEVO.ListIndex = -1 Then
        faltantes = faltantes & "+ UBICACION NUEVA" & vbCrLf
    End If
    
    If Me.UBICACION_GENERAL_NUEVO.ListIndex = -1 Then
        faltantes = faltantes & "+ UBICACION GENERAL NUEVA" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_ESPECIFICA_NUEVO.Value) = "" Then
        faltantes = faltantes & "+ UBICACION ESPECIFICA NUEVA" & vbCrLf
    End If
    
    If Trim(Me.FECHA.Value) = "" Then
        faltantes = faltantes & "+ FECHA" & vbCrLf
    End If
    
    If Trim(Me.MEMO.Value) = "" Then
        faltantes = faltantes & "+ MEMO" & vbCrLf
    End If
    
    ' Validar al menos una opciOn URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        faltantes = faltantes & "+ URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If
    
    ValidarTrasladoYNombramiento = faltantes
End Function

' ============================================
' VALIDACION PARA TRASLADO
' ============================================
Private Function ValidarTraslado() As String
    Dim faltantes As String
    
    ' Verificar campos obligatorios
    If Trim(Me.CLASIFICACION_CARGO.Value) = "" Then
        faltantes = faltantes & "+ CLASIFICACION CARGO ACTUAL" & vbCrLf
    End If
    
    If Me.UBICACION_NUEVO.ListIndex = -1 Then
        faltantes = faltantes & "+ UBICACION NUEVA" & vbCrLf
    End If
    
    If Me.UBICACION_GENERAL_NUEVO.ListIndex = -1 Then
        faltantes = faltantes & "+ UBICACION GENERAL NUEVA" & vbCrLf
    End If
    
    If Trim(Me.UBICACION_ESPECIFICA_NUEVO.Value) = "" Then
        faltantes = faltantes & "+ UBICACION ESPECIFICA NUEVA" & vbCrLf
    End If
    
    If Trim(Me.FECHA.Value) = "" Then
        faltantes = faltantes & "+ FECHA" & vbCrLf
    End If
    
    If Trim(Me.MEMO.Value) = "" Then
        faltantes = faltantes & "+ MEMO" & vbCrLf
    End If
    
    ' Validar al menos una opciOn URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        faltantes = faltantes & "+ URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If
    
    ValidarTraslado = faltantes
End Function

' ============================================
' VALIDACION PARA NOMBRAMIENTO
' ============================================
Private Function ValidarNombramiento() As String
    Dim faltantes As String
    
    ' Verificar campos obligatorios
    If Trim(Me.CLASIFICACION_CARGO.Value) = "" Then
        faltantes = faltantes & "+ CLASIFICACION CARGO ACTUAL" & vbCrLf
    End If
    
    If Trim(Me.CARGO_NUEVO.Value) = "" Then
        faltantes = faltantes & "+ CARGO FUNCIONAL NUEVO" & vbCrLf
    End If
    
    If Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
        faltantes = faltantes & "+ CARGO OCUPACIONAL NUEVO" & vbCrLf
    End If
    
    If Me.CLASIFICACION_CARGO_NUEVO.ListIndex = -1 Then
        faltantes = faltantes & "+ CLASIFICACION DE CARGO NUEVO" & vbCrLf
    End If
    
    If Trim(Me.FECHA.Value) = "" Then
        faltantes = faltantes & "+ FECHA" & vbCrLf
    End If
    
    If Trim(Me.MEMO.Value) = "" Then
        faltantes = faltantes & "+ MEMO" & vbCrLf
    End If
    
    ' Validar al menos una opciOn URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        faltantes = faltantes & "+ URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If
    
    ValidarNombramiento = faltantes
End Function

' ============================================
' FUNCION PARA ENFOCAR EL PRIMER CAMPO FALTANTE
' ============================================
Private Sub EnfocarPrimerCampoFaltante(MOTIVO As String)
    ' Esta funciOn puede enfocar el primer campo que falta
    ' segun el motivo, para facilitar la correcciOn
    
    Me.COMPARATIVA.Value = 1
    Select Case MOTIVO
        Case "TRASLADO Y NOMBRAMIENTO"
            If Trim(Me.CARGO_NUEVO.Value) = "" Then
                Me.CARGO_NUEVO.SetFocus
            ElseIf Trim(Me.CARGO_OCUPACIONAL_NUEVO.Value) = "" Then
                Me.CARGO_OCUPACIONAL_NUEVO.SetFocus
            ElseIf Me.CLASIFICACION_CARGO_NUEVO.ListIndex = -1 Then
                Me.CLASIFICACION_CARGO_NUEVO.SetFocus
            ElseIf Me.UBICACION_NUEVO.ListIndex = -1 Then
                Me.UBICACION_NUEVO.SetFocus
            End If
    End Select
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
    Range(CeldaMotivo).End(xlDown).Offset(0, 6).Value = Me.CLASIFICACION_CARGO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 7).Value = Me.UBICACION.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 8).Value = Me.UBICACION_GENERAL.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 9).Value = Me.UBICACION_ESPECIFICA.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 10).Value = Me.CARGO_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 11).Value = Me.CLASIFICACION_CARGO_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 12).Value = Me.UBICACION_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 13).Value = Me.UBICACION_GENERAL_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 14).Value = Me.UBICACION_ESPECIFICA_NUEVO.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 15).Value = Me.FECHA.Value
        If Me.FEMENINO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 16).Value = "F"
        ElseIf Me.MASCULINO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 16).Value = "M"
        End If
    Range(CeldaMotivo).End(xlDown).Offset(0, 17).Value = Me.MEMO.Value
        If Me.URBANO.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 18).Value = "URBANO"
        ElseIf Me.RURAL.Value = True Then
            Range(CeldaMotivo).End(xlDown).Offset(0, 18).Value = "RURAL"
        End If
    Range(CeldaMotivo).End(xlDown).Offset(0, 19).Value = Me.RESPONSABLE.Value
    Range(CeldaMotivo).End(xlDown).Offset(0, 20).Value = Me.OBSERVACIONES.Value


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

MsgBox "El " & Me.MOTIVO.Value & " fue guardado con Exito." & vbCrLf & _
       "Y Datos del Trabajador fueron editados en NOMINA", _
       vbInformation, "Exito"

Call UserForm_Initialize
Call UserForm_Activate

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

