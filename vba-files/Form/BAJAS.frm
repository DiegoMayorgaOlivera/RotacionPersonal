VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BAJAS 
   Caption         =   "BAJAS"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18255
   OleObjectBlob   =   "BAJAS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "BAJAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'====================================================================================================================================
' Formulario para gestionar las bajas de empleados. 
' Permite registrar nuevas bajas y consultar las ya registradas.
'====================================================================================================================================



Dim Borrador As Boolean
Option Explicit
' ------------------------------------------------
' Constantes compartidas
' ------------------------------------------------
Private Const PREFIJO        As String = "C$ "
Private Const LONG_PREFIJO   As Long = 3

' Variables para SALARIO_BASE
Private m_Salario As Double
Private m_SalarioEnFoco As Boolean

' Variables para ANTIGUEDAD_SALARIAL
Private m_Antiguedad As Double
Private m_AntiguedadEnFoco As Boolean

'=====================================
' RUTA BASE DE LAS FOTOS
'=====================================
Public Function ObtenerRutaFotos() As String
    ObtenerRutaFotos = ThisWorkbook.Path & "\FOTOS\"
End Function


Private Sub UserForm_Activate()

Me.Top = 0
Me.Left = 0
Me.ScrollBars = fmScrollBarsNone
Me.EMP.SetFocus

End Sub

Private Sub UserForm_Initialize()

    LimpiarTodo
    Me.EMP.Value = ""
    
    Dim rutaBase As String
    rutaBase = ObtenerRutaFotos()
    OcultarTodo

End Sub
Private Sub EMP_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    Dim tblBj, tblNom As ListObject
    Dim c, d As Range
    Dim empVal, nombreEmpleado As String
    
    'Referencia a la tabla
    Set tblBj = Hoja3.ListObjects("BAJAS")
    
    empVal = Trim(Me.EMP.Value)
    
    If empVal = "" Then
        OcultarTodo
        LimpiarTodo
        Cancel = True   '  impide que el foco salga del control
        Exit Sub
    End If
    
    'Buscar empleado en Tabla de Bajas
    Set c = tblBj.ListColumns(5).DataBodyRange.Find( _
                What:=empVal, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
            
                    'Referencia a la tabla
                Set tblNom = Hoja25.ListObjects("NOMINA")
                
                'Buscar empleado en la Tabla de Nomina
                Set d = tblNom.ListColumns(1).DataBodyRange.Find( _
                            What:=empVal, _
                            LookAt:=xlWhole, _
                            MatchCase:=False)
                
                'Si NO existe
                If d Is Nothing Then
                    LimpiarTodo
                    OcultarTodo
                    Me.EMP.Value = ""
                    MsgBox "Empleado no existe en la nomina", vbExclamation
                    Cancel = True   '  impide que el foco salga del control
                    Exit Sub
                    
                End If
        
        Call NuevaBaja
    Else
        nombreEmpleado = c.Offset(0, 1).Value              'Col H
        MsgBox "Ya Existe una Baja Registrada para el Empleado: " & vbNewLine & vbNewLine & empVal & " " & nombreEmpleado, vbExclamation
        BajaRegistrada
        
    End If
    
    
    
End Sub

Private Sub BajaRegistrada()

    Dim tbl As ListObject
    Dim c As Range
    Dim empVal As String
    
    'Referencia a la tabla
    Set tbl = Hoja3.ListObjects("BAJAS")
    
    empVal = Trim(Me.EMP.Value)
    
    If empVal = "" Then
        OcultarTodo
        LimpiarTodo
        Exit Sub
    End If
    
    'Buscar empleado en columna 5 (Col G = No. EMP)
    Set c = tbl.ListColumns(5).DataBodyRange.Find( _
                What:=empVal, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    '=========================
    'Agregar ComboBoxes
    '=========================
    LimpiarTodo
    MostrarTodo
    Me.DATOS.Value = 0
    
    Call CargarComboDesdeTabla(Me.CLASIFICACION_CARGO, "AUX", "CLASIFICACION_CARGO", , True)
    Call CargarComboDesdeTabla(Me.UBICACION, "AUX", "UBICACION", , True)
    Call CargarComboDesdeTabla(Me.UBICACION_GENERAL, "AUX", "DEPENDENCIAS", "DEPENDENCIA CORTA", True)
    Call CargarComboDesdeTabla(Me.ESTADO_CIVIL, "AUX", "ESTADO_CIVIL", , True)
    Call CargarComboDesdeTabla(Me.NIVEL_ACADEMICO, "AUX", "NIVEL_ACADEMICO", , True)
    Call CargarComboDesdeTabla(Me.TIPO_BAJA, "AUX", "TIPO_BAJA", , True)
    Call CargarComboDesdeTabla(Me.RESPONSABLE, "AUX", "RESPONSABLE", , True)
    Call CargarComboDesdeTabla(Me.RESPONSABLE_ANULACION, "AUX", "RESPONSABLE", , True)
    
    
    '=================================
    'Cargar Datos desde Nomina
    '=================================


    Me.NOMBRES.Value = c.Offset(0, 1).Value              'Col H
    Me.CEDULA.Value = c.Offset(0, 2).Value               'Col I
    Me.EDAD.Value = c.Offset(0, 18).Value                'Col Y
    
    SeleccionarGenero c.Offset(0, 17).Value              'Col X

    Me.CARGO.Value = c.Offset(0, 10).Value               'Col Q
    Me.CLASIFICACION_CARGO.Value = c.Offset(0, 20).Value 'Col AA
    Me.CLASIFICACION_CARGO.Locked = True
    Me.UBICACION.Value = c.Offset(0, 7).Value            'Col N
    Me.UBICACION_GENERAL.Value = c.Offset(0, 8).Value    'Col O
    Me.UBICACION_ESPECIFICA.Value = c.Offset(0, 9).Value 'Col P
    
    Call CargarFoto(Me.NOMBRES.Value)
    
    Me.ESTADO_CIVIL.Value = c.Offset(0, 15).Value        'COL V
    SeleccionarDependientes c.Offset(0, 16).Value        'Col W
    SeleccionarDomicilio c.Offset(0, 21).Value           'Col AB
    Me.NIVEL_ACADEMICO.Value = c.Offset(0, 13).Value     'COL T
    Me.ESTUDIO.Value = c.Offset(0, 14).Value             'COL U
    
    Me.FECHA_INGRESO.Value = c.Offset(0, 11).Value       'Col R
    Me.ANTIGUEDAD_LABORAL.Value = c.Offset(0, 12).Value  'Col S
    
    Dim SalarioBase, AntiguedadSalarial As Variant
    
    SalarioBase = c.Offset(0, 5).Value         'COL L
    Me.SALARIO_BASE.Value = "C$ " & Format(SalarioBase, "#,##0.00")
    AntiguedadSalarial = c.Offset(0, 6).Value  'COL M
    Me.ANTIGUEDAD_SALARIAL.Value = "C$ " & Format(AntiguedadSalarial, "#,##0.00")
    Me.FECHA_BAJA.Value = c.Offset(0, 3).Value           'COL J
    Me.TIPO_BAJA.Value = c.Offset(0, 22).Value           'COL AC
    Me.MOTIVO_BAJA.Value = c.Offset(0, 23).Value         'COL AD
    Me.RESPONSABLE.Value = c.Offset(0, 24).Value         'COL AE
    Me.FECHA_REGISTRO.Value = c.Offset(0, 4).Value       'COL K
    
    '=========================
    'Preparar tabla para Antiguedad Laboral
    '=========================
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 0).NumberFormat = "dd/mm/yyyy"
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 0).Value = CDate(Me.FECHA_INGRESO.Value)
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 1).NumberFormat = "dd/mm/yyyy"
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 1).Value = CDate(Me.FECHA_BAJA.Value)
    
    
    Me.ANULAR.Visible = True
    Me.height = 540
    Me.RESPONSABLE_ANULACION_ETIQUETA.Visible = True
    Me.RESPONSABLE_ANULACION.Visible = True
    
    
End Sub

Private Sub NuevaBaja()

    Dim tbl As ListObject
    Dim c As Range
    Dim empVal As String
    
    'Referencia a la tabla
    Set tbl = Hoja25.ListObjects("NOMINA")
    
    empVal = Trim(Me.EMP.Value)
    
    'Si esta vacio ? limpiar y salir
    If empVal = "" Then
        OcultarTodo
        LimpiarTodo
        
        Exit Sub
    End If
    
    'Buscar empleado en columna 1 (Col A = No. EMP)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=empVal, _
                LookAt:=xlWhole, _
                MatchCase:=False)
    
    'Si NO existe
    If c Is Nothing Then
        LimpiarTodo
        OcultarTodo
        Me.EMP.Value = ""
        MsgBox "Empleado no existe en la nomina", vbExclamation
        Me.EMP.SetFocus
        Exit Sub
        
    End If
    
    
    '=================================
    'Cargar Datos desde Nomina
    '=================================

    LimpiarTodo
    MostrarTodo
    Me.DATOS.Value = 0
    Me.NOMBRES.Value = c.Offset(0, 1).Value              'Col B
    Me.CEDULA.Value = c.Offset(0, 12).Value              'Col M
    Me.EDAD.Value = c.Offset(0, 15).Value                'Col P
    
    SeleccionarGenero c.Offset(0, 14).Value              'Col 0

    Me.CARGO.Value = c.Offset(0, 8).Value                'Col I
    Me.CLASIFICACION_CARGO.Value = c.Offset(0, 9).Value  'Col J
    Me.UBICACION.Value = c.Offset(0, 3).Value            'Col D
    Me.UBICACION_GENERAL.Value = c.Offset(0, 5).Value    'Col F
    Me.UBICACION_ESPECIFICA.Value = c.Offset(0, 6).Value 'Col G
    Me.FECHA_INGRESO.Value = c.Offset(0, 10).Value       'Col K
    Me.ANTIGUEDAD_LABORAL.Value = c.Offset(0, 11).Value  'Col L
    Call CargarFoto(Me.NOMBRES.Value)
        
    '=========================
    'Preparar tabla para Antiguedad Laboral
    '=========================
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 0).NumberFormat = "dd/mm/yyyy"
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 0).Value = CDate(Me.FECHA_INGRESO.Value)
    Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 1).Value = ""
    
    '=========================
    'Agregar ComboBoxes
    '=========================
    
    Call CargarComboDesdeTabla(Me.UBICACION, "AUX", "UBICACION", , True)
    Call CargarComboDesdeTabla(Me.CLASIFICACION_CARGO, "AUX", "CLASIFICACION_CARGO", , True)
    Call CargarComboDesdeTabla(Me.UBICACION_GENERAL, "AUX", "DEPENDENCIAS", "DEPENDENCIA CORTA", True)
    Call CargarComboDesdeTabla(Me.ESTADO_CIVIL, "AUX", "ESTADO_CIVIL", , True)
    Call CargarComboDesdeTabla(Me.NIVEL_ACADEMICO, "AUX", "NIVEL_ACADEMICO", , True)
    Call CargarComboDesdeTabla(Me.TIPO_BAJA, "AUX", "TIPO_BAJA", , True)
    Call CargarComboDesdeTabla(Me.RESPONSABLE, "AUX", "RESPONSABLE", , True)
    
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

Private Sub URBANO_Click()

    If Me.URBANO.Value = True Then
        Me.URBANO.ForeColor = &HFF0000
        Me.RURAL.ForeColor = &H0&
    Else
        Me.URBANO.ForeColor = &H0&
        Me.RURAL.ForeColor = &HFF0000
    End If
    
End Sub
Private Sub RURAL_Click()

    If Me.RURAL.Value = True Then
        Me.RURAL.ForeColor = &HFF0000
        Me.URBANO.ForeColor = &H0&
    Else
        Me.RURAL.ForeColor = &H0&
        Me.URBANO.ForeColor = &HFF0000
    End If

End Sub

'=========================
'SelecciOn de Dependientes
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

Private Sub SI_Click()

    If Me.SI.Value = True Then
        Me.SI.ForeColor = &HFF0000
        Me.NO.ForeColor = &H0&
    Else
        Me.SI.ForeColor = &H0&
        Me.NO.ForeColor = &HFF0000
    End If
    
End Sub
Private Sub NO_Click()

    If Me.NO.Value = True Then
        Me.NO.ForeColor = &HFF0000
        Me.SI.ForeColor = &H0&
    Else
        Me.NO.ForeColor = &H0&
        Me.SI.ForeColor = &HFF0000
    End If

End Sub



' ============================================
' FUNCION PARA CARGAR DATOS DE TABLA A COMBOBOX
' ============================================

Public Sub CargarComboDesdeTabla( _
        ByVal cbo As MSForms.ComboBox, _
        ByVal NombreHoja As String, _
        ByVal NombreTabla As String, _
        Optional ByVal NombreColumna As String = "", _
        Optional ByVal QuitarDuplicados As Boolean = False)

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim CELDA As Range
    Dim dict As Object
    
    'Validar hoja
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Worksheets(NombreHoja)
    Set tbl = ws.ListObjects(NombreTabla)
    On Error GoTo 0
    
    'Limpiar combo
    cbo.Clear
    
    'Si la tabla esta vacia
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    
    'Si no se especifica columna ? usar la primera
    If NombreColumna = "" Then
        Set rng = tbl.ListColumns(1).DataBodyRange
    Else
        Set rng = tbl.ListColumns(NombreColumna).DataBodyRange
    End If
    
    'Si quitar duplicados
    If QuitarDuplicados Then
        Set dict = CreateObject("Scripting.Dictionary")
    End If
    
    'Recorrer rango
    For Each CELDA In rng
        
        If Trim(CELDA.Value) <> "" Then
            
            If QuitarDuplicados Then
                
                If Not dict.exists(CELDA.Value) Then
                    dict.Add CELDA.Value, Nothing
                    cbo.AddItem CELDA.Value
                End If
                
            Else
                cbo.AddItem CELDA.Value
            End If
            
        End If
        
    Next CELDA
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al cargar ComboBox." & vbCrLf & _
           "Verifique hoja, tabla o columna.", vbCritical

End Sub

Private Sub EMP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, backspace, delete, enter, tab y .

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub ESTUDIO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub

Private Sub MOTIVO_BAJA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
    KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub

'===================================================================================================================================
'===================================================================================================================================
'Para Salario y Antiguedad

' =============================================================================
'                           SALARIO_BASE
' =============================================================================

Private Sub SALARIO_BASE_GotFocus()
    m_SalarioEnFoco = True
    MostrarSinFormatoMiles SALARIO_BASE, m_Salario
End Sub

Private Sub SALARIO_BASE_LostFocus()
    m_SalarioEnFoco = False
    MostrarConFormato SALARIO_BASE, m_Salario
End Sub

Private Sub SALARIO_BASE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Bloqueo de borrado del prefijo
    If KeyAscii = 8 Then   ' Backspace
        If SALARIO_BASE.SelStart <= LONG_PREFIJO Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    ' Solo permitir: digitos, punto (.), backspace, tab, enter
    If Not (KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 13 Or _
            (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' Control de punto decimal
    If KeyAscii = 46 Then   ' .
        Dim txt As String: txt = Mid(SALARIO_BASE.Text, LONG_PREFIJO + 1)
        If InStr(txt, ".") > 0 Then
            KeyAscii = 0    ' ya existe un punto
            Exit Sub
        End If
    End If
End Sub

Private Sub SALARIO_BASE_Change()
    Dim texto As String: texto = SALARIO_BASE.Text
    
    ' Proteger prefijo
    If Left(texto, LONG_PREFIJO) <> PREFIJO Then
        texto = PREFIJO & Mid(texto, LONG_PREFIJO + 1)
        SALARIO_BASE.Text = texto
        SALARIO_BASE.SelStart = Len(texto)
    End If
    
    ' Extraer solo la parte numerica
    Dim numStr As String
    numStr = Replace(Mid(texto, LONG_PREFIJO + 1), ",", "")   ' por si acaso
    
    If numStr = "" Or numStr = "." Then
        m_Salario = 0
    ElseIf IsNumeric(numStr) Then
        m_Salario = CDbl(numStr)
        ' Limitar a 2 decimales
        m_Salario = Int(m_Salario * 100 + 0.0000001) / 100   ' redondeo correcto
    Else
        ' Valor invalido ? mantener anterior
        MostrarSinFormatoMiles SALARIO_BASE, m_Salario
        Exit Sub
    End If
    
    ' Si esta en foco ? mantener sin separadores de miles
    If m_SalarioEnFoco Then
        MostrarSinFormatoMiles SALARIO_BASE, m_Salario
    End If
End Sub


' =============================================================================
'                       ANTIGUEDAD_SALARIAL
' =============================================================================

Private Sub ANTIGUEDAD_SALARIAL_GotFocus()
    m_AntiguedadEnFoco = True
    MostrarSinFormatoMiles ANTIGUEDAD_SALARIAL, m_Antiguedad
End Sub

Private Sub ANTIGUEDAD_SALARIAL_Lostfocus()
    m_AntiguedadEnFoco = False
    MostrarConFormato ANTIGUEDAD_SALARIAL, m_Antiguedad
End Sub

Private Sub ANTIGUEDAD_SALARIAL_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Mismo control que SALARIO_BASE
    If KeyAscii = 8 Then
        If ANTIGUEDAD_SALARIAL.SelStart <= LONG_PREFIJO Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    If Not (KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 13 Or _
            (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) Then
        KeyAscii = 0
        Exit Sub
    End If
    
    If KeyAscii = 46 Then
        Dim txt As String: txt = Mid(ANTIGUEDAD_SALARIAL.Text, LONG_PREFIJO + 1)
        If InStr(txt, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub ANTIGUEDAD_SALARIAL_Change()
    Dim texto As String: texto = ANTIGUEDAD_SALARIAL.Text
    
    If Left(texto, LONG_PREFIJO) <> PREFIJO Then
        texto = PREFIJO & Mid(texto, LONG_PREFIJO + 1)
        ANTIGUEDAD_SALARIAL.Text = texto
        ANTIGUEDAD_SALARIAL.SelStart = Len(texto)
    End If
    
    Dim numStr As String
    numStr = Replace(Mid(texto, LONG_PREFIJO + 1), ",", "")
    
    If numStr = "" Or numStr = "." Then
        m_Antiguedad = 0
    ElseIf IsNumeric(numStr) Then
        m_Antiguedad = CDbl(numStr)
        m_Antiguedad = Int(m_Antiguedad * 100 + 0.0000001) / 100
    Else
        MostrarSinFormatoMiles ANTIGUEDAD_SALARIAL, m_Antiguedad
        Exit Sub
    End If
    
    If m_AntiguedadEnFoco Then
        MostrarSinFormatoMiles ANTIGUEDAD_SALARIAL, m_Antiguedad
    End If
End Sub


' =============================================================================
'           Funciones auxiliares (compartidas)
' =============================================================================

Private Sub MostrarSinFormatoMiles(ctrl As MSForms.TextBox, ByVal valor As Double)
    If valor = 0 Then
        ctrl.Text = PREFIJO
    Else
        ' Siempre mostramos con punto decimal y exactamente 2 decimales
        ' sin separador de miles
        Dim txt As String
        txt = Format(valor, "0.00")               '  "1234.50"
        txt = Replace(txt, ",", "")               ' quitamos coma si el Format regional la puso
        ctrl.Text = PREFIJO & txt
    End If
End Sub

Private Sub MostrarConFormato(ctrl As MSForms.TextBox, ByVal valor As Double)
    If valor = 0 Then
        ctrl.Text = PREFIJO
    Else
        ' Formato final: separador de miles (coma) + siempre 2 decimales
        Dim txt As String
        txt = Format(valor, "#,##0.00")           '  "1,234.50"  (o "1.234,50" segun locale)
        ctrl.Text = PREFIJO & txt
    End If
End Sub


' ------------------------------------------------
' Metodos publicos para cargar y obtener valores
' ------------------------------------------------

Public Sub CargarSalario(valor As Variant)
    If IsNumeric(valor) Then
        m_Salario = Round(CDbl(valor), 2)
        MostrarConFormato SALARIO_BASE, m_Salario
    Else
        m_Salario = 0
        SALARIO_BASE.Text = PREFIJO
    End If
End Sub

Public Function ObtenerSalario() As Double
    ObtenerSalario = m_Salario
End Function

Public Sub CargarAntiguedad(valor As Variant)
    If IsNumeric(valor) Then
        m_Antiguedad = Round(CDbl(valor), 2)
        MostrarConFormato ANTIGUEDAD_SALARIAL, m_Antiguedad
    Else
        m_Antiguedad = 0
        ANTIGUEDAD_SALARIAL.Text = PREFIJO
    End If
End Sub

Public Function ObtenerAntiguedad() As Double
    ObtenerAntiguedad = m_Antiguedad
End Function
'===================================================================================================================================
'===================================================================================================================================


Private Sub FECHA_REGISTRO_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    If Not EsFechaValidaYFormatear(Me.FECHA_REGISTRO) Then
        Cancel = True   '  impide que el foco salga del control
        Exit Sub
    End If
    
 End Sub

Private Sub FECHA_REGISTRO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, /, backspace, delete, enter, tab

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub FECHA_REGISTRO_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 8 Then
       Borrador = True
    Else
       Borrador = False
    End If
    
End Sub

Private Sub FECHA_REGISTRO_Change()
        If Borrador = False Then
        
            If Len(Me.FECHA_REGISTRO.Value) > 10 Then
                
                Me.FECHA_REGISTRO.Value = Mid(Me.FECHA_REGISTRO.Value, 1, 10)
                MsgBox "Fecha de Registro Incorrecta"
            
            Else
                
                If Len(Me.FECHA_REGISTRO.Value) = 2 Then
                Me.FECHA_REGISTRO.Value = Me.FECHA_REGISTRO.Value & "/"
                End If
                
                If Len(Me.FECHA_REGISTRO.Value) = 5 Then
                Me.FECHA_REGISTRO.Value = Me.FECHA_REGISTRO.Value & "/"
                End If
                            
            End If
        End If
End Sub


Private Sub FECHA_BAJA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

' Permite: numeros, /, backspace, delete, enter, tab

    If Not (Chr(KeyAscii) Like "[0-9]" Or _
            KeyAscii = 8 Or KeyAscii = 127 Or _
            KeyAscii = 13 Or KeyAscii = 9) Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub FECHA_BAJA_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 8 Then
       Borrador = True
    Else
       Borrador = False
    End If
    
End Sub

Private Sub FECHA_BAJA_Change()
        If Borrador = False Then
        
            If Len(Me.FECHA_BAJA.Value) > 10 Then
                
                Me.FECHA_BAJA.Value = Mid(Me.FECHA_BAJA.Value, 1, 10)
                MsgBox "Fecha de Baja Incorrecta", vbExclamation, "Error"
            
            Else
                
                If Len(Me.FECHA_BAJA.Value) = 2 Then
                Me.FECHA_BAJA.Value = Me.FECHA_BAJA.Value & "/"
                End If
                
                If Len(Me.FECHA_BAJA.Value) = 5 Then
                Me.FECHA_BAJA.Value = Me.FECHA_BAJA.Value & "/"
                End If
                            
            End If
        End If
End Sub

Private Sub FECHA_BAJA_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 Dim CeldaFechaBaja As String
 
    CeldaFechaBaja = Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Address(False, False)
    
    
    
    If Not EsFechaValidaYFormatear(Me.FECHA_BAJA) Then
        Me.ANTIGUEDAD_LABORAL.Value = Range(CeldaFechaBaja).Offset(1, 2).Value
        Cancel = True   '  impide que el foco salga del control
        Exit Sub
    End If
 
 
                
        'Agregar las fechas para el calculo
        Hoja24.Range(CeldaFechaBaja).Offset(1, 0).NumberFormat = "dd/mm/yyyy"
        Hoja24.Range(CeldaFechaBaja).Offset(1, 0).Value = CDate(Me.FECHA_INGRESO.Value)
                
                
        If Me.FECHA_BAJA.Value = "" Then
        
            Me.ANTIGUEDAD_LABORAL.Value = Range(CeldaFechaBaja).Offset(1, 2).Value
            Exit Sub
        Else
            
            Range(CeldaFechaBaja).Offset(1, 1).NumberFormat = "dd/mm/yyyy"
            Range(CeldaFechaBaja).Offset(1, 1).Value = CDate(Me.FECHA_BAJA.Value)
            
            'Sustituir Antiguedad Laboral
            Me.ANTIGUEDAD_LABORAL.Value = Range(CeldaFechaBaja).Offset(1, 3).Value
        
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

Private Sub MostrarTodo()


'=================================
'Valores de Nomina
'=================================
Me.EMP.Left = 12
Me.EMP_ETIQUETA.Left = 12
Me.T�TULO.Visible = False
Me.NOMBRES.Visible = True
Me.NOMBRES_ETIQUETA.Visible = True
Me.EDAD.Visible = True
Me.EDAD_ETIQUETA.Visible = True
Me.FEMENINO.Visible = True
Me.SEXO.Visible = True
Me.SEXO_ETIQUETA.Visible = True
Me.MASCULINO.Visible = True
Me.CARGO.Visible = True
Me.CARGO_ETIQUETA.Visible = True
Me.CLASIFICACION_CARGO.Visible = True
Me.CLASIFICACION_CARGO_ETIQUETA.Visible = True
Me.CEDULA.Visible = False
Me.CEDULA_ETIQUETA.Visible = False
Me.UBICACION.Visible = True
Me.UBICACION_ETIQUETA.Visible = True
Me.UBICACION_GENERAL.Visible = True
Me.UBICACION_GENERAL_ETIQUETA.Visible = True
Me.UBICACION_ESPECIFICA.Visible = True
Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = True
Me.FOTO_TRABAJADOR.Visible = True
Me.SIN_FOTO.Visible = True
Me.ACTUALIZAR_FOTO.Visible = True
Me.AGREGAR_FOTO.Visible = True
Me.ELIMINAR_FOTO.Visible = True
Me.CARPETA_FOTO.Visible = True

'=================================
'Datos personales
'=================================

Me.ESTADO_CIVIL.Visible = True
Me.ESTADO_CIVIL_ETIQUETA.Visible = True
Me.DEPENDIENTES.Visible = True
Me.DEPENDIENTES_ETIQUETA.Visible = True
Me.SI.Visible = True
Me.NO.Visible = True
Me.DOMICILIO.Visible = True
Me.DOMICILIO_ETIQUETA.Visible = True
Me.URBANO.Visible = True
Me.RURAL.Visible = True
Me.NIVEL_ACADEMICO.Visible = True
Me.NIVEL_ACADEMICO_ETIQUETA.Visible = True
Me.ESTUDIO.Visible = True
Me.ESTUDIO_ETIQUETA.Visible = True

'=================================
'Datos de Baja
'=================================

Me.FECHA_INGRESO.Visible = True
Me.FECHA_INGRESO_ETIQUETA.Visible = True
Me.ANTIGUEDAD_LABORAL.Visible = True
Me.ANTIGUEDAD_LABORAL_ETIQUETA.Visible = True
Me.SALARIO_BASE.Visible = True
Me.SALARIO_BASE_ETIQUETA.Visible = True
Me.ANTIGUEDAD_SALARIAL.Visible = True
Me.ANTIGUEDAD_SALARIAL_ETIQUETA.Visible = True
Me.FECHA_BAJA.Visible = True
Me.FECHA_BAJA_ETIQUETA.Visible = True
Me.TIPO_BAJA.Visible = True
Me.TIPO_BAJA_ETIQUETA.Visible = True
Me.MOTIVO_BAJA.Visible = True
Me.MOTIVO_BAJA_ETIQUETA.Visible = True

'=================================
'Datos de Registro
'=================================

Me.RESPONSABLE.Visible = True
Me.RESPONSABLE_ETIQUETA.Visible = True
Me.FECHA_REGISTRO.Visible = True
Me.FECHA_REGISTRO_ETIQUETA.Visible = True
Me.REGISTRAR.Visible = False
Me.ANULAR.Visible = False
Me.RESPONSABLE_ANULACION_ETIQUETA.Visible = False
Me.RESPONSABLE_ANULACION.Visible = False


'==============================================
'Estado inicial de Salario y Antiguedad
'==============================================
    SALARIO_BASE.Text = PREFIJO
    m_Salario = 0
    m_SalarioEnFoco = False
    
    ANTIGUEDAD_SALARIAL.Text = PREFIJO
    m_Antiguedad = 0
    m_AntiguedadEnFoco = False
'==============================================

Me.height = 490
Me.width = 925
Me.ScrollBars = fmScrollBarsVertical

End Sub


Private Sub OcultarTodo()


'=================================
'Valores de Nomina
'=================================
Me.EMP.Left = 390
Me.EMP_ETIQUETA.Left = 390
Me.T�TULO.Visible = True
Me.NOMBRES.Visible = False
Me.NOMBRES_ETIQUETA.Visible = False
Me.EDAD.Visible = False
Me.EDAD_ETIQUETA.Visible = False
Me.FEMENINO.Visible = False
Me.SEXO.Visible = False
Me.SEXO_ETIQUETA.Visible = False
Me.MASCULINO.Visible = False
Me.CARGO.Visible = False
Me.CARGO_ETIQUETA.Visible = False
Me.CLASIFICACION_CARGO.Visible = False
Me.CLASIFICACION_CARGO_ETIQUETA.Visible = False
Me.CEDULA.Visible = False
Me.CEDULA_ETIQUETA.Visible = False
Me.UBICACION.Visible = False
Me.UBICACION_ETIQUETA.Visible = False
Me.UBICACION_GENERAL.Visible = False
Me.UBICACION_GENERAL_ETIQUETA.Visible = False
Me.UBICACION_ESPECIFICA.Visible = False
Me.UBICACION_ESPECIFICA_ETIQUETA.Visible = False
Me.FOTO_TRABAJADOR.Visible = False
Me.SIN_FOTO.Visible = False
Me.ACTUALIZAR_FOTO.Visible = False
Me.AGREGAR_FOTO.Visible = False
Me.ELIMINAR_FOTO.Visible = False
Me.CARPETA_FOTO.Visible = False

'=================================
'Datos personales
'=================================

Me.ESTADO_CIVIL.Visible = False
Me.ESTADO_CIVIL_ETIQUETA.Visible = False
Me.DEPENDIENTES.Visible = False
Me.DEPENDIENTES_ETIQUETA.Visible = False
Me.SI.Visible = False
Me.NO.Visible = False
Me.DOMICILIO.Visible = False
Me.DOMICILIO_ETIQUETA.Visible = False
Me.URBANO.Visible = False
Me.RURAL.Visible = False
Me.NIVEL_ACADEMICO.Visible = False
Me.NIVEL_ACADEMICO_ETIQUETA.Visible = False
Me.ESTUDIO.Visible = False
Me.ESTUDIO_ETIQUETA.Visible = False

'=================================
'Datos de Baja
'=================================

Me.FECHA_INGRESO.Visible = False
Me.FECHA_INGRESO_ETIQUETA.Visible = False
Me.ANTIGUEDAD_LABORAL.Visible = False
Me.ANTIGUEDAD_LABORAL_ETIQUETA.Visible = False
Me.SALARIO_BASE.Visible = False
Me.SALARIO_BASE_ETIQUETA.Visible = False
Me.ANTIGUEDAD_SALARIAL.Visible = False
Me.ANTIGUEDAD_SALARIAL_ETIQUETA.Visible = False
Me.FECHA_BAJA.Visible = False
Me.FECHA_BAJA_ETIQUETA.Visible = False
Me.TIPO_BAJA.Visible = False
Me.TIPO_BAJA_ETIQUETA.Visible = False
Me.MOTIVO_BAJA.Visible = False
Me.MOTIVO_BAJA_ETIQUETA.Visible = False

'=================================
'Datos de Registro
'=================================

Me.RESPONSABLE.Visible = False
Me.RESPONSABLE_ETIQUETA.Visible = False
Me.FECHA_REGISTRO.Visible = False
Me.FECHA_REGISTRO_ETIQUETA.Visible = False
Me.REGISTRAR.Visible = False
Me.ANULAR.Visible = False
Me.RESPONSABLE_ANULACION_ETIQUETA.Visible = False
Me.RESPONSABLE_ANULACION.Visible = False

'Vaciar las celdas de calculo de Antiguedad Laboral

Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 0).Value = ""
Hoja24.ListObjects("ANTIGUEDAD_LABORAL").HeaderRowRange.Find("FECHA DE INGRESO").Offset(1, 1).Value = ""

Me.height = 100
Me.width = 500
Me.ScrollBars = fmScrollBarsNone

End Sub


Private Sub LimpiarTodo()


'=================================
'Valores de Nomina
'=================================
Me.NOMBRES.Value = ""
Me.EDAD.Value = ""
Me.FEMENINO.Value = False
Me.MASCULINO.Value = False
Me.CARGO.Value = ""
Me.CLASIFICACION_CARGO.Value = ""
Me.CLASIFICACION_CARGO.Clear
Me.CEDULA.Value = ""
Me.UBICACION.Value = ""
Me.UBICACION.Clear
Me.UBICACION_GENERAL.Value = ""
Me.UBICACION_GENERAL.Clear
Me.UBICACION_ESPECIFICA.Value = ""
MostrarSinFoto

'=================================
'Datos personales
'=================================

Me.ESTADO_CIVIL.Value = ""
Me.ESTADO_CIVIL.Clear
Me.SI.Value = False
Me.SI.ForeColor = &H0&
Me.NO.Value = False
Me.NO.ForeColor = &H0&
Me.URBANO.Value = False
Me.URBANO.ForeColor = &H0&
Me.RURAL.Value = False
Me.RURAL.ForeColor = &H0&
Me.NIVEL_ACADEMICO.Value = ""
Me.NIVEL_ACADEMICO.Clear
Me.ESTUDIO.Value = ""

'=================================
'Datos de Baja
'=================================

Me.FECHA_INGRESO.Value = ""
Me.ANTIGUEDAD_LABORAL.Value = ""
Me.SALARIO_BASE.Value = ""
Me.ANTIGUEDAD_SALARIAL.Value = ""
Me.FECHA_BAJA.Value = ""
Me.TIPO_BAJA.Value = ""
Me.TIPO_BAJA.Clear
Me.MOTIVO_BAJA.Value = ""

'=================================
'Datos de Registro
'=================================

Me.RESPONSABLE.Value = ""
Me.RESPONSABLE.Clear
Me.RESPONSABLE_ANULACION.Value = ""
Me.RESPONSABLE_ANULACION.Clear
Me.FECHA_REGISTRO.Value = ""
Me.DATOS.Value = 0

End Sub

'=================================
'Validar datos en Listas de ComboBox
'=================================


Private Function EsValorValido(ByVal cbo As MSForms.ComboBox, _
                               ByVal Hoja As String, _
                               ByVal Tabla As String, _
                               Optional ByVal Columna As String = "") As Boolean
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim c As Range
    
    If Trim(cbo.Value) = "" Then
        EsValorValido = True
        Exit Function
    End If
    
    Set ws = ThisWorkbook.Worksheets(Hoja)
    Set tbl = ws.ListObjects(Tabla)
    
    If Columna = "" Then
        Set rng = tbl.ListColumns(1).DataBodyRange
    Else
        Set rng = tbl.ListColumns(Columna).DataBodyRange
    End If
    
    Set c = rng.Find(What:=cbo.Value, LookAt:=xlWhole, MatchCase:=False)
    
    If c Is Nothing Then
        
        MsgBox "El valor ingresado en " & cbo.Name & " no es valido." & vbCrLf & _
               "Seleccione una opcion de la lista.", _
               vbExclamation, "Valor incorrecto"
        
        cbo.Value = ""
        EsValorValido = False
    Else
        EsValorValido = True
    End If
    
End Function

Private Sub CLASIFICACION_CARGO_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If Not EsValorValido(Me.CLASIFICACION_CARGO, "AUX", "CLASIFICACION_CARGO") Then
        Cancel = True
    End If

End Sub

Private Sub UBICACION_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If Not EsValorValido(Me.UBICACION, "AUX", "UBICACION") Then
        Cancel = True
    End If

End Sub

Private Sub UBICACION_GENERAL_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If Not EsValorValido(Me.UBICACION_GENERAL, _
                         "AUX", _
                         "DEPENDENCIAS", _
                         "DEPENDENCIA CORTA") Then
        Cancel = True
    End If

End Sub

Private Sub ESTADO_CIVIL_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If Not EsValorValido(Me.ESTADO_CIVIL, "AUX", "ESTADO_CIVIL") Then
        Cancel = True
    End If

End Sub

Private Sub NIVEL_ACADEMICO_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If Not EsValorValido(Me.NIVEL_ACADEMICO, "AUX", "NIVEL_ACADEMICO") Then
        Cancel = True
    End If

End Sub


Private Sub TIPO_BAJA_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    If Not EsValorValido(Me.TIPO_BAJA, "AUX", "TIPO_BAJA") Then
        Cancel = True
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
        Cancel = True
        Exit Sub
    End If
    
    'Buscar responsable en columna 1 (Col A = UbicaciOn)
    Set c = tbl.ListColumns(1).DataBodyRange.Find( _
                What:=Resp, _
                LookAt:=xlWhole, _
                MatchCase:=False)
        Me.REGISTRAR.Visible = True
        Me.REGISTRAR.Enabled = True
        
    'Si NO existe
    If c Is Nothing Then
        
        MsgBox "Responsable no Existe", vbExclamation
        Me.REGISTRAR.Visible = False
        Me.REGISTRAR.Enabled = False
        Me.RESPONSABLE.Value = ""
        Me.RESPONSABLE.SetFocus
        Exit Sub
        
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

Private Sub REGISTRAR_Click()

    Dim MOTIVO As String
    Dim camposFaltantes As String
    Dim totalCampos As Integer
    Dim msgTitulo As String
    
    
camposFaltantes = ValidarBajas()
    
    ' Si hay campos faltantes, mostrar mensaje con formato
    If camposFaltantes <> "" Then
        ' Contar lineas para el mensaje
        totalCampos = UBound(Split(camposFaltantes, vbCrLf))
        
        MsgBox "CAMPOS OBLIGATORIOS FALTANTES: " & totalCampos & vbCrLf & vbCrLf & _
               camposFaltantes & vbCrLf & _
               "--------------------------------------" & vbCrLf & _
               "Complete todos los campos mencionados", _
               vbExclamation, "ValidaciOn de datos"
        
        ' Opcional: Enfocar el primer campo faltante
        EnfocarPrimerCampoFaltante
    ElseIf Me.ANTIGUEDAD_LABORAL.Value = "ERROR FECHA BAJA" Then
            MsgBox "Error en la Fecha de Baja", vbExclamation, "Error en Fecha"
            Me.FECHA_BAJA.SetFocus
            Exit Sub
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
Private Function ValidarBajas() As String
    Dim faltantes As String
    Dim listaCampos As String
    
    ' Verificar campos obligatorios
    If Me.CLASIFICACION_CARGO.ListIndex = -1 Then
        faltantes = faltantes & "- CLASIFICACION CARGO" & vbCrLf
    End If
    
    If Me.ESTADO_CIVIL.ListIndex = -1 Then
        faltantes = faltantes & "- ESTADO CIVIL" & vbCrLf
    End If
        ' Validar al menos una opciOn SI/NO
    If Me.SI.Value = False And Me.NO.Value = False Then
        faltantes = faltantes & "- DEPENDIENTES: SI o NO (debe seleccionar uno)" & vbCrLf
    End If
    
    ' Validar al menos una opciOn URBANO/RURAL
    If Me.URBANO.Value = False And Me.RURAL.Value = False Then
        faltantes = faltantes & "- URBANO o RURAL (debe seleccionar uno)" & vbCrLf
    End If
    
    If Me.NIVEL_ACADEMICO.ListIndex = -1 Then
        faltantes = faltantes & "- NIVEL ACADEMICO" & vbCrLf
    End If
    
    If Trim(Me.ESTUDIO.Value) = "" Then
        faltantes = faltantes & "- ESTUDIO" & vbCrLf
    End If
    
    If Me.SALARIO_BASE.Value = "C$ " Then
        faltantes = faltantes & "- SALARIO BASE" & vbCrLf
    End If
    
    If Me.ANTIGUEDAD_SALARIAL.Value = "C$ " Then
        faltantes = faltantes & "- ANTIGUEDAD SALARIAL" & vbCrLf
    End If
    
    If Trim(Me.FECHA_BAJA.Value) = "" Then
        faltantes = faltantes & "- FECHA BAJA" & vbCrLf
    End If
    
    If Me.TIPO_BAJA.ListIndex = -1 Then
        faltantes = faltantes & "- TIPO BAJA" & vbCrLf
    End If
    
    If Trim(Me.MOTIVO_BAJA.Value) = "" Then
        faltantes = faltantes & "- MOTIVO BAJA" & vbCrLf
    End If
    
    If Trim(Me.FECHA_REGISTRO.Value) = "" Then
        faltantes = faltantes & "- FECHA REGISTRO" & vbCrLf
    End If
    
    ValidarBajas = faltantes
    
End Function

' ============================================
' FUNCION PARA ENFOCAR EL PRIMER CAMPO FALTANTE
' ============================================
Private Sub EnfocarPrimerCampoFaltante()
    ' Esta funciOn puede enfocar el primer campo que falta
    ' segun el motivo, para facilitar la correccion
        
    If Me.CLASIFICACION_CARGO.ListIndex = -1 Then
        Me.CLASIFICACION_CARGO.SetFocus
    ElseIf Me.ESTADO_CIVIL.ListIndex = -1 Then
        Me.DATOS.Value = 0
        Me.ESTADO_CIVIL.SetFocus
    ElseIf Me.SI.Value = False And Me.NO.Value = False Then
        Me.DATOS.Value = 0
        Me.SEXO.SetFocus
    ElseIf Me.URBANO.Value = False And Me.RURAL.Value = False Then
        Me.DATOS.Value = 0
        Me.DOMICILIO.SetFocus
    ElseIf Me.NIVEL_ACADEMICO.ListIndex = -1 Then
        Me.DATOS.Value = 0
        Me.NIVEL_ACADEMICO.SetFocus
    ElseIf Trim(Me.ESTUDIO.Value) = "" Then
        Me.DATOS.Value = 0
        Me.ESTUDIO.SetFocus
    ElseIf Trim(Me.SALARIO_BASE.Value) = "" Then
        Me.DATOS.Value = 1
        Me.SALARIO_BASE.SetFocus
    ElseIf Trim(Me.ANTIGUEDAD_SALARIAL.Value) = "" Then
        Me.DATOS.Value = 1
        Me.ANTIGUEDAD_SALARIAL.SetFocus
    ElseIf Trim(Me.FECHA_BAJA.Value) = "" Then
        Me.DATOS.Value = 1
        Me.FECHA_BAJA.SetFocus
    ElseIf Me.TIPO_BAJA.ListIndex = -1 Then
        Me.DATOS.Value = 1
        Me.TIPO_BAJA.SetFocus
    ElseIf Trim(Me.MOTIVO_BAJA.Value) = "" Then
        Me.DATOS.Value = 1
        Me.MOTIVO_BAJA.SetFocus
    ElseIf Trim(Me.FECHA_REGISTRO.Value) = "" Then
        Me.FECHA_REGISTRO.SetFocus
    End If
End Sub

Private Sub GuardarRegistro()

Dim CeldaBaja, CodEmpleado As String

Hoja3.Select

CeldaBaja = Hoja3.ListObjects("BAJAS").HeaderRowRange.Find("No. EMP").Address(False, False)
CodEmpleado = Me.EMP.Value
    
    If Range(CeldaBaja).Offset(1, 0).Value = "" Then
    
        'Registrar cuando no hay ningun registro
        Range(CeldaBaja).Offset(1, 0).Value = CodEmpleado
               
    Else
      
        'Registrar cuando ya hay uno o mas registros en la tabla
    
        Range(CeldaBaja).End(xlDown).Offset(1, 0).Value = CodEmpleado
        
    End If
    
    'Registrar datos del Traslado y Nombramiento
    
    Range(CeldaBaja).End(xlDown).Offset(0, 1).Value = Me.NOMBRES.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 2).Value = Me.CEDULA.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 3).Value = Me.FECHA_BAJA.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 4).Value = Me.FECHA_REGISTRO.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 5).Value = ObtenerSalario
    Range(CeldaBaja).End(xlDown).Offset(0, 6).Value = ObtenerAntiguedad
    Range(CeldaBaja).End(xlDown).Offset(0, 7).Value = Me.UBICACION.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 8).Value = Me.UBICACION_GENERAL.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 9).Value = Me.UBICACION_ESPECIFICA.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 10).Value = Me.CARGO.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 11).Value = Me.FECHA_INGRESO.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 12).Value = Me.ANTIGUEDAD_LABORAL.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 13).Value = Me.NIVEL_ACADEMICO.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 14).Value = Me.ESTUDIO.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 15).Value = Me.ESTADO_CIVIL.Value
    
        If Me.SI.Value = True Then
            Range(CeldaBaja).End(xlDown).Offset(0, 16).Value = "SI"
        ElseIf Me.NO.Value = True Then
            Range(CeldaBaja).End(xlDown).Offset(0, 16).Value = "NO"
        End If
        
        If Me.FEMENINO.Value = True Then
            Range(CeldaBaja).End(xlDown).Offset(0, 17).Value = "F"
        ElseIf Me.MASCULINO.Value = True Then
            Range(CeldaBaja).End(xlDown).Offset(0, 17).Value = "M"
        End If
    Range(CeldaBaja).End(xlDown).Offset(0, 18).Value = Me.EDAD.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 20).Value = Me.CLASIFICACION_CARGO.Value
        If Me.URBANO.Value = True Then
            Range(CeldaBaja).End(xlDown).Offset(0, 21).Value = "URBANO"
        ElseIf Me.RURAL.Value = True Then
            Range(CeldaBaja).End(xlDown).Offset(0, 21).Value = "RURAL"
        End If
    Range(CeldaBaja).End(xlDown).Offset(0, 22).Value = Me.TIPO_BAJA.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 23).Value = Me.MOTIVO_BAJA.Value
    Range(CeldaBaja).End(xlDown).Offset(0, 24).Value = Me.RESPONSABLE.Value



    'Actualizar los datos del trabajador en NOMINA
    
    Dim Ultimo, f, empVal As Integer
                        
    empVal = Me.EMP.Value
    
    Ultimo = Hoja25.Range("A" & rows.Count).End(xlUp).row
    
    For f = 2 To Ultimo
        If empVal = Hoja25.Cells(f, "A").Value Then
            
            Hoja25.Cells(f, "Q").Value = "INACTIVO"
            Hoja25.Cells(f, "R").Value = Me.FECHA_BAJA.Value
            
        End If
    Next f
    

Call LimpiarTodo
ActiveWorkbook.Save

MsgBox "La BAJA del Empleado " & Me.NOMBRES.Value & " fue guardado con exito." & vbCrLf & _
       "Y Datos del Trabajador fueron editados en NOMINA", _
       vbInformation, "exito"

Call UserForm_Initialize
Call UserForm_Activate

End Sub

Private Sub ANULAR_Click()
Dim nombreTrabajador As String
Dim respuesta As VbMsgBoxResult


    nombreTrabajador = Me.NOMBRES.Value

    If Me.RESPONSABLE_ANULACION.Value = "" Then
    MsgBox "Selecciona un Responsable para la Anulacion", , "Datos Incompletos"
    Me.RESPONSABLE_ANULACION.SetFocus
    Exit Sub
    End If
    
    ' Muestra el mensaje de alerta con botones Si y No
    respuesta = MsgBox("Esta por Anular el registro de baja de: " & vbCrLf & vbCrLf & nombreTrabajador & vbCrLf & vbCrLf & "¿Desea eliminar este registro?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmar Accion")
    
    ' Evalua la respuesta
    If respuesta = vbYes Then
        ' Accion si el usuario presiona Si
        
            '=====================================================
            'Encuentra la fila del COdigo de Empleado
                    
            Dim tbl As ListObject
            Dim c As Range
            Dim empVal As String
        
            'Referencia a la tabla
            Set tbl = Hoja3.ListObjects("BAJAS")
            
            empVal = Trim(Me.EMP.Value)
            
            'Buscar empleado en columna 5 (Col G = No. EMP)
            Set c = tbl.ListColumns(5).DataBodyRange.Find( _
                        What:=empVal, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
            '=====================================================
        Hoja3.Select
        
        c.Offset(0, 1).Value = ""
        c.Offset(0, 2).Value = ""
        c.Offset(0, 3).Value = ""
        c.Offset(0, 4).Value = ""
        c.Offset(0, 5).Value = ""
        c.Offset(0, 6).Value = ""
        c.Offset(0, 7).Value = ""
        c.Offset(0, 8).Value = ""
        c.Offset(0, 9).Value = ""
        c.Offset(0, 10).Value = ""
        c.Offset(0, 11).Value = ""
        c.Offset(0, 12).Value = ""
        c.Offset(0, 13).Value = ""
        c.Offset(0, 14).Value = ""
        c.Offset(0, 15).Value = ""
        c.Offset(0, 16).Value = ""
        c.Offset(0, 17).Value = ""
        c.Offset(0, 18).Value = ""
        c.Offset(0, 20).Value = ""
        c.Offset(0, 21).Value = ""
        c.Offset(0, 22).Value = ""
        c.Offset(0, 23).Value = ""
        c.Offset(0, 24).Value = Me.RESPONSABLE_ANULACION.Value
        
        
        '=====================================================
            
        'Actualizar los datos del trabajador en NOMINA

                    Dim Ultimo, f As Integer
                                        
                    Ultimo = Hoja25.Range("A" & rows.Count).End(xlUp).row
                    
                    For f = 2 To Ultimo
                        If empVal = Hoja25.Cells(f, "A").Value Then
                            
                            Hoja25.Cells(f, "Q").Value = "ACTIVO"
                            Hoja25.Cells(f, "R").Value = ""
                            
                        End If
                    Next f
                    
        '=====================================================
        c.Offset(0, 1).Value = nombreTrabajador & " (" & empVal & ")"
        c.Offset(0, 0).Value = "BAJA ANULADA"
        
        'MENSAJE DE BAJA ANULADA
        MsgBox "Se ha Anulado Correctamente la Baja de:" & vbCrLf & vbCrLf & nombreTrabajador, vbInformation, "Baja Anulada"
        Me.EMP.Value = ""
        Call EMP_AfterUpdate
        
    Else
        ' AcciOn si el usuario presiona No
        MsgBox "Operacion Cancelada.", vbInformation, "Anulacion Cancelada"
        Exit Sub
        
    End If

End Sub

