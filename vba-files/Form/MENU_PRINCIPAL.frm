VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MENU_PRINCIPAL 
   Caption         =   "MENU PRINCIPAL"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "MENU_PRINCIPAL.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "MENU_PRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Activate()
Me.Top = 0
Me.Left = 0
Me.height = 330

End Sub

Private Sub UserForm_Initialize()

Inicio

End Sub

Sub Inicio()
    
      Me.ALTA.Visible = True
      Me.ALTA.Top = 50
      Me.BAJA.Visible = True
      Me.BAJA.Top = 100
      Me.REASIGNACION.Visible = True
      Me.REASIGNACION.Top = 150
      Me.CARNETS.Visible = True
      Me.CARNETS.Top = 200
      Me.PASANTIAS.Visible = True
      Me.PASANTIAS.Top = 250
      
      Me.REGISTRO_ALTAS.Visible = False
      Me.DOCUMENTACION.Visible = False
      Me.ATRAS.Visible = False
    
End Sub

Sub OcultarpRrincipales()
    
    Me.ALTA.Visible = False
    Me.BAJA.Visible = False
    Me.REASIGNACION.Visible = False
    Me.CARNETS.Visible = False
    Me.PASANTIAS.Visible = False
    

End Sub

Sub BAJA_CLICK()
      'Mostrar el formulario de BAJAS
      Unload Me
      BAJAS.Show
      Load BAJAS

End Sub

Sub ALTA_CLICK()
      'Mostrar el formulario de ALTAS
      OcultarpRrincipales
      
      Me.REGISTRO_ALTAS.Visible = True
      Me.REGISTRO_ALTAS.Top = 70
      Me.DOCUMENTACION.Visible = True
      Me.DOCUMENTACION.Top = 120
      Me.ATRAS.Visible = True
      Me.ATRAS.Top = 170
      Me.ATRAS.Left = 100
      Me.height = 250

End Sub

Sub REGISTRO_ALTAS_CLICK()
      'Mostrar el formulario de REGISTRO ALTAS
      Unload Me
      ALTAS.Show
      Load ALTAS
      
End Sub

Sub DOCUMENTACION_CLICK()
      'Mostrar el formulario de DOCUMENTACION
      Unload Me
      DOCUMENTOS.Show
      Load DOCUMENTOS
      
End Sub

Sub REASIGNACION_CLICK()
      'Mostrar el formulario de REASIGNACION
      Unload Me
      TRAS_NOM.Show
      Load TRAS_NOM

End Sub

Sub CARNETS_CLICK()
      'Mostrar el formulario de CARNETS
      Unload Me
      CONTROL_CARNETS.Show
      Load CONTROL_CARNETS

End Sub

Sub PASANTIAS_CLICK()
      'Mostrar el formulario de PASANTIAS
      Unload Me
      PASANTIA.Show
      Load PASANTIA

End Sub

Private Sub ATRAS_Click()

UserForm_Initialize
UserForm_Activate

End Sub
