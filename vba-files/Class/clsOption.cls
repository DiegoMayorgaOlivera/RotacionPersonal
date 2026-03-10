VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Opt As MSForms.OptionButton
Attribute Opt.VB_VarHelpID = -1

Private Sub Opt_Click()

    Dim ctrl As Control
    
    'Todos en negro
    For Each ctrl In Opt.Parent.Controls
        If TypeOf ctrl Is MSForms.OptionButton Then
            ctrl.ForeColor = vbBlack
        End If
    Next ctrl

    'Marcado en azul
    Opt.ForeColor = vbBlue

    'Volver a negro
    For Each ctrl In Opt.Parent.Controls
        If TypeOf ctrl Is MSForms.Label Then
            ctrl.ForeColor = vbBlack
        End If
    Next ctrl

DOCUMENTOS.ActualizarBarraProgreso

End Sub


