Attribute VB_Name = "Modulo3"
Sub MOVIMIENTO_Haga_clic_en()

Load TRAS_NOM
TRAS_NOM.Show

End Sub

Sub FICHA_Haga_clic_en()

Load DOCUMENTOS

DOCUMENTOS.Show

End Sub

Sub IMAGEN_1_Haga_clic_en()

Load CONTROL_CARNETS
CONTROL_CARNETS.Show

End Sub

Sub Baja_Llamar_Haga_clic_en()

Load BAJAS
BAJAS.Show

End Sub

Sub MENU_FORMULARIOS()
Attribute MENU_FORMULARIOS.VB_ProcData.VB_Invoke_Func = "f\n14"

' MENU_FORMULARIOS Macro
' Acceso directo: CTRL+f

    Load MENU_PRINCIPAL
    MENU_PRINCIPAL.Show
    
End Sub
