VERSION 5.00
Begin VB.Form Registro 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call RegisterComponent(App.Path & "\Componentes\" & "dao360.dll", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "dao360.dll", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSADODC.OCX", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSADODC.OCX", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "Comdlg32.ocx", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "Comdlg32.ocx", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSHFLXGD.OCX", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSHFLXGD.OCX", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSDE.DLL", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSDE.DLL", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSDBRPT.DLL", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSDBRPT.DLL", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSHTMPGD.DLL", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "MSHTMPGD.DLL", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "WBCLSDSR.OCX", True, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "WBCLSDSR.OCX", False, True)
    Call RegisterComponent(App.Path & "\Componentes\" & "msado15.dll", True, True)
    Call RegisterComponent("C:\Program Files\Common Files\System\ado\msado15.dll", False, True)
    End
End Sub
