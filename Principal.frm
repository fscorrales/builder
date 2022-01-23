VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "BUILDER"
   ClientHeight    =   3360
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5475
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgMultifuncion 
      Left            =   4920
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuNuevaDB 
         Caption         =   "&Nueva DB"
      End
      Begin VB.Menu Line01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuConectar 
         Caption         =   "&Conectar"
      End
      Begin VB.Menu MnuDesconectar 
         Caption         =   "&Desconectar"
      End
      Begin VB.Menu Line02 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu MnuTablas 
      Caption         =   "&Tablas"
      Begin VB.Menu MnuNuevaTabla 
         Caption         =   "&Nueva"
      End
      Begin VB.Menu MnuListadoTablas 
         Caption         =   "&Listado"
      End
   End
   Begin VB.Menu MnuCampos 
      Caption         =   "&Campos"
      Begin VB.Menu MnuNuevoCampo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu MnuListadoCampos 
         Caption         =   "&Listado"
      End
      Begin VB.Menu Line03 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIndice 
         Caption         =   "&Indice"
      End
   End
   Begin VB.Menu MnuRelaciones 
      Caption         =   "&Relaciones"
      Begin VB.Menu MnuNuevaRelacion 
         Caption         =   "&Nueva"
      End
      Begin VB.Menu MnuListadoRelaciones 
         Caption         =   "&Listado"
      End
   End
   Begin VB.Menu MnuRegistros 
      Caption         =   "Re&gistros"
      Begin VB.Menu MnuListadoRegistros 
         Caption         =   "&Listado"
      End
      Begin VB.Menu MnuImportarXLS 
         Caption         =   "&Importar XLS"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

    Conectado (False)

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Conectado (False)

End Sub

Private Sub MnuConectar_Click()

    Conectar ("")
    
End Sub

Private Sub MnuDesconectar_Click()

    Conectado (False)

End Sub

Private Sub MnuImportarXLS_Click()

    ImportarXLS.Show
    CargarcmbTablaImportar

End Sub

Private Sub MnuIndice_Click()

    CargaIndice.Show
    Call CargarcmbNombreTabla(CargaIndice)
    VaciarListBox

End Sub

Private Sub MnuListadoCampos_Click()

    ListadoCampos.Show
    ConfigurardgCampos
    Call CargarcmbNombreTabla(ListadoCampos)

End Sub

Private Sub MnuListadoRegistros_Click()

    ListadoRegistros.Show
    Call CargarcmbNombreTabla(ListadoRegistros)

End Sub

Private Sub MnuListadoRelaciones_Click()

    ListadoRelaciones.Show
    ConfigurardgRelaciones
    CargardgRelaciones

End Sub

Private Sub MnuListadoTablas_Click()

    ListadoTablas.Show
    ConfigurardgTablas
    CargardgTablas
    
End Sub

Private Sub MnuNuevaDB_Click()
    
    NuevaBD
    
End Sub

Private Sub MnuNuevaRelacion_Click()

    CargaRelacion.Show
    CargarcmbTablaOrigenyDestino

End Sub

Private Sub MnuNuevaTabla_Click()

    CargaTabla.Show

End Sub

Private Sub MnuNuevoCampo_Click()

    CargaCampo.Show
    Call CargarcmbNombreTabla(CargaCampo)
    CargarcmbTipoCampo
    
End Sub

Private Sub MnuSalir_Click()
    Set rstCargadgRegistros = Nothing
    End

End Sub

