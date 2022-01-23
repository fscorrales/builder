VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoCampos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Campos"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Acciones Posibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   6000
      Width           =   7650
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Campo"
         Height          =   375
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H008080FF&
         Caption         =   "Editar Campo"
         Height          =   375
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Campo"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Tabla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7650
      Begin VB.ComboBox cmbNombreTabla 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Campos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   7650
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCampos 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7400
         _ExtentX        =   13044
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbNombreTabla_Change()

    ConfigurardgCampos

End Sub

Private Sub cmbNombreTabla_Click()
    
    ConfigurardgCampos
    If Not cmbNombreTabla.Text = "" Then
        Call CargardgCampos(cmbNombreTabla.Text)
    End If
    
End Sub

Private Sub cmdAgregar_Click()

    CargaCampo.Show
    Call CargarcmbNombreTabla(CargaCampo)
    CargarcmbTipoCampo

End Sub

Private Sub cmdEditar_Click()

    EditarCampo

End Sub

Private Sub cmdEliminar_Click()

    EliminarCampo

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoCampos, 7750, 7350)

End Sub
