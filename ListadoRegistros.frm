VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoRegistros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registros"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Registros"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   7650
      Begin VB.TextBox txtEdicion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgRegistros 
         Height          =   4455
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   7400
         _ExtentX        =   13044
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7650
      Begin VB.ComboBox cmbOrdenCampo 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   3615
      End
      Begin VB.CommandButton cmdActualizar 
         BackColor       =   &H008080FF&
         Caption         =   "Actualizar"
         Height          =   855
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox cmbNombreTabla 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Ordenar por"
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
         TabIndex        =   7
         Top             =   840
         Width           =   1455
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
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ListadoRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Variable para la clase
Dim EditarGrilla As CEditarFlexGrid

Private Sub cmdActualizar_Click()

    If Not cmbNombreTabla.Text = "" Or cmbOrdenCampo.Text = "" Then
        ConfigurardgRegistros
        CargardgRegistros
    End If
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoRegistros, 7750, 6900)

End Sub

Private Sub cmbNombreTabla_Change()

    cmbOrdenCampo.Clear

End Sub

Private Sub cmbNombreTabla_Click()
    
    If Not cmbNombreTabla.Text = "" Then
        cmbOrdenCampo.Clear
        Call CargarcmbOrdenCampo(cmbNombreTabla.Text)
        'Nueva instancia
        Set EditarGrilla = New CEditarFlexGrid
        'Inicia ( se le envia el Flex y el text )
        EditarGrilla.Iniciar dgRegistros, txtEdicion
    End If
    
End Sub
