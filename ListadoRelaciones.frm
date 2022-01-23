VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoRelaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relaciones"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
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
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   9015
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Relación"
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Relación"
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Relaciones"
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
      Top             =   0
      Width           =   9060
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgRelaciones 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8800
         _ExtentX        =   15531
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoRelaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    CargaRelacion.Show
    CargarcmbTablaOrigenyDestino
    Unload ListadoRelaciones

End Sub

Private Sub cmdEliminar_Click()

    EliminarRelacion

End Sub

Private Sub Form_Load()

    Call CenterMe(ListadoRelaciones, 9200, 6300)
    
End Sub
