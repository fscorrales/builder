VERSION 5.00
Begin VB.Form CargaIndice 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Indice"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Indice"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Campos con Indice Principal"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6615
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar Campo"
         Height          =   375
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4440
         Width           =   2175
      End
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H008080FF&
         Caption         =   "Agregar Campo"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4440
         Width           =   2175
      End
      Begin VB.ListBox lstCampoDisponible 
         Height          =   3570
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3000
      End
      Begin VB.ListBox lstCampoIndice 
         Height          =   3570
         Left            =   3480
         TabIndex        =   4
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Campos a Indexar"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Campos Disponibles"
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
         TabIndex        =   6
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tabla a Indexar"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cmbNombreTabla 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   4935
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
Attribute VB_Name = "CargaIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    
    Call PasarDatosListBox(lstCampoDisponible, lstCampoIndice)

End Sub

Private Sub cmdEliminar_Click()

    Call PasarDatosListBox(lstCampoIndice, lstCampoDisponible)

End Sub

Private Sub cmdGuardar_Click()

    If Not lstCampoDisponible.ListCount = "0" Then
        GenerarIndice
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaIndice, 6950, 7050)

End Sub

Private Sub cmbNombreTabla_Click()
    
    VaciarListBox
    If Not cmbNombreTabla.Text = "" Then
        Call CargarlstCampoDisponible(cmbNombreTabla.Text)
    End If
    
End Sub

Private Sub cmbNombreTabla_Change()

    VaciarListBox

End Sub
