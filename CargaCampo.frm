VERSION 5.00
Begin VB.Form CargaCampo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Campos"
   ClientHeight    =   3840
   ClientLeft      =   -165
   ClientTop       =   -105
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Campo"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Atributos del Campo"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   6615
      Begin VB.CheckBox chkLongitud 
         Alignment       =   1  'Right Justify
         Caption         =   "Permitir Long. Cero"
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
         Height          =   315
         Left            =   3480
         TabIndex        =   5
         Top             =   1560
         Width           =   2350
      End
      Begin VB.CheckBox chkRequerido 
         Alignment       =   1  'Right Justify
         Caption         =   "Requerido"
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
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2350
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtNombreCampo 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.TextBox txtTamaño 
         Height          =   285
         Left            =   4920
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo"
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
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Tamaño"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tabla a la que pertence el Campo"
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
      TabIndex        =   7
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cmbNombreTabla 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
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
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()

    If strEditandoCampo = "" Then
        GenerarCampo
    Else
        EditarCampo
    End If

End Sub

Private Sub Form_Load()

    Call CenterMe(CargaCampo, 6950, 4200)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoCampo = ""

End Sub
