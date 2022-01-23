VERSION 5.00
Begin VB.Form CargaRelacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Relaciones"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Atributos de la Relación"
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
      TabIndex        =   12
      Top             =   3240
      Width           =   6615
      Begin VB.TextBox txtNombreRelacion 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label3 
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
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Campo"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Campos a Relacionar"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   6615
      Begin VB.ComboBox cmbCampoOrigen 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   4815
      End
      Begin VB.ComboBox cmbCampoDestino 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Campo Destino"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Campo Origen"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tablas a Relacionar"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.ComboBox cmbTablaDestino 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   4815
      End
      Begin VB.ComboBox cmbTablaOrigen 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Tabla Destino"
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
      Begin VB.Label Label2 
         Caption         =   "Tabla Origen"
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
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "CargaRelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    
    If ValidaRelacion = True Then
        GenerarRelacion
    End If
    
End Sub

Private Sub Form_Load()

    Call CenterMe(CargaRelacion, 6950, 5150)

End Sub

Private Sub cmbTablaOrigen_Change()

    cmbCampoOrigen.Clear

End Sub

Private Sub cmbTablaOrigen_Click()
    
    If Not cmbTablaOrigen.Text = "" Then
        cmbCampoOrigen.Clear
        Call CargarcmbCampoOrigen(cmbTablaOrigen.Text)
    End If
    
End Sub

Private Sub cmbTablaDestino_Change()

    cmbCampoDestino.Clear

End Sub

Private Sub cmbTablaDestino_Click()
    
    If Not cmbTablaDestino.Text = "" Then
        cmbCampoDestino.Clear
        Call CargarcmbCampoDestino(cmbTablaDestino.Text)
    End If
    
End Sub
