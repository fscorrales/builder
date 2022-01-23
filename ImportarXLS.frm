VERSION 5.00
Begin VB.Form ImportarXLS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar Registros"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImportar 
      BackColor       =   &H008080FF&
      Caption         =   "Importar Datos"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tabla Destino"
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
      TabIndex        =   8
      Top             =   1560
      Width           =   6975
      Begin VB.ComboBox cmbTabla 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Tabla"
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
   Begin VB.Frame Frame2 
      Caption         =   "Origen de Datos"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtCeldaFin 
         Height          =   285
         Left            =   4800
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtCeldaInicio 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   255
         Left            =   6480
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Celda Fin"
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
         Left            =   3360
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Celda Inicio"
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
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Ruta de Acceso"
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
Attribute VB_Name = "ImportarXLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()

    DefinirRutaAcceso

End Sub

Private Sub cmdImportar_Click()
    
    ImportarRegistrosXLS
    
End Sub

Private Sub Form_Load()

    Call CenterMe(ImportarXLS, 7350, 3450)

End Sub
