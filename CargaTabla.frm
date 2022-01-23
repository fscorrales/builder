VERSION 5.00
Begin VB.Form CargaTabla 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tablas"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1566.102
   ScaleMode       =   0  'User
   ScaleWidth      =   4620.148
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Nueva Tabla"
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
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtTabla 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "Guardar Datos"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "CargaTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()

    If strEditandoTabla = "" Then
        GenerarTabla (txtTabla.Text)
    Else
        EditarTabla
    End If
    
End Sub

Private Sub Form_Load()

    Call CenterMe(CargaTabla, 4650, 2050)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    strEditandoTabla = ""

End Sub
