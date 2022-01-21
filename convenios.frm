VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form convenios 
   Caption         =   "Convenios"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   8505
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Buscar socio por Nº Cobro"
      Height          =   495
      Left            =   1560
      TabIndex        =   23
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buscar socio por Nombre"
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "C O N F I R M A"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   2535
   End
   Begin MSMask.MaskEdBox MaskEdBox5 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Elimina"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modifica"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ingresa"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.PictureBox DBGrid1 
      Height          =   6615
      Left            =   2880
      ScaleHeight     =   6555
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   1440
      Width           =   7335
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox conv_nrocom 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   720
      Width           =   5415
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      Caption         =   "Valor del convenio"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Nombres"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Apellidos"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Nº de Socio"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Cierre"
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "RAZON SOCIAL"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Nº del Comercio"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "convenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
Unload Me
End Sub

