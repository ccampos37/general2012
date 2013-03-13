VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormGuiRemDev 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Salir"
      Height          =   855
      Left            =   6720
      Picture         =   "FormGuiRemDev.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Eliminar"
      Height          =   855
      Left            =   3600
      Picture         =   "FormGuiRemDev.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Grabar"
      Height          =   855
      Left            =   5160
      Picture         =   "FormGuiRemDev.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Modificar"
      Height          =   855
      Left            =   2040
      Picture         =   "FormGuiRemDev.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Adicionar"
      Height          =   855
      Left            =   480
      Picture         =   "FormGuiRemDev.frx":0CD8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4680
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   6240
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   6240
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   6240
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   6240
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   6240
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Almacen Ref"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Forma  Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Orden Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Razon Social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Doc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Transaccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Doc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. Doc. Ref"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "R.U.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   3
      FormatString    =   "  Codigo "
   End
End
Attribute VB_Name = "FormGuiRemDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
