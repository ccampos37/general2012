VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frPrmQC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel de 5ta. Categoria"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frPrmQC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "2do. Tope para Impuestos de 4ta. Categoria"
      Height          =   915
      Left            =   135
      TabIndex        =   61
      Top             =   1875
      Width           =   6030
      Begin AplisetControlText.Aplitext xPorcentaje2 
         Height          =   285
         Left            =   4935
         TabIndex        =   8
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   2
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xPorcentaje3 
         Height          =   285
         Left            =   4935
         TabIndex        =   10
         Top             =   540
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   2
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext XUIT2 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   255
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext XUIT3 
         Height          =   285
         Left            =   3630
         TabIndex        =   7
         Top             =   255
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext XUIT4 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Top             =   570
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de deducción (Exceso de           UIT )"
         Height          =   195
         Left            =   105
         TabIndex        =   63
         Top             =   600
         Width           =   3075
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de deducción (Exceso de           UIT hasta            UIT)"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   285
         Width           =   4275
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "1er. Tope para Impuestos de 5ta. Categoria"
      Height          =   585
      Left            =   135
      TabIndex        =   59
      Top             =   1245
      Width           =   6015
      Begin AplisetControlText.Aplitext xPorcentaje 
         Height          =   285
         Left            =   4935
         TabIndex        =   5
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   2
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext XUIT1 
         Height          =   285
         Left            =   2130
         TabIndex        =   4
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tasa de deducción (Hasta              UIT)"
         Height          =   195
         Left            =   135
         TabIndex        =   60
         Top             =   285
         Width           =   2835
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1020
      Left            =   6270
      TabIndex        =   43
      Top             =   1710
      Width           =   1785
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "      De Acuerdo al     TUO D.S. 054-99-EF Reglam. D.S. 122-94-EF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   75
         TabIndex        =   44
         Top             =   210
         Width           =   1650
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Retenciones Anuales"
      Height          =   4875
      Left            =   120
      TabIndex        =   40
      Top             =   2820
      Width           =   7950
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   0
         Left            =   3780
         TabIndex        =   11
         Top             =   405
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   1
         Left            =   3780
         TabIndex        =   12
         Top             =   645
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   2
         Left            =   3780
         TabIndex        =   13
         Top             =   885
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   4
         Left            =   3780
         TabIndex        =   15
         Top             =   1380
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   5
         Left            =   3780
         TabIndex        =   16
         Top             =   1620
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   3
         Left            =   3780
         TabIndex        =   14
         Top             =   1125
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   6
         Left            =   3780
         TabIndex        =   17
         Top             =   1860
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   7
         Left            =   3780
         TabIndex        =   18
         Top             =   2100
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   8
         Left            =   3780
         TabIndex        =   19
         Top             =   2340
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   9
         Left            =   3780
         TabIndex        =   20
         Top             =   2580
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   240
         Index           =   11
         Left            =   3780
         TabIndex        =   22
         Top             =   3075
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   423
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   210
         Index           =   10
         Left            =   3780
         TabIndex        =   21
         Top             =   2820
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   0
         Left            =   6105
         TabIndex        =   23
         Top             =   405
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   1
         Left            =   6105
         TabIndex        =   24
         Top             =   645
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   2
         Left            =   6105
         TabIndex        =   25
         Top             =   885
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   4
         Left            =   6105
         TabIndex        =   27
         Top             =   1380
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   5
         Left            =   6105
         TabIndex        =   28
         Top             =   1620
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   3
         Left            =   6105
         TabIndex        =   26
         Top             =   1125
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   6
         Left            =   6105
         TabIndex        =   29
         Top             =   1860
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   7
         Left            =   6105
         TabIndex        =   30
         Top             =   2100
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   8
         Left            =   6105
         TabIndex        =   31
         Top             =   2340
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   9
         Left            =   6105
         TabIndex        =   32
         Top             =   2580
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   240
         Index           =   11
         Left            =   6105
         TabIndex        =   34
         Top             =   3075
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   423
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xAcumula 
         Height          =   210
         Index           =   10
         Left            =   6105
         TabIndex        =   33
         Top             =   2820
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   370
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin VB.Label Label11 
         Caption         =   "M. /"
         Height          =   195
         Left            =   6120
         TabIndex        =   58
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label10 
         Caption         =   "M. x"
         Height          =   195
         Left            =   3810
         TabIndex        =   57
         Top             =   150
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NOV (Ro x 02) + A + GN + Ra              r11=(i-d)/4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   135
         TabIndex        =   56
         Top             =   2820
         Width           =   5460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ABR (Ro x 09) + A + GF + GN + Ra         r4=(i-a)/9"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   55
         Top             =   1125
         Width           =   5355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ENE (Ro x 12) + A + GF + GN              r1=i/12"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   54
         Top             =   405
         Width           =   5040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "FEB (Ro x 11) + A + GF + GN + Ra         r2=i/12"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   53
         Top             =   645
         Width           =   5040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "MAR (Ro x 10) + A + GF + GN + Ra         r3=i/12"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   52
         Top             =   885
         Width           =   5040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "MAY (Ro x 08) + A + GF + GN + Ra         r5=(i-b)/8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   51
         Top             =   1380
         Width           =   5355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "JUN (Ro x 07) + A + GF + GN + Ra         r6=(i-b)/8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   5
         Left            =   135
         TabIndex        =   50
         Top             =   1620
         Width           =   5355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "JUL (Ro x 06) + A + GN + Ra              r7=(i-b)/8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   49
         Top             =   1860
         Width           =   5355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "AGO (Ro x 05) + A + GN + Ra              r8=(i-c)/5"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   7
         Left            =   135
         TabIndex        =   48
         Top             =   2100
         Width           =   5355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SET (Ro x 04) + A + GN + Ra              r9=(i-d)/4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   8
         Left            =   135
         TabIndex        =   47
         Top             =   2340
         Width           =   5355
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "OCT (Ro x 03) + A + GN + Ra              r10=(i-d)/4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   135
         TabIndex        =   46
         Top             =   2580
         Width           =   5460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DIC (Ro x 01) + A + Ra                   r12= i-e"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   135
         TabIndex        =   45
         Top             =   3075
         Width           =   5145
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frPrmQC.frx":08CA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   105
         TabIndex        =   42
         Top             =   4380
         Width           =   7755
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frPrmQC.frx":0978
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   105
         TabIndex        =   41
         Top             =   3390
         Width           =   7755
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   6870
      TabIndex        =   36
      Top             =   660
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   6870
      TabIndex        =   35
      Top             =   195
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros de Cálculo"
      Height          =   1185
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   6030
      Begin AplisetControlText.Aplitext xMinimo 
         Height          =   285
         Left            =   4965
         TabIndex        =   3
         Top             =   825
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xUit 
         Height          =   285
         Left            =   4965
         TabIndex        =   2
         Top             =   525
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xNumUIT 
         Height          =   285
         Left            =   4965
         TabIndex        =   1
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Valor mínimo de deducción"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Valor de la UIT (en Nuevos Soles)"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   540
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero de U.I.T."
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   240
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frPrmQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS5TA As New ADODB.Recordset

Private Sub Command1_Click()
    With RS5TA
        If RS5TA.RecordCount = 0 Then
            .AddNew
        End If
        !NUMUIT = Val(xNumUIT.Text)
        !VALORUIT = Val(xUit.Text)
        !VALORMIN = Val(xMinimo.Text)
        
        !UIT1 = CDbl(XUIT1.Text)
        !UIT2 = CDbl(XUIT2.Text)
        !UIT3 = CDbl(XUIT3.Text)
        !UIT4 = CDbl(XUIT4.Text)
        
        !PORCENTAJE = Val(xPorcentaje.Text)
        !PORCENTAJE2 = Val(xPorcentaje2.Text)
        !PORCENTAJE3 = CDbl(xPorcentaje3.Text)
        !MES01 = Val(xMes(0).Text)
        !MES02 = Val(xMes(1).Text)
        !MES03 = Val(xMes(2).Text)
        !MES04 = Val(xMes(3).Text)
        !MES05 = Val(xMes(4).Text)
        !MES06 = Val(xMes(5).Text)
        !MES07 = Val(xMes(6).Text)
        !MES08 = Val(xMes(7).Text)
        !MES09 = Val(xMes(8).Text)
        !MES10 = Val(xMes(9).Text)
        !MES11 = Val(xMes(10).Text)
        !MES12 = Val(xMes(11).Text)
        !ACUMULA01 = Val(xAcumula(0).Text)
        !ACUMULA02 = Val(xAcumula(1).Text)
        !ACUMULA03 = Val(xAcumula(2).Text)
        !ACUMULA04 = Val(xAcumula(3).Text)
        !ACUMULA05 = Val(xAcumula(4).Text)
        !ACUMULA06 = Val(xAcumula(5).Text)
        !ACUMULA07 = Val(xAcumula(6).Text)
        !ACUMULA08 = Val(xAcumula(7).Text)
        !ACUMULA09 = Val(xAcumula(8).Text)
        !ACUMULA10 = Val(xAcumula(9).Text)
        !ACUMULA11 = Val(xAcumula(10).Text)
        !ACUMULA12 = Val(xAcumula(11).Text)
        .Update
    End With
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ACTUALIZABD
    RS5TA.Open "CONFIG5TA", DBSYSTEM, adOpenDynamic, adLockOptimistic
    If RS5TA.RecordCount > 0 Then
        xNumUIT.Text = RS5TA!NUMUIT
        xUit.Text = RS5TA!VALORUIT
        xMinimo.Text = RS5TA!VALORMIN
        
        XUIT1.Text = ESNULO(RS5TA!UIT1, 0)
        XUIT2.Text = ESNULO(RS5TA!UIT2, 0)
        XUIT3.Text = ESNULO(RS5TA!UIT3, 0)
        XUIT4.Text = ESNULO(RS5TA!UIT4, 0)
        
        xPorcentaje.Text = RS5TA!PORCENTAJE
        xPorcentaje2.Text = RS5TA!PORCENTAJE2
        xPorcentaje3.Text = ESNULO(RS5TA!PORCENTAJE3, 0)
        
        
        xMes(0).Text = "" & RS5TA!MES01
        xMes(1).Text = "" & RS5TA!MES02
        xMes(2).Text = "" & RS5TA!MES03
        xMes(3).Text = "" & RS5TA!MES04
        xMes(4).Text = "" & RS5TA!MES05
        xMes(5).Text = "" & RS5TA!MES06
        xMes(6).Text = "" & RS5TA!MES07
        xMes(7).Text = "" & RS5TA!MES08
        xMes(8).Text = "" & RS5TA!MES09
        xMes(9).Text = "" & RS5TA!MES10
        xMes(10).Text = "" & RS5TA!MES11
        xMes(11).Text = "" & RS5TA!MES12
        xAcumula(0).Text = "" & RS5TA!ACUMULA01
        xAcumula(1).Text = "" & RS5TA!ACUMULA02
        xAcumula(2).Text = "" & RS5TA!ACUMULA03
        xAcumula(3).Text = "" & RS5TA!ACUMULA04
        xAcumula(4).Text = "" & RS5TA!ACUMULA05
        xAcumula(5).Text = "" & RS5TA!ACUMULA06
        xAcumula(6).Text = "" & RS5TA!ACUMULA07
        xAcumula(7).Text = "" & RS5TA!ACUMULA08
        xAcumula(8).Text = "" & RS5TA!ACUMULA09
        xAcumula(9).Text = "" & RS5TA!ACUMULA10
        xAcumula(10).Text = "" & RS5TA!ACUMULA11
        xAcumula(11).Text = "" & RS5TA!ACUMULA12
    End If
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RS5TA = Nothing
End Sub

Public Sub ACTUALIZABD()
Dim I As Single
 For I = 1 To 12
   If Not ExisteCampo("MES" & Format(I, "00"), "CONFIG5TA", DBSYSTEM) Then
        DBSYSTEM.Execute "ALTER TABLE CONFIG5TA ADD COLUMN MES" & Format(I, "00") & " TEXT(4)  "
    End If
 Next I
End Sub

