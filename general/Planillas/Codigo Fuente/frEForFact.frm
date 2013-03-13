VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frEForFact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmulas de Facturación"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frEForFact.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2468
      TabIndex        =   5
      Top             =   2130
      Width           =   1260
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   953
      TabIndex        =   4
      Top             =   2130
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fórmula"
      Height          =   1845
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin AplisetControlText.Aplitext xFormula 
         Height          =   300
         Left            =   1245
         TabIndex        =   3
         Top             =   1320
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   529
         MaxLength       =   200
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   300
         Left            =   1245
         TabIndex        =   2
         Top             =   866
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   529
         MaxLength       =   35
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   300
         Left            =   1245
         TabIndex        =   1
         Top             =   412
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         MaxLength       =   8
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fórmula"
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   1373
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   919
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   465
         Width           =   495
      End
   End
End
Attribute VB_Name = "frEForFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMACEPTAR_CLICK()
    Dim X As Integer
    If xCodigo.Text = "" Then
        MsgBox "EL CÓDIGO DEBE TENER AL MENOS 4 CARACTERES. FALTA COMPLETAR DATOS", vbCritical
        Exit Sub
    End If
    If vpTarea = "NUEVO" Then
        DbSystem.Execute "UPDATE FORMFACT SET CODIGO=CODIGO WHERE CODIGO='" & xCodigo.Text & "'", X
        If X <> 0 Then
            MsgBox "EL CÓDIGO QUE HA INGRESADO YA EXISTE, POR FAVOR CAMBIELO", vbCritical
            xCodigo.SetFocus
            Exit Sub
        End If
    End If
    If xNombre.Text = "" Then
        MsgBox "FALTA AGREGAR EL NOMBRE DE LA FORMULA DE FACTURACIÓN", vbCritical
        xNombre.SetFocus
        Exit Sub
    End If
    If xFormula.Text = "" Then
        MsgBox "FALTA ESCRIBIR LA FORMULA DE ACCIÓN PARA LA FORMULA DE FACTURACIÓN, POR FAVOR INGRESELA CORRECTAMENTE", vbCritical
        xFormula.SetFocus
        Exit Sub
    End If
    If vpTarea = "EDITAR" Then DbSystem.Execute "DELETE FROM FORMFACT WHERE CODIGO='" & xCodigo.Text & "'"
    DbSystem.Execute "INSERT INTO FORMFACT (CODIGO,NOMBRE,FORMULA) VALUES ('" & xCodigo.Text & "','" & xNombre.Text & "','" & xFormula.Text & "')"
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub FORM_LOAD()
    If Not ExisteTabla("FORMFACT") Then
        MsgBox "NO EXISTE LA TABLA FORMFACT LA CUAL ES NECESARIA PARA EJECUTAR LAS TAREAS DEL SISTEMA, POR FAVOR COMUNICARSE CON ENTERPRISE SOLUTIONS S.A.", vbCritical
        Unload Me
    End If
    If vpTarea = "EDITAR" Then
        xCodigo.Locked = True
    End If
End Sub

