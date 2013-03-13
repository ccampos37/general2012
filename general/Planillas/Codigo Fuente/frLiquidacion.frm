VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frLiquidacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de Trabajadores"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frLiquidacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   5775
      Top             =   6255
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3353
      TabIndex        =   9
      Top             =   6270
      Width           =   1290
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   1778
      TabIndex        =   8
      Top             =   6270
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabajador"
      Height          =   6090
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6180
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagar"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   2100
         TabIndex        =   59
         Top             =   2610
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Base Cálculo"
         Height          =   450
         Left            =   135
         TabIndex        =   56
         Top             =   1425
         Width           =   1860
      End
      Begin AplisetControlText.Aplitext DiasGrati 
         Height          =   300
         Left            =   4725
         TabIndex        =   29
         Top             =   4080
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext DiasVac 
         Height          =   300
         Left            =   3390
         TabIndex        =   28
         Top             =   4080
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext DiasCTS 
         Height          =   300
         Left            =   2040
         TabIndex        =   27
         Top             =   4080
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext MesGrati 
         Height          =   300
         Left            =   4725
         TabIndex        =   25
         Top             =   3450
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext MesVac 
         Height          =   300
         Left            =   3390
         TabIndex        =   24
         Top             =   3450
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext MesCTS 
         Height          =   300
         Left            =   2040
         TabIndex        =   23
         Top             =   3450
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext AnnoGrati 
         Height          =   300
         Left            =   4725
         TabIndex        =   21
         Top             =   2820
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext AnnoVac 
         Height          =   300
         Left            =   3390
         TabIndex        =   20
         Top             =   2820
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext AnnoCTS 
         Height          =   300
         Left            =   2040
         TabIndex        =   19
         Top             =   2820
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         Text            =   "0.00"
         TipoDato        =   "N"
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagar"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4755
         TabIndex        =   14
         Top             =   2610
         Width           =   1035
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagar"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3420
         TabIndex        =   13
         Top             =   2610
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker xFechaCese 
         Height          =   300
         Left            =   4725
         TabIndex        =   7
         Top             =   1020
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36840
      End
      Begin MSComCtl2.DTPicker xFechaGrati 
         Height          =   300
         Left            =   4725
         TabIndex        =   6
         Top             =   2295
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36840
      End
      Begin MSComCtl2.DTPicker xFechaVac 
         Height          =   300
         Left            =   3390
         TabIndex        =   5
         Top             =   2295
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36840
      End
      Begin MSComCtl2.DTPicker xFechaCTS 
         Height          =   300
         Left            =   2055
         TabIndex        =   4
         Top             =   2295
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62193665
         CurrentDate     =   36840
      End
      Begin MSComCtl2.DTPicker xFechaIng 
         Height          =   300
         Left            =   2055
         TabIndex        =   3
         Top             =   1020
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   62193665
         CurrentDate     =   36840
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label xBase2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   58
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label xBase3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   57
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label xBase1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2055
         TabIndex        =   55
         Top             =   1980
         Width           =   1320
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Base de Cálculo"
         Height          =   300
         Left            =   120
         TabIndex        =   54
         Top             =   1980
         Width           =   1920
      End
      Begin VB.Label Neto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   53
         Top             =   5730
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Neto de Liquidacion de Tiempo de Servicios"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   12
         Left            =   90
         TabIndex        =   52
         Top             =   5730
         Width           =   4620
      End
      Begin VB.Label xAFP 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4725
         TabIndex        =   51
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   11
         Left            =   105
         TabIndex        =   50
         Top             =   5340
         Width           =   1920
      End
      Begin VB.Label Tot3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   49
         Top             =   5340
         Width           =   1320
      End
      Begin VB.Label Tot2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   48
         Top             =   5340
         Width           =   1320
      End
      Begin VB.Label Tot1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2040
         TabIndex        =   47
         Top             =   5340
         Width           =   1320
      End
      Begin VB.Label AFP2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   46
         Top             =   5025
         Width           =   1320
      End
      Begin VB.Label AFP1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   45
         Top             =   5025
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Retensión por Pensiones"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   10
         Left            =   105
         TabIndex        =   44
         Top             =   5025
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sub Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   9
         Left            =   105
         TabIndex        =   43
         Top             =   4710
         Width           =   1920
      End
      Begin VB.Label Sub3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   42
         Top             =   4710
         Width           =   1320
      End
      Begin VB.Label Sub2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   41
         Top             =   4710
         Width           =   1320
      End
      Begin VB.Label Sub1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2040
         TabIndex        =   40
         Top             =   4710
         Width           =   1320
      End
      Begin VB.Label Dias3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   39
         Top             =   4395
         Width           =   1320
      End
      Begin VB.Label Dias2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   38
         Top             =   4395
         Width           =   1320
      End
      Begin VB.Label Dias1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2040
         TabIndex        =   37
         Top             =   4395
         Width           =   1320
      End
      Begin VB.Label Mes3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   36
         Top             =   3765
         Width           =   1320
      End
      Begin VB.Label Mes2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   35
         Top             =   3765
         Width           =   1320
      End
      Begin VB.Label Mes1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2040
         TabIndex        =   34
         Top             =   3765
         Width           =   1320
      End
      Begin VB.Label Anno3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   4725
         TabIndex        =   33
         Top             =   3135
         Width           =   1320
      End
      Begin VB.Label Anno2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   3390
         TabIndex        =   32
         Top             =   3135
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fondo Pensiones"
         Height          =   195
         Left            =   4725
         TabIndex        =   31
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label Anno1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   2040
         TabIndex        =   30
         Top             =   3135
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Días por Cancelar"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   8
         Left            =   105
         TabIndex        =   26
         Top             =   4080
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Meses por Cancelar"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   7
         Left            =   105
         TabIndex        =   22
         Top             =   3450
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Años por Cancelar"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   6
         Left            =   105
         TabIndex        =   18
         Top             =   2820
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ultima Fecha de Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2295
         Width           =   1920
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha de Ingreso"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1020
         Width           =   1920
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha de Cese"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   3390
         TabIndex        =   15
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gratificaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   4725
         TabIndex        =   12
         Top             =   1665
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vacaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   3390
         TabIndex        =   11
         Top             =   1665
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.T.S."
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   2055
         TabIndex        =   10
         Top             =   1665
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Trabajador"
         Height          =   195
         Left            =   135
         TabIndex        =   1
         Top             =   375
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ANNOCTS_CHANGE()
    CALCULOSOLES
End Sub
Private Sub ANNOGRATI_CHANGE()
    CALCULOSOLES
End Sub
Private Sub ANNOVAC_CHANGE()
    CALCULOSOLES
End Sub
Private Sub CMACEPTAR_CLICK()
    If xTrab.Tag = "" Then Exit Sub
    If MsgBox("Desea aceptar los valores de Liquidacion", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim Z1 As Long
    DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO WHERE CODTRAB='" & xTrab.Tag & "' AND SALDO<>0 AND TIPOGRUPO=1", Z1
    If Z1 > 0 Then
        MsgBox "El trabajador " & xTrab.Text & " presenta Ingresos en su Cuenta Corriente. Elabore una boleta de remuneraciones para debitar por completo sus pendientes de la empresa", vbInformation
        Exit Sub
    End If
    If ExisteTablaAux(" [##TMPLIQCTA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPLIQCTA" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO WHERE CODTRAB='" & xTrab.Tag & "' AND SALDO<>0 AND TIPOGRUPO=2", Z1
    If Z1 > 0 Then
        Z1 = MsgBox("El trabajador " & xTrab.Text & " presenta Egresos en su Cuenta Corriente. Desea cobrar esos Egresos del Trabajador", vbYesNoCancel)
        If Z1 = vbCancel Then Exit Sub
        If Z1 = vbYes Then
            DBSYSTEM.Execute "SELECT CODMOV, DESCRIPCION, SALDO, SALDO AS MONTO INTO  [##TMPLIQCTA" & VGL_COMPUTER & "]  FROM MOVICTA WHERE TIPOGRUPO=2 AND CODTRAB='" & xTrab.Tag & "'"
            frPagoCtaCteLiq.Show 1
            If VPTAREA = "NO" Then Exit Sub
        End If
    End If
    If Val(Neto.Caption) <= 0 Then Exit Sub
    If Check2.Value Or Check3.Value = 3 Then
        RegLiquida.ActVaca = IIf(Check2.Value = 1, True, False)
        RegLiquida.ActGrati = IIf(Check3.Value = 1, True, False)
        frTrasLiquidac.Show 1
        If RegLiquida.CANCEL Then Exit Sub
        If Check1.Value = 1 Then
            DBSYSTEM.Execute "INSERT INTO INGMOV2000 (CODTRAB,CONCEPTO,VALOR,CODNOMBOL ) VALUES ('" & xTrab.Tag & "','REMUVAC'," & Sub2.Caption & "," & RegLiquida.CronoVac & ")"
        End If
        If Check1.Value = 2 Then
            DBSYSTEM.Execute "INSERT INTO INGMOV2000 (CODTRAB,CONCEPTO,VALOR,CODNOMBOL ) VALUES ('" & xTrab.Tag & "','REMUGRAT'," & Sub3.Caption & "," & RegLiquida.CronoGrat & ")"
        End If
    End If
    Dim RS1 As New ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    RS1.Open "LIQUIDACIONES", DBSYSTEM, adOpenDynamic, adLockOptimistic
    'LO HE HECHO DE ESTA MANERA, PUES EXEDIA EL LIMITE DE CADENA
    With RS1
        .AddNew
        !CODTRAB = xTrab.Tag
        !FECHAING = xFechaIng.Value
        !CARGO = DevuelveValor("SELECT CARGO FROM TRABAJADORES WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
        !FECHACESE = xFechaCese.Value
        !BASECTS = Val(xBase1.Caption)
        !BASEVAC = Val(xBase2.Caption)
        !BASEGRATI = Val(xBase3.Caption)
        !FECCTS = xFechaCTS.Value
        !FECVAC = xFechaVac.Value
        !FECGRATI = xFechaGrati.Value
        !a1 = AnnoCTS.Text
        !A2 = AnnoVac.Text
        !A3 = AnnoGrati.Text
        !M1 = MesCTS.Text
        !M2 = MesVac.Text
        !M3 = MesGrati.Text
        !D1 = DiasCTS.Text
        !D2 = DiasVac.Text
        !D3 = DiasGrati.Text
        .Fields("CODAFP") = Trim(xAFP.ToolTipText)
        !AFP1 = Val(AFP1.Caption)
        !AFP2 = Val(AFP2.Caption)
        !Neto = Val(Neto.Caption)
        !CronoVac = RegLiquida.CronoVac
        !CronoGrat = RegLiquida.CronoGrat
        .Update
        DBSYSTEM.Execute "UPDATE TRABAJADORES SET SITUACIÓN='2', FECHACESE=" & DateSQL(xFechaCese.Value) & " WHERE CODTRAB='" & xTrab.Tag & "'"
    End With
    Dim RSAUX As New ADODB.Recordset
    If ExisteTablaAux(" [##TMPLIQCTA" & VGL_COMPUTER & "] ") Then
        RSAUX.Open " [##TMPLIQCTA" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic, adLockReadOnly
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "INSERT INTO PAGOSCTA (CODMOV, NUMBOL, CODNOMBOL, TIPOBOLETA, MONTO, DOLAR, CODTRAB, TIPO) VALUES (" & RSAUX!CODMOV & "," & DevuelveValor("SELECT MAX(CODIGO) AS ULTIMO FROM LIQUIDACIONES", DBSYSTEM) & ",0,'L'," & RSAUX!MONTO & ",0,'" & xTrab.Tag & "',2)"
            RSAUX.MoveNext
        Loop
        Set RSAUX = Nothing
    End If
    DBSYSTEM.Execute "UPDATE TRABAJADORES SET SITUACIÓN='2', FECHACESE=" & DateSQL(xFechaCese.Value) & " WHERE CODTRAB='" & xTrab.Tag & "'"
    If MsgBox("Desea imprimir la hoja de Liquidaciones", vbQuestion + vbYesNo) = vbNo Then
        Unload Me
        Exit Sub
    End If
    cmAceptar.Enabled = False
    If ExisteTablaAux(" [##TMPVOUCHER" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE ##TMPVOUCHER" & VGL_COMPUTER
    
On Error GoTo ERRPRINT
    DBSYSTEM.Execute "CREATE TABLE  [##TMPVOUCHER" & VGL_COMPUTER & "] ( CODTRAB VARCHAR(8))"
    DBSYSTEM.Execute "INSERT INTO  [##TMPVOUCHER" & VGL_COMPUTER & "]  VALUES('" & Trim(xTrab.Tag) & "')"
    DBSYSTEM.Execute "CREATE INDEX CODTRAB ON  [##TMPVOUCHER" & VGL_COMPUTER & "]  (CODTRAB)"
    
    'REPORTE
    With Reporte
        .Reset
        .WindowTitle = "PLAN0074.RPT - RESUMEN"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=" & VGL_BASE & ""
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0074.RPT"
        .StoredProcParam(0) = REGSISTEMA.BASESQL
        .StoredProcParam(1) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        
        .Formulas(0) = "xEmpresa='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "xRuc='" & REGSISTEMA.RUC & "'"
        .Formulas(2) = "xCts=" & xBase1.Caption
        .Formulas(3) = "xAnnoCts=" & AnnoCTS.Text
        .Formulas(4) = "xMesCts=" & MesCTS.Text
        .Formulas(5) = "xDiasCts=" & DiasCTS.Text
        .Formulas(6) = "Anno1=" & Anno1.Caption
        .Formulas(7) = "Mes1=" & Mes1.Caption
        .Formulas(8) = "Dias1=" & Dias1.Caption
        
        .Formulas(9) = "xVac=" & xBase2.Caption
        .Formulas(10) = "xAnnoVac=" & AnnoVac.Text
        .Formulas(11) = "xMesVac=" & MesVac.Text
        .Formulas(12) = "xDiasVac=" & DiasVac.Text
        .Formulas(13) = "Anno2=" & Anno2.Caption
        .Formulas(14) = "Mes2=" & Mes2.Caption
        .Formulas(15) = "Dias2=" & Dias2.Caption
        
        .Formulas(16) = "xGrat=" & xBase3.Caption
        .Formulas(17) = "xAnnoGrati=" & AnnoGrati.Text
        .Formulas(18) = "xMesGrati=" & MesGrati.Text
        .Formulas(19) = "xDiasGrati=" & DiasGrati.Text
        .Formulas(20) = "Anno3=" & Anno3.Caption
        .Formulas(21) = "Mes3=" & Mes3.Caption
        .Formulas(22) = "Dias3=" & Dias3.Caption
        
        'Sección de AFP
        .Formulas(23) = "xAFP='" & xAFP.Caption & "(" & xAFP.Tag & " %)'"
        .Formulas(24) = "xAFPVac=" & AFP1.Caption
        .Formulas(25) = "xAFPGrat=" & AFP2.Caption
        
        If .Status <> 2 Then .Action = 1
    End With
    
    Screen.MousePointer = 11
    Screen.MousePointer = 1
 Exit Sub
 
ERRPRINT:
    Resume Next
    
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Command1_Click()
    If xFechaCTS.Value > xFechaCese.Value Then
        MsgBox "Existe una Inconsistencia de fechas. La fecha de C.T.S. no puede ser mayor que la Fecha de Cese", vbInformation
        Exit Sub
    End If
    If xFechaVac.Value > xFechaCese.Value Then
        MsgBox "Existe una Inconsistencia de fechas. La fecha de Vacaciones no puede ser mayor que la Fecha de Cese", vbInformation
        Exit Sub
    End If
    If xFechaGrati.Value > xFechaCese.Value Then
        MsgBox "Existe una Inconsistencia de fechas. La fecha de Gratificaciones no puede ser mayor que la Fecha de Cese", vbInformation
        Exit Sub
    End If
    If xFechaIng.Value >= xFechaCese.Value Then
        MsgBox "Existe una Inconsistencia de fechas. La fecha de Ingreso no puede ser mayor que la Fecha de Cese", vbInformation
        Exit Sub
    End If
    Load frmBaseCalcLiq
    With frmBaseCalcLiq
        .xTrab.Tag = xTrab.Tag
        .xTrab.Caption = xTrab.Text
        .xFecCTS.Value = xFechaCTS.Value
        .xFecVac.Value = xFechaVac.Value
        .xFecGrat.Value = xFechaGrati.Value
        .xFechaCese.Value = xFechaCese.Value
        .xFec3.Value = xFechaCese.Value
        .xFecha2.Value = xFechaCese.Value
        If Check2.Value = 0 Then .Tab1.TabEnabled(1) = False
        If Check3.Value = 0 Then .Tab1.TabEnabled(2) = False
        .REALIZARCALCULOCTS
        .Show 1
        CALCULOSOLES
    End With
End Sub
Private Sub DIASCTS_CHANGE()
    CALCULOSOLES
End Sub
Private Sub DIASGRATI_CHANGE()
    CALCULOSOLES
End Sub
Private Sub DIASVAC_CHANGE()
    CALCULOSOLES
End Sub
Private Sub Form_Load()
    xFechaCese.Value = Date
End Sub
Private Sub MESCTS_CHANGE()
    CALCULOSOLES
End Sub
Private Sub MESGRATI_CHANGE()
    CALCULOSOLES
End Sub
Private Sub MESVAC_CHANGE()
    CALCULOSOLES
End Sub
Private Sub XFECHACESE_CHANGE()
    CALCULODIAS
End Sub
Private Sub XFECHACTS_CHANGE()
    If VALIDARFECHA(xFechaCTS) Then Exit Sub
    CALCULODIAS
End Sub
Private Sub XFECHAGRATI_CHANGE()
    If VALIDARFECHA(xFechaGrati) Then Exit Sub
    CALCULODIAS
End Sub
Private Sub XFECHAVAC_CHANGE()
    If VALIDARFECHA(xFechaVac) Then Exit Sub
    CALCULODIAS
End Sub
Private Function VALIDARFECHA(FECHA As DTPicker) As Boolean
    VALIDARFECHA = True
    If FECHA.Value < xFechaIng.Value Then
        MsgBox "La ultima Fecha de Pago no puede ser menor que la Fecha de Ingreso"
        FECHA.Value = xFechaIng.Value
        Exit Function
    End If
    If FECHA.Value > xFechaCese.Value Then
        MsgBox "La ultima Fecha de Pago no puede ser mayor que la Fecha de Cese"
        FECHA.Value = xFechaCese.Value
        Exit Function
    End If
    VALIDARFECHA = False
End Function
Private Sub XTRAB_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT CODTRAB, NOMBRES, FECHAING, CODCCOSTO,CENTRO, NOMBREAFP, FONDOPENS FROM VWTRABAJ WHERE SITUACIÓN <'2' AND CODTRAB NOT IN (SELECT CODTRAB FROM HISTOVAC WHERE CERRADO=0)", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSAUX.RecordCount = 0 Or RSAUX.EOF Then
        MsgBox "No se han encontrado Trabajadores", vbInformation
        Set RSAUX = Nothing
        cmAceptar.Enabled = False
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xTrab.Text = RSAUX!NOMBRES
        xTrab.Tag = RSAUX!CODTRAB
        xFechaIng.Value = RSAUX!FECHAING
        xAFP.Caption = RSAUX!NOMBREAFP
        xAFP.ToolTipText = RSAUX!FONDOPENS
        xAFP.Tag = DevuelveValor("SELECT (APOROBLI+SEGURO+COMISIONRA) AS PORC1 FROM AFPS WHERE CODAFP='" & RSAUX!FONDOPENS & "'", DBSYSTEM)
        Dim XFEC As Date
        If IsNull(DevuelveValor("SELECT MAX(FECHAFIN) AS T1 FROM CTS, PLANCTS WHERE CTS.CODIGO=PLANCTS.CODIGO AND PLANCTS.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)) Then
            MsgBox "No presenta ningún Deposito de CTS", vbInformation
            xFechaCTS.Value = xFechaIng.Value
        Else
            xFechaCTS.Value = DevuelveValor("SELECT MAX(FECHAFIN) AS T1 FROM CTS, PLANCTS WHERE CTS.CODIGO=PLANCTS.CODIGO AND PLANCTS.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            If xFechaCTS.Value < xFechaIng.Value Then xFechaCTS.Value = xFechaIng
        End If
        If IsNull(DevuelveValor("SELECT MAX(FECHAFIN) AS T1 FROM HISTOVAC WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)) Then
            MsgBox "No presenta ningún Pago de Vacaciones", vbInformation
            xFechaVac.Value = xFechaIng.Value
        Else
            xFechaVac.Value = DevuelveValor("SELECT MAX(FECHAFIN) AS T1 FROM HISTOVAC WHERE CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            If xFechaVac.Value < xFechaIng.Value Then xFechaVac.Value = xFechaIng
        End If
        If IsNull(DevuelveValor("SELECT MAX(FECHAFIN) AS T1 FROM GRATIFICACION, PLANGRATI WHERE GRATIFICACION.CODIGO=PLANGRATI.CODIGO AND PLANGRATI.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)) Then
            MsgBox "No presenta ninguna gratificación PAGADA", vbInformation
            xFechaGrati.Value = xFechaIng.Value
        Else
            xFechaGrati.Value = DevuelveValor("SELECT MAX(FECHAFIN) AS T1 FROM GRATIFICACION, PLANGRATI WHERE GRATIFICACION.CODIGO=PLANGRATI.CODIGO AND PLANGRATI.CODTRAB='" & xTrab.Tag & "'", DBSYSTEM)
            If xFechaGrati.Value < xFechaIng.Value Then xFechaGrati.Value = xFechaIng
        End If
        CALCULODIAS
    End If
    Set RSAUX = Nothing
End Sub
Public Sub CALCULODIAS()
    Dim A As Integer, B As Integer, C As Integer
    TiempoTrans xFechaCTS.Value, xFechaCese.Value, A, B, C
    AnnoCTS.Text = A
    MesCTS.Text = B
    DiasCTS.Text = C
    TiempoTrans xFechaVac.Value, xFechaCese.Value, A, B, C
    AnnoVac.Text = A
    MesVac.Text = B
    DiasVac.Text = C
    TiempoTrans xFechaGrati.Value, xFechaCese.Value, A, B, C
    AnnoGrati.Text = A
    MesGrati.Text = B
    DiasGrati.Text = C
    CALCULOSOLES
End Sub
Public Sub CALCULOSOLES()
    Anno1.Caption = Format(Val(xBase1.Caption) * Val(AnnoCTS.Text), "0.00")
    Mes1.Caption = Format(Val(xBase1.Caption) / 12 * Val(MesCTS.Text), "0.00")
    Dias1.Caption = Format(Val(xBase1.Caption) / 360 * Val(DiasCTS.Text), "0.00")
    Anno2.Caption = Format(Val(xBase2.Caption) * Val(AnnoVac.Text), "0.00")
    Mes2.Caption = Format(Val(xBase2.Caption) / 12 * Val(MesVac.Text), "0.00")
    Dias2.Caption = Format(Val(xBase2.Caption) / 360 * Val(DiasVac.Text), "0.00")
    Anno3.Caption = Format(Val(xBase3.Caption) * Val(AnnoGrati.Text), "0.00")
    Mes3.Caption = Format(Val(xBase3.Caption) / 12 * Val(MesGrati.Text), "0.00")
    Dias3.Caption = Format(Val(xBase3.Caption) / 360 * Val(DiasGrati.Text), "0.00")
    Sub1.Caption = Format(Val(Mes1.Caption) + Val(Anno1.Caption) + Val(Dias1.Caption), "0.00")
    Sub2.Caption = Format(Val(Mes2.Caption) + Val(Anno2.Caption) + Val(Dias2.Caption), "0.00")
    Sub3.Caption = Format(Val(Mes3.Caption) + Val(Anno3.Caption) + Val(Dias3.Caption), "0.00")
    AFP1.Caption = Format(Val(Sub2.Caption) * Val(xAFP.Tag) / 100, "0.00")
    AFP2.Caption = Format(Val(Sub3.Caption) * Val(xAFP.Tag) / 100, "0.00")
    Tot1.Caption = Sub1.Caption
    Tot2.Caption = Format(Val(Sub2.Caption) - Val(AFP1.Caption), "0.00")
    Tot3.Caption = Format(Val(Sub3.Caption) - Val(AFP2.Caption), "0.00")
    Neto.Caption = Val(Tot1.Caption) + Val(Tot2.Caption) + Val(Tot3.Caption)
End Sub

