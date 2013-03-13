VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPrGenAsiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Asientos de Venta"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   Icon            =   "FrmProGenAsiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9015
   Begin VB.Frame Frame1 
      Height          =   1020
      Index           =   2
      Left            =   1380
      TabIndex        =   0
      Top             =   5355
      Width           =   6615
      Begin VB.CommandButton CmdCon 
         Caption         =   "&Consulta"
         Height          =   675
         Left            =   1665
         Picture         =   "FrmProGenAsiento.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones1 
         Caption         =   "&Enviar"
         Height          =   675
         Index           =   0
         Left            =   495
         Picture         =   "FrmProGenAsiento.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones1 
         Caption         =   "&Imprimir"
         CausesValidation=   0   'False
         Height          =   675
         Index           =   2
         Left            =   4020
         Picture         =   "FrmProGenAsiento.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones1 
         Caption         =   "&Eliminar"
         CausesValidation=   0   'False
         Height          =   675
         Index           =   1
         Left            =   2850
         Picture         =   "FrmProGenAsiento.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones1 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   675
         Index           =   3
         Left            =   5205
         Picture         =   "FrmProGenAsiento.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Consulta de Comprobante"
      Height          =   5220
      Left            =   90
      TabIndex        =   16
      Top             =   105
      Visible         =   0   'False
      Width           =   8775
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1770
         Left            =   225
         TabIndex        =   28
         Top             =   2295
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   3122
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label LbTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label25"
         Height          =   255
         Left            =   6600
         TabIndex        =   45
         Top             =   1140
         Width           =   1620
      End
      Begin VB.Label LbTasa 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label24"
         Height          =   255
         Left            =   4290
         TabIndex        =   44
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label LbDocumento 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label23"
         Height          =   255
         Left            =   6600
         TabIndex        =   43
         Top             =   765
         Width           =   1650
      End
      Begin VB.Label LbFecDoc 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label22"
         Height          =   255
         Left            =   4290
         TabIndex        =   42
         Top             =   765
         Width           =   1020
      End
      Begin VB.Label LbFecha 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label21"
         Height          =   255
         Left            =   6600
         TabIndex        =   41
         Top             =   375
         Width           =   1020
      End
      Begin VB.Label LbComprobante 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label20"
         Height          =   255
         Left            =   4290
         TabIndex        =   40
         Top             =   375
         Width           =   1020
      End
      Begin VB.Label LbIgv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label19"
         Height          =   255
         Left            =   1410
         TabIndex        =   39
         Top             =   1140
         Width           =   1590
      End
      Begin VB.Label LbCliente 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label18"
         Height          =   255
         Left            =   1410
         TabIndex        =   38
         Top             =   765
         Width           =   1605
      End
      Begin VB.Label LbSubdiario 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label17"
         Height          =   255
         Left            =   1410
         TabIndex        =   37
         Top             =   375
         Width           =   495
      End
      Begin VB.Label LbCambio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label18"
         Height          =   255
         Left            =   6600
         TabIndex        =   36
         Top             =   1515
         Width           =   1590
      End
      Begin VB.Label LbConversion 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label17"
         Height          =   255
         Left            =   1410
         TabIndex        =   35
         Top             =   1515
         Width           =   495
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   8340
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Label LbDiferencia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label19"
         Height          =   255
         Left            =   6645
         TabIndex        =   34
         Top             =   4560
         Width           =   1650
      End
      Begin VB.Label LbHaber 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label18"
         Height          =   255
         Left            =   3705
         TabIndex        =   33
         Top             =   4560
         Width           =   1650
      End
      Begin VB.Label LbDebe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label17"
         Height          =   255
         Left            =   930
         TabIndex        =   32
         Top             =   4560
         Width           =   1650
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8340
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label16 
         Caption         =   "Diferencia  :"
         Height          =   240
         Left            =   5685
         TabIndex        =   31
         Top             =   4575
         Width           =   1395
      End
      Begin VB.Label Label15 
         Caption         =   "Haber   :"
         Height          =   240
         Left            =   2985
         TabIndex        =   30
         Top             =   4575
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Debe   :"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   4575
         Width           =   645
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Cambio  :"
         Height          =   240
         Left            =   5445
         TabIndex        =   27
         Top             =   1530
         Width           =   1875
      End
      Begin VB.Label Label12 
         Caption         =   "Conversión   :"
         Height          =   240
         Left            =   285
         TabIndex        =   26
         Top             =   1530
         Width           =   1875
      End
      Begin VB.Label Label11 
         Caption         =   "Monto Total  :"
         Height          =   255
         Left            =   5445
         TabIndex        =   25
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label10 
         Caption         =   "Tasa               :"
         Height          =   255
         Left            =   3105
         TabIndex        =   24
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label9 
         Caption         =   "Monto I.G.V. :"
         Height          =   255
         Left            =   285
         TabIndex        =   23
         Top             =   1140
         Width           =   1545
      End
      Begin VB.Label Label8 
         Caption         =   "Documento   :"
         Height          =   240
         Left            =   5445
         TabIndex        =   22
         Top             =   780
         Width           =   2010
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Doc.    :"
         Height          =   240
         Left            =   3105
         TabIndex        =   21
         Top             =   780
         Width           =   2010
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente           :"
         Height          =   240
         Left            =   285
         TabIndex        =   20
         Top             =   780
         Width           =   2010
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha           :"
         Height          =   210
         Left            =   5445
         TabIndex        =   19
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Comprobante  :"
         Height          =   210
         Left            =   3105
         TabIndex        =   18
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Subdiario       :"
         Height          =   210
         Left            =   285
         TabIndex        =   17
         Top             =   420
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4500
      Left            =   60
      TabIndex        =   11
      Top             =   840
      Width           =   8760
      Begin VB.CommandButton cmdMarc2 
         Caption         =   "&Marcar Todo"
         Height          =   285
         Left            =   105
         TabIndex        =   49
         Top             =   2310
         Width           =   1050
      End
      Begin VB.CommandButton cmdDesm2 
         Caption         =   "&Desm. Todo"
         Height          =   285
         Left            =   1320
         TabIndex        =   48
         Top             =   2310
         Width           =   1050
      End
      Begin VB.CommandButton cmdMarc1 
         Caption         =   "&Marcar Todo"
         Height          =   285
         Left            =   105
         TabIndex        =   47
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdDesm1 
         Caption         =   "&Desm. Todo"
         Height          =   285
         Left            =   1350
         TabIndex        =   46
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton CmdEnviar 
         Caption         =   "Generar    >>"
         Height          =   285
         Left            =   7125
         TabIndex        =   12
         Top             =   2310
         Width           =   1500
      End
      Begin MSFlexGridLib.MSFlexGrid Flex2 
         Height          =   1755
         Left            =   105
         TabIndex        =   14
         Top             =   2640
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   3096
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid Flex1 
         Height          =   1755
         Left            =   105
         TabIndex        =   15
         Top             =   510
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   3096
         _Version        =   393216
         ForeColorSel    =   16777215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   105
      Width           =   8760
      Begin VB.CommandButton CmdAceptar 
         Caption         =   ">>"
         Height          =   225
         Left            =   7725
         TabIndex        =   10
         Top             =   300
         Width           =   480
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   4935
         TabIndex        =   9
         Top             =   240
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36753
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1665
         TabIndex        =   8
         Top             =   240
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24772609
         CurrentDate     =   36753
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final  :"
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial   :"
         Height          =   255
         Left            =   390
         TabIndex        =   6
         Top             =   270
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FrmPrGenAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cVGDBT As ADODB.Connection
Dim Ado1 As ADODB.Recordset
Dim Ado2 As ADODB.Recordset
Dim cAsiento As String
Dim cSubdiario As String
Dim nTra As Integer, nTra2 As Integer

Private Sub CmdAceptar_Click()
Dim nI As Integer
On Error GoTo ErrAcep
nTra = 1
cConexCom.BeginTrans
cConexCom.Execute "Update FacCab Set CFCOMPROB = ' '  Where ISNULL(CFCOMPROB)"
cConexCom.Execute "Update FacCab Set CFSUBDIAR = ' '  Where ISNULL(CFSUBDIAR)"
cConexCom.CommitTrans
nTra = 0


Set Ado1 = New ADODB.Recordset
Ado1.Open "Select CFTD,CFNUMSER , CFNUMDOC,CFFECDOC,CFFECVEN,CFCODMON,CFIMPORTE  From " & _
"FACCAB WHERE CFFECDOC >= #" & Format(DTPicker1, "MM/DD/YYYY") & "# AND CFFECDOC <= #" & Format(DTPicker2, "MM/DD/YYYY") & "#  AND  " & _
"TRIM(CFCOMPROB) = '' And Trim(CFSUBDIAR) = ''", cConexCom, adOpenStatic

Set_Flex1

If Ado1.RecordCount > 0 Then
    Do While Not Ado1.EOF
        Flex1.AddItem ("" & vbTab & Ado1(0) & vbTab & Ado1(1) & vbTab & Ado1(2) & vbTab & Ado1(3) & vbTab & Ado1(4) & vbTab & Ado1(5) & vbTab & Ado1(6))
        Ado1.MoveNext
        If Ado1.EOF Then Exit Do
    Loop
End If
Ado1.Close
Exit Sub
ErrAcep:
            MsgBox err.Description
            If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub cmdBotones1_Click(Index As Integer)
Dim nI As Integer
Select Case Index
Case 0: 'Envia
             For nI = 0 To Flex2.Rows - 1
                    If Flex2.TextMatrix(nI, 0) = "»" Then
                            Call Enviar(Flex2.TextMatrix(nI, 2), Flex2.TextMatrix(nI, 8), Flex2.TextMatrix(nI, 4), Flex2.TextMatrix(nI, 5), Flex2.TextMatrix(nI, 6))
                    End If
            Next nI

            For nI = (Flex2.Rows - 1) To 2 Step -1
                    If Flex2.TextMatrix(nI, 0) = "»" Then
                            Flex2.RemoveItem nI
                    End If
            Next nI

            If nI = 1 Then
                    If Flex2.TextMatrix(nI, 0) = "»" Then
                            If (Flex2.Rows - 1) = 1 Then
                                Set_Flex2
                            Else
                                    Flex2.RemoveItem nI
                            End If
                    End If
            End If
Case 1: 'Elimina
            For nI = 0 To Flex2.Rows - 1
                    If Flex2.TextMatrix(nI, 0) = "»" Then
                            Call Elimina(Flex2.TextMatrix(nI, 2), Flex2.TextMatrix(nI, 8), Flex2.TextMatrix(nI, 4), Flex2.TextMatrix(nI, 5))
                    End If
            Next nI

            For nI = (Flex2.Rows - 1) To 2 Step -1
                    If Flex2.TextMatrix(nI, 0) = "»" Then
                            Flex2.RemoveItem nI
                    End If
            Next nI

            If nI = 1 Then
                    If Flex2.TextMatrix(nI, 0) = "»" Then
                            If (Flex2.Rows - 1) = 1 Then
                                Set_Flex2
                            Else
                                    Flex2.RemoveItem nI
                            End If
                    End If
            End If
            
Case 2:
               ' reportes
Case 3:
            If cmdBotones1(1).Enabled = False Then
                    Frame4.Visible = False
                    cmdBotones1(0).Enabled = True
                    cmdBotones1(1).Enabled = True
                    cmdBotones1(2).Enabled = True
                    cmdBotones1(3).Enabled = True
                    CmdCon.Enabled = True
            Else
                    Unload Me
            End If
End Select
End Sub

Private Sub CmdCon_Click()
Dim Ad1 As ADODB.Recordset
Dim Ad2 As ADODB.Recordset
Dim Ad3 As ADODB.Recordset
Dim Ad4 As ADODB.Recordset

Set Ad1 = New ADODB.Recordset
Set Ad2 = New ADODB.Recordset
Set Ad3 = New ADODB.Recordset
Set Ad4 = New ADODB.Recordset

LimpiaLabel
Init_ControlDataGrid DataGrid1

If Flex2.TextMatrix(Flex2.Row, 0) = "»" Then
        Ad1.Open "Select * From ContCab Where SUBDIAR_CODIGO = '" & Flex2.TextMatrix(Flex2.Row, 2) & "' and CMOV_NUM = '" & Flex2.TextMatrix(Flex2.Row, 8) & "'", cConexCom, adOpenStatic
        Ad2.Open "Select DMOV_SECUE,DMOV_CUENT,DMOV_ANEXO,IIF(CMOV_MONED = 'MN',DMOV_DEBE,DMOV_DEBUS) AS COL1,IIF(CMOV_MONED = 'MN',DMOV_HABER,DMOV_HABUS) AS COL2  " & _
                         "From ContDet A  Inner Join ContCab  B on A.SUBDIAR_CODIGO = B.SUBDIAR_CODIGO and  A.DMOV_NUM = B.CMOV_NUM Where A.SUBDIAR_CODIGO = '" & Flex2.TextMatrix(Flex2.Row, 2) & "' " & _
                         "and DMOV_NUM = '" & Flex2.TextMatrix(Flex2.Row, 8) & "'", cConexCom, adOpenStatic
                         
        Ad3.Open "Select * From ContVentas Where CO_C_SUBDI = '" & Flex2.TextMatrix(Flex2.Row, 2) & "' and CO_NUM = '" & Flex2.TextMatrix(Flex2.Row, 8) & "'", cConexCom, adOpenStatic
        
        If Ad1.RecordCount > 0 Then
                LbSubdiario = Ad1("SUBDIAR_CODIGO")
                LbComprobante = IIf(Not IsNull(Ad1("CMOV_C_COMPR")), Ad1("CMOV_C_COMPR"), " ")
                LbFecha = Ad1("CMOV_FECHA")
                LbConversion = Ad1("CMOV_CONVE")
                LbCambio = Format(Ad1("CMOV_TIPCA"), "0.#0")
                
                Ad4.Open "Select Sum (IIF(CMOV_MONED = 'MN',DMOV_DEBE,DMOV_DEBUS))  AS COL1,Sum(IIF(CMOV_MONED = 'MN',DMOV_HABER,DMOV_HABUS))  AS COL2 " & _
                                 "From ContDet A  Inner Join ContCab  B on A.SUBDIAR_CODIGO = B.SUBDIAR_CODIGO and  A.DMOV_NUM = B.CMOV_NUM Where A.SUBDIAR_CODIGO = '" & Flex2.TextMatrix(Flex2.Row, 2) & "' " & _
                                  "and DMOV_NUM = '" & Flex2.TextMatrix(Flex2.Row, 8) & "'", cConexCom, adOpenStatic
                
                LbDebe = Format(Ad4("Col1"), "0.#0")
                LbHaber = Format(Ad4("Col2"), "0.#0")
                LbDiferencia = Format(Ad4("COL1") - Ad4("COL2"), "0.#0")
        End If
        Ad1.Close
        If Ad3.RecordCount > 0 Then
                LbCliente = Ad3("CO_C_CLIEN")
                LbFecDoc = Ad3("CO_D_FECDC")
                LbDocumento = Ad3("CO_C_TPDOC") & Ad3("CO_C_DOCUM")
                LbIgv = Format(IIf(Ad3("CO_C_MONED") = "MN", Ad3("CO_N_IGV"), Ad3("CO_N_IGVUS")), "0.#0")
                LbTasa = Format(Ad3("CO_N_TASA"), "0.#0")
                LbTotal = Format(IIf(Ad3("CO_C_MONED") = "MN", Ad3("CO_N_MONTO"), Ad3("CO_N_MTOUS")), "0.#0")
        End If
        Ad3.Close
        Set DataGrid1.DataSource = Ad2
        
        With DataGrid1
                .Columns(0).Caption = "   Sec."
                .Columns(1).Caption = "         Cuenta"
                .Columns(2).Caption = "         Anexo"
                .Columns(3).Caption = "         Debe"
                .Columns(3).Alignment = dbgRight
                .Columns(3).NumberFormat = "#0.#0 "
                .Columns(4).Caption = "         Haber"
                .Columns(4).Alignment = dbgRight
                .Columns(4).NumberFormat = "#0.#0 "
        End With
        
        DataGrid1.Refresh
        
        Frame4.Visible = True
        cmdBotones1(0).Enabled = False
        cmdBotones1(1).Enabled = False
        cmdBotones1(2).Enabled = False
        cmdBotones1(3).Enabled = True
        CmdCon.Enabled = False
End If
End Sub

Private Sub cmdDesm1_Click()
Dim I As Integer
For I = 1 To Flex1.Rows - 1
            Flex1.TextMatrix(I, 0) = ""
Next I
End Sub

Private Sub cmdDesm2_Click()
Dim I As Integer
For I = 1 To Flex2.Rows - 1
            Flex2.TextMatrix(I, 0) = ""
Next I
End Sub

Private Sub CmdEnviar_Click()
Dim nI As Integer
For nI = 0 To Flex1.Rows - 1
        If Flex1.TextMatrix(nI, 0) = "»" Then
                Call Gen_Asiento(Flex1.TextMatrix(nI, 1), Flex1.TextMatrix(nI, 2), Flex1.TextMatrix(nI, 3), Flex1.TextMatrix(nI, 6))
        End If
Next nI

For nI = (Flex1.Rows - 1) To 2 Step -1
        If Flex1.TextMatrix(nI, 0) = "»" Then
                   Flex1.RemoveItem nI
        End If
Next nI

If nI = 1 Then
        If Flex1.TextMatrix(nI, 0) = "»" Then
            If (Flex1.Rows - 1) = 1 Then
                Set_Flex1
            Else
                Flex1.RemoveItem nI
            End If
        End If
End If

Set_Flex2

Set Ado2 = New ADODB.Recordset
Ado2.Open "Select CO_C_CUENT,CO_C_SUBDI,CO_C_COMPR,CO_C_TPDOC,CO_C_DOCUM,CO_C_MONED,CO_N_MONTO,CO_N_MTOUS,CO_NUM    From CONTVENTAS Where trim(CO_C_COMPR) = ''  or isnull(CO_C_COMPR)   and CO_PER >= '" & Year(DTPicker1) & Month(DTPicker1) & "' and CO_PER <= '" & Year(DTPicker2) & Month(DTPicker2) & "'", cConexCom, adOpenStatic
If Ado2.RecordCount > 0 Then
        Do While Not Ado2.EOF
                Flex2.AddItem ("" & vbTab & Ado2(0) & vbTab & Ado2(1) & vbTab & Ado2(2) & vbTab & Ado2(3) & vbTab & Ado2(4) & vbTab & Ado2(5) & vbTab & IIf(Ado2(5) = "MN", Ado2(6), Ado2(7)) & vbTab & Ado2(8))
                Ado2.MoveNext
                If Ado2.EOF Then Exit Do
        Loop
End If
End Sub

Private Sub cmdMarc1_Click()
Dim I As Integer
For I = 1 To Flex1.Rows - 1
            Flex1.TextMatrix(I, 0) = "»"
Next I
End Sub

Private Sub cmdMarc2_Click()
Dim I As Integer
For I = 1 To Flex2.Rows - 1
            Flex2.TextMatrix(I, 0) = "»"
Next I
End Sub

Private Sub Flex1_Click()
If Flex1.Row <> 0 Then
    If Flex1.TextMatrix(Flex1.Row, 0) = "»" Then
        Flex1.TextMatrix(Flex1.Row, 0) = ""
    Else
        Flex1.TextMatrix(Flex1.Row, 0) = "»"
    End If
End If
End Sub

Private Sub Flex1_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeySpace Then
          Flex1_Click
 End If
 End Sub

Private Sub Flex2_Click()
If Flex2.Row <> 0 Then
    If Flex2.TextMatrix(Flex2.Row, 0) = "»" Then
        Flex2.TextMatrix(Flex2.Row, 0) = ""
    Else
        Flex2.TextMatrix(Flex2.Row, 0) = "»"
    End If
End If
End Sub

Private Sub Flex2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
          Flex2_Click
 End If
End Sub

Private Sub Form_Load()
CentForm Me
ADOConectar

Set_Flex1
Set_Flex2

DTPicker1 = Date
DTPicker1.Day = 1
DTPicker2 = Date

cAsiento = Devolver_Dato(2, vGEmpresa, "Parametros", "ctccod", "CTTIPOASIEN")
cSubdiario = Devolver_Dato(2, vGEmpresa, "Parametros", "ctccod", "CTSUBVEN")

Set Ado1 = New ADODB.Recordset
Ado1.Open "Select CFTD,CFNUMSER , CFNUMDOC,CFFECDOC,CFFECVEN,CFCODMON,CFIMPORTE  From " & _
"FACCAB WHERE CFFECDOC >= #" & Format(DTPicker1, "MM/DD/YYYY") & "# AND CFFECDOC <= #" & Format(DTPicker2, "MM/DD/YYYY") & "#  AND  " & _
"TRIM(CFCOMPROB) = '' And Trim(CFSUBDIAR) = ''", cConexCom, adOpenStatic

Set_Flex1

If Ado1.RecordCount > 0 Then
    Do While Not Ado1.EOF
        Flex1.AddItem ("" & vbTab & Ado1(0) & vbTab & Ado1(1) & vbTab & Ado1(2) & vbTab & Ado1(3) & vbTab & Ado1(4) & vbTab & Ado1(5) & vbTab & Ado1(6))
        Ado1.MoveNext
        If Ado1.EOF Then Exit Do
    Loop
End If
Ado1.Close

Set Ado2 = New ADODB.Recordset
Ado2.Open "Select CO_C_CUENT,CO_C_SUBDI,CO_C_COMPR,CO_C_TPDOC,CO_C_DOCUM,CO_C_MONED,CO_N_MONTO,CO_N_MTOUS,CO_NUM    From CONTVENTAS Where trim(CO_C_COMPR) = ''  or isnull(CO_C_COMPR)   and CO_PER >= '" & Year(DTPicker1) & Month(DTPicker1) & "' and CO_PER <= '" & Year(DTPicker2) & Month(DTPicker2) & "'", cConexCom, adOpenStatic
If Ado2.RecordCount > 0 Then
        Do While Not Ado2.EOF
                Flex2.AddItem ("" & vbTab & Ado2(0) & vbTab & Ado2(1) & vbTab & Ado2(2) & vbTab & Ado2(3) & vbTab & Ado2(4) & vbTab & Ado2(5) & vbTab & IIf(Ado2(5) = "MN", Ado2(6), Ado2(7)) & vbTab & Ado2(8))
                Ado2.MoveNext
                If Ado2.EOF Then Exit Do
        Loop
End If
Ado2.Close
End Sub

Private Sub Gen_Asiento(cTipo As String, cNumeroSerie As String, cNumero As String, cMoneda As String)
Dim cAdo2 As ADODB.Recordset
Dim cAdo3 As ADODB.Recordset
Dim cSql11 As String, cSql22 As String, cSql33 As String
Dim cCuentaCliente As String
Dim cCuentaTributo As String
Dim cCuentaVentas As String
Dim nPorIgv As Integer
Dim nI As Integer
Dim cFam  As String
Dim cComAux As String
Dim nTotDet As Double

On Error GoTo ErrGen

Set cAdo2 = New ADODB.Recordset
Set cAdo3 = New ADODB.Recordset

cCuentaCliente = "": cCuentaTributo = "": cCuentaVentas = ""

cAdo2.Open "Select * from Parametros_Ctas Where TIPO_FAC = '" & cTipo & "'", cConexCom, adOpenStatic
If cAdo2.RecordCount > 0 Then
If Trim(cMoneda) = "MN" Then
        cCuentaCliente = cAdo2("CTA_SOLES")
Else
        cCuentaCliente = cAdo2("CTA_DOLA")
End If
cCuentaTributo = cAdo2("CTA_1")
cCuentaVentas = cAdo2("CTA_2")
cAdo2.Close

cAdo3.Open "Select * From FacCab Where CFTD = '" & cTipo & "' and CFNUMSER = '" & cNumeroSerie & "' AND CFNUMDOC = '" & cNumero & "'", cConexCom, adOpenStatic
If cAdo3.RecordCount > 0 Then
        cComAux = Numera(cSubdiario)
        
        If cAdo3("CFESTADO") <> "A" Then
                cSql11 = "Insert Into ContVentas (CO_C_Cuent,CO_C_Mes,CO_C_Subdi,CO_C_CLIEN,CO_C_TpDoc,CO_C_Docum,CO_C_TpDRf,CO_C_DCRef,CO_N_DCTO,CO_N_DCTUS,CO_N_Igv,"
                cSql11 = cSql11 & "CO_N_IgvUs,CO_N_Tasa,CO_N_Monto,CO_N_MtoUS,CO_C_Moned,CO_C_Conve,CO_N_TipCa,CO_C_Ruc,CO_A_Razon,CO_PER,CO_NUM,CO_D_FECHA,CO_D_FECDC) Values  "
                cSql11 = cSql11 & "('" & cCuentaCliente & "','" & Format(Month(Date), "00") & "','" & cSubdiario & "','" & cAdo3("CFCODCLI") & "','" & cAdo3("cftd") & "','" & cAdo3("cfnumser") & cAdo3("cfnumdoc") & "',"
                cSql11 = cSql11 & "'" & IIf(Trim(cAdo3("CFRFTD")) <> "", cAdo3("CFRFTD"), "  ") & "','" & IIf(Trim(cAdo3("CFRFNUMSER")) <> "", cAdo3("CFRFNUMSER"), "  ") & IIf(Trim(cAdo3("CFRFNUMDOC")) <> "", cAdo3("CFRFNUMDOC"), "  ") & "',"
                If Trim(cMoneda) = "MN" Then
                    cSql11 = cSql11 & "" & cAdo3("CFDESCVAL") & ","
                Else
                    cSql11 = cSql11 & "" & Val(Format(cAdo3("CFDESCVAL") * cAdo3("CFTIPCAM"), ".00")) & ","
                End If
                If Trim(cMoneda) = "ME" Then
                    cSql11 = cSql11 & "" & cAdo3("CFDESCVAL") & ","
                Else
                    cSql11 = cSql11 & "" & Val(Format(cAdo3("CFDESCVAL") / cAdo3("CFTIPCAM"), ".00")) & ","
                End If
                If Trim(cMoneda) = "MN" Then
                    cSql11 = cSql11 & "" & cAdo3("CFIGV") & ","
                Else
                    cSql11 = cSql11 & "" & Val(Format(cAdo3("CFIGV") * cAdo3("CFTIPCAM"), ".00")) & ","
                End If
                If Trim(cMoneda) = "ME" Then
                    cSql11 = cSql11 & "" & cAdo3("CFIGV") & ","
                Else
                    cSql11 = cSql11 & "" & Val(Format(cAdo3("CFIGV") / cAdo3("CFTIPCAM"), ".00")) & ","
                End If
                nPorIgv = Devolver_Dato(2, vGEmpresa, "Parametros", "CTCCOD", "CTVALORIGV")
                cSql11 = cSql11 & "" & nPorIgv & ","
                If cMoneda = "MN" Then
                    cSql11 = cSql11 & "" & cAdo3("CFIMPORTE") & ","
                Else
                    cSql11 = cSql11 & "" & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & ","
                End If
                If cMoneda = "ME" Then
                    cSql11 = cSql11 & "" & cAdo3("CFIMPORTE") & ","
                Else
                    cSql11 = cSql11 & "" & Val(Format(cAdo3("CFIMPORTE") / cAdo3("CFTIPCAM"), ".00")) & ","
                End If
                cSql11 = cSql11 & "'" & cMoneda & "','VTA',"
                'If cMONEDA = "MN" Then
                '        cSql11 = cSql11 & "" & Round(1 / cAdo3("CFTIPCAM"), 6) & ","
                'Else
                        cSql11 = cSql11 & "" & cAdo3("CFTIPCAM") & ","
                'End If
                cSql11 = cSql11 & "'" & cAdo3("CFRUC") & "','" & cAdo3("CFNOMBRE") & "','" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                
                
                cSql22 = "Insert Into ContCab (SUBDIAR_CODIGO,CMOV_VENTA,CMOV_MONED,CMOV_CONVE,CMOV_TIPCA,"
                If cAdo3("CFDH") = "H" Then
                    cSql22 = cSql22 & "CMOV_HABER , CMOV_HABUS,CMOV_PER,CMOV_NUM,CMOV_FECHA)"
                Else
                    cSql22 = cSql22 & "CMOV_DEBE,CMOV_DEBUS,CMOV_PER,CMOV_NUM,CMOV_FECHA)"
                End If
                cSql22 = cSql22 & "  Values ('" & cSubdiario & "',TRUE,'" & cMoneda & "','VTA',"
                'If cMONEDA = "MN" Then
                '        cSql22 = cSql22 & "" & Round(1 / cAdo3("CFTIPCAM"), 6) & ","
                'Else
                        cSql22 = cSql22 & "" & cAdo3("CFTIPCAM") & ","
                'End If
                        
                If cAdo3("CFDH") = "H" Then
                        If cMoneda = "MN" Then
                                cSql22 = cSql22 & "" & cAdo3("CFIMPORTE") & "," & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#)"
                        Else
                                cSql22 = cSql22 & "" & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#)"
                        End If
                ElseIf cAdo3("CFDH") = "D" Then
                        If cMoneda = "MN" Then
                                cSql22 = cSql22 & "" & cAdo3("CFIMPORTE") & "," & Val(Format(cAdo3("CFIMPORTE") / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#)"
                        Else
                                cSql22 = cSql22 & "" & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#)"
                        End If
                End If
                
                nTra = 1
                cConexCom.BeginTrans
                cConexCom.Execute cSql11
                cConexCom.Execute cSql22
                cConexCom.CommitTrans
                nTra = 0
                
                If Trim(cAsiento) = "N" Then
                        For nI = 1 To 3
                                cSql33 = "Insert Into CONTDET (SUBDIAR_CODIGO,DMOV_SECUE,DMOV_VENTA,DMOV_CUENT,DMOV_DOCUM,DMOV_ANEXO,"
                                cSql33 = cSql33 & "DMOV_DEBE,DMOV_DEBUS,DMOV_HABER,DMOV_HABUS,DMOV_PER,DMOV_NUM,DMOV_FECHA,DMOV_FECDC) VALUES ('" & cSubdiario & "','" & Format(nI, "0000") & "',TRUE,"
                                If nI = 1 Then
                                    cSql33 = cSql33 & "'" & cCuentaCliente & "',"
                                ElseIf nI = 2 Then
                                    cSql33 = cSql33 & "'" & cCuentaTributo & "',"
                                ElseIf nI = 3 Then
                                    cSql33 = cSql33 & "'" & cCuentaVentas & "',"
                                End If
                                cSql33 = cSql33 & "'" & cTipo & cNumeroSerie & cNumero & "','" & "02" & cAdo3("CFCODCLI") & "' ,"
                                If cAdo3("CFDH") = "H" Then
                                        If cMoneda = "MN" Then
                                                If nI = 1 Then
                                                        cSql33 = cSql33 & "0,0," & cAdo3("CFIMPORTE") & "," & Val(Format(cAdo3("CFIMPORTE") / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                ElseIf nI = 2 Then
                                                        cSql33 = cSql33 & "" & cAdo3("CFIGV") & "," & Val(Format(cAdo3("CFIGV") / cAdo3("CFTIPCAM"), ".00")) & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                Else
                                                        cSql33 = cSql33 & "" & Val(Format(cAdo3("CFIMPORTE") - cAdo3("CFIGV"), ".00")) & "," & Val(Format((cAdo3("CFIMPORTE") - cAdo3("CFIGV")) / cAdo3("CFTIPCAM"), ".00")) & ",0,0, '" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                End If
                                        Else
                                                If nI = 1 Then
                                                        cSql33 = cSql33 & "0,0," & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                ElseIf nI = 2 Then
                                                        cSql33 = cSql33 & "" & Val(Format(cAdo3("CFIGV") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIGV") & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                Else
                                                        cSql33 = cSql33 & "" & Val(Format((cAdo3("CFIMPORTE") - cAdo3("CFIGV")) * cAdo3("CFTIPCAM"), ".00")) & "," & (cAdo3("CFIMPORTE") - cAdo3("CFIGV")) & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                End If
                                        End If
                                ElseIf cAdo3("CFDH") = "D" Then
                                        If cMoneda = "MN" Then
                                                If nI = 1 Then
                                                        cSql33 = cSql33 & "" & cAdo3("CFIMPORTE") & "," & Val(Format(cAdo3("CFIMPORTE") / cAdo3("CFTIPCAM"), ".00")) & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                ElseIf nI = 2 Then
                                                        cSql33 = cSql33 & "0,0," & cAdo3("CFIGV") & "," & Val(Format(cAdo3("CFIGV") / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                Else
                                                        cSql33 = cSql33 & "0,0," & cAdo3("CFIMPORTE") - cAdo3("CFIGV") & "," & Val(Format((cAdo3("CFIMPORTE") - cAdo3("CFIGV")) / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                End If
                                        Else
                                                If nI = 1 Then
                                                        cSql33 = cSql33 & "" & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                ElseIf nI = 2 Then
                                                        cSql33 = cSql33 & "0,0," & Val(Format(cAdo3("CFIGV") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIGV") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                Else
                                                        cSql33 = cSql33 & "0,0," & Val(Format((cAdo3("CFIMPORTE") - cAdo3("CFIGV")) * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") - cAdo3("CFIGV") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                End If
                                        End If
                                End If
                                nTra = 1
                                cConexCom.BeginTrans
                                cConexCom.Execute cSql33
                                cConexCom.CommitTrans
                                nTra = 0
                        Next nI
                Else
                        Set cAdo2 = New ADODB.Recordset
                        cAdo2.Open "Select * From FacDet Where DFTD = '" & cAdo3("CFTD") & "' AND DFNUMSER = '" & cAdo3("CFNUMSER") & "' AND DFNUMDOC = '" & cAdo3("CFNUMDOC") & "'", cConexCom, adOpenStatic
                        If cAdo2.RecordCount > 0 Then
                                For nI = 1 To 2
                                        cSql33 = "Insert Into CONTDET (SUBDIAR_CODIGO,DMOV_SECUE,DMOV_VENTA,DMOV_CUENT,DMOV_DOCUM,DMOV_ANEXO,"
                                        cSql33 = cSql33 & "DMOV_DEBE,DMOV_DEBUS,DMOV_HABER,DMOV_HABUS,DMOV_PER,DMOV_NUM,DMOV_FECHA,DMOV_FECDC) VALUES ('" & cSubdiario & "','" & Format(nI, "0000") & "',TRUE,"
                                        If nI = 1 Then
                                                cSql33 = cSql33 & "'" & cCuentaCliente & "',"
                                        ElseIf nI = 2 Then
                                                cSql33 = cSql33 & "'" & cCuentaTributo & "',"
                                        End If
                                        cSql33 = cSql33 & "'" & cTipo & cNumeroSerie & cNumero & "','" & "02" & cAdo3("CFCODCLI") & "' ,"
                                        If cAdo3("CFDH") = "H" Then
                                                If cMoneda = "MN" Then
                                                        If nI = 1 Then
                                                                    cSql33 = cSql33 & "0,0," & cAdo3("CFIMPORTE") & "," & Val(Format(cAdo3("CFIMPORTE") / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        ElseIf nI = 2 Then
                                                                    cSql33 = cSql33 & "" & cAdo3("CFIGV") & "," & Val(Format(cAdo3("CFIGV") / cAdo3("CFTIPCAM"), ".00")) & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        End If
                                                Else
                                                        If nI = 1 Then
                                                                    cSql33 = cSql33 & "0,0," & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        ElseIf nI = 2 Then
                                                                    cSql33 = cSql33 & "" & Val(Format(cAdo3("CFIGV") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIGV") & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        End If
                                                End If
                                        Else
                                                If cMoneda = "MN" Then
                                                        If nI = 1 Then
                                                                    cSql33 = cSql33 & "" & cAdo3("CFIMPORTE") & "," & Val(Format(cAdo3("CFIMPORTE") / cAdo3("CFTIPCAM"), ".00")) & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        ElseIf nI = 2 Then
                                                                    cSql33 = cSql33 & "0,0," & cAdo3("CFIGV") & "," & Val(Format(cAdo3("CFIGV") / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        End If
                                                Else
                                                        If nI = 1 Then
                                                                    cSql33 = cSql33 & "" & Val(Format(cAdo3("CFIMPORTE") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIMPORTE") & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        ElseIf nI = 2 Then
                                                                    cSql33 = cSql33 & "0,0," & Val(Format(cAdo3("CFIGV") * cAdo3("CFTIPCAM"), ".00")) & "," & cAdo3("CFIGV") & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                        End If
                                                End If
                                        End If
                                        nTra = 1
                                        cConexCom.BeginTrans
                                        cConexCom.Execute cSql33
                                        cConexCom.CommitTrans
                                        nTra = 0
                                Next nI
                                Do While Not cAdo2.EOF
                                        cSql33 = "Insert Into CONTDET (SUBDIAR_CODIGO,DMOV_SECUE,DMOV_VENTA,DMOV_CUENT,DMOV_DOCUM,DMOV_ANEXO,"
                                        cSql33 = cSql33 & "DMOV_DEBE,DMOV_DEBUS,DMOV_HABER,DMOV_HABUS,DMOV_PER,DMOV_NUM,DMOV_FECHA,DMOV_FECDC) VALUES ('" & cSubdiario & "','" & Format(nI, "0000") & "',TRUE,"
                                        If nI >= 3 Then
                                                cFam = Devolver_Dato(1, cAdo2("DFCODIGO"), "MaeArt", "ACODIGO", "AFAMILIA")
                                                cCuentaVentas = Devolver_Dato(1, cFam, "Familia", "FAM_CODIGO", "FAM_CTA")
                                                If Trim(cCuentaVentas) = "" Then
                                                        MsgBox "No hay Cuenta Contable Asignada a la Familia de Artículos", vbInformation, "Información"
                                                        GoTo ErrGen
                                                End If
                                                
                                                cSql33 = cSql33 & "'" & cCuentaVentas & "',"
                                                nTotDet = 0
                                                If cAdo2("DFARTIGV") Then
                                                        If cMoneda = "MN" Then
                                                                nTotDet = cAdo2("DFIMPMN")
                                                        Else
                                                                nTotDet = cAdo2("DFIMPUS")
                                                        End If
                                                Else
                                                         If cMoneda = "MN" Then
                                                                nTotDet = cAdo2("DFIMPMN") - cAdo2("DFIGV")
                                                        Else
                                                                nTotDet = cAdo2("DFIMPUS") - cAdo2("DFIGV")
                                                        End If
                                                End If
                                        End If
                                        cSql33 = cSql33 & "'" & cTipo & cNumeroSerie & cNumero & "','" & "02" & cAdo3("CFCODCLI") & "' ,"
                                        If cAdo3("CFDH") = "H" Then
                                                If cMoneda = "MN" Then
                                                        cSql33 = cSql33 & "" & nTotDet & "," & Val(Format(nTotDet / cAdo3("CFTIPCAM"), ".00")) & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                Else
                                                        cSql33 = cSql33 & "" & Val(Format(nTotDet * cAdo3("CFTIPCAM"), ".00")) & "," & nTotDet & ",0,0,'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                End If
                                        Else
                                                If cMoneda = "MN" Then
                                                        cSql33 = cSql33 & "0,0," & nTotDet & "," & Val(Format(nTotDet / cAdo3("CFTIPCAM"), ".00")) & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                Else
                                                        cSql33 = cSql33 & "0,0," & Val(Format(nTotDet * cAdo3("CFTIPCAM"), ".00")) & "," & nTotDet & ",'" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#)"
                                                End If
                                        End If
                                        nTra = 1
                                        cConexCom.BeginTrans
                                        cConexCom.Execute cSql33
                                        cConexCom.CommitTrans
                                        nTra = 0
                                        nI = nI + 1
                                        cAdo2.MoveNext
                                        If cAdo2.EOF Then Exit Do
                                Loop
                        End If
                        cAdo2.Close
                End If
                
                nTra = 1
                cConexCom.BeginTrans
                cConexCom.Execute "Update FacCab Set CFSUBDIAR = '" & cSubdiario & "' Where  CFTD = '" & cAdo3("CFTD") & "' AND CFNUMSER = '" & cAdo3("CFNUMSER") & "' AND CFNUMDOC = '" & cAdo3("CFNUMDOC") & "' "
                cConexCom.CommitTrans
                nTra = 0
        Else
                cSql11 = "Insert Into ContVentas (CO_C_Cuent,CO_C_Mes,CO_C_Subdi,CO_C_CLIEN,CO_C_TpDoc,CO_C_Docum,CO_C_TpDRf,CO_C_DCRef,CO_N_DCTO,CO_N_DCTUS,CO_N_Igv,"
                cSql11 = cSql11 & "CO_N_IgvUs,CO_N_Tasa,CO_N_Monto,CO_N_MtoUS,CO_C_Moned,CO_C_Conve,CO_N_TipCa,CO_C_Ruc,CO_A_Razon,CO_PER,CO_NUM,CO_D_FECHA,CO_D_FECDC,CO_L_ANULA) Values  "
                cSql11 = cSql11 & "('" & cCuentaCliente & "','" & Format(Month(Date), "00") & "','" & cSubdiario & "','" & cAdo3("CFCODCLI") & "','" & cAdo3("cftd") & "','" & cAdo3("cfnumser") & cAdo3("cfnumdoc") & "',"
                cSql11 = cSql11 & "'" & IIf(Trim(cAdo3("CFRFTD")) <> "", cAdo3("CFRFTD"), "  ") & "','" & IIf(Trim(cAdo3("CFRFNUMSER")) <> "", cAdo3("CFRFNUMSER"), "  ") & IIf(Trim(cAdo3("CFRFNUMDOC")) <> "", cAdo3("CFRFNUMDOC"), "  ") & "',"
                cSql11 = cSql11 & "0,0,0,0,0,0,0,"
                cSql11 = cSql11 & "'" & cMoneda & "','VTA',0,"
                cSql11 = cSql11 & "' ','ANULADO','" & Year(Date) & Month(Date) & "','" & cComAux & "',#" & Format(Date, "mm/dd/yyyy") & "#,#" & cAdo3("CFFECDOC") & "#,TRUE)"
                
                nTra = 1
                cConexCom.BeginTrans
                cConexCom.Execute cSql11
                cConexCom.Execute "Update FacCab Set CFSUBDIAR = '" & cSubdiario & "' Where  CFTD = '" & cAdo3("CFTD") & "' AND CFNUMSER = '" & cAdo3("CFNUMSER") & "' AND CFNUMDOC = '" & cAdo3("CFNUMDOC") & "' "
                cConexCom.CommitTrans
                nTra = 0
        End If
End If
cAdo3.Close
Else
        MsgBox "No ha definido los Parámetros de Asiento de Venta", vbInformation, "Información"
End If
Exit Sub
ErrGen:
        MsgBox err.Description
        If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Function Numera(ASUB As String) As String
Dim ADOReg As New ADODB.Recordset
Dim Fanulado As New ADODB.Recordset
Dim xAux As String

Set ADOReg = New ADODB.Recordset
Set Fanulado = New ADODB.Recordset

ADOReg.Open "SELECT MAX(CMOV_NUM) FROM CONTCAB  WHERE SUBDIAR_CODIGO='" & ASUB & "' AND CMOV_PER ='" & Year(Date) & Month(Date) & "'", cConexCom, adOpenStatic
Fanulado.Open "SELECT MAX(CO_NUM) FROM CONTVENTAS WHERE CO_C_SUBDI='" & ASUB & "' AND CO_C_Mes = '" & Format(Month(Date), "00") & "' AND CO_L_ANULA AND CO_PER = '" & Year(Date) & Month(Date) & "' ", cConexCom, adOpenStatic
If ADOReg.RecordCount <> 0 Then
     Numera = Format(Val(IIf(IsNull(ADOReg.Fields(0)), 0, ADOReg.Fields(0))) + 1, "0000")
Else
     Numera = 1
End If
If Fanulado.RecordCount <> 0 Then
    xAux = IIf(IsNull(Fanulado(0)), 0, Fanulado(0))
    If xAux >= Val(Numera) Then
            Numera = Format(xAux + 1, "0000")
    End If
End If
ADOReg.Close
Fanulado.Close
End Function

Private Sub Set_Flex1()
Init_Flex Flex1
Flex1.Rows = 1
Flex1.FormatString = "^ Selec.|Doc.|Nro. Ser.|Nro. Doc.|Fec. Doc.|Fec. Ven.|Cod. Mon.|Importe"

Flex1.ColWidth(0) = 600
Flex1.ColWidth(1) = 800
Flex1.ColWidth(2) = 1000
Flex1.ColWidth(3) = 1500
Flex1.ColWidth(4) = 1500
Flex1.ColWidth(5) = 1500
Flex1.ColWidth(6) = 800
Flex1.ColWidth(7) = 2000
End Sub

Private Sub Set_Flex2()
Init_Flex Flex2
Flex2.Rows = 1
Flex2.FormatString = "^ Selec.|Cuenta|SubDiario|Comprobante|Tipo Doc.|Nro. Doc.|Cod. Mon.|Importe|Comp.Aux."

Flex2.ColWidth(0) = 600
Flex2.ColWidth(1) = 1000
Flex2.ColWidth(2) = 1000
Flex2.ColWidth(3) = 1000
Flex2.ColWidth(4) = 1000
Flex2.ColWidth(5) = 1500
Flex2.ColWidth(6) = 1000
Flex2.ColWidth(7) = 2000
Flex2.ColWidth(8) = 10
End Sub

Private Sub Elimina(Subdi As String, Compr As String, TIPO As String, NroDoc As String)
On Error GoTo ErrEli
nTra = 1
cConexCom.BeginTrans
cConexCom.Execute "Delete From ContVentas Where CO_NUM = '" & Compr & "' and CO_C_SUBDI = '" & Subdi & "'"
cConexCom.Execute "Delete From ContCab Where SUBDIAR_CODIGO = '" & Subdi & "' And CMOV_NUM = '" & Compr & "'"
cConexCom.Execute "Delete From ContDet Where SUBDIAR_CODIGO = '" & Subdi & "' And DMOV_NUM = '" & Compr & "'"
cConexCom.Execute "Update  FacCab Set CFSUBDIAR = '' Where CFSUBDIAR = '" & Subdi & "' and CFTD = '" & TIPO & "' and CFNUMSER & CFNUMDOC = '" & NroDoc & "'"
cConexCom.CommitTrans
nTra = 0
Exit Sub
ErrEli:
        MsgBox err.Description
        If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub LimpiaLabel()
LbSubdiario = "": LbComprobante = "": LbFecha = "": LbCliente = "": LbFecDoc = ""
LbDocumento = "": LbIgv = "0.00": LbTasa = "0.00": LbTotal = "0.00": LbConversion = "": LbCambio = "0.00"
LbDebe = "0.00": LbHaber = "0.00": LbDiferencia = "0.00"
End Sub

Private Sub ADOConectar()
Set cVGDBT = New ADODB.Connection
With cVGDBT 'para Movimientos
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.3.51"
        .ConnectionString = "Data Source=" & cRutPath & "\BDCONT" & Year(Date) & ".MDB"
        .Open
End With
End Sub

Public Function ANumeracion(ASUB As String) As String
Dim ADOReg As ADODB.Recordset, Fanulado As ADODB.Recordset
Set ADOReg = New ADODB.Recordset
Set Fanulado = New ADODB.Recordset
Dim xAux As Integer

ADOReg.Open "SELECT MAX(CMOV_C_COMPR) FROM CABMOV" & Format(Month(Date), "00") & " WHERE SUBDIAR_CODIGO='" & ASUB & "' AND MONTH(CMOV_FECHA)=" & Month(Date), cVGDBT, adOpenStatic
Fanulado.Open "SELECT MAX(CO_C_COMPR) FROM VENTAS WHERE CO_C_SUBDI='" & ASUB & "' AND CO_C_Mes = '" & Format(Month(Date), "00") & "' AND CO_L_ANULA", cVGDBT, adOpenStatic
If ADOReg.RecordCount <> 0 Then
        ANumeracion = Format(Val(IIf(IsNull(ADOReg.Fields(0)), 0, ADOReg.Fields(0))) + 1, "0000")
Else
        ANumeracion = 1
End If
If Fanulado.RecordCount <> 0 Then
        xAux = IIf(IsNull(Fanulado(0)), 0, Fanulado(0))
        If xAux >= Val(ANumeracion) Then
                ANumeracion = Format(xAux + 1, "0000")
        End If
End If
ADOReg.Close
Fanulado.Close
End Function

Private Sub Enviar(Subd As String, cComp As String, TIPO As String, Ndoc As String, cMond As String)
Dim cSqlE1 As String
Dim cSqlE2 As String
Dim cSqlE3 As String
Dim cSqlE4 As String
Dim cSqlE5 As String
Dim cSqlE6 As String
Dim cSqlE7 As String
Dim cAnulado As String

Dim cNUMDOC As String

On Error GoTo ErrEnv

cAnulado = Devolver_Dato(1, TIPO, "FacCab", "CFTD", "CFESTADO", Ndoc, "CFNUMSER & CFNUMDOC")

cNUMDOC = ANumeracion(cSubdiario)

If Trim(cAnulado) <> "A" Then

        cSqlE1 = "Insert Into Ventas (CO_C_CUENT,CO_C_MES,CO_C_SUBDI,CO_C_COMPR,CO_D_FECHA,CO_C_CLIEN,CO_C_TPDOC,CO_C_DOCUM,CO_D_FECDC,"
        cSqlE1 = cSqlE1 & "CO_C_TPDRF,CO_C_DCREF,CO_C_MONED,CO_N_DCTO,CO_N_DCTUS,CO_N_IGV,CO_N_IGVUS,CO_N_TASA,CO_N_MONTO,CO_N_MTOUS,"
        cSqlE1 = cSqlE1 & "CO_C_CONVE,CO_N_TIPCA,CO_L_ANULA,CO_C_RUC,CO_A_RAZON)  SELECT CO_C_CUENT,CO_C_MES,CO_C_SUBDI,'" & cNUMDOC & "',CO_D_FECHA,CO_C_CLIEN,CO_C_TPDOC,CO_C_DOCUM,CO_D_FECDC,"
        cSqlE1 = cSqlE1 & "CO_C_TPDRF,CO_C_DCREF,CO_C_MONED,CO_N_DCTO,CO_N_DCTUS,CO_N_IGV,CO_N_IGVUS,CO_N_TASA,CO_N_MONTO,CO_N_MTOUS,"
        cSqlE1 = cSqlE1 & "CO_C_CONVE,iif('" & cMond & "' ='MN',val(format(1/ CO_N_TIPCA,'.000000')),CO_N_TIPCA),CO_L_ANULA,CO_C_RUC,CO_A_RAZON  FROM  " & cRutPath & "BDCOMUN.MDB.CONTVENTAS WHERE CO_C_SUBDI = '" & Subd & "' and CO_NUM = '" & cComp & "'"
        
        cSqlE2 = "Insert Into CabMov" & Format(Month(Date), "00") & " (SUBDIAR_CODIGO,CMOV_C_COMPR,CMOV_FECHA,CMOV_MONED,CMOV_CONVE,CMOV_TIPCA,CMOV_DEBE,CMOV_HABER,CMOV_DEBUS,CMOV_HABUS,"
        cSqlE2 = cSqlE2 & "CMOV_VENTA) Select SUBDIAR_CODIGO,'" & cNUMDOC & "',CMOV_FECHA,CMOV_MONED,CMOV_CONVE,iif('" & cMond & "' ='MN',val(format(1/CMOV_TIPCA,'.000000')),CMOV_TIPCA),CMOV_DEBE,CMOV_HABER,CMOV_DEBUS,CMOV_HABUS,"
        cSqlE2 = cSqlE2 & "CMOV_VENTA From  " & cRutPath & "BDCOMUN.MDB.CONTCAB WHERE SUBDIAR_CODIGO = '" & Subd & "' and CMOV_NUM = '" & cComp & "'"
        
        cSqlE3 = "Insert Into DetMov" & Format(Month(Date), "00") & " (SUBDIAR_CODIGO,DMOV_SECUE,DMOV_C_COMPR,DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO,DMOV_DOCUM,DMOV_FECDC,"
        cSqlE3 = cSqlE3 & "DMOV_DEBE,DMOV_HABER,DMOV_DEBUS,DMOV_HABUS,DMOV_VENTA) Select SUBDIAR_CODIGO,DMOV_SECUE,'" & cNUMDOC & "',DMOV_FECHA,DMOV_CUENT,DMOV_ANEXO,DMOV_DOCUM,DMOV_FECDC,"
        cSqlE3 = cSqlE3 & "DMOV_DEBE,DMOV_HABER,DMOV_DEBUS,DMOV_HABUS,DMOV_VENTA From  " & cRutPath & "BDCOMUN.MDB.CONTDET  WHERE SUBDIAR_CODIGO = '" & Subd & "' and DMOV_NUM = '" & cComp & "'"
        
        cSqlE4 = "Update CONTVENTAS Set CO_C_COMPR = '" & cNUMDOC & "'   WHERE CO_C_SUBDI = '" & Subd & "' and CO_NUM = '" & cComp & "'"
        cSqlE5 = "Update CONTCAB Set CMOV_C_COMPR = '" & cNUMDOC & "'  WHERE SUBDIAR_CODIGO = '" & Subd & "' and CMOV_NUM = '" & cComp & "'"
        cSqlE6 = "Update CONTDET  set DMOV_C_COMPR =  '" & cNUMDOC & "'   WHERE SUBDIAR_CODIGO = '" & Subd & "' and DMOV_NUM = '" & cComp & "'"
        cSqlE7 = "Update FacCab Set CFCOMPROB =  '" & cNUMDOC & "'  Where CFTD = '" & TIPO & "' and CFNUMSER & CFNUMDOC = '" & Ndoc & "'   "
        
        
        nTra2 = 1
        cVGDBT.BeginTrans
        cVGDBT.Execute cSqlE1
        cVGDBT.Execute cSqlE2
        cVGDBT.Execute cSqlE3
        cVGDBT.CommitTrans
        nTra2 = 0
        
        nTra = 1
        cConexCom.BeginTrans
        cConexCom.Execute cSqlE4
        cConexCom.Execute cSqlE5
        cConexCom.Execute cSqlE6
        cConexCom.Execute cSqlE7
        cConexCom.CommitTrans
        nTra = 0
Else
        cSqlE1 = "Insert Into Ventas (CO_C_CUENT,CO_C_MES,CO_C_SUBDI,CO_C_COMPR,CO_D_FECHA,CO_C_CLIEN,CO_C_TPDOC,CO_C_DOCUM,CO_D_FECDC,"
        cSqlE1 = cSqlE1 & "CO_C_TPDRF,CO_C_DCREF,CO_C_MONED,CO_N_DCTO,CO_N_DCTUS,CO_N_IGV,CO_N_IGVUS,CO_N_TASA,CO_N_MONTO,CO_N_MTOUS,"
        cSqlE1 = cSqlE1 & "CO_C_CONVE,CO_N_TIPCA,CO_L_ANULA,CO_C_RUC,CO_A_RAZON)  SELECT CO_C_CUENT,CO_C_MES,CO_C_SUBDI,'" & cNUMDOC & "',CO_D_FECHA,CO_C_CLIEN,CO_C_TPDOC,CO_C_DOCUM,CO_D_FECDC,"
        cSqlE1 = cSqlE1 & "CO_C_TPDRF,CO_C_DCREF,CO_C_MONED,0,0,0,0,0,0,0,"
        cSqlE1 = cSqlE1 & "CO_C_CONVE,0,CO_L_ANULA,CO_C_RUC,CO_A_RAZON  FROM  " & cRutPath & "BDCOMUN.MDB.CONTVENTAS WHERE CO_C_SUBDI = '" & Subd & "' and CO_NUM = '" & cComp & "'"
        
        nTra2 = 1
        cVGDBT.BeginTrans
        cVGDBT.Execute cSqlE1
        cVGDBT.CommitTrans
        nTra2 = 0
        
        cSqlE4 = "Update CONTVENTAS Set CO_C_COMPR = '" & cNUMDOC & "'   WHERE CO_C_SUBDI = '" & Subd & "' and CO_NUM = '" & cComp & "'"
        
        nTra = 1
        cConexCom.BeginTrans
        cConexCom.Execute cSqlE4
        cConexCom.CommitTrans
        nTra = 0
End If

Exit Sub
ErrEnv:
       MsgBox err.Description
       If nTra2 = 1 Then cVGDBT.RollbackTrans
       If nTra = 1 Then cConexCom.RollbackTrans
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If cVGDBT.State = 1 Then cVGDBT.Close
End Sub
