VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmRptOcEmisión 
   BackColor       =   &H000000C0&
   Caption         =   "Emisión de Ordenes de Compra"
   ClientHeight    =   8592
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11400
   ForeColor       =   &H000000FF&
   Icon            =   "FrmRptOcEmisión.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8592
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.Timer Tm 
      Interval        =   1000
      Left            =   2655
      Top             =   45
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8175
      Left            =   45
      TabIndex        =   0
      Top             =   855
      Width           =   9855
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label LblTm 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   240
      Left            =   2475
      TabIndex        =   1
      Top             =   135
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "FrmRptOcEmisión"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim Veri As Integer
Dim xMsg As String

Private Sub Form_Load()
On Error Resume Next
Dim rs As New ADODB.Recordset

CADENA = "jh_spOcImprimir " & Val(FrmOrdenCompra.LblParte.Caption) & ", " & Val(FrmOrdenCompra.Tipcom)
Set rs = DEData.CnxVg.Execute(CADENA)
Report.Database.SetDataSource rs
Report.DisplayProgressDialog = True
Report.TxtEmpresa.SetText ("Area de Sistemas")
Report.TxtEmpresa.SetText VGParametros.NomEmpresa

datos = "RUC: " & VGParametros.RucEmpresa & " Dirección: poner  Tlf:  Fax: "
Report.TxtEmprDato.SetText datos

Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
FrmImpresora.Show vbModal

If xmImpresora = "" Then Exit Sub
UseDefault = False
Report.SelectPrinter xmControlador, xmImpresora, xmPuerto
Report.PaperSize = xmTamaño
Report.PaperOrientation = xmOrientacion
Report.PrintOut
End Sub

Private Sub Tm_Timer()
On Error Resume Next
If LblTm.Caption = 0 Then CRViewer1.Refresh: LblTm.Caption = 1
End Sub
