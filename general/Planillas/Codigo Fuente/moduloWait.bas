Attribute VB_Name = "moduloWait"
Enum progresColl
 un_coolbar = 0
 dos_coolbar = 1
End Enum
Type datosCoolbar
  progres1Min As Double
  progres2Min As Double
  progres1max As Double
  progres2max As Double
  count1 As Double
  count2 As Double
End Type
Public datosCoolbarX As datosCoolbar
Public Sub inciarCoolbarX(Optional MaxcoolBar1 As Double, Optional MaxCoolbar2 As Double, Optional xcount1 As Double, Optional xcount2 As Double)
Form1.CoolBar1.Visible = True

If xcount1 = 0 Then xcount1 = 1
If xcount2 = 0 Then xcount2 = 1
    datosCoolbarX.count1 = xcount1
    datosCoolbarX.count2 = xcount2

    datosCoolbarX.progres1Min = 0
    datosCoolbarX.progres2Min = 0
    datosCoolbarX.progres1max = MaxcoolBar1
    datosCoolbarX.progres2max = MaxCoolbar2

    
Call iniciarCoolBar1
Call iniciarCoolBar2
End Sub
Public Sub EtiquetasCoolbar1(Optional Caption_label1 As String)
    Form1.lblprimero.Caption = Caption_label1
End Sub
Public Sub EtiquetasCoolbar2(Optional Caption_label2 As String)
    Form1.lblSegundo.Caption = Caption_label2
End Sub
Public Sub iniciarCoolBar1()
    Form1.pbarCoolbar1.Min = 0
    Form1.pbarCoolbar1.Max = datosCoolbarX.progres1max
    Form1.pbarCoolbar1.Value = 0
End Sub
Public Sub iniciarCoolBar2()
    Form1.pbarCoolbar2.Min = 0
    Form1.pbarCoolbar2.Max = datosCoolbarX.progres2max
    Form1.pbarCoolbar2.Value = 0
End Sub
Public Sub CoolbarProgresPrincipal(Optional etiquetaPrincipal As String)
On Error GoTo handler
     Form1.pbarCoolbar1.Value = Form1.pbarCoolbar1.Value + datosCoolbarX.count1
     Call EtiquetasCoolbar1(etiquetaPrincipal)
     DoEvents
Exit Sub
handler:
    MsgBox ERR.Description & Chr(13) & " Sobrepaso el maximo valor"
End Sub
Public Sub CoolbarProgresSecundario(Optional etiquetaSecundaria As String)
On Error GoTo handler
     Form1.pbarCoolbar2.Value = Form1.pbarCoolbar2.Value + datosCoolbarX.count2
     Call EtiquetasCoolbar2(etiquetaSecundaria)
     DoEvents
Exit Sub
handler:
    MsgBox ERR.Description & Chr(13) & " Sobrepaso el maximo valor"
End Sub
Public Sub terminarCoolbar()
    Form1.CoolBar1.Visible = False
End Sub

