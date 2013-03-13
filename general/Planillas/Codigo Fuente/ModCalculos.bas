Attribute VB_Name = "ModCalculos"
Option Explicit
Enum TipoCalculo
    PROMEDIO = 1
    ULTIMOVALOR = 2
    PRIMERVALOR = 3
    SUMA = 4
    MEDIA = 5
    PROMEDIOVALOR = 6
    PRIMERO = 7
    ULTIMO = 8
    MAYORVALOR = 9
    MENORVALOR = 10
    NUMERO = 11
    NSECUENCIA = 12
End Enum
Public VpTipoQuery As TipoCalculo
Type TypeLiquida
    ActVaca As Boolean
    ActGrati As Boolean
    CronoVac As Long
    CronoGrat As Long
    CANCEL As Boolean
End Type
Public RegLiquida As TypeLiquida

Public Function CALCULOCONCEPTOS(CODTRAB As String, FECHAINI As Date, FECHAFIN As Date, TIPO As TipoCalculo, CONCEPTO As String, General3 As Boolean) As Double
    Dim XNUMMES As Integer, X As Integer, NUMOCURRE As Integer, SUMATOTAL As Double
    Dim FEC1 As Date, FEC2 As Date, STRMES As String, VALOR As Double, RESULTADO As Double
    FEC1 = CDate("01/" & Month(FECHAINI) & "/" & Year(FECHAINI))
    FEC2 = CDate("01/" & Month(FECHAFIN) & "/" & Year(FECHAFIN))
    XNUMMES = DateDiff("m", FEC1, FEC2) + 1
    NUMOCURRE = 0
    RESULTADO = 0
    SUMATOTAL = 0
    Dim ACUM As String
    ACUM = ""
    If Not General3 Then
        CONCEPTO = "'" + CONCEPTO + "'"
        For X = 1 To Len(CONCEPTO)
            ACUM = ACUM + Mid(CONCEPTO, X, 1)
            If Mid(CONCEPTO, X + 1, 1) = "," Then
                ACUM = ACUM + "'"
            End If
            If Mid(CONCEPTO, X, 1) = "," Then
                ACUM = ACUM + "'"
            End If
       Next
    Else
        For X = 1 To Len(CONCEPTO)
            ACUM = ACUM + Mid(CONCEPTO, X, 1)
            If Mid(CONCEPTO, X + 1, 1) = "," Then
                ACUM = ACUM + "+"
            End If
            If Mid(CONCEPTO, X, 1) = "," Then
                ACUM = ACUM + "+"
            End If
       Next
    End If
    CONCEPTO = ACUM
    'Para Calcular la Secuencia de un valor
    Dim IX As Integer, SX As Integer
    Dim CADX As String
    Dim RSSECX As New ADODB.Recordset
    SX = 0
    RSSECX.Fields.Append "Numero", adInteger
    RSSECX.Open
    For X = 1 To XNUMMES
        STRMES = Format(Month(FEC1), "00") & Year(FEC1)
        If ExisteTabla("BOL" & STRMES) Then
            If Not General3 Then
                VALOR = Round(DevuelveValor("SELECT SUM(MONTO) AS SUMADEMONTO FROM BOL" & UCase(STRMES) & " BOL INNER JOIN MOV" & UCase(STRMES) & " MOV ON BOL.INUMBOL = MOV.INUMBOL Where (((MOV.CONCEPTO) In(" & CONCEPTO & ")) AND ((BOL.CODTRAB)='" & CODTRAB & "'))", DBSYSTEM), 2)
            Else
                VALOR = Round(DevuelveValor("SELECT SUM(" & CONCEPTO & ") AS SUMADEMONTO FROM BOL" & UCase(STRMES) & " BOL Where BOL.CODTRAB='" & CODTRAB & "'", DBSYSTEM), 2)
            End If
            Select Case TIPO
                Case PRIMERVALOR
                    If RESULTADO = 0 And VALOR <> 0 Then
                        RESULTADO = VALOR
                        Exit For
                    End If
                Case ULTIMOVALOR
                    If VALOR <> 0 Then RESULTADO = VALOR
                Case MAYORVALOR
                    If X = 1 Then RESULTADO = VALOR Else If VALOR > RESULTADO Then RESULTADO = VALOR
                Case MENORVALOR
                    If RESULTADO = 0 Then RESULTADO = VALOR Else If VALOR < RESULTADO Then RESULTADO = VALOR
                Case NSECUENCIA
                    If VALOR <> 0 Then
                        SX = SX + 1
                       Else: SX = 0
                    End If
                    If SX <> 0 Or XNUMMES = X Then
                        RSSECX.AddNew
                        RSSECX!NUMERO = SX: RSSECX.Update
                    End If
            End Select
            If VALOR <> 0 Then NUMOCURRE = NUMOCURRE + 1
            SUMATOTAL = SUMATOTAL + VALOR
        End If
        FEC1 = DateAdd("m", 1, FEC1)
    Next
    Select Case TIPO
        Case MEDIA
            RESULTADO = SUMATOTAL / 2
        Case PROMEDIO
            RESULTADO = SUMATOTAL / XNUMMES
        Case PROMEDIOVALOR
            If SUMATOTAL = 0 Then RESULTADO = 0 Else RESULTADO = SUMATOTAL / NUMOCURRE
        Case SUMA
            RESULTADO = SUMATOTAL
        Case NUMERO
            RESULTADO = NUMOCURRE
        Case NSECUENCIA
            RSSECX.Sort = "Numero Desc"
            If RSSECX.RecordCount > 0 Then
               RSSECX.MoveFirst
               RESULTADO = RSSECX!NUMERO
            End If
    End Select
    CALCULOCONCEPTOS = Round(RESULTADO, 2)
End Function

Public Sub TiempoTrans(ByVal FECHAINI As Date, ByVal FECHAFIN As Date, Optional ByRef Annos As Integer, Optional ByRef Meses As Integer, Optional ByRef Dias As Integer)
Attribute TiempoTrans.VB_Description = "Tiempo Transcurrido en años, meses y dias entre dos fechas"
    Dim XFEC As Date
    Annos = DateDiff("yyyy", FECHAINI, FECHAFIN)
    XFEC = DateAdd("yyyy", 0 - Annos, FECHAFIN)
    Meses = DateDiff("m", FECHAINI, XFEC)
    XFEC = DateAdd("m", 0 - Meses, XFEC)
    Dias = DateDiff("d", FECHAINI, XFEC)
    If Meses < 0 Then
        Annos = Annos - 1
        Meses = 12 + Meses
    End If
    If Dias < 0 Then
        Meses = Meses - 1
        Dias = 30 + Dias
    End If
End Sub

Public Function CALCULOMES(CODTRAB As String, CONCEPTO As String, Optional MES As String = "NONE", Optional FECHAINI As Date, Optional FECHAFIN As Date) As Double
    Dim XNUMMES As Integer, X As Integer, SUMATOTAL As Double
    Dim FEC1 As Date, FEC2 As Date, STRMES As String, VALOR As Double, RESULTADO As Double
    FEC1 = CDate("01/" & Month(FECHAINI) & "/" & Year(FECHAINI))
    FEC2 = CDate("01/" & Month(FECHAFIN) & "/" & Year(FECHAFIN))
    XNUMMES = DateDiff("M", FEC1, FEC2) + 1
    SUMATOTAL = 0
    If MES = "NONE" Then
        For X = 1 To XNUMMES
            STRMES = Format(Month(FEC1), "00") & Year(FEC1)
            If ExisteTabla("BOL" & STRMES) Then
                VALOR = Round(DevuelveValor("SELECT SUM(" & CONCEPTO & ") AS SUMADEMONTO FROM BOL" & STRMES & " BOL WHERE BOL.CODTRAB='" & CODTRAB & "'", DBSYSTEM), 2)
                SUMATOTAL = SUMATOTAL + VALOR
            End If
            FEC1 = DateAdd("M", 1, FEC1)
        Next
    Else
        SUMATOTAL = Round(DevuelveValor("SELECT SUM(" & CONCEPTO & ") AS SUMADEMONTO FROM BOL" & MES & " BOL WHERE BOL.CODTRAB='" & CODTRAB & "'", DBSYSTEM), 2)
    End If
    If IsNull(SUMATOTAL) Then SUMATOTAL = 0
    CALCULOMES = SUMATOTAL
End Function
Public Function CALCULOMES2(CODTRAB As String, CONCEPTO As String, Optional MES As String = "NONE", Optional FECHAINI As Date, Optional FECHAFIN As Date) As Double
    Dim XNUMMES As Integer, X As Integer, SUMATOTAL As Double
    Dim FEC1 As Date, FEC2 As Date, STRMES As String, VALOR As Double, RESULTADO As Double
    Dim RS_TAB_AUX As ADODB.Recordset
    'VERIFICA COLUMNA DE PLANILLA DEL CONCEPTO A SUMAR
    Dim XCOLUMNA  As String
    Set RS_TAB_AUX = New ADODB.Recordset
    RS_TAB_AUX.Open "SELECT * FROM CONCEPTOS WHERE CODIGO='" & CONCEPTO & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RS_TAB_AUX.RecordCount > 0 Then
        XCOLUMNA = RS_TAB_AUX!COLPLANILLA
    Else
        CALCULOMES2 = 0
        Exit Function
    End If
    FEC1 = CDate("01/" & Month(FECHAINI) & "/" & Year(FECHAINI))
    FEC2 = CDate("01/" & Month(FECHAFIN) & "/" & Year(FECHAFIN))
    XNUMMES = DateDiff("M", FEC1, FEC2) + 1
    SUMATOTAL = 0
    If MES = "NONE" Then
        For X = 1 To XNUMMES
            STRMES = Format(Month(FEC1), "00") & Year(FEC1)
                VALOR = Round(DevuelveValor("SELECT SUM(" & XCOLUMNA & ") AS SUMADEMONTO FROM PLAN2000 WHERE CODTRAB='" & CODTRAB & "' AND MES>=" & DateSQL(FEC1) & " AND MES<=" & DateSQL(FEC2), DBSYSTEM), 2)
                SUMATOTAL = SUMATOTAL + VALOR
            FEC1 = DateAdd("M", 1, FEC1)
        Next
    Else
        VALOR = Round(DevuelveValor("SELECT SUM(" & XCOLUMNA & ") AS SUMADEMONTO FROM PLAN2000 WHERE CODTRAB='" & CODTRAB & "' AND MES=" & DateSQL(STRMES) & "", DBSYSTEM), 2)
    End If
    If IsNull(SUMATOTAL) Then SUMATOTAL = 0
    CALCULOMES2 = SUMATOTAL
End Function


