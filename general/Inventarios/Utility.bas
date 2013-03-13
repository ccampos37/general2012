Attribute VB_Name = "UTILITY"
Option Explicit

'***************************************************
'  Declaración API para Escribir y Leer un (*.INI)
'***************************************************

Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
        
Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

        
Public Function sGetIni(sIniFile As String, sSection As String, sKey As String, sDefault As String) As String
 Dim sTemp As String * 256
 Dim nLength As Integer
 sTemp = Space$(256)
 nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sIniFile)
 sGetIni = Left$(sTemp, nLength)
End Function

Public Sub WriteIni(sIniFile As String, sSection As String, sKey _
                    As String, sValue As String)
 Dim sTemp As String
 Dim n As Integer
 
 sTemp = sValue
 For n = 1 To Len(sValue)
  If Mid$(sValue, n, 1) = vbCr Or Mid$(sValue, n, 1) = vbLf Then
   Mid$(sValue, n) = " "
  End If
 Next n
 n = WritePrivateProfileString(sSection, sKey, sTemp, sIniFile)
End Sub

Public Function numero(Number) As String
   Dim aValor As Double
   If IsNull(Number) Or Len(Trim(Number)) = 0 Then
     numero = 0
   Else
     numero = Number
   End If
End Function

Public Function ComputerName() As Variant
    Dim sName As String
    Dim iRetVal As Long
    Dim ipos As Integer
    sName = Space$(255)
    iRetVal = GetComputerName(sName, 255&)
    If iRetVal = 0 Then
      ComputerName = ""
      Exit Function
    End If
    ipos = InStr(sName, Chr$(0))
    ComputerName = Left$(sName, ipos - 1)
End Function

Sub ImpresionRptCad(Reporte As Crystal.CrystalReport, cNombreReporte As String, PFormulas(), Param(), Optional ORDEN As String, Optional titulo As String, Optional Seleccion As String)
Dim I As Integer
Dim sServer As String
Dim sBase As String
Dim sUsuario As String
Dim sPwd As String
Dim sRutaReportes As String
On Error GoTo procImpresionRptError
    'Leo ini  sección de proc(s) almacenados
    sServer = sGetIni(App.Path & "\wenco.ini", "Bstore", "dserver", "?")
    If Trim(sServer) = "?" Then sServer = "(local)"
        
    sBase = sGetIni(App.Path & "\wenco.ini", "Bstore", "dbase", "?")
    If Trim(sBase) = "?" Then sBase = "BDMAIN"
    
    sUsuario = sGetIni(App.Path & "\wenco.ini", "Bstore", "duser", "?")
    If Trim(sUsuario) = "?" Then sUsuario = "sa"
    
    sPwd = sGetIni(App.Path & "\wenco.ini", "Bstore", "dpass", "?")
    If Trim(sPwd) = "?" Then sPwd = ""

    'Leo la ruta en donde se encuentra los archivos de reportes
    sRutaReportes = sGetIni(App.Path & "\wenco.ini", "CONFIG", "rpt ", "?")
    If Trim(sRutaReportes) = "?" Then sRutaReportes = "C:\WENCO\REPORTES\"
    Screen.MousePointer = 11
    With Reporte
        .Reset
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = titulo
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .ReportFileName = sRutaReportes & cNombreReporte
'        .LogOnServer "pdssql.dll", sServer, sBase, sUsuario, sPwd
        .Connect = "DSN=" & sServer & ";DSQ=" & sBase & ";UID=" & sUsuario & ";PWD=" & sPwd
        If UBound(PFormulas) > 0 Then
            For I = 0 To UBound(PFormulas) - 1
                .formulas(2 + I) = PFormulas(I)
            Next
        End If
        .DiscardSavedData = True
        If UBound(Param) > 0 Then
            For I = 0 To UBound(Param) - 1
                .StoredProcParam(I) = Param(I)
            Next
        End If
   '     If orden <> "" Then Call CrystOrden(MDIPrincipal.CryRptProc, orden)
        If Seleccion <> "" Then .SelectionFormula = Seleccion
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
    Exit Sub
    
procImpresionRptError:
  If Err.Number = 9 Then Resume Next
  Screen.MousePointer = 1
  MsgBox "Error inesperado: " & Err.Number & "  " & Err.Description, vbExclamation
End Sub
Sub ImpresionRptdefault(Reporte As Crystal.CrystalReport, titulo)
Dim I As Integer
Dim sServer As String
Dim sBase As String
Dim sUsuario As String
Dim sPwd As String
Dim sRutaReportes As String
Screen.MousePointer = 11
With Reporte
     .Reset
     .Destination = crptToWindow
     .WindowState = crptMaximized
     .WindowTitle = titulo
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
End With
End Sub

Sub Captura_error()
    If Err.Number <> 0 Then
        MsgBox Str(Err.Number) + "," + Err.Description, vbCritical, "Mensaje"
        
    End If
End Sub

Public Function NUMLET(num As String)
Dim cLET As String
Dim cWork As String
Dim cUNIDAD As String
Dim cDECENA As String
Dim cCENTENA As String
Dim nMODULUS As Integer
Dim nI As Integer
Dim nK As Integer
Dim Lit1 As String
Dim Lit2 As String
Dim Lit3 As String
Dim Lit4 As String
Dim Lit5 As String
Lit1 = "Uno    Doc    Trec   Cuatroc  Quin   Seisc  Setec  Ochoc  Novec  "
Lit2 = "Diez     Veinte   Treinta  Cuarenta CincuentaSesenta  Setenta  Ochenta  Noventa  "
Lit3 = "Once      Doce      Trece     Catorce   Quince    Dieciseis DiecisieteDieciocho Diecinueve"
Lit4 = "Uno   Dos   Tres  CuatroCinco Seis  Siete Ocho  Nueve "
Lit5 = "Millon    Billon    Trillon   CuatrillonQuintillon"
'Proceso Input = Num , Output = Let

cLET = ""

'Dim NUM As Double
'NUM = Val(NUMx)

If num > 0.99 Then
    'Separa los Enteros en una Cadena de Caracteres
     If InStr(1, Trim(Str(num)), ".", 0) > 0 Then
        cWork = Mid(Trim(Str(num)), 1, InStr(1, Trim(Str(num)), ".", 0) - 1)
     Else
        cWork = Str(num)
     End If
     nMODULUS = Int(Len(Trim(cWork)) / 3)
     nMODULUS = Len(Trim(cWork)) - (nMODULUS * 3)
     
     If nMODULUS > 0 Then
        cWork = String(3 - nMODULUS, "0") & Trim(cWork)
     End If
     
     nK = (Len(Trim(cWork)) / 3) - 1
    'Procesa de Mil en Mil
     nI = 1
     Do While nI < Len(Trim(cWork)) - 1
        cCENTENA = Mid(Trim(cWork), nI, 1)
        cDECENA = Mid(Trim(cWork), nI + 1, 1)
        cUNIDAD = Mid(Trim(cWork), nI + 2, 1)
        'Centenas
        If cCENTENA <> "0" Then
            If cCENTENA = "1" Then
                cLET = cLET & "Cien "
                If cDECENA <> "0" Or cUNIDAD <> "0" Then
                    cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "to "
                End If
            Else
                cLET = cLET & Trim(Mid(Lit1, ((Val(cCENTENA) - 1) * 7) + 1, 7)) & "ientos "
            End If
        End If
        'Decenas
        If cDECENA <> "0" Then
            If cDECENA = "1" And cUNIDAD <> "0" Then
                If ((Val(cUNIDAD) - 1) * 10) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit3, ((Val(cUNIDAD) - 1) * 10) + 1, 10))
            Else
                If ((Val(cDECENA) - 1) * 9) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit2, ((Val(cDECENA) - 1) * 9) + 1, 9))
            End If
        End If
        'Unidades
        If cUNIDAD <> "0" Then
            If cDECENA > "1" Then
                cLET = Mid(cLET, 1, (Len(cLET) - 1)) & "i"
                If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + LCase(Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6)))
            Else
                If cDECENA < "1" Then
                    If ((Val(cUNIDAD) - 1) * 6) + 1 > 0 Then cLET = cLET + Trim(Mid(Lit4, ((Val(cUNIDAD) - 1) * 6) + 1, 6))
                End If
            End If
        End If
        cLET = cLET & " "
        'Pone Miles o Millones
        If nK > 0 Then
            If cCENTENA & cDECENA & cUNIDAD = "001" Then
                cLET = Mid(cLET, 1, Len(cLET) - 2) & " "
            End If
            nMODULUS = Int(nK / 2)
            nMODULUS = nK - (nMODULUS * 2)
            If nMODULUS = 0 Then
                cLET = cLET + Trim(Mid(Lit5, (((nK / 2) - 1) * 10) + 1, 10))
                If cCENTENA & cDECENA & cUNIDAD = "001" Or num > 1999999 Then
                    cLET = cLET & "es "
                Else
                    cLET = cLET & " "
                End If
            Else
                If cCENTENA & cDECENA & cUNIDAD > "000" Then
                    cLET = cLET & "Mil "
                End If
            End If
            nK = nK - 1
        End If
        nI = nI + 3
    Loop
    cLET = cLET & "con "
End If
If InStr(1, Trim(Str(num)), ".", 0) > 0 Then
    cLET = cLET + Mid(Trim(Str(num)), InStr(1, Trim(Str(num)), ".", 0) + 1, 2) & "/100" & " "
Else
    cLET = cLET + "00/100" & " "
End If
NUMLET = cLET
End Function

Public Function Seguir(MBox As Object, ntecla As Integer)
    If ntecla = 13 Then
        SendKeys "{tab}"
    End If
End Function
