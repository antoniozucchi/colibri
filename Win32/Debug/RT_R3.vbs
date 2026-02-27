Option Explicit

Dim SapGuiAuto, application, connection, session
Dim LOGFILE

LOGFILE = "C:\Users\kbeb\Documents\Colibri\sap_log_rts.txt"

' =========================
' LOG robusto
' =========================
Sub LogLine(ByVal line)
  On Error Resume Next
  Dim fso, folderPath, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  folderPath = fso.GetParentFolderName(LOGFILE)
  If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
  Set f = fso.OpenTextFile(LOGFILE, 8, True)
  If Err.Number <> 0 Then
    WScript.Echo "ERRO ao abrir log: " & Err.Number & " - " & Err.Description & vbCrLf & LOGFILE
    Err.Clear
    Exit Sub
  End If
  f.WriteLine Now & " | " & line
  f.Close
  On Error GoTo 0
End Sub

Function GetStatusText(sess)
  On Error Resume Next
  GetStatusText = sess.findById("wnd[0]/sbar").Text
  On Error GoTo 0
End Function

Function GetStatusType(sess)
  On Error Resume Next
  GetStatusType = sess.findById("wnd[0]/sbar").MessageType ' E/W/S
  On Error GoTo 0
End Function

Sub FechaPopupSeExistir(sess)
  On Error Resume Next
  If sess.Children.Count > 1 Then
    sess.findById("wnd[1]/tbar[0]/btn[0]").press
  End If
  On Error GoTo 0
End Sub

' =========================
' Resolve "Campo Contrato obrigatório." (seleciona 2ª linha do matchcode)
' =========================
Function ResolverContratoObrigatorio(ByVal idExec, ByVal idRT, ByVal linha, ByVal contratoFix)
  ResolverContratoObrigatorio = False
  Dim pathContrato, st, stType
  pathContrato = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                 "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CONTRATO[19,0]"
  LogLine idExec & "|" & idRT & "|" & linha & "|INFO_CONTRATO_OBRIGATORIO|Campo Contrato obrigatório."
  If Len(Trim(contratoFix)) > 0 Then
    session.findById(pathContrato).text = contratoFix
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 300
  Else
    session.findById(pathContrato).setFocus
    session.findById("wnd[0]").sendVKey 4
    WScript.Sleep 300
    session.findById("wnd[1]/usr/lbl[1,5]").setFocus
    session.findById("wnd[1]/usr/lbl[1,5]").caretPosition = 0
    session.findById("wnd[1]").sendVKey 14 ' seta para baixo = 2ª linha
    WScript.Sleep 100
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    WScript.Sleep 250
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 300
  End If
  st = GetStatusText(session)
  stType = GetStatusType(session)
  If stType = "E" Then
    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_CONTRATO|" & st
    ResolverContratoObrigatorio = False
  Else
    LogLine idExec & "|" & idRT & "|" & linha & "|OK_CONTRATO|" & st
    ResolverContratoObrigatorio = True
  End If
End Function

' =========================
' Processa 1 RT (CRIAR/MODIFICAR) por string acao
' =========================
Function ProcessarRT(ByVal idExec, ByVal idRT, ByVal linha, ByVal acao, ByVal rtNumero, ByVal tipoRT, _
                     ByVal pernr, ByVal contratoFix, ByVal qmtxt, ByVal requis, ByVal pessoaContato, ByVal foneContato, _
                     ByVal iwerk, ByVal ingrp, ByVal codTrasg, ByVal classePas, _
                     ByVal origem, ByVal destino, ByVal dtViagem, ByVal hViagem, _
                     ByVal ccusto, ByVal cobertura, ByVal salvar)
  ProcessarRT = False
  On Error Resume Next
  Dim st, stType, pathPernr
  LogLine idExec & "|" & idRT & "|" & linha & "|INICIO_ITEM|ACAO=" & acao & "|RT=" & rtNumero & "|PERNR=" & pernr

  If UCase(acao) = "MODIFICAR" Then
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 250
    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero
    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").caretPosition = Len(rtNumero)
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 250
  Else
    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw51"
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 250
    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = tipoRT
    session.findById("wnd[0]").sendVKey 0
    WScript.Sleep 200
  End If

  session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/subNOTIF_TYPE:SAPLIQS0:1061/txtVIQMEL-QMTXT").text = qmtxt
  session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/subNOTIF_TYPE:SAPLIQS0:1061/txtVIQMEL-QMTXT").caretPosition = 24
  session.findById("wnd[0]").sendVKey 0
  WScript.Sleep 150

  ' ======== COBERTURA (aplica quando cobertura=1/TRUE) ========
  If (Trim(cobertura) = "1") Or (UCase(Trim(cobertura)) = "TRUE") Then
    On Error Resume Next
    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/btnANWENDERSTATUS").press
    WScript.Sleep 150
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").selected = True
    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").setFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    On Error GoTo 0
    WScript.Sleep 200
  End If

  session.findById("wnd[0]/shellcont/shell").selectItem "010","Column01"
  session.findById("wnd[0]/shellcont/shell").clickLink "010","Column01"
  session.findById("wnd[0]/usr/ctxtYSCSRTCB-REQUIS").text = requis
  session.findById("wnd[0]/usr/txtYSCSRTCB-PESSOA_CONTATO").text = pessoaContato
  session.findById("wnd[0]/usr/txtYSCSRTCB-TELEFONE_CONTATO").text = foneContato
  session.findById("wnd[0]/usr/ctxtVIQMEL-IWERK").text = iwerk
  session.findById("wnd[0]/usr/ctxtVIQMEL-INGRP").text = ingrp
  session.findById("wnd[0]").sendVKey 0

  pathPernr = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
            "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-PERNR[3,0]"
  session.findById(pathPernr).text = pernr
  session.findById("wnd[0]").sendVKey 0
  WScript.Sleep 300

  st = GetStatusText(session)
  stType = GetStatusType(session)
  If stType = "E" Then
    If InStr(1, st, "não está ativo", vbTextCompare) > 0 Then
      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_PAX|" & st
      FechaPopupSeExistir session
      Exit Function
    End If
    If InStr(1, st, "Contrato obrigatório", vbTextCompare) > 0 Then
      On Error GoTo 0
      If Not ResolverContratoObrigatorio(idExec, idRT, linha, contratoFix) Then
        FechaPopupSeExistir session
        Exit Function
      End If
      On Error Resume Next
    Else
      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_PERNR|" & st
      FechaPopupSeExistir session
      Exit Function
    End If
  End If

  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-CODTRASG[8,0]").text = codTrasg
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CLASSEPAS[9,0]").text = classePas
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_ORIGEM[28,0]").text = origem
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_DESTINO[38,0]").text = destino
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-DTVIAGEM[52,0]").text = dtViagem
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-HVIAGEM[53,0]").text = hViagem
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _
                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CCUSTO[63,0]").Text = ccusto

  If salvar Then
    session.findById("wnd[0]").sendVKey 11
    WScript.Sleep 500
    session.findById("wnd[0]").sendVKey 11
    WScript.Sleep 500
  End If

  WScript.Sleep 400
  st = GetStatusText(session)
  If Len(Trim(st)) = 0 Then
    WScript.Sleep 600
    st = GetStatusText(session)
  End If

  stType = GetStatusType(session)
  If stType = "E" Then
    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_FINAL|" & st
    FechaPopupSeExistir session
    Exit Function
  End If

  LogLine idExec & "|" & idRT & "|" & linha & "|OK|" & st
  ProcessarRT = True
  On Error GoTo 0
End Function

LogLine "INICIO_LOTE"
LogLine "LOG_USADO|" & LOGFILE

If Not IsObject(application) Then
  Set SapGuiAuto  = GetObject("SAPGUI")
  Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
  Set connection = application.Children(0)
End If
If Not IsObject(session) Then
  Set session = connection.Children(0)
End If
session.findById("wnd[0]").maximize

Dim itens, i, okCount, errCount, salvar
salvar = True

itens = Array( _
  Array("1126119","45","1","CRIAR","","R3","9623477","","TMIB: ANTONIO AUGUSTO ZUCHI LEITE","9623477","Antônio Zucchi","(21) 99830-5159","2620","T31","M","TT","TMIB","PCM-9","25.02.2026","07:00","ES38PZOC5Z","1"))

okCount = 0
errCount = 0

For i = 0 To UBound(itens)
  Dim a, n
  a = itens(i)
  n = UBound(a)
  If n < 21 Then
    LogLine "ERRO_ITEM_TAMANHO|" & i & "|UBOUND=" & n
    errCount = errCount + 1
  Else
    If ProcessarRT(a(0),a(1),a(2),a(3),a(4),a(5),a(6),a(7),a(8),a(9), _
                   a(10),a(11),a(12),a(13),a(14),a(15),a(16),a(17),a(18),a(19), _
                   a(20),a(21), salvar) Then
      okCount = okCount + 1
    Else
      errCount = errCount + 1
    End If
  End If
  WScript.Sleep 250
Next

LogLine "FIM_LOTE|OK=" & okCount & "|ERRO=" & errCount
