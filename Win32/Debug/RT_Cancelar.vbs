Option Explicit
Dim SapGuiAuto, application, connection, session
Dim LOGFILE
LOGFILE = "C:\Dev\Projetos\Colibri\Win32\Debug\sap_log_cancel.txt"

Sub LogLine(ByVal line)
  On Error Resume Next
  Dim fso, folderPath, f
  Set fso = CreateObject("Scripting.FileSystemObject")
  folderPath = fso.GetParentFolderName(LOGFILE)
  If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
  Set f = fso.OpenTextFile(LOGFILE, 8, True)
  f.WriteLine Now & " | " & line
  f.Close
  On Error GoTo 0
End Sub

Function StatusText()
  On Error Resume Next
  StatusText = session.findById("wnd[0]/sbar").Text
  On Error GoTo 0
End Function

Function StatusType()
  On Error Resume Next
  StatusType = session.findById("wnd[0]/sbar").MessageType
  On Error GoTo 0
End Function

Sub FechaPopupSeExistir()
  On Error Resume Next
  If session.Children.Count > 1 Then session.findById("wnd[1]/tbar[0]/btn[0]").press
  On Error GoTo 0
End Sub

Sub CancelarRT(ByVal idProgRT, ByVal rtNumero)
  On Error Resume Next
  Dim st, tp
  LogLine idProgRT & "|CANCELAR|INICIO|RT=" & rtNumero

  session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"
  session.findById("wnd[0]").sendVKey 0
  WScript.Sleep 250

  session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero
  session.findById("wnd[0]").sendVKey 0
  WScript.Sleep 250

  ' >>> ajuste aqui se seu caminho de menu for diferente <<<
  session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[11]/menu[0]").select
  WScript.Sleep 200
  session.findById("wnd[0]/tbar[0]/btn[11]").press
  WScript.Sleep 400

  st = StatusText()
  If Len(Trim(st)) = 0 Then
    WScript.Sleep 600
    st = StatusText()
  End If
  tp = StatusType()

  If tp = "E" Then
    LogLine idProgRT & "|CANCELAR|ERRO|" & st
    FechaPopupSeExistir
  Else
    LogLine idProgRT & "|CANCELAR|OK|" & st
  End If
  On Error GoTo 0
End Sub

LogLine "INICIO_LOTE"
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

CancelarRT 45, "325"
WScript.Sleep 250
LogLine "FIM_LOTE"
