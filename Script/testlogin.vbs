Dim SapGuiAuto
Dim application
Dim connection
Dim session
Dim WshShell
Dim i

Set WshShell = CreateObject("WScript.Shell")

' Helper: Day tat ca cua so SAP xuong sau cac ung dung khac
Sub PushSAPToBack()
    Dim sTmpPS : sTmpPS = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\sap_to_back.ps1"
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = fso.OpenTextFile(sTmpPS, 2, True)
    f.WriteLine "Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class W32H{[DllImport(""user32.dll"")]public static extern bool SetWindowPos(IntPtr h,IntPtr i,int x,int y,int w,int ht,uint f);}'"
    f.WriteLine "Get-Process | Where-Object {$_.MainWindowTitle -match 'SAP' -and $_.MainWindowHandle -ne 0} | ForEach-Object {[W32H]::SetWindowPos($_.MainWindowHandle,[IntPtr]1,0,0,0,0,3)}"
    f.Close
    WshShell.Run "powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File """ & sTmpPS & """", 0, True
End Sub

Sub HideSAPLogon()
        Dim sTmpPS2 : sTmpPS2 = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\sap_hide_logon.ps1"
        Dim fso2 : Set fso2 = CreateObject("Scripting.FileSystemObject")
        Dim f2 : Set f2 = fso2.OpenTextFile(sTmpPS2, 2, True)
        f2.WriteLine "Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class W32H2{[DllImport(" & Chr(34) & "user32.dll" & Chr(34) & ")]public static extern bool ShowWindow(IntPtr h,int n);}'"
        f2.WriteLine "Get-Process | Where-Object {$_.MainWindowTitle -match 'SAP Logon' -and $_.MainWindowHandle -ne 0} | ForEach-Object {[W32H2]::ShowWindow($_.MainWindowHandle, 6)}"
        f2.Close
        WshShell.Run "powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File """ & sTmpPS2 & """", 0, True
End Sub
Sub HandleMultipleLogonPopup(ByRef sapSession, ByRef sapConnection)
    Dim nTry
    For nTry = 1 To 5
        WScript.Sleep 1000
        On Error Resume Next
        Set sapSession = sapConnection.Sessions(0)
        sapSession.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select
        If Err.Number = 0 Then
            sapSession.findById("wnd[1]/tbar[0]/btn[0]").press
            Err.Clear
            On Error GoTo 0
            WScript.Sleep 1000
            Exit Sub
        End If
        Err.Clear
        On Error GoTo 0
    Next
End Sub

'==============================
' MO SAP LOGON
'==============================
WshShell.Run """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe""", 0, False

'==============================
' DOI SAP GUI SAN SANG
'==============================
For i = 1 To 30
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit For
    End If
    Err.Clear
    Set SapGuiAuto = Nothing
    On Error GoTo 0
    WScript.Sleep 1000
Next

If SapGuiAuto Is Nothing Then
    Dim oLog : Set oLog = CreateObject("Scripting.FileSystemObject").OpenTextFile("D:\4.DEV\Python\XuatDuLieuSAP\Data\MB52_log.txt", 8, True)
    oLog.WriteLine Now & " [ERROR] Khong bat duoc SAPGUI"
    oLog.Close
    WScript.Quit
End If

Set application = SapGuiAuto.GetScriptingEngine

'==============================
' MO ENV
'==============================
' application.OpenConnection "2.APOLLON QAS AA3", True
application.OpenConnection "1.APOLLON PRD AA0 Group", True
For i = 1 To 30
    If application.Children.Count > 0 Then Exit For
    WScript.Sleep 1000
Next

Set connection = application.Children(0)
 HideSAPLogon()

' Doi session khoi tao xong
For i = 1 To 30
    If connection.Sessions.Count > 0 Then Exit For
    WScript.Sleep 1000
Next

'==============================
' DANG NHAP (SSO)
'==============================
WScript.Sleep 2000

' Refresh session
Set session = connection.Sessions(0)

' Kiem tra co man hinh dang nhap khong (chi ton tai tren login screen)
Dim bOnLoginScreen
On Error Resume Next
session.findById("wnd[0]/usr/txtRSYST-MANDT").setFocus
bOnLoginScreen = (Err.Number = 0)
Err.Clear
On Error GoTo 0

If bOnLoginScreen Then
    ' Kiem tra co bang SSO User Selection khong
    Dim oSSOTable
    On Error Resume Next
    Set oSSOTable = session.findById("wnd[0]/usr/tblSAPLSSO2TC_USR_SEL")
    Dim bHasSSO
    bHasSSO = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If bHasSSO Then
        ' Chon dong co Client = 280 trong bang SSO
        Dim nRow, sClient
        For nRow = 0 To oSSOTable.RowCount - 1
            On Error Resume Next
            sClient = session.findById("wnd[0]/usr/tblSAPLSSO2TC_USR_SEL/txtRSYST-MANDT[0," & nRow & "]").Text
            If Err.Number = 0 And Trim(sClient) = "280" Then
                On Error GoTo 0
                oSSOTable.selectedRows = CStr(nRow)
                session.findById("wnd[0]").sendVKey 0
                Exit For
            End If
            On Error GoTo 0
        Next
    Else
        ' Man hinh dang nhap thuong - nhap Client va User
        session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "280"
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "249533"
        session.findById("wnd[0]").sendVKey 0
    End If

    ' Xu ly popup Multiple Logon neu xuat hien
    HandleMultipleLogonPopup session, connection

    ' Doi SAP Easy Access menu load xong
    For i = 1 To 60
        On Error Resume Next
        Set session = connection.Sessions(0)
        Dim sOKCod
        sOKCod = session.findById("wnd[0]/tbar[0]/okcd").Text
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
        WScript.Sleep 1000
    Next
End If

Set session = connection.Sessions(0)
HandleMultipleLogonPopup session, connection
