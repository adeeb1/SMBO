Attribute VB_Name = "modSecurity"
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, _
uProcess As PROCESSENTRY32) As Long

Private Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" ( _
ByVal lFlags As Long, lProcessID As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const TH32CS_SNAPPROCESS As Long = 2&

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Public Sub KillProcess(ProcessName As String)
    Dim uProcess As PROCESSENTRY32
    Dim RProcessFound As Long, hSnapshot As Long, ExitCode As Long, MyProcess As Long
    Dim SzExename As String, WinDirEnv As String
    Dim AppKill As Boolean
    Dim i As Integer, AppCount As Integer
        
    If ProcessName <> "" Then
        AppCount = 0

        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        RProcessFound = ProcessFirst(hSnapshot, uProcess)
  
        Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            WinDirEnv = Environ("Windir") + "\"
            WinDirEnv = LCase$(WinDirEnv)
        
            If Right$(SzExename, Len(ProcessName)) = LCase$(ProcessName) Then
               AppCount = AppCount + 1
               MyProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
               AppKill = TerminateProcess(MyProcess, ExitCode)
               Call CloseHandle(MyProcess)
            End If
            
            RProcessFound = ProcessNext(hSnapshot, uProcess)
            
            DoEvents
        Loop While RProcessFound
        
        Call CloseHandle(hSnapshot)
    End If
End Sub

Public Sub StopCheatingHacking(ByVal ProcessNum As Integer)
    If ProcessNum = 0 Then
        Call KillProcess("RawSockets.exe")
        Call KillProcess("VB6.exe")
        Call KillProcess("studio.exe")
        Call KillProcess("capsa.exe")
        Call KillProcess("prorat.exe")
        Call KillProcess("PacketEditor.exe")
        Call KillProcess("wireshark.exe")
        Call KillProcess("TSearch.exe")
        Call KillProcess("programme test.exe")
        Call KillProcess("WPE PRO - Modified.exe")
        Call KillProcess("PermEdit.exe")
        Call KillProcess("cv.exe")
        Call KillProcess("nettools5.exe")
        Call KillProcess("Nsauditor.exe")
        Call KillProcess("EtherD.exe")
        Call KillProcess("EHSniffer.exe")
        Call KillProcess("APS.exe")
    Else
        Select Case ProcessNum
            Case 1
                Call KillProcess("RawSockets.exe")
            Case 2
                Call KillProcess("VB6.exe")
            Case 3
                Call KillProcess("studio.exe")
            Case 4
                Call KillProcess("capsa.exe")
            Case 5
                Call KillProcess("prorat.exe")
            Case 6
                Call KillProcess("PacketEditor.exe")
            Case 7
                Call KillProcess("wireshark.exe")
            Case 8
                Call KillProcess("TSearch.exe")
            Case 9
                Call KillProcess("programme test.exe")
            Case 10
                Call KillProcess("WPE PRO - Modified.exe")
            Case 11
                Call KillProcess("PermEdit.exe")
            Case 12
                Call KillProcess("cv.exe")
            Case 13
                Call KillProcess("nettools5.exe")
            Case 14
                Call KillProcess("Nsauditor.exe")
            Case 15
                Call KillProcess("EtherD.exe")
            Case 16
                Call KillProcess("EHSniffer.exe")
            Case 17
                Call KillProcess("APS.exe")
        End Select
    End If
End Sub
