option explicit

dim oFSO, LogFile_full, LogFile_cur, oShell, cur, ver, sTitle
ver = "1.2"

sTitle = "GetCPUUsage v." & ver & " by Dragokas"

Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

' Make me Admin :) Получаем права Администратора
if WScript.Arguments.Count = 0 then
    if not isAdminRights() then
        Elevate()
        WScript.Quit
	end if
end if

cur = oFSO.GetParentFolderName(WScript.ScriptFullName)

LogFile_full = cur & "\ProcessCPU_Average.csv"
LogFile_cur  = cur & "\ProcessCPU_Current.csv"
if oFSO.FileExists(LogFile_full) then RemoveFile LogFile_full
if oFSO.FileExists(LogFile_cur)  then RemoveFile LogFile_cur

CPUTimeToLog
msgbox "Готово." & vblf & "Выложите в теме, где Вам оказывают помощь, файлы:" & vblf & vblf &_
	"1. ProcessCPU_Current.csv" & vblf & "2. ProcessCPU_Average.csv" & vblf & vblf &_
	"упаковав в архив формата zip.", vbInformation, sTitle
WScript.Quit

Set oFSO = Nothing: Set oShell = Nothing


Sub CPUTimeToLog()
	dim Kernel_t1, User_t1, Total_t1
	dim Kernel_t2, User_t2, Total_t2
	dim oSCR_t1, oSCR_t2, oSCR_PID, oSCR_path, oSCR_Serv, oSCR_parentPID, oTS, WMI, oProcesses, oProcess, Key
	dim Proc_t1, Proc_t2, Delta_Proc, Delta_System, oServices, oService, Service_Name, ParentPID, ParentPath
	Dim aLines(), aField(), aPos(), i, sName

	'PID -> TotalTime
	set oSCR_t1 = CreateObject("Scripting.Dictionary")
	set oSCR_t2 = CreateObject("Scripting.Dictionary")
	'PID -> Name
	set oSCR_PID = CreateObject("Scripting.Dictionary")
	'PID -> Путь и параметры командной строки
	set oSCR_path = CreateObject("Scripting.Dictionary")
	'PID -> Service
	set oSCR_Serv = CreateObject("Scripting.Dictionary")
	'PID <-> ParentPID
	set oSCR_parentPID = CreateObject("Scripting.Dictionary")

	Set WMI = GetObject("winmgmts:\root\cimv2")

	Set oServices = WMI.ExecQuery("SELECT * FROM Win32_Service")
	For each oService in oServices
		if oSCR_Serv.Exists(oService.ProcessID) then
			oSCR_Serv(oService.ProcessID) = oSCR_Serv(oService.ProcessID) & _
				oService.Name & " (" & oService.Caption & "), "
		else
			oSCR_Serv.Add oService.ProcessID, oService.Name & " (" & oService.Caption & "), "
		end if
	Next

	WScript.Sleep(500) ' Нормализация % скачка, вызванного этим скриптом

	' 1-я засечка
	Set oProcesses = WMI.ExecQuery("SELECT * FROM Win32_Process")
	For each oProcess in oProcesses
		with oProcess
			Kernel_t1 = Kernel_t1 + cdbl(.KernelModeTime)
			User_t1   = User_t1   + cdbl(.UserModeTime)
			oSCR_t1.Add        .ProcessID, cdbl(.KernelModeTime) + cdbl(.UserModeTime)
			oSCR_PID.Add       .ProcessID, .Caption         'PID <-> Name
			oSCR_path.Add      .ProcessID, .ExecutablePath  'PID <-> Path
			oSCR_parentPID.Add .ProcessID, .ParentProcessId 'PID <-> ParentPID
		end with
	Next
	'Всего времени всех процессов
	Total_t1 = Kernel_t1 + User_t1
	
	Redim aLines(0), aField(0), aPos(0)
	
	For each Key in oSCR_t1.Keys 'Process Name, PID, CPU Time, CPU (%), Path, Service
	
		sName = oSCR_PID(Key)
	
		If sName <> "System Idle Process" Then
			Proc_t1 = oScr_t1(Key)
			if (oSCR_Serv.Exists(Key) and Key <> 0) then Service_Name = oSCR_Serv(Key) else Service_Name = ""
			ParentPID = oSCR_parentPID(Key)
			if (oSCR_path.Exists(ParentPID) and Key <> 0) then ParentPath = oSCR_path(ParentPID) else ParentPath = ""
			
			ArrayAdd aLines, sName & ";" &_
				Key & ";" &_
				Replace(Proc_t1, ".", ",") & ";" &_
				Replace(round(Proc_t1 / Total_t1 * 100, 2), ".", ",") & ";" &_
				oScr_path(Key) & ";" &_
				ParentPath & ";" &_
				Service_Name
			
			ArrayAdd aField, Proc_t1 / Total_t1
			ArrayAdd aPos, Ubound(aField)
		End If
	Next
	
	QuickSortSpecial aField, aPos, 1, UBound(aField)
	
	set oTS = oFSO.CreateTextFile(LogFile_full, true)
	oTS.WriteLine "Process Name;PID;CPU Time;CPU (%);Path;ParentPath;Service"
	For i = Ubound(aPos) to 1 step -1
		oTS.WriteLine aLines(aPos(i))
	Next
	oTS.Close
	
	WScript.Sleep(2000) 'выжидаю 2 сек.

	' 2-я засечка
	Set oProcesses = WMI.ExecQuery("SELECT * FROM Win32_Process")
	For each oProcess in oProcesses
		with oProcess
			Kernel_t2 = Kernel_t2 + cdbl(.KernelModeTime)
			User_t2   = User_t2   + cdbl(.UserModeTime)
			oSCR_t2.Add .ProcessID, cdbl(.KernelModeTime) + cdbl(.UserModeTime)
			if not oSCR_PID.Exists(.ProcessID) then
				oSCR_PID.Add       .ProcessID, .Caption         'PID <-> Name (если появились новые)
				oSCR_path.Add      .ProcessID, .ExecutablePath  'PID <-> Path (если появились новые)
				oSCR_parentPID.Add .ProcessID, .ParentProcessId 'PID <-> ParentPID
			end if
		end with
	Next
	'Всего времени всех процессов
	Total_t2 = Kernel_t2 + User_t2
	
	Redim aLines(0), aField(0), aPos(0)
	
	' Записываю разницу по формуле:
	' % нагрузки процесса = Дельта времени процесса / дельта времени системы * 100
	For each Key in oSCR_t2.Keys 'Process Name, PID, CPU Time, CPU (%), Path, Service
		
		sName = oSCR_PID(Key)
		
		If sName <> "System Idle Process" Then
			Proc_t1 = oScr_t1(Key)
			Proc_t2 = oScr_t2(Key)
			Delta_Proc   = Proc_t2  - Proc_t1
			Delta_System = Total_t2 - Total_t1
			if (oSCR_Serv.Exists(Key) and Key <> 0) then Service_Name = oSCR_Serv(Key) else Service_Name = ""
			ParentPID = oSCR_parentPID(Key)
			if (oSCR_path.Exists(ParentPID) and Key <> 0) then ParentPath = oSCR_path(ParentPID) else ParentPath = ""
			
			ArrayAdd aLines, sName & ";" &_
				Key & ";" &_
				Replace(Proc_t2 - Proc_t1, ".", ",") & ";" &_
				Replace(round(Delta_Proc / Delta_System * 100, 2), ".", ",") & ";" &_
				oScr_path(Key) & ";" &_
				ParentPath & ";" &_
				Service_Name
				
			ArrayAdd aField, Delta_Proc / Delta_System
			ArrayAdd aPos, Ubound(aField)
		End If
	Next
	
	QuickSortSpecial aField, aPos, 1, UBound(aField)
	
	set oTS = oFSO.CreateTextFile(LogFile_cur, true)
	oTS.WriteLine "Process Name;PID;CPU Time;CPU (%);Path;ParentPath;Service"
	For i = Ubound(aPos) to 1 step -1
		oTS.WriteLine aLines(aPos(i))
	Next
	oTS.Close

	Set oProcess = Nothing: set oProcesses = Nothing: set WMI = Nothing: set oTS = Nothing
	Set oSCR_PID = Nothing: set oSCR_t1 = Nothing: set oSCR_t2 = Nothing: set oSCR_path = Nothing
	Set oSCR_Serv = Nothing: set oSCR_parentPID = Nothing
End Sub


Sub Elevate()
	Dim colOS, oOS, strOSLong, oShellApp
	Const DQ = """"
	Set colOS = GetObject("winmgmts:\root\cimv2").ExecQuery("Select * from Win32_OperatingSystem")
	For Each oOS In colOS: strOSLong = oOS.Version: Next
	If CLng(Split(Replace(strOSLong, ",", "."), ".")(0)) >= 6 and Not isAdminRights Then
		Set oShellApp = CreateObject("Shell.Application")
		oShellApp.ShellExecute WScript.FullName, DQ & WScript.ScriptFullName & DQ & " " & DQ & "Twice" & DQ, "", "runas", 1
		WScript.Quit
	End If
	set oOS = Nothing: set colOS = Nothing: set oShellApp = Nothing
End Sub

Function isAdminRights()
	Dim oReg, strKey, intErrNum, flagAccess
    Const KQV = &H1, KSV = &H2, HKCU = &H80000001, HKLM = &H80000002
    Set oReg = GetObject("winmgmts:root\default:StdRegProv")
    strKey = "System\CurrentControlSet\Control\Session Manager"
    intErrNum = oReg.CheckAccess(HKLM, strKey, KQV + KSV, flagAccess)
    isAdminRights = flagAccess
    Set oReg = Nothing
End Function

Sub ArrayAdd(arr, value)
	Redim Preserve arr(Ubound(arr) + 1)
	arr(Ubound(arr)) = value
End Sub

'Sort user type arrays by any field (c) Dragokas
'input - j(), k() arrays
'output - k() array with indeces of j() array, in sorted order
Public Sub QuickSortSpecial(j(), k(), ByVal low, ByVal high)
    Dim i, L, M, wsp
    i = low: L = high: M = j((i + L) \ 2)
    Do Until i > L: Do While j(i) < M: i = i + 1: Loop: Do While j(L) > M: L = L - 1: Loop
        If (i <= L) Then wsp = j(i): j(i) = j(L): j(L) = wsp: wsp = k(i): k(i) = k(L): k(L) = wsp: i = i + 1: L = L - 1
    Loop
    If low < L Then QuickSortSpecial j, k, low, L
    If i < high Then QuickSortSpecial j, k, i, high
End Sub

Sub RemoveFile(sFile)
	On Error Resume Next
	oFSO.DeleteFile sFile
	if Err.Number <> 0 then
		MsgBox "Файл " & oFSO.GetFileName(sFile) & " заблокирован! Сперва, закройте все программы, которые его используют.", vbExclamation, sTitle
		WScript.Quit 5
	end if
End Sub