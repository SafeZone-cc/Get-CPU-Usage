option explicit

dim oFSO, LogFile_full, LogFile_cur, oShell, cur, ver
ver = "1.1"

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
if oFSO.FileExists(LogFile_full) then oFSO.DeleteFile(LogFile_full)
if oFSO.FileExists(LogFile_cur)  then oFSO.DeleteFile(LogFile_cur)

CPUTimeToLog
msgbox "Готово." & vblf & "Выложите в теме, где Вам оказывают помощь, файлы:" & vblf & vblf &_
	"1. ProcessCPU_Current.csv" & vblf & "2. ProcessCPU_Average.csv" & vblf & vblf &_
	"упаковав в архив формата zip.", vbInformation, "GetCPUUsage v." & ver & " by Dragokas"
WScript.Quit

Set oFSO = Nothing: Set oShell = Nothing


Sub CPUTimeToLog()
	dim Kernel_t1, User_t1, Total_t1
	dim Kernel_t2, User_t2, Total_t2
	dim oSCR_t1, oSCR_t2, oSCR_PID, oSCR_path, oSCR_Serv, oSCR_parentPID, oTS, WMI, oProcesses, oProcess, Key
	dim Proc_t1, Proc_t2, Delta_Proc, Delta_System, oServices, oService, Service_Name, ParentPID, ParentPath

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

	set oTS = oFSO.CreateTextFile(LogFile_full, true)
	oTS.WriteLine "Process Name;PID;CPU Time;CPU (%);Service;ParentPath;Path"
	
	For each Key in oSCR_t1.Keys 'Process Name, PID, CPU Time, CPU (%), Path, Service
		Proc_t1 = oScr_t1(Key)
		if (oSCR_Serv.Exists(Key) and Key <> 0) then Service_Name = oSCR_Serv(Key) else Service_Name = ""
		ParentPID = oSCR_parentPID(Key)
		if (oSCR_path.Exists(ParentPID) and Key <> 0) then ParentPath = oSCR_path(ParentPID) else ParentPath = ""
		oTS.WriteLine oSCR_PID(Key) & ";" &_
			Key & ";" &_
			Proc_t1 & ";" &_
			round(Proc_t1 / Total_t1 * 100, 2) & ";" &_
			Service_Name & ";" &_
			ParentPath & ";" &_
			oScr_path(Key)
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

	set oTS = oFSO.CreateTextFile(LogFile_cur, true)
	oTS.WriteLine "Process Name;PID;CPU Time;CPU (%);Service;ParentPath;Path"

	' Записываю разницу по формуле:
	' % нагрузки процесса = Дельта времени процесса / дельта времени системы * 100
	For each Key in oSCR_t2.Keys 'Process Name, PID, CPU Time, CPU (%), Path, Service
		Proc_t1 = oScr_t1(Key)
		Proc_t2 = oScr_t2(Key)
		Delta_Proc   = Proc_t2  - Proc_t1
		Delta_System = Total_t2 - Total_t1
		if (oSCR_Serv.Exists(Key) and Key <> 0) then Service_Name = oSCR_Serv(Key) else Service_Name = ""
		ParentPID = oSCR_parentPID(Key)
		if (oSCR_path.Exists(ParentPID) and Key <> 0) then ParentPath = oSCR_path(ParentPID) else ParentPath = ""
		oTS.WriteLine oSCR_PID(Key) & ";" &_
			Key & ";" &_
			Proc_t2 - Proc_t1 & ";" &_
			round(Delta_Proc / Delta_System * 100, 2) & ";" &_
			Service_Name & ";" &_
			ParentPath & ";" &_
			oScr_path(Key)
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
	If Left(strOSLong, 1) = "6" and Not isAdminRights Then
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