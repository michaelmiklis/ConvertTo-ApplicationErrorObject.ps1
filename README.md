# ConvertTo-ApplicationErrorObject.ps1

This PowerShell script accepts multiple Application Error Eventlog entries (EventID 1000) via pipeline input. The typical application error message:

			Faulting application name: VSIXAutoUpdate.exe, version: 15.5.100.41289, time stamp: 0x5a0b3ffb
			Faulting module name: KERNELBASE.dll, version: 10.0.17133.1, time stamp: 0x6b02ce16
			Exception code: 0xe0434352
			Fault offset: 0x0010d722
			Faulting process id: 0x1be0
			Faulting application start time: 0x01d3f057d6476cd6
			Faulting application path: C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe
			Faulting module path: C:\WINDOWS\System32\KERNELBASE.dll
			Report Id: e192bc75-ee60-4475-9070-e34f26620b91
			Faulting package full name: 
			Faulting package-relative application ID: 

will be parsed into a full PowerShell Object:

			MachineName      : DESKTOP-KUGK40G
			UserId           : 
			TimeCreated      : 20.05.2018 18:30:30
			ProcessName      : VSIXAutoUpdate.exe
			ProcessVersion   : 15.5.100.41289
			ProcessTimestamp : 0x5a0b3ffb
			ModuleName       : KERNELBASE.dll
			ModuleVersion    : 10.0.17133.1
			ModuleTimestamp  : 0x6b02ce16
			Exception        : 0xe0434352
			FailureOffset    : 0x0010d722
			ProcessID        : 0x1be0
			ProcessStartTime : 0x01d3f057d6476cd6
			ProcessPath      : C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe
			ModulePath       : C:\WINDOWS\System32\KERNELBASE.dll
			ReportID         : e192bc75-ee60-4475-9070-e34f26620b91
			FullPackageName  : 
			AppID            : 


If you pipe the output into "ft *" or "Export-CSV" you will get a great overview of all application crashes:

|MachineName     |UserId |TimeCreated         |ProcessName        |ProcessVersion |ProcessTimestamp |ModuleName     |ModuleVersion |ModuleTimestamp |Exception  |FailureOffset |ProcessID |ProcessStartTime   |ProcessPath                                                                                     |ModulePath                         |ReportID              |
|-----------     |------ |-----------         |-----------        |-------------- |---------------- |----------     |------------- |--------------- |---------  |------------- |--------- |----------------   |-----------                                                                                     |----------                         |--------              |
|DESKTOP-KUGK40G |       |20.05.2018 18:30:30 |VSIXAutoUpdate.exe |15.5.100.41289 |0x5a0b3ffb       |KERNELBASE.dll |10.0.17133.1  |0x6b02ce16      |0xe0434352 |0x0010d722    |0x1be0    |0x01d3f057d6476cd6 |C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe |C:\WINDOWS\System32\KERNELBASE.dll |e192bc75-ee60-4475-...|
|DESKTOP-KUGK40G |       |17.05.2018 09:49:49 |VSIXAutoUpdate.exe |15.5.100.41289 |0x5a0b3ffb       |KERNELBASE.dll |10.0.17133.1  |0x6b02ce16      |0xe0434352 |0x0010d722    |0x1560    |0x01d3edb374d5c779 |C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe |C:\WINDOWS\System32\KERNELBASE.dll |93a23c66-b0dd-4911-...|
|DESKTOP-KUGK40G |       |14.05.2018 08:40:20 |VSIXAutoUpdate.exe |15.5.100.41289 |0x5a0b3ffb       |KERNELBASE.dll |10.0.17133.1  |0x6b02ce16      |0xe0434352 |0x0010d722    |0x2578    |0x01d3eb4e65ece40d |C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe |C:\WINDOWS\System32\KERNELBASE.dll |f385582d-8b7e-4ea5-...|
|DESKTOP-KUGK40G |       |11.05.2018 05:47:34 |VSIXAutoUpdate.exe |15.5.100.41289 |0x5a0b3ffb       |KERNELBASE.dll |10.0.17133.1  |0x6b02ce16      |0xe0434352 |0x0010d722    |0x2b4c    |0x01d3e8dac7d0d000 |C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe |C:\WINDOWS\System32\KERNELBASE.dll |a699017c-97f7-484c-...|
|DESKTOP-KUGK40G |       |10.05.2018 12:42:45 |VSIXAutoUpdate.exe |15.5.100.41289 |0x5a0b3ffb       |KERNELBASE.dll |10.0.17133.1  |0x6b02ce16      |0xe0434352 |0x0010d722    |0xaf4     |0x01d3e84b9c6bc801 |C:\Program Files (x86)\Microsoft Visual Studio\2017\Professional\Common7\IDE\VSIXAutoUpdate.exe |C:\WINDOWS\System32\KERNELBASE.dll |4cf41990-93e7-4b70-...|


Currently English and German language eventlog entries are supported.


Usage
========
Execute the following commandline in a PowerShell Window after you have loaded the script:

	Get-WinEvent -LogName Application | Where-Object { ($_.ID -eq 1000) -and ($_.ProviderName -eq "Application Error") } | ConvertTo-ApplicationErrorObject | ft *
	
