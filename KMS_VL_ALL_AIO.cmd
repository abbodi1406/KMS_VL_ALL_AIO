<!-- : Begin batch script
@setlocal DisableDelayedExpansion
@set uivr=v52
@echo off
:: ### Configuration Options ###

:: change to 1 to enable debug mode (can be used with unattended options)
set _Debug=0

:: change to 0 to turn OFF Windows or Office activation processing via the script
set ActWindows=1
set ActOffice=1

:: change to 0 to turn OFF auto conversion for Office C2R Retail to Volume
set AutoR2V=1

:: change to 0 to keep Office C2R vNext license (subscription or lifetime)
set vNextOverride=1

:: change to 0 to revert Windows 10/11 KMS38 to normal KMS
set SkipKMS38=1

:: ### Unattended Options ###

:: change to 1 and set KMS_IP address to activate via external KMS server unattended
set External=0
set KMS_IP=172.16.0.2

:: change to 1 to run Manual activation mode unattended
set uManual=0

:: change to 1 to run AutoRenewal activation mode unattended
set uAutoRenewal=0

:: change to 1 to suppress any output
set Silent=0

:: change to 1 to redirect output to a text file, works only with Silent=1
set Logger=0

:: ### Advanced KMS Options ###

:: change KMS auto renewal schedule for activated clients, range in minutes: from 15 to 43200
:: example: 10080 = weekly, 1440 = daily, 43200 = monthly
set KMS_RenewalInterval=10080

:: change KMS reattempt schedule for unactivated clients, range in minutes: from 15 to 43200
set KMS_ActivationInterval=120

:: change Hardware Hash for KMS emulator server (only affect Windows 8.1 and 10)
set KMS_HWID=0x3A1C049600B60076

:: change KMS TCP port
set KMS_Port=1688

:: change to 1 to use VBScript to access WMI
:: automatically enabled if wmic.exe is not installed
set WMI_VBS=0

:: change to 1 to use Windows PowerShell to access WMI
:: automatically enabled if wmic.exe and VBScript are not installed
set WMI_PS=0

:: ###################################################################
:: # NORMALLY THERE IS NO NEED TO CHANGE ANYTHING BELOW THIS COMMENT #
:: ###################################################################

set KMS_Emulation=1
set Unattend=0
set _uIP=172.16.0.2

set "_Null=1>nul 2>nul"

set _args=
set _elev=
set _batf=
set _batp=
set fAUR=
set rAUR=
set _args=%*
if not defined _args goto :NoProgArgs

set _args=%_args:"=%
for %%A in (%_args%) do (
if /i "%%A"=="-elevated" (set _elev=1
) else if /i "%%A"=="-wow" (set _rel1=1
) else if /i "%%A"=="-arm" (set _rel2=1
) else if /i "%%A"=="/d" (set _Debug=1
) else if /i "%%A"=="/u" (set Unattend=1
) else if /i "%%A"=="/s" (set Silent=1
) else if /i "%%A"=="/l" (set Logger=1
) else if /i "%%A"=="/o" (set ActOffice=1&set ActWindows=0
) else if /i "%%A"=="/w" (set ActOffice=0&set ActWindows=1
) else if /i "%%A"=="/c" (set AutoR2V=0
) else if /i "%%A"=="/v" (set vNextOverride=0
) else if /i "%%A"=="/x" (set SkipKMS38=0
) else if /i "%%A"=="/e" (set fAUR=0&set External=1&set uManual=0&set uAutoRenewal=0
) else if /i "%%A"=="/m" (set fAUR=0&set External=0&set uAutoRenewal=0
) else if /i "%%A"=="/a" (set fAUR=1&set External=0&set uManual=0
) else if /i "%%A"=="/r" (set rAUR=1
) else (set "KMS_IP=%%A")
)

:NoProgArgs
set "_cmdf=%~f0"
if exist "%SystemRoot%\Sysnative\cmd.exe" if not defined _rel1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" -wow %*"
exit /b
)
if exist "%SystemRoot%\SysArm32\cmd.exe" if /i %PROCESSOR_ARCHITECTURE%==AMD64 if not defined _rel2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" -arm %*"
exit /b
)
if %External% EQU 1 (if "%KMS_IP%"=="%_uIP%" (set fAUR=0&set External=0) else (set fAUR=0))
if %uManual% EQU 1 (set fAUR=0&set External=0&set uAutoRenewal=0)
if %uAutoRenewal% EQU 1 (set fAUR=1&set External=0&set uManual=0)
if defined fAUR set Unattend=1
if defined rAUR set Unattend=1
if %Silent% EQU 1 set Unattend=1
set _run=nul
if %Logger% EQU 1 set _run="%~dpn0_Silent.log"

set "SysPath=%SystemRoot%\System32"
set "Path=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "SysPath=%SystemRoot%\Sysnative"
set "Path=%SystemRoot%\Sysnative;%SystemRoot%;%SystemRoot%\Sysnative\Wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%Path%"
)
set "_psc=powershell -nop -c"
set "_buf={$W=$Host.UI.RawUI.WindowSize;$B=$Host.UI.RawUI.BufferSize;$W.Height=31;$B.Height=300;$Host.UI.RawUI.WindowSize=$W;$Host.UI.RawUI.BufferSize=$B;}"
set "_err===== ERROR ===="
set "o_x64=1a6a3e40f610a6394ef539a039308dbe2f526ac1"
set "o_x86=1b6ee9e99b1dbcfc9427bb3a61c65ed53667fc22"
set "o_arm=3d158e9cbd6de13f954e8dba356e369b557fe2d9"
set "_bit=64"
set "_wow=1"
if /i "%PROCESSOR_ARCHITECTURE%"=="amd64" set "xBit=x64"&set "xOS=x64"&set "_orig=%o_x64%"
if /i "%PROCESSOR_ARCHITECTURE%"=="arm64" set "xBit=x86"&set "xOS=A64"&set "_orig=%o_arm%"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" set "xBit=x86"&set "xOS=x86"&set "_orig=%o_x86%"&set "_wow=0"&set "_bit=32"
if /i "%PROCESSOR_ARCHITEW6432%"=="amd64" set "xBit=x64"&set "xOS=x64"&set "_orig=%o_x64%"
if /i "%PROCESSOR_ARCHITEW6432%"=="arm64" set "xBit=x86"&set "xOS=A64"&set "_orig=%o_arm%"

set _invpth=0
set "param=%~f0"
cmd /v:on /c echo(^^!param^^!| findstr /R "[| ` ~ ! @ %% \^ & ( ) \[ \] { } + = ; ' , |]*^" 1>nul 2>nul
if %errorlevel% EQU 0 set _invpth=1

reg query HKLM\SYSTEM\CurrentControlSet\Services\WinMgmt /v Start 2>nul | find /i "0x4" 1>nul && (goto :E_WMS)

set _cwmi=0
for %%# in (wmic.exe) do @if not "%%~$PATH:#"=="" (
wmic path Win32_ComputerSystem get CreationClassName /value 2>nul | find /i "ComputerSystem" 1>nul && set _cwmi=1
)
set _pwsh=1
for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" set _pwsh=0
if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" set _pwsh=0
if %_pwsh% equ 0 goto :E_PWS

set psfull=1
2>nul %_psc% $ExecutionContext.SessionState.LanguageMode | find /i "Full" 1>nul || set psfull=0
if %psfull% equ 0 goto :E_PLM

set _dllPath=%SystemRoot%\System32
if %xOS%==A64 %_psc% $env:PROCESSOR_ARCHITECTURE 2>nul | find /i "x86" 1>nul && set _dllPath=%SystemRoot%\Sysnative
set _dllNum=1
if %xOS%==x64 set _dllNum=2
if %xOS%==A64 set _dllNum=3
set preparedcolor=0

1>nul 2>nul reg query HKU\S-1-5-19 && (
  goto :Passed
  ) || (
  if defined _elev goto :E_Admin
)

set _PSarg="""%~f0""" %_args% -elevated
set _PSarg=%_PSarg:'=''%

(1>nul 2>nul cscript //NoLogo "%~f0?.wsf" //job:ELAV /File:"%~f0" %_args% -elevated) && (
  exit /b
  ) || (
  call setlocal EnableDelayedExpansion
  1>nul 2>nul %SysPath%\WindowsPowerShell\v1.0\%_psc% "start cmd.exe -arg '/c \"!_PSarg!\"' -verb runas" && (
    exit /b
    ) || (
    goto :E_Admin
  )
)

:Passed
if not exist "%SystemRoot%\Temp\" mkdir "%SystemRoot%\Temp" 1>nul 2>nul
set "_onat=HKLM\SOFTWARE\Microsoft\Office"
set "_owow=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office"
set "_batf=%~f0"
set "_batp=%_batf:'=''%"
set "_utemp=%TEMP%"
set "_Local=%LocalAppData%"
set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
:: set _UNC=0
:: if "%_work:~0,2%"=="\\" (
:: set _UNC=1
:: ) else (
:: net use %~d0 %_Null%
:: if not errorlevel 1 set _UNC=1
:: )
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%PUBLIC%\Desktop\desktop.ini" set "_dsk=%PUBLIC%\Desktop"
for %%A in (14,15,16,19,21,24) do call :officeMsg %%A
set "_mOuwp=Detected Office 365/2016 UWP is not supported by KMS_VL_ALL"
set DO15Ids=ProPlus,Standard,Access,Lync,Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set DO16Ids=ProPlus,Standard,Access,SkypeforBusiness,Excel,Outlook,PowerPoint,Publisher,Word
set LV16Ids=Mondo,ProPlus,ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness,OneNote,Excel,Outlook,PowerPoint,Publisher,Word
set LR16Ids=%LV16Ids%,Professional,HomeBusiness,HomeStudent,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set "ESUEditions=Enterprise,EnterpriseE,EnterpriseN,Professional,ProfessionalE,ProfessionalN,Ultimate,UltimateE,UltimateN"
if %winbuild% GEQ 9200 set "ESUEditions=Enterprise,EnterpriseN,Professional,ProfessionalN"
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum" (
set "ESUEditions=ServerDatacenter,ServerDatacenterCore,ServerDatacenterV,ServerDatacenterVCore,ServerStandard,ServerStandardCore,ServerStandardV,ServerStandardVCore,ServerEnterprise,ServerEnterpriseCore,ServerEnterpriseV,ServerEnterpriseVCore"
if %winbuild% GEQ 9200 set "ESUEditions=ServerDatacenter,ServerStandard"
)
set UBR=0
if %winbuild% GEQ 7601 for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v UBR 2^>nul') do if not errorlevel 1 set /a UBR=%%b
set _WSH=1
reg query "HKCU\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
reg query "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled 2>nul | find /i "0x0" 1>nul && (set _WSH=0)
if %_cwmi% EQU 0 if %WMI_PS% EQU 0 if %_WSH% EQU 1 if exist "%SysPath%\vbscript.dll" set WMI_VBS=1
if %_cwmi% EQU 0 if %WMI_VBS% EQU 0 if %_pwsh% EQU 1 set WMI_PS=1
if %_cwmi% EQU 0 if %WMI_VBS% EQU 0 if %WMI_PS% EQU 0 goto :E_WMI
set _acc=WMIC
if %WMI_VBS% NEQ 0 if %WMI_PS% EQU 0 (
if %_WSH% EQU 0 goto :E_WSH
if not exist "%SysPath%\vbscript.dll" goto :E_VBS
if %_invpth% EQU 1 goto :E_PTH
set _cwmi=0
set _acc=VBS
)
if %WMI_PS% NEQ 0 (
set _cwmi=0
set WMI_VBS=0
set _acc=PS
)

set "_csg=cscript.exe //NoLogo //Job:WmiMulti "%~nx0?.wsf""
set "_csq=cscript.exe //NoLogo //Job:WmiQuery "%~nx0?.wsf""
set "_csm=cscript.exe //NoLogo //Job:WmiMethod "%~nx0?.wsf""
set "_csp=cscript.exe //NoLogo //Job:WmiPKey "%~nx0?.wsf""
set "_csd=cscript.exe //NoLogo //Job:MPS "%~nx0?.wsf""
set "_csx=cscript.exe //NoLogo //Job:XPDT "%~nx0?.wsf""

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 2>nul | find /i "0x0" 1>nul && (set _NCS=0)
setlocal EnableDelayedExpansion
set "_oem=!_work!"
copy /y nul "!_work!\#.rw" 1>nul 2>nul && (
if exist "!_work!\#.rw" del /f /q "!_work!\#.rw"
) || (
set "_oem=!_dsk!"
set "_log=!_dsk!\%~n0"
if %Logger% EQU 1 set _run="!_dsk!\%~n0_Silent.log"
)
pushd "!_work!"
set "_suf="
call :qrSingle Win32_OperatingSystem LocalDateTime
if %_Debug% EQU 1 if exist "!_log!_Debug.log" (
for /f "tokens=2 delims==." %%# in ('%_qr%') do set "_date=%%#"
set "_suf=_!_date:~8,6!"
)

if %_Debug% EQU 0 (
  set "_Nul1=1>nul"
  set "_Nul2=2>nul"
  set "_Nul6=2^>nul"
  set "_Nul3=1>nul 2>nul"
  set "_Pause=pause >nul"
  if %Unattend% EQU 1 set "_Pause="
  if %Silent% EQU 0 (call :Begin) else (call :Begin >!_run! 2>&1)
) else (
  set "_Nul1="
  set "_Nul2="
  set "_Nul6="
  set "_Nul3="
  set "_Pause="
  if %Silent% EQU 0 (
  echo.
  echo Running in Debug Mode...
  if not defined _args (echo The window will be closed when finished) else (echo please wait...)
  echo.
  echo writing debug log to:
  echo "!_log!_Debug!_suf!.log"
  )
  @echo on
  @prompt $G
  @call :Begin >"!_log!_tmp.log" 2>&1 &cmd /u /c type "!_log!_tmp.log">"!_log!_Debug!_suf!.log"&del "!_log!_tmp.log"
)
@color 07
@title %ComSpec%
@echo off
@exit /b

:Begin
if %_Debug% EQU 1 (
if defined _args echo %_args%
echo "!_batf!"
)
if exist "%PUBLIC%\ReadMeAIO.html" del /f /q "%PUBLIC%\ReadMeAIO.html"
if exist "%_temp%\'" del /f /q "%_temp%\'"
if exist "%_temp%\`.txt" del /f /q "%_temp%\`.txt"
set _verb=0
set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _Hook="%SysPath%\SppExtComObjHook.dll"
set _Hops=%SysPath%\SppExtComObjHook.dll
set w7inf=%SystemRoot%\Migration\WTR\KMS_VL_ALL.inf
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"
set "_TaskOs=\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTaskLogon"
set "line1============================================================="
set "line2=************************************************************"
set "line3=____________________________________________________________"
set "line4=__________________________________________________"
set SSppHook=0
for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\skus') do (
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %_Nul3%
if %errorlevel% EQU 1060 set OsppHook=0

if %winbuild% LSS 9200 (set "sppxtra=%SysPath%\spp\tokens\channels") else (set "sppxtra=%SysPath%\spp\tokens\addons")
set ESU_KMS=0
set isAddon=0
for /f %%A in ('dir /b /ad %sppxtra% %_Nul6%') do (
  if %winbuild% LSS 9200 (
    if exist "%sppxtra%\%%A\*ESU-*-VL-BYPASS*.xrm-ms" set ESU_KMS=1
    if exist "%sppxtra%\%%A\*VL-DMAK*.xrm-ms" set isAddon=1
  )
  if %winbuild% GEQ 9200 (
    if exist "%sppxtra%\%%A\*ESU-*-GVLK*.xrm-ms" set ESU_KMS=1
    if exist "%sppxtra%\%%A\*Volume-MAK*.xrm-ms" set isAddon=1
  )
)
if %ESU_KMS% EQU 1 set isAddon=1
if %isAddon% EQU 1 (set "adoff=and LicenseDependsOn is NULL"&set "adonn=and LicenseDependsOn is not NULL") else (set "adoff="&set "adonn=")
set ESU_EDT=0
if %ESU_KMS% EQU 1 for %%A in (%ESUEditions%) do (
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\Security-SPP-Component-SKU-%%A\*.xrm-ms" set ESU_EDT=1
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*.xrm-ms" set ESU_EDT=1
)
:: if %ESU_EDT% EQU 1 if %winbuild% LSS 9200 set SSppHook=1
set ESU_ADD=0

if %winbuild% GEQ 9200 (
  set OSType=Win8
  set SppVer=SppExtComObj.exe
) else if %winbuild% GEQ 7600 (
  set OSType=Win7
  set SppVer=sppsvc.exe
) else (
  goto :UnsupportedVersion
)
if %OSType% EQU Win8 reg query "%IFEO%\sppsvc.exe" %_Nul3% && (
reg delete "%IFEO%\sppsvc.exe" /f %_Nul3%
call :StopService sppsvc
)

set _uRI=%KMS_RenewalInterval%
set _uAI=%KMS_ActivationInterval%
set _dDbg=No
if %ActWindows% EQU 0 if %ActOffice% EQU 0 set ActWindows=1
if %_Debug% EQU 1 if not defined fAUR set fAUR=0&set External=0
if %Unattend% EQU 1 if not defined fAUR set fAUR=0&set External=0
if not defined fAUR if not defined rAUR goto :MainMenu
if defined rAUR (set _verb=1&cls&call :RemoveHook&goto :cCache)
call :subOffice
set Unattend=1
set _ReAR=0
set _AUR=0
if exist %_Hook% dir /b /al %_Hook% %_Nul3% || (
  reg query "%IFEO%\%SppVer%" /v VerifierFlags %_Nul3% && (set _AUR=1) || (reg query "%IFEO%\osppsvc.exe" /v VerifierFlags %_Nul3% && set _AUR=1)
)
if %fAUR% EQU 1 (set _ReAR=1&if %_AUR% EQU 0 (set _AUR=1&set _verb=1&set _rtr=DoActivate&cls&goto :InstallHook) else (set _verb=0&set _rtr=DoActivate&cls&goto :InstallHook))
if %External% EQU 0 (set _AUR=0&cls&goto :DoActivate)
cls&goto :DoActivate

:MainMenu
cls
mode con cols=80 lines=34
color 07
set "_title=KMS_VL_ALL_AIO %uivr%"
title %_title%
call :subOffice
set _dMode=Manual
set _ReAR=0
set _AUR=0
if exist %_Hook% dir /b /al %_Hook% %_Nul3% || (
  reg query "%IFEO%\%SppVer%" /v VerifierFlags %_Nul3% && (set _AUR=1&set "_dMode=Auto Renewal") || (reg query "%IFEO%\osppsvc.exe" /v VerifierFlags %_Nul3% && (set _AUR=1&set "_dMode=Auto Renewal"))
)
if %_AUR% EQU 0 (set "_dHook=Not Installed") else (set "_dHook=Installed")
if %ActWindows% EQU 0 (set _dAwin=No) else (set _dAwin=Yes)
if %ActOffice% EQU 0 (set _dAoff=No) else (set _dAoff=Yes)
if %AutoR2V% EQU 0 (set _dArtv=No) else (set _dArtv=Yes)
if %SkipKMS38% EQU 0 (set _dWXKMS=No) else (set _dWXKMS=Yes)
if %_Debug% EQU 0 (set _dDbg=No) else (set _dDbg=Yes)
if %vNextOverride% EQU 0 (set _dNxt=No) else (set _dNxt=Yes)
set _el=
set _quit=
if %preparedcolor%==0 call :colorprep
if %_NCS% EQU 0 (
pushd %_temp%
if not exist "'" (<nul >"'" set /p "=.")
)
echo.
echo           %line3%
echo.
rem echo                [1] Activate [%_dMode% Mode]
if %_AUR% EQU 1 (
call :Cfgbg %_cWht% "               [1] Activate " %_cGrn% "[%_dMode% Mode]"
) else (
call :Cfgbg %_cWht% "               [1] Activate " %_cBlu% "[%_dMode% Mode]"
)
echo.
if %_AUR% EQU 1 (
call :Cfgbg %_cWht% "               [2] Install Activation Auto-Renewal " %_cGrn% "[%_dHook%]"
) else (
echo                [2] Install Activation Auto-Renewal
)
echo                [3] Uninstall Completely
echo                %line4%
echo.
echo                    Configuration:
echo.
if %_dDbg%==No (
echo                [4] Enable Debug Mode         [%_dDbg%]
) else (
call :Cfgbg %_cWht% "               [4] Enable Debug Mode         " %_cRed% "[%_dDbg%]"
)
if %_dAwin%==Yes (
echo                [5] Process Windows           [%_dAwin%]
) else (
call :Cfgbg %_cWht% "               [5] Process Windows           " %_cYel% "[%_dAwin%]"
)
if %_dAoff%==Yes (
echo                [6] Process Office            [%_dAoff%]
) else (
call :Cfgbg %_cWht% "               [6] Process Office            " %_cYel% "[%_dAoff%]"
)
if %_dArtv%==Yes (
echo                [7] Convert Office C2R-R2V    [%_dArtv%]
) else (
call :Cfgbg %_cWht% "               [7] Convert Office C2R-R2V    " %_cYel% "[%_dArtv%]"
)
if %_dNxt%==No (
if %sub_next% EQU 1 (
call :Cfgbg %_cYel% "               [V] Override Office C2R vNext " %_cYel% "[%_dNxt%]"
  ) else (
echo                [V] Override Office C2R vNext [%_dNxt%]
  )
) else (
if %sub_next% EQU 1 (
call :Cfgbg %_cYel% "               [V] Override Office C2R vNext " %_cRed% "[%_dNxt%]"
  ) else (
echo                [V] Override Office C2R vNext [%_dNxt%]
  )
)
if %winbuild% GEQ 10240 (
if %_dWXKMS%==Yes (
echo                [X] Skip Windows KMS38        [%_dWXKMS%]
) else (
call :Cfgbg %_cWht% "               [X] Skip Windows KMS38        " %_cYel% "[%_dWXKMS%]"
))
echo                %line4%
echo.
echo                    Miscellaneous:
echo.
echo                [8] Check Activation Status
echo                [S] Create $OEM$ Folder
echo                [D] Decode Embedded Binary Files
echo                [R] Read Me
echo                [E] Activate {External Mode}
echo           %line3%
echo.
if %_NCS% EQU 0 (
popd
)
choice /c 1234567890EDRSVX /n /m ">           Choose a menu option, or press 0 to Exit: "
set _el=%errorlevel%
if %_el%==16 if %winbuild% GEQ 10240 (if %SkipKMS38% EQU 0 (set SkipKMS38=1) else (set SkipKMS38=0))&goto :MainMenu
if %_el%==15 (if %vNextOverride% EQU 0 (set vNextOverride=1) else (set vNextOverride=0))&goto :MainMenu
if %_el%==14 (call :CreateOEM)&goto :MainMenu
if %_el%==13 (call :CreateReadMe)&goto :MainMenu
if %_el%==12 (call :CreateBIN)&goto :MainMenu
if %_el%==11 goto :E_IP
if %_el%==10 (set _quit=1&goto :TheEnd)
if %_el%==9 (call :casWm)&goto :MainMenu
if %_el%==8 (call :casWm)&goto :MainMenu
if %_el%==7 (if %AutoR2V% EQU 0 (set AutoR2V=1) else (set AutoR2V=0))&goto :MainMenu
if %_el%==6 (if %ActOffice% EQU 0 (set ActOffice=1) else (set ActWindows=1&set ActOffice=0))&goto :MainMenu
if %_el%==5 (if %ActWindows% EQU 0 (set ActWindows=1) else (set ActWindows=0&set ActOffice=1))&goto :MainMenu
if %_el%==4 (if %_Debug% EQU 0 (set _Debug=1) else (set _Debug=0))&goto :MainMenu
if %_el%==3 (if %_dDbg%==No (set _verb=1&cls&call :RemoveHook&goto :cCache) else (set _verb=1&cls&goto :RemoveHook))
if %_el%==2 (set _ReAR=1&if %_AUR% EQU 0 (set _AUR=1&set _verb=1&set _rtr=DoActivate&cls&goto :InstallHook) else (set _verb=0&set _rtr=DoActivate&cls&goto :InstallHook))
if %_el%==1 (cls&goto :DoActivate)
goto :MainMenu

:colorprep
set preparedcolor=1

if %_NCS% EQU 1 (
for /f "tokens=1,2 delims=#" %%A in ('"prompt #$H#$E# & echo on & for %%B in (1) do rem"') do set _EC=%%B

set "_cBlu="44;97m""
set "_cRed="40;91m""
set "_cGrn="40;92m""
set "_cYel="40;93m""
set "_cWht="40;37m""
exit /b
)

for /f %%A in ('"prompt $H&for %%B in (1) do rem"') do set "_BS=%%A %%A"

set "_cBlu="1F""
set "_cRed="0C""
set "_cGrn="0A""
set "_cYel="0E""
set "_cWht="07""
exit /b

:Cfgbg
if %_NCS% EQU 1 (
echo %_EC%[%~1%~2%_EC%[%~3%~4%_EC%[0m
exit /b
)
setlocal
set "s=%~2"
set "t=%~4"
call :Pfgbg %1 s %3 t
exit /b

:Pfgbg
setlocal EnableDelayedExpansion
set "s=!%~2!"
set "t=!%~4!"
for /f delims^=^ eol^= %%i in ("!s!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~1 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
setlocal EnableDelayedExpansion
for /f delims^=^ eol^= %%i in ("!t!") do (
  if "!" equ "" setlocal DisableDelayedExpansion
    >`.txt (echo %%i\..\')
    findstr /a:%~3 /f:`.txt "."
    <nul set /p "=%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%%_BS%"
)
echo(
exit /b

:E_IP
cls
set kip=
echo.
echo Enter / Paste the external KMS Server address, or just press Enter to return:
echo.
set /p kip=
if not defined kip goto :MainMenu
set "kip=%kip: =%"
set "KMS_IP=%kip%"
set External=1
cls

:DoActivate
if %_dDbg%==Yes (
set "_para=/d"
if %ActWindows% EQU 0 set "_para=!_para! /o"
if %ActOffice% EQU 0 set "_para=!_para! /w"
if %vNextOverride% EQU 0 set "_para=!_para! /v"
if %SkipKMS38% EQU 0 set "_para=!_para! /x"
if %External% EQU 1 set "_para=!_para! /e %KMS_IP%"
if %External% EQU 0 if %_AUR% EQU 0 set "_para=!_para! /m"
if %External% EQU 0 if %_AUR% EQU 1 set "_para=!_para! /a"
goto :DoDebug
)
if %External% EQU 1 (
if "%KMS_IP%"=="%_uIP%" set External=0
)
if %External% EQU 1 (
set _AUR=1
)
if %External% EQU 0 (
set KMS_IP=%_uIP%
)
if %_AUR% EQU 0 (
set KMS_RenewalInterval=43200
set KMS_ActivationInterval=43200
) else (
set KMS_RenewalInterval=%_uRI%
set KMS_ActivationInterval=%_uAI%
)
if %External% EQU 1 (
color 8F&set "mode=External ^(%KMS_IP%^)"
) else (
if %_AUR% EQU 0 (color 1F&set "mode=Manual") else (color 07&set "mode=Auto Renewal")
)
if %Unattend% EQU 0 (
if %_Debug% EQU 0 (title %_title%) else (set "_title=KMS_VL_ALL_AIO %uivr% : %mode%"&title KMS_VL_ALL_AIO %uivr% : %mode%)
) else (
echo.
echo Running KMS_VL_ALL_AIO %uivr%
)
if %Silent% EQU 0 if %_Debug% EQU 0 (
%_Nul3% %_psc% "&%_buf%"
if %Unattend% EQU 0 title %_title%
)
if %winbuild% GEQ 9600 (
  reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoGenTicket /t REG_DWORD /d 1 /f %_Nul3%
  if %winbuild% EQU 14393 reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v NoAcquireGT /t REG_DWORD /d 1 /f %_Nul3%
)
echo.
echo Activation Mode: %mode%
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %External% EQU 0 if %_ReAR% EQU 0 (set _verb=0&set _rtr=ReturnHook&goto :InstallHook)

:ReturnHook
if %External% EQU 0 if %_AUR% EQU 1 (
call :UpdateIFEOEntry %SppVer%
call :UpdateIFEOEntry osppsvc.exe
)
if %External% EQU 1 if %_AUR% EQU 1 (
call :UpdateOSPPEntry osppsvc.exe
)

SET Win10Gov=0
SET "EditionWMI="
SET "EditionID="
IF %winbuild% LSS 14393 if %SSppHook% NEQ 0 GOTO :Main
SET "RegKey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
SET "Pattern=Microsoft-Windows-*Edition~31bf3856ad364e35"
SET "EditionPKG=FFFFFFFF"
FOR /F "TOKENS=8 DELIMS=\" %%A IN ('REG QUERY "%RegKey%" /f "%Pattern%" /k %_Nul6% ^| FIND /I "CurrentVersion"') DO (
  REG QUERY "%RegKey%\%%A" /v "CurrentState" %_Nul2% | FIND /I "0x70" %_Nul1% && (
    FOR /F "TOKENS=3 DELIMS=-~" %%B IN ('ECHO %%A') DO SET "EditionPKG=%%B"
  )
)
IF /I "%EditionPKG:~-7%"=="Edition" (
SET "EditionID=%EditionPKG:~0,-7%"
) ELSE (
FOR /F "TOKENS=3 DELIMS=: " %%A IN ('DISM /English /Online /Get-CurrentEdition %_Nul6% ^| FIND /I "Current Edition :"') DO SET "EditionID=%%A"
)
net start sppsvc /y %_Nul3%
call :qrQuery SoftwareLicensingProduct "ApplicationID='%_wApp%' %adoff% AND PartialProductKey is not NULL" LicenseFamily
FOR /F "TOKENS=2 DELIMS==" %%A IN ('%_qr% %_Nul6%') DO SET "EditionWMI=%%A"
IF "%EditionWMI%"=="" (
IF %winbuild% GEQ 17063 FOR /F "SKIP=2 TOKENS=2*" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%B"
IF %winbuild% LSS 14393 (
  FOR /F "SKIP=2 TOKENS=2*" %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%B"
  GOTO :Main
  )
)
IF NOT "%EditionWMI%"=="" SET "EditionID=%EditionWMI%"
IF /I "%EditionID%"=="IoTEnterprise" SET "EditionID=Enterprise"
IF /I "%EditionID%"=="IoTEnterpriseS" IF %winbuild% LSS 22610 (
SET "EditionID=EnterpriseS"
IF %winbuild% GEQ 19041 IF %UBR% GEQ 2788 SET "EditionID=IoTEnterpriseS"
)
IF /I "%EditionID%"=="ProfessionalSingleLanguage" SET "EditionID=Professional"
IF /I "%EditionID%"=="ProfessionalCountrySpecific" SET "EditionID=Professional"
IF /I "%EditionID%"=="EnterpriseG" SET Win10Gov=1
IF /I "%EditionID%"=="EnterpriseGN" SET Win10Gov=1

:Main
if defined EditionID (set "_winos=Windows %EditionID% edition") else (set "_winos=Detected Windows")
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %_Nul6%') do if not errorlevel 1 set "_winos=%%b"
set "nKMS=does not support KMS activation..."
set "nEval=Evaluation Editions cannot be activated. Please install full Windows OS."
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum" set _eval=1
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalEdition~*.mum" set "nEval=Server Evaluation cannot be activated. Please convert to full Server OS."
if exist "%SystemRoot%\Servicing\Packages\Microsoft-Windows-Server*EvalCorEdition~*.mum" set _eval=1&set "nEval=Server Evaluation cannot be activated. Please convert to full Server OS."
set "_C16R="
reg query %_onat%\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
reg query %_onat%\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=%_onat%\ClickToRun\Configuration"
)
if not defined _C16R reg query %_owow%\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
reg query %_owow%\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=%_owow%\ClickToRun\Configuration"
)
set "_C15R="
reg query %_onat%\15.0\ClickToRun /v InstallPath %_Nul3% && for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
reg query %_onat%\15.0\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && call set "_C15R=%_onat%\15.0\ClickToRun\Configuration"
if not defined _C15R reg query %_onat%\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && call set "_C15R=%_onat%\15.0\ClickToRun\propertyBag"
)
set "_C14R="
if %_wow%==0 (reg query %_onat%\14.0\CVH /f Click2run /k %_Nul3% && set "_C14R=1") else (reg query %_owow%\14.0\CVH /f Click2run /k %_Nul3% && set "_C14R=1")
for %%A in (14,15,16,19,21,24) do call :officeLoc %%A
if %_O14MSI% EQU 1 set "_C14R="

set S_OK=1
call :RunSPP
if %ActOffice% NEQ 0 call :RunOSPP
if %ActOffice% EQU 0 (echo.&echo Office activation is OFF...)
if %S_OK% EQU 0 if %External% EQU 0 call :CheckFR

if exist "!_temp!\crv*.txt" del /f /q "!_temp!\crv*.txt"
if exist "!_temp!\*chk.txt" del /f /q "!_temp!\*chk.txt"
if exist "!_temp!\slmgr.vbs" del /f /q "!_temp!\slmgr.vbs"
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc

if %_AUR% EQU 0 call :RemoveHook

set "d1=$t=[AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2, $False).DefineType(0);"
set "d2=[void]$t.DefinePInvokeMethod('SLpTriggerServiceWorker', 'sppc.dll', 22, 1, [Int32], @([UInt32], [IntPtr], [String], [UInt32]), 1, 3);"
set "d3=[void]$t.CreateType()::SLpTriggerServiceWorker(0, 0, 'reeval', 0);"
if %winbuild% GEQ 9200 (
if %_pwsh% equ 1 %_psc% "!d1! !d2! !d3!"
if %_pwsh% equ 0 sc start sppsvc trigger=reeval;sessionid=0 %_Nul3%
)

if %_verb% EQU 1 (
echo.&echo %line3%&echo.
if %External% EQU 0 if "%_rtr%"=="DoActivate" (
echo.
echo Make sure to exclude this file in the Antivirus protection.
echo %SystemRoot%\System32\SppExtComObjHook.dll)
)
set External=0
set KMS_IP=%_uIP%
if %uManual% EQU 1 timeout 5
if %uAutoRenewal% EQU 1 timeout 5
if %Unattend% NEQ 0 goto :TheEnd
echo.
echo Press any key to continue . . .
pause >nul
goto :MainMenu

:RunSPP
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set W1nd0ws=1
set WinPerm=0
set WinVL=0
set Off1ce=0
set RanR2V=0
for %%A in (15,16,19,21,24) do set aC2R%%A=0
if %winbuild% GEQ 9200 if %ActOffice% NEQ 0 call :sppoff
call :qrQuery %spp% "Description like '%%%%KMSCLIENT%%%%'" Name
%_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set WinVL=1)
if %WinVL% EQU 0 (
if %ActWindows% EQU 0 (
  echo.&echo Windows activation is OFF...
  ) else (
  if %SSppHook% EQU 0 (
    echo.&echo %_winos% %nKMS%
    if defined _eval echo %nEval%
    ) else (
    echo.&echo Error: Failed checking KMS Activation ID^(s^) for Windows.&echo Either sppsvc service or SppExtComObjHook.dll is not functional.&call :CheckWS
    exit /b
    )
  )
)
if %WinVL% EQU 0 if %Off1ce% EQU 0 exit /b
if %_AUR% EQU 0 (
reg delete "HKLM\%SPPk%\%_wApp%" /f %_Null%
rem reg delete "HKLM\%SPPk%\%_oApp%" /f %_Null%
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f %_Null%
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f %_Null%
)
set _gvlk=0
call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%' and PartialProductKey is not NULL" Name
if %winbuild% GEQ 10240 %_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
set gpr=0
call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%' and PartialProductKey is not NULL" GracePeriodRemaining
if %winbuild% GEQ 10240 if %SkipKMS38% NEQ 0 if %_gvlk% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set "gpr=%%A"
call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%' and PartialProductKey is not NULL" LicenseFamily
if %gpr% NEQ 0 if %gpr% GTR 259200 (
set W1nd0ws=0
%_qr% %_Nul2% | findstr /i EnterpriseG %_Nul1% && (call set W1nd0ws=1)
)
call :qrSingle %sps% Version
for /f "tokens=2 delims==" %%A in ('%_qr%') do set slsv=%%A
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32 %_Nul3%
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" /reg:32 %_Nul3%
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32 %_Null%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" /reg:32 %_Nul3%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" /reg:32 %_Nul3%
)
reg delete "HKLM\%SPPk%\%_oApp%" /f %_Null%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
)
call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%'" ID
if %W1nd0ws% EQU 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkwin)
call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%' %adoff%" ID
if %W1nd0ws% EQU 1 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkwin)
:: call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%' %adonn%" ID
:: if %ESU_EDT% EQU 1 if %winbuild% GEQ 9200 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :esu8chk)
:: if %ESU_EDT% EQU 1 if %winbuild% LSS 9200 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :esu7chk)
if %W1nd0ws% EQU 1 if %ActWindows% EQU 0 (echo.&echo Windows activation is OFF...)
call :qrQuery %spp% "ApplicationID='%_oApp%' and Description like '%%%%KMSCLIENT%%%%'" ID
if %Off1ce% EQU 1 if %ActOffice% NEQ 0 for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkoff 1)
if %_AUR% EQU 0 (
call :cREG %_Nul3%
) else (
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing %_Null%
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching %_Null%
)
exit /b

:sppoff
set OffUWP=0
if %winbuild% GEQ 10240 reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msoxmled.exe" %_Nul3% && (
dir /b "%ProgramFiles%\WindowsApps\Microsoft.Office.Desktop*" %_Nul3% && set OffUWP=1
if not %xOS%==x86 dir /b "%ProgramW6432%\WindowsApps\Microsoft.Office.Desktop*" %_Nul3% && set OffUWP=1
)
rem nothing installed
if %loc_off24% EQU 0 if %loc_off21% EQU 0 if %loc_off19% EQU 0 if %loc_off16% EQU 0 if %loc_off15% EQU 0 (
if %winbuild% GEQ 9200 (
  if %OffUWP% EQU 0 (echo.&echo No Installed Office 2013-2024 Product Detected...) else (echo.&echo %_mOuwp%)
  exit /b
  )
if %winbuild% LSS 9200 (if %loc_off14% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b))
)
if %vNextOverride% EQU 1 if %AutoR2V% EQU 1 (
set sub_o365=0
set sub_proj=0
set sub_vsio=0
if %sub_next% EQU 1 (
  reg delete HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing /f %_Nul3%
  rmdir /s /q "!_Local!\Microsoft\Office\Licenses\" %_Nul3%
  rmdir /s /q "!ProgramData!\Microsoft\Office\Licenses\" %_Nul3%
  )
)
set Off1ce=1
set _sC2R=sppoff
set _fC2R=ReturnSPP

call :qrQuery %spp% "Description like '%%%%KMSCLIENT%%%%' AND NOT Name like '%%%%MondoR_KMS_Automation%%%%'" Name
%_qr% > "!_temp!\sppchk.txt" 2>&1
for %%A in (14,15,16,19,21,24) do (
set vol_off%%A=0
if !loc_off%%A! EQU 1 find /i "Office %%A" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off%%A=1)
)
call :qrQuery %spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%%%'" LicenseFamily
if %vol_off16% EQU 1 find /i "Office16MondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set vol_off16=0)
)
call :qrQuery %spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%%%'" LicenseFamily
if %vol_off15% EQU 1 find /i "OfficeMondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% || (set vol_off15=0)
)

call :qrQuery %spp% "ApplicationID='%_oApp%' AND NOT Name like '%%%%O365%%%%'" Name
%_qr% > "!_temp!\sppchk.txt" 2>&1
for %%A in (14,15,16,19,21,24) do (
set ret_off%%A=0
find /i "R_Retail" "!_temp!\sppchk.txt" %_Nul2% | find /i "Office %%A" %_Nul1% && (set ret_off%%A=1)
)
call :qrQuery %spp% "ApplicationID='%_oA14%'" Description
if %winbuild% LSS 9200 if %vol_off14% EQU 0 %_qr% %_Nul2% | find /i "channel" %_Nul1% && (set ret_off14=1)

set run_off24=0&set prr_off24=0&set prv_off24=0
if %loc_off24% EQU 1 if %ret_off24% EQU 1 if %_O16MSI% EQU 0 if %vol_off24% EQU 0 set run_off24=1
if %loc_off24% EQU 1 if %ret_off24% EQU 1 if %_O16MSI% EQU 0 if %vol_off24% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office24%%a2024R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off24+=1
  find /i "Office24%%a2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off24+=1
  )
for %%a in (Professional) do find /i "Office24%%a2024R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off24+=1
  find /i "Office24ProPlus2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off24+=1
  )
for %%a in (HomeBusiness,HomeStudent,Home) do find /i "Office24%%a2024R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off24+=1
  find /i "Office24Standard2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off24+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office24%%a2024R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off24+=1
  find /i "Office24%%a2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off24+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office24%%a2024R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off24+=1
  find /i "Office24%%a2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off24+=1
  )
)
if %loc_off24% EQU 1 if %ret_off24% EQU 1 if %_O16MSI% EQU 0 if %vol_off24% EQU 1 if %prv_off24% LSS %prr_off24% (set vol_off24=0&set run_off24=1)

set run_off21=0&set prr_off21=0&set prv_off21=0
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 0 set run_off21=1
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
for %%a in (Professional) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21ProPlus2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21Standard2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office21%%a2021R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off21+=1
  find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off21+=1
  )
)
if %loc_off21% EQU 1 if %ret_off21% EQU 1 if %_O16MSI% EQU 0 if %vol_off21% EQU 1 if %prv_off21% LSS %prr_off21% (set vol_off21=0&set run_off21=1)

set run_off19=0&set prr_off19=0&set prv_off19=0
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 0 set run_off19=1
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 1 (
for %%a in (%DO16Ids%) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
for %%a in (Professional) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19ProPlus2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19Standard2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office19%%a2019R" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off19+=1
  find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off19+=1
  )
)
if %loc_off19% EQU 1 if %ret_off19% EQU 1 if %_O16MSI% EQU 0 if %vol_off19% EQU 1 if %prv_off19% LSS %prr_off19% (set vol_off19=0&set run_off19=1)

set run_off16=0&set prr_off16=0&set prv_off16=0
if %loc_off16% EQU 1 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if defined _C16R (
for %%a in (%DO16Ids%) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off24% EQU 0 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off24% EQU 1 find /i "Office24%%a2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
for %%a in (Professional) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off24% EQU 0 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16ProPlusVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off24% EQU 1 find /i "Office24ProPlus2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21ProPlus2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19ProPlus2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off24% EQU 0 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16StandardVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off24% EQU 1 find /i "Office24Standard2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21Standard2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19Standard2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off24% EQU 0 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off24% EQU 1 find /i "Office24%%a2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office16%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off16+=1
  if %vol_off16% EQU 1 if %vol_off24% EQU 0 if %vol_off21% EQU 0 if %vol_off19% EQU 0 find /i "Office16%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off24% EQU 1 find /i "Office24%%a2024VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off21% EQU 1 find /i "Office21%%a2021VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  if %vol_off16% EQU 0 if %vol_off19% EQU 1 find /i "Office19%%a2019VL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off16+=1
  )
)
if %loc_off16% EQU 1 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if defined _C16R if %prv_off16% LSS %prr_off16% (set vol_off16=0&set run_off16=1)
call :qrQuery %spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%%%'" LicenseFamily
if %loc_off16% EQU 1 if %run_off16% EQU 0 if %sub_o365% EQU 0 if defined _C16R %_qr% %_Nul2% | find /i "O365" %_Nul1% && (
find /i "Office16MondoVL" "!_temp!\sppchk.txt" %_Nul1% || set run_off16=1
)

set run_off15=0&set prr_off15=0&set prv_off15=0
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 0 if defined _C15R set run_off15=1
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 1 if defined _C15R (
for %%a in (%DO15Ids%) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
for %%a in (Professional) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "OfficeProPlusVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
for %%a in (HomeBusiness,HomeStudent) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "OfficeStandardVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
if %sub_proj% EQU 0 for %%a in (ProjectPro,ProjectStd) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
if %sub_vsio% EQU 0 for %%a in (VisioPro,VisioStd) do find /i "Office%%aR" "!_temp!\sppchk.txt" %_Nul1% && (
  call set /a prr_off15+=1
  find /i "Office%%aVL" "!_temp!\sppchk.txt" %_Nul1% && call set /a prv_off15+=1
  )
)
if %loc_off15% EQU 1 if %ret_off15% EQU 1 if %_O15MSI% EQU 0 if %vol_off15% EQU 1 if defined _C15R if %prv_off15% LSS %prr_off15% (set vol_off15=0&set run_off15=1)
call :qrQuery %spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%%%'" LicenseFamily
if %loc_off15% EQU 1 if %run_off15% EQU 0 if defined _C15R %_qr% %_Nul2% | find /i "O365" %_Nul1% && (
find /i "OfficeMondoVL" "!_temp!\sppchk.txt" %_Nul1% || set run_off15=1
)

set vol_offgl=1
if %vol_off24% EQU 0 if %vol_off21% EQU 0 if %vol_off19% EQU 0 if %vol_off16% EQU 0 if %vol_off15% EQU 0 (
if %winbuild% GEQ 9200 set vol_offgl=0
if %winbuild% LSS 9200 if %vol_off14% EQU 0 set vol_offgl=0
)
rem mixed Volume + Retail
if %run_off24% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off21% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off19% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off16% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
if %run_off15% EQU 1 if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
rem all supported Volume + message for unsupported
if %loc_off16% EQU 0 if %ret_off16% EQU 1 if %_O16MSI% EQU 0 if %OffUWP% EQU 1 (echo.&echo %_mOuwp%)
if %vol_offgl% EQU 1 (
if %ret_off16% EQU 1 if %_O16MSI% EQU 1 (echo.&echo %_mO16m%)
if %ret_off15% EQU 1 if %_O15MSI% EQU 1 (echo.&echo %_mO15m%)
if %winbuild% LSS 9200 if %loc_off14% EQU 1 if %vol_off14% EQU 0 (if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%))
exit /b
)
set Off1ce=0
rem Retail C2R
if %AutoR2V% EQU 1 if %RanR2V% EQU 0 goto :C2RR2V
:ReturnSPP
rem Retail MSI/C2R or failed C2R-R2V
if %loc_off24% EQU 1 if %vol_off24% EQU 0 (
if %aC2R24% EQU 1 (echo.&echo %_mO24a%) else (echo.&echo %_mO24c%)
)
if %loc_off21% EQU 1 if %vol_off21% EQU 0 (
if %aC2R21% EQU 1 (echo.&echo %_mO21a%) else (echo.&echo %_mO21c%)
)
if %loc_off19% EQU 1 if %vol_off19% EQU 0 (
if %aC2R19% EQU 1 (echo.&echo %_mO19a%) else (echo.&echo %_mO19c%)
)
if %loc_off16% EQU 1 if %vol_off16% EQU 0 (
if defined _C16R (if %aC2R16% EQU 1 (echo.&echo %_mO16a%) else (if %sub_o365% EQU 0 echo.&echo %_mO16c%)) else if %_O16MSI% EQU 1 (if %ret_off16% EQU 1 echo.&echo %_mO16m%)
)
if %loc_off15% EQU 1 if %vol_off15% EQU 0 (
if defined _C15R (if %aC2R15% EQU 1 (echo.&echo %_mO15a%) else (echo.&echo %_mO15c%)) else if %_O15MSI% EQU 1 (if %ret_off15% EQU 1 echo.&echo %_mO15m%)
)
if %winbuild% LSS 9200 if %loc_off14% EQU 1 if %vol_off14% EQU 0 (
if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%)
)
exit /b

:sppchkoff
call :qrQuery %spp% "ID='%app%'" Name
%_qr% > "!_temp!\sppchk.txt"
set _eof=0
for %%A in (14,15,16,19,21,24) do (
find /i "Office %%A" "!_temp!\sppchk.txt" %_Nul1% && (if !loc_off%%A! EQU 0 set _eof=1)
)
if %_eof% EQU 1 exit /b
if %1 EQU 1 (set _officespp=1) else (set _officespp=0)
rem call :qrQuery %spp% "ID='%app%'" Name
for /f "tokens=3 delims==, " %%G in ('%_qr%') do set OffVer=%%G
call :qrQuery %spp% "PartialProductKey is not NULL" ID

:: skip activating 2024 Preview ids if 2024 RTM keys installed
if /i '%app%' EQU 'fceda083-1203-402a-8ec4-3d7ed9f3648c' (
%_qr% %_Nul2% | findstr /i "8d368fc1-9470-4be2-8d66-90e836cbb051" %_Nul3% && (exit /b)
)
if /i '%app%' EQU 'aaea0dc8-78e1-4343-9f25-b69b83dd1bce' (
%_qr% %_Nul2% | findstr /i "f510af75-8ab7-4426-a236-1bfb95c34ff8" %_Nul3% && (exit /b)
)
if /i '%app%' EQU '4ab4d849-aabc-43fb-87ee-3aed02518891' (
%_qr% %_Nul2% | findstr /i "fa187091-8246-47b1-964f-80a0b1e5d69a" %_Nul3% && (exit /b)
)

%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)

:: skip installing 2024 RTM keys if 2024 Preview keys installed
if /i '%app%' EQU '8d368fc1-9470-4be2-8d66-90e836cbb051' (
%_qr% %_Nul2% | findstr /i "fceda083-1203-402a-8ec4-3d7ed9f3648c" %_Nul3% && (exit /b)
)
if /i '%app%' EQU 'f510af75-8ab7-4426-a236-1bfb95c34ff8' (
%_qr% %_Nul2% | findstr /i "aaea0dc8-78e1-4343-9f25-b69b83dd1bce" %_Nul3% && (exit /b)
)
if /i '%app%' EQU 'fa187091-8246-47b1-964f-80a0b1e5d69a' (
%_qr% %_Nul2% | findstr /i "4ab4d849-aabc-43fb-87ee-3aed02518891" %_Nul3% && (exit /b)
)

call :offchk%OffVer%
exit /b

:sppchkwin
set _officespp=0
call :qrQuery %spp% "ApplicationID='%_wApp%' and Description like '%%%%KMSCLIENT%%%%' and PartialProductKey is not NULL" Name
if %winbuild% GEQ 14393 if %WinPerm% EQU 0 if %_gvlk% EQU 0 %_qr% %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
call :qrQuery %spp% "ID='%app%'" LicenseStatus
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
call :qrQuery %spp% "PartialProductKey is not NULL" ID
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
if %winbuild% GEQ 14393 if %_gvlk% EQU 1 exit /b
if %WinPerm% EQU 1 exit /b
if %winbuild% LSS 10240 (call :winchk&exit /b)
set _eof=0
for %%A in (
b71515d9-89a2-4c60-88c8-656fbcca7f3a,af43f7f0-3b1e-4266-a123-1fdb53f4323b,075aca1f-05d7-42e5-a3ce-e349e7be7078
11a37f09-fb7f-4002-bd84-f3ae71d11e90,43f2ab05-7c87-4d56-b27c-44d0f9a3dabd,2cf5af84-abab-4ff0-83f8-f040fb2576eb
6ae51eeb-c268-4a21-9aae-df74c38b586d,ff808201-fec6-4fd4-ae16-abbddade5706,34260150-69ac-49a3-8a0d-4a403ab55763
4dfd543d-caa6-4f69-a95f-5ddfe2b89567,5fe40dd6-cf1f-4cf2-8729-92121ac2e997,903663f7-d2ab-49c9-8942-14aa9e0a9c72
2cc171ef-db48-4adc-af09-7c574b37f139,5b2add49-b8f4-42e0-a77c-adad4efeeeb1
) do (
if /i "%app%" EQU "%%A" set _eof=1
)
if %_eof% EQU 1 exit /b
if not defined EditionID (call :winchk&exit /b)
if %winbuild% LSS 14393 (call :winchk&exit /b)
if /i '%app%' EQU '32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee' if /i %EditionID% NEQ EnterpriseS exit /b
if /i '%app%' EQU 'ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69' if /i %EditionID% NEQ CloudEdition exit /b
if /i '%app%' EQU 'd30136fc-cb4b-416e-a23d-87207abc44a9' if /i %EditionID% NEQ CloudEditionN exit /b
if /i '%app%' EQU '0df4f814-3f57-4b8b-9a9d-fddadcd69fac' if /i %EditionID% NEQ CloudE exit /b
if /i '%app%' EQU 'e0c42288-980c-4788-a014-c080d2e1926e' if /i %EditionID% NEQ Education exit /b
if /i '%app%' EQU '73111121-5638-40f6-bc11-f1d7b0d64300' if /i %EditionID% NEQ Enterprise exit /b
if /i '%app%' EQU '2de67392-b7a7-462a-b1ca-108dd189f588' if /i %EditionID% NEQ Professional exit /b
if /i '%app%' EQU '3f1afc82-f8ac-4f6c-8005-1d233e606eee' if /i %EditionID% NEQ ProfessionalEducation exit /b
if /i '%app%' EQU '82bbc092-bc50-4e16-8e18-b74fc486aec3' if /i %EditionID% NEQ ProfessionalWorkstation exit /b
if /i '%app%' EQU '3c102355-d027-42c6-ad23-2e7ef8a02585' if /i %EditionID% NEQ EducationN exit /b
if /i '%app%' EQU 'e272e3e2-732f-4c65-a8f0-484747d0d947' if /i %EditionID% NEQ EnterpriseN exit /b
if /i '%app%' EQU 'a80b5abf-76ad-428b-b05d-a47d2dffeebf' if /i %EditionID% NEQ ProfessionalN exit /b
if /i '%app%' EQU '5300b18c-2e33-4dc2-8291-47ffcec746dd' if /i %EditionID% NEQ ProfessionalEducationN exit /b
if /i '%app%' EQU '4b1571d3-bafb-4b40-8087-a961be2caf65' if /i %EditionID% NEQ ProfessionalWorkstationN exit /b
if /i '%app%' EQU '58e97c99-f377-4ef1-81d5-4ad5522b5fd8' if /i %EditionID% NEQ Core exit /b
if /i '%app%' EQU 'cd918a57-a41b-4c82-8dce-1a538e221a83' if /i %EditionID% NEQ CoreSingleLanguage exit /b
if /i '%app%' EQU 'ec868e65-fadf-4759-b23e-93fe37f2cc29' if /i %EditionID% NEQ ServerRdsh exit /b
if /i '%app%' EQU 'e4db50ea-bda1-4566-b047-0ca50abc6f07' if /i %EditionID% NEQ ServerRdsh exit /b
call :qrQuery %spp% "Description like '%%%%KMSCLIENT%%%%'" ID
if /i "%app%" EQU "e4db50ea-bda1-4566-b047-0ca50abc6f07" (
%_qr% | findstr /i "ec868e65-fadf-4759-b23e-93fe37f2cc29" %_Nul3% && (exit /b)
)
call :winchk
exit /b

:winchk
if not defined tok (if %winbuild% GEQ 9200 (set "tok=4") else (set "tok=7"))
call :qrQuery %spp% "LicenseStatus='1' and Description like '%%%%KMSCLIENT%%%%' %adoff%" Name
%_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && (exit /b)
echo.
call :qrQuery %spp% "LicenseStatus='1' and GracePeriodRemaining='0' %adoff% and PartialProductKey is not NULL" Name
%_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && (
set WinPerm=1
)
set WinOEM=0
call :qrQuery %spp% "ApplicationID='%_wApp%' and LicenseStatus='1' %adoff%" Name
if %WinPerm% EQU 0 %_qr% %_Nul2% | findstr /i "Windows" %_Nul3% && set WinOEM=1
call :qrQuery %spp% "ApplicationID='%_wApp%' and LicenseStatus='1' %adoff%" Description
if %WinOEM% EQU 1 (
for /f "tokens=%tok% delims=, " %%G in ('%_qr%') do set "channel=%%G"
for %%A in (VOLUME_MAK, RETAIL, OEM_DM, OEM_SLP, OEM_COA, OEM_COA_SLP, OEM_COA_NSLP, OEM_NONSLP, OEM) do if /i "%%A"=="!channel!" set WinPerm=1
)
if %WinPerm% EQU 0 (
copy /y %SysPath%\slmgr.vbs "!_temp!\slmgr.vbs" %_Nul3%
cscript.exe //NoLogo "!_temp!\slmgr.vbs" /xpr %_Nul2% | findstr /i "permanently" %_Nul3% && set WinPerm=1
)
call :qrQuery %spp% "ApplicationID='%_wApp%' and LicenseStatus='1' %adoff%" Name
if %WinPerm% EQU 1 (
for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Checking: %%x
echo Product is Permanently Activated.
exit /b
)
call :insKey
exit /b

:esuchk
set ls=0
call :qrQuery %spp% "ID='%~1'" LicenseStatus
for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls=%%A
if "%ls%"=="1" (
echo Checking: %~2
echo Product is Permanently Activated.
exit /b
)
call :qrQuery %spp% "PartialProductKey is not NULL" ID
%_qr% %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
call :insKey
exit /b

:esu8chk
set _officespp=0
set ESU_ADD=1
call :qrQuery %spp% "ID='%app%'" LicenseStatus
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
for %%# in (
f57b5b6b-80c2-46e4-ae9d-9fe98e032cb7:c0a2ea62-12ad-435b-ab4f-c9bfab48dbc4:Server-ESU-Year1
b1b1ef19-a088-4962-aedb-2a647a891104:e3e2690b-931c-4c80-b1ff-dffba8a81988:Server-ESU-Year2
1a716f14-0607-425f-a097-5f2f1f091315:55b1dd2d-2209-4ea0-a805-06298bad25b3:Server-ESU-Year3
) do for /f "tokens=1-3 delims=:" %%A in ("%%#") do (
if /i "%app%" EQU "%%A" call :esuchk "%%B" "%%C"
)
exit /b

:esu7chk
set _officespp=0
set ESU_ADD=1
call :qrQuery %spp% "ID='%app%'" LicenseStatus
%_qr% %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
for %%# in (
e7cce015-33d6-41c1-9831-022ba63fe1da:8e7bfb1e-acc1-4f56-abae-b80fce56cd4b:Server-ESU-PA
f2b21bfc-a6b0-4413-b4bb-9f06b55f2812:553673ed-6ddf-419c-a153-b760283472fd:Server-ESU-Year1
bfc078d0-8c7f-475c-8519-accc46773113:04fa0286-fa74-401e-bbe9-fbfbb158010d:Server-ESU-Year2
23c6188f-c9d8-457e-81b6-adb6dacb8779:16c08c85-0c8b-4009-9b2b-f1f7319e45f9:Server-ESU-Year3
1f631693-b509-4624-8715-83d2a0020395:32163ff8-e96d-40b1-973c-44b9bf096d83:Server-ESU-Year4
c5a229b1-b1ab-47c5-aa5b-3da35bfc5f0c:338554b7-6f0a-48b6-aef3-ff6b42f1398b:Server-ESU-Year5
0f40a233-f300-4109-8394-ee0259811566:0ed79ef1-052c-4362-9cdc-5141491d06dd:Server-ESU-Year6
3fcc2df2-f625-428d-909a-1f76efc849b6:77db037b-95c3-48d7-a3ab-a9c6d41093e0:Client-ESU-Year1
dadfcd24-6e37-47be-8f7f-4ceda614cece:0e00c25d-8795-4fb7-9572-3803d91b6880:Client-ESU-Year2
0c29c85e-12d7-4af8-8e4d-ca1e424c480c:4220f546-f522-46df-8202-4d07afd26454:Client-ESU-Year3
2878c146-cf02-449e-8c7b-567546d5a68a:fb7960ec-99e3-4e57-9e3c-3412547eaa02:Client-ESU-Year4
c83ccefc-3d93-42a3-8961-dd2b93256f93:aced6a60-dd77-4f39-99ad-f1f8b97a7962:Client-ESU-Year5
de0a2119-f778-4c83-a358-d0a45ec96e05:7e94be23-b161-4956-a682-146ab291774c:Client-ESU-Year6
) do for /f "tokens=1-3 delims=:" %%A in ("%%#") do (
if /i "%app%" EQU "%%A" call :esuchk "%%B" "%%C"
)
exit /b

:RunOSPP
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
set Off1ce=0
set RanR2V=0
for %%A in (15,16,19,21,24) do set aC2R%%A=0
if %winbuild% LSS 9200 (set "aword=2010-2024") else (set "aword=2010")
if %OsppHook% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
if %winbuild% GEQ 9200 if %loc_off14% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
set err_offsvc=0
net start osppsvc /y %_Nul3% || (
sc start osppsvc %_Nul3%
if !errorlevel! EQU 1053 set err_offsvc=1
)
if %err_offsvc% EQU 1 (echo.&echo Error: osppsvc service is not running...&exit /b)
if %winbuild% GEQ 9200 call :oppoff
if %winbuild% LSS 9200 call :sppoff
if %Off1ce% EQU 0 exit /b
if %_AUR% EQU 0 (
reg delete "HKLM\%OPPk%\%_oA14%" /f %_Null%
reg delete "HKLM\%OPPk%\%_oApp%" /f %_Null%
)
set "vPrem="&set "vProf="
call :qrQuery %spp% "LicenseFamily='OfficeVisioPrem-MAK'" LicenseStatus
if %loc_off14% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set vPrem=%%A
call :qrQuery %spp% "LicenseFamily='OfficeVisioPro-MAK'" LicenseStatus
if %loc_off14% EQU 1 for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set vProf=%%A
call :qrSingle %sps% Version
for /f "tokens=2 delims==" %%A in ('%_qr%') do set slsv=%%A
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
call :qrQuery %spp% "Description like '%%%%KMSCLIENT%%%%'" ID
for /f "tokens=2 delims==" %%G in ('%_qr%') do (set app=%%G&call :sppchkoff 2)
if %_AUR% EQU 0 (
call :cREG %_Nul3%
) else (
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing %_Null%
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching %_Null%
)
exit /b

:oppoff
call :qrQuery %spp% "Description is not NULL" Description
%_qr% %_Nul2% | find /i "KMSCLIENT" %_Nul1% && (
set Off1ce=1
exit /b
)
set ret_off14=0
%_qr% %_Nul2% | find /i "channel" %_Nul1% && (set ret_off14=1)
if defined _C14R (echo.&echo %_mO14c%) else if %_O14MSI% EQU 1 (if %ret_off14% EQU 1 echo.&echo %_mO14m%)
exit /b

:offoem
set _orv=16
if "%OffVer%"=="15" set _orv=15
if "%OffVer%"=="14" exit /b
reg delete "HKLM\SOFTWARE\Microsoft\Office\%_orv%.0\Common\OEM" /f %_Null%
reg delete "HKLM\SOFTWARE\Microsoft\Office\%_orv%.0\Common\OEM" /f /reg:32 %_Null%
exit /b

:offchk
set ls=0
set ls2=0
set ls3=0
set ls4=0
call :qrQuery %spp% "LicenseFamily='Office%~1'" LicenseStatus
for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls=%%A
call :qrQuery %spp% "LicenseFamily='Office%~3'" LicenseStatus
if /i not "%~3"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls2=%%A
call :qrQuery %spp% "LicenseFamily='Office%~5'" LicenseStatus
if /i not "%~5"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls3=%%A
call :qrQuery %spp% "LicenseFamily='Office%~7'" LicenseStatus
if /i not "%~7"=="" for /f "tokens=2 delims==" %%A in ('%_qr% %_Nul6%') do set /a ls4=%%A
if "%ls4%"=="1" (
echo Checking: %~8
echo Product is Permanently Activated.
exit /b
)
if "%ls3%"=="1" (
echo Checking: %~6
echo Product is Permanently Activated.
exit /b
)
if "%ls2%"=="1" (
echo Checking: %~4
echo Product is Permanently Activated.
exit /b
)
if "%ls%"=="1" (
echo Checking: %~2
echo Product is Permanently Activated.
exit /b
)

:: skip installing 2024 Preview keys if 2024 RTM ids detected
call :qrQuery %spp% "ApplicationID='%_oApp%'" ID
if /i '%app%' EQU 'fceda083-1203-402a-8ec4-3d7ed9f3648c' (
%_qr% %_Nul2% | findstr /i "8d368fc1-9470-4be2-8d66-90e836cbb051" %_Nul3% && (exit /b)
)
if /i '%app%' EQU 'aaea0dc8-78e1-4343-9f25-b69b83dd1bce' (
%_qr% %_Nul2% | findstr /i "f510af75-8ab7-4426-a236-1bfb95c34ff8" %_Nul3% && (exit /b)
)
if /i '%app%' EQU '4ab4d849-aabc-43fb-87ee-3aed02518891' (
%_qr% %_Nul2% | findstr /i "fa187091-8246-47b1-964f-80a0b1e5d69a" %_Nul3% && (exit /b)
)

call :insKey
exit /b

:offchk24
if /i '%app%' EQU 'fceda083-1203-402a-8ec4-3d7ed9f3648c' (
call :offchk "24ProPlus2024PreviewVL_MAK_AE" "Office ProPlus 2024 Preview"
exit /b
)
if /i '%app%' EQU 'aaea0dc8-78e1-4343-9f25-b69b83dd1bce' (
call :offchk "24ProjectPro2024PreviewVL_MAK_AE" "Project Pro 2024 Preview"
exit /b
)
if /i '%app%' EQU '4ab4d849-aabc-43fb-87ee-3aed02518891' (
call :offchk "24VisioPro2024PreviewVL_MAK_AE" "Visio Pro 2024 Preview"
exit /b
)
if /i '%app%' EQU '8d368fc1-9470-4be2-8d66-90e836cbb051' (
call :offchk "24ProPlus2024PreviewVL_MAK_AE" "Office ProPlus 2024 Preview" "24ProPlus2024VL_MAK_AE1" "Office ProPlus 2024" "24ProPlus2024VL_MAK_AE2" "Office ProPlus 2024" "24ProPlus2024VL_MAK_AE3" "Office ProPlus 2024"
exit /b
)
if /i '%app%' EQU 'bbac904f-6a7e-418a-bb4b-24c85da06187' (
call :offchk "24Standard2024VL_MAK_AE1" "Office Standard 2024" "24Standard2024VL_MAK_AE2" "Office Standard 2024"
exit /b
)
if /i '%app%' EQU 'f510af75-8ab7-4426-a236-1bfb95c34ff8' (
call :offchk "24ProjectPro2024PreviewVL_MAK_AE" "Project Pro 2024 Preview" "24ProjectPro2024VL_MAK_AE1" "Project Pro 2024" "24ProjectPro2024VL_MAK_AE2" "Project Pro 2024"
exit /b
)
if /i '%app%' EQU '9f144f27-2ac5-40b9-899d-898c2b8b4f81' (
call :offchk "24ProjectStd2024VL_MAK_AE" "Project Standard 2024"
exit /b
)
if /i '%app%' EQU 'fa187091-8246-47b1-964f-80a0b1e5d69a' (
call :offchk "24VisioPro2024PreviewVL_MAK_AE" "Visio Pro 2024 Preview" "24VisioPro2024VL_MAK_AE" "Visio Pro 2024"
exit /b
)
if /i '%app%' EQU '923fa470-aa71-4b8b-b35c-36b79bf9f44b' (
call :offchk "24VisioStd2024VL_MAK_AE" "Visio Standard 2024"
exit /b
)
call :insKey
exit /b

:offchk21
if /i '%app%' EQU 'f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b' exit /b
if /i '%app%' EQU '76093b1b-7057-49d7-b970-638ebcbfd873' exit /b
if /i '%app%' EQU 'a3b44174-2451-4cd6-b25f-66638bfb9046' exit /b
if /i '%app%' EQU 'fbdb3e18-a8ef-4fb3-9183-dffd60bd0984' (
call :offchk "21ProPlus2021VL_MAK_AE1" "Office ProPlus 2021" "21ProPlus2021VL_MAK_AE2" "Office ProPlus 2021"
exit /b
)
if /i '%app%' EQU '080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3' (
call :offchk "21Standard2021VL_MAK_AE" "Office Standard 2021"
exit /b
)
if /i '%app%' EQU '76881159-155c-43e0-9db7-2d70a9a3a4ca' (
call :offchk "21ProjectPro2021VL_MAK_AE1" "Project Pro 2021" "21ProjectPro2021VL_MAK_AE2" "Project Pro 2021"
exit /b
)
if /i '%app%' EQU '6dd72704-f752-4b71-94c7-11cec6bfc355' (
call :offchk "21ProjectStd2021VL_MAK_AE" "Project Standard 2021"
exit /b
)
if /i '%app%' EQU 'fb61ac9a-1688-45d2-8f6b-0674dbffa33c' (
call :offchk "21VisioPro2021VL_MAK_AE" "Visio Pro 2021"
exit /b
)
if /i '%app%' EQU '72fce797-1884-48dd-a860-b2f6a5efd3ca' (
call :offchk "21VisioStd2021VL_MAK_AE" "Visio Standard 2021"
exit /b
)
call :insKey
exit /b

:offchk19
if /i '%app%' EQU '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%app%' EQU 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%app%' EQU '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
if /i '%app%' EQU '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
call :offchk "19ProPlus2019VL_MAK_AE" "Office ProPlus 2019"
exit /b
)
if /i '%app%' EQU '6912a74b-a5fb-401a-bfdb-2e3ab46f4b02' (
call :offchk "19Standard2019VL_MAK_AE" "Office Standard 2019"
exit /b
)
if /i '%app%' EQU '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
call :offchk "19ProjectPro2019VL_MAK_AE" "Project Pro 2019"
exit /b
)
if /i '%app%' EQU '1777f0e3-7392-4198-97ea-8ae4de6f6381' (
call :offchk "19ProjectStd2019VL_MAK_AE" "Project Standard 2019"
exit /b
)
if /i '%app%' EQU '5b5cf08f-b81a-431d-b080-3450d8620565' (
call :offchk "19VisioPro2019VL_MAK_AE" "Visio Pro 2019"
exit /b
)
if /i '%app%' EQU 'e06d7df3-aad0-419d-8dfb-0ac37e2bdf39' (
call :offchk "19VisioStd2019VL_MAK_AE" "Visio Standard 2019"
exit /b
)
call :insKey
exit /b

:offchk16
if /i '%app%' EQU 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
call :offchk "16ProPlusVL_MAK" "Office ProPlus 2016"
exit /b
)
if /i '%app%' EQU 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
call :offchk "16StandardVL_MAK" "Office Standard 2016"
exit /b
)
if /i '%app%' EQU '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
call :offchk "16ProjectProVL_MAK" "Project Pro 2016"
exit /b
)
if /i '%app%' EQU 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
call :offchk "16ProjectStdVL_MAK" "Project Standard 2016"
exit /b
)
if /i '%app%' EQU '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
call :offchk "16VisioProVL_MAK" "Visio Pro 2016"
exit /b
)
if /i '%app%' EQU 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
call :offchk "16VisioStdVL_MAK" "Visio Standard 2016"
exit /b
)
if /i '%app%' EQU '829b8110-0e6f-4349-bca4-42803577788d' (
call :offchk "16ProjectProXC2RVL_MAKC2R" "Project Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU 'cbbaca45-556a-4416-ad03-bda598eaa7c8' (
call :offchk "16ProjectStdXC2RVL_MAKC2R" "Project Standard 2016 C2R"
exit /b
)
if /i '%app%' EQU 'b234abe3-0857-4f9c-b05a-4dc314f85557' (
call :offchk "16VisioProXC2RVL_MAKC2R" "Visio Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU '361fe620-64f4-41b5-ba77-84f8e079b1f7' (
call :offchk "16VisioStdXC2RVL_MAKC2R" "Visio Standard 2016 C2R"
exit /b
)
call :insKey
exit /b

:offchk15
if /i '%app%' EQU 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
call :offchk "ProPlusVL_MAK" "Office ProPlus 2013"
exit /b
)
if /i '%app%' EQU 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
call :offchk "StandardVL_MAK" "Office Standard 2013"
exit /b
)
if /i '%app%' EQU '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
call :offchk "ProjectProVL_MAK" "Project Pro 2013"
exit /b
)
if /i '%app%' EQU '427a28d1-d17c-4abf-b717-32c780ba6f07' (
call :offchk "ProjectStdVL_MAK" "Project Standard 2013"
exit /b
)
if /i '%app%' EQU 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
call :offchk "VisioProVL_MAK" "Visio Pro 2013"
exit /b
)
if /i '%app%' EQU 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
call :offchk "VisioStdVL_MAK" "Visio Standard 2013"
exit /b
)
call :insKey
exit /b

:offchk14
if /i '%app%' EQU '6f327760-8c5c-417c-9b61-836a98287e0c' (
call :offchk "ProPlus-MAK" "Office ProPlus 2010" "ProPlusAcad-MAK" "Office Professional Academic 2010"
exit /b
)
if /i '%app%' EQU '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
call :offchk "Standard-MAK" "Office Standard 2010" "StandardAcad-MAK"  "Office Standard Academic 2010"
exit /b
)
if /i '%app%' EQU 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
call :offchk "SmallBusBasics-MAK" "Office Small Business Basics 2010"
exit /b
)
if /i '%app%' EQU 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
call :offchk "ProjectPro-MAK" "Project Pro 2010"
exit /b
)
if /i '%app%' EQU '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
call :offchk "ProjectStd-MAK" "Project Standard 2010" "ProjectStd-MAK2" "Project Standard 2010"
exit /b
)
if /i '%app%' EQU '92236105-bb67-494f-94c7-7f7a607929bd' (
call :offchk "VisioPrem-MAK" "Visio Premium 2010" "VisioPro-MAK" "Visio Pro 2010"
exit /b
)
if defined vPrem exit /b
if /i '%app%' EQU 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
call :offchk "VisioPro-MAK" "Visio Pro 2010" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
if defined vProf exit /b
if /i '%app%' EQU '9ed833ff-4f92-4f36-b370-8683a4f13275' (
call :offchk "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
call :insKey
exit /b

:officeLoc
set loc_off%1=0
set _O%1MSI=0
if %1 EQU 19 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set loc_off%1=1
exit /b
)
if %1 EQU 21 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2021 %_Nul1% && set loc_off%1=1
exit /b
)
if %1 EQU 24 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2024 %_Nul1% && set loc_off%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" (
set loc_off%1=1
set _O%1MSI=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" (
set loc_off%1=1
set _O%1MSI=1
)

if %1 EQU 16 if defined _C16R (
for /f "skip=2 tokens=2*" %%a in ('reg query %_C16R% /v ProductReleaseIds') do echo %%b> "!_temp!\c2rchk.txt"
for %%a in (%LV16Ids%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
for %%a in (%LR16Ids%) do (
  findstr /I /C:"%%aRetail" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
exit /b
)

if %1 EQU 15 if defined _C15R (
set loc_off%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if not %xOS%==x86 if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if not %xOS%==x86 if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
exit /b

:subOffice
set kNext=HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext
set sub_next=0
set sub_o365=0
set sub_proj=0
set sub_vsio=0
set _Identity=0
dir /b /s /a:-d "!_Local!\Microsoft\Office\Licenses\*" %_Nul3% && (set _Identity=1&set sub_next=1)
dir /b /s /a:-d "!ProgramData!\Microsoft\Office\Licenses\*" %_Nul3% && (set _Identity=1&set sub_next=1)
if %_Identity% EQU 0 call :officeSub
exit /b

:officeSub
reg query %kNext% %_Nul3% || exit /b
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*retail" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x2" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r ".*volume" %_Nul2% | findstr /i /v "project visio" %_Nul2% | find /i "0x3" %_Nul1% && (set sub_o365=1)
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "project.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_proj=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x2" %_Nul1% && set sub_vsio=1
reg query %kNext% | findstr /i /r "visio.*" %_Nul2% | find /i "0x3" %_Nul1% && set sub_vsio=1
if %sub_o365% EQU 1 set sub_next=1
if %sub_proj% EQU 1 set sub_next=1
if %sub_vsio% EQU 1 set sub_next=1
exit /b

:officeMsg
set ov=%1
if %1 EQU 14 set ov=10
if %1 EQU 15 set ov=13
set "_mO%1a=Detected Office 20%ov% C2R Retail is activated"
set "_mO%1c=Detected Office 20%ov% C2R Retail could not be converted to Volume"
if %1 EQU 14 set "_mO%1c=Detected Office 20%ov% C2R Retail is not supported by KMS_VL_ALL"
if %1 EQU 19 exit /b
if %1 EQU 21 exit /b
if %1 EQU 24 exit /b
set "_mO%1m=Detected Office 20%ov% MSI Retail is not supported by KMS_VL_ALL"
exit /b

:insKey
set S_OK=1
echo.
set "_key="
call :qrQuery %spp% "ID='%app%'" Name
if %ESU_ADD% EQU 0 for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Installing Key: %%x
if %ESU_ADD% EQU 1 for /f "tokens=2 delims==f" %%x in ('%_qr%') do echo Installing Key: %%x
set ESU_ADD=0
call :keys %app%
if "%_key%"=="" (echo No associated KMS Client key found&exit /b)
call :qrPKey %sps% %slsv% %_key%
%_qr% %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
set S_OK=0
exit /b
)
call :qrMethod %sps% Version %slsv% RefreshLicenseStatus
if %sps% EQU SoftwareLicensingService %_qr% %_Nul3%

:activate
set S_OK=1
if %sps% EQU SoftwareLicensingService (
if %_officespp% EQU 0 (
  reg delete "HKLM\%SPPk%\%_wApp%\%app%" /f %_Null%
  ) else (
  reg delete "HKLM\%SPPk%\%_oApp%\%app%" /f %_Null%
  call :offoem
  )
if %winbuild% GEQ 9600 reg delete "HKU\S-1-5-20\%SPPk%\PersistedSystemState" /f %_Null%
) else (
reg delete "HKLM\%OPPk%\%_oA14%\%app%" /f %_Null%
reg delete "HKLM\%OPPk%\%_oApp%\%app%" /f %_Null%
call :offoem
)
call :qrQuery %spp% "ID='%app%'" Name
if %W1nd0ws% EQU 0 if %_officespp% EQU 0 if %sps% EQU SoftwareLicensingService (
reg add "HKLM\%SPPk%\%_wApp%\%app%" /f /v KeyManagementServiceName /t REG_SZ /d "127.0.0.2" %_Nul3%
reg add "HKLM\%SPPk%\%_wApp%\%app%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
reg add "HKU\S-1-5-20\%SPPk%\%_wApp%\%app%" /f /v DiscoveredKeyManagementServiceIpAddress /t REG_SZ /d "127.0.0.2" %_Nul3%
for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Checking: %%x
echo Product is KMS 2038 Activated.
exit /b
)
rem call :qrQuery %spp% "ID='%app%'" Name
if %ESU_ADD% EQU 0 for /f "tokens=2 delims==" %%x in ('%_qr%') do echo Activating: %%x
if %ESU_ADD% EQU 1 for /f "tokens=2 delims==f" %%x in ('%_qr%') do echo Activating: %%x
set ESU_ADD=0
call :qrMethod %spp% ID %app% Activate
%_qr% %_Nul3%
call set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% EQU -1073418187 (
echo Product Activation Failed: 0xC004F035
if %OSType% EQU Win7 echo Windows 7 cannot be KMS-activated on this computer due to unqualified OEM BIOS.
echo See Read Me for details.
exit /b
)
if %ERRORCODE% EQU -1073417728 (
echo Product Activation Failed: 0xC004F200
echo Windows needs to rebuild the activation-related files.
echo See KB2736303 for details.
exit /b
)
if %ERRORCODE% EQU -1073422315 (
echo Product Activation Failed: 0xC004E015
echo Running slmgr.vbs /rilc to mitigate.
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); ReinstallLicenses"
  ) else (
  cscript.exe //NoLogo //B %SysPath%\slmgr.vbs /rilc
  )
)
if %ERRORCODE% NEQ 0 (
if %sps% EQU SoftwareLicensingService (call :StopService sppsvc) else (call :StopService osppsvc)
%_qr% %_Nul3%
call set ERRORCODE=!ERRORLEVEL!
)
set gpr=0
set gpr2=0
call :qrQuery %spp% "ID='%app%'" GracePeriodRemaining
for /f "tokens=2 delims==" %%x in ('%_qr%') do (set gpr=%%x&set /a "gpr2=(%%x+1440-1)/1440")
if %ERRORCODE% EQU 0 if %gpr% EQU 0 (
echo Product Activation succeeded, but Remaining Period failed to increase.
if %OSType% EQU Win7 echo This could be related to the error described in KB4487266
exit /b
)
set Act_OK=0
if %gpr% EQU 43200 if %_officespp% EQU 0 if %winbuild% GEQ 9200 set Act_OK=1
if %gpr% EQU 64800 set Act_OK=1
if %gpr% GTR 259200 if %Win10Gov% EQU 1 set Act_OK=1
if %gpr% EQU 259200 set Act_OK=1
if %ERRORCODE% EQU 0 if %Act_OK% EQU 1 (
echo Product Activation Successful
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
exit /b
)
cmd /c exit /b %ERRORCODE%
if %ERRORCODE% NEQ 0 (
echo Product Activation Failed: 0x!=ExitCode!
) else (
echo Product Activation Failed
)
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
set S_OK=0
exit /b

:StopService
sc query %1 | find /i "STOPPED" %_Nul1% || net stop %1 /y %_Nul3%
sc query %1 | find /i "STOPPED" %_Nul1% || sc stop %1 %_Nul3%
goto :eof

:InstallHook
if %_dDbg%==Yes (
set "_para=/d /a"
if %ActWindows% EQU 0 set "_para=!_para! /o"
if %ActOffice% EQU 0 set "_para=!_para! /w"
if %vNextOverride% EQU 0 set "_para=!_para! /v"
if %SkipKMS38% EQU 0 set "_para=!_para! /x"
goto :DoDebug
)
if %_verb% EQU 1 (
if %Silent% EQU 0 if %_Debug% EQU 0 (
mode con cols=100 lines=34
%_Nul3% %_psc% "&%_buf%"
if %Unattend% EQU 0 title %_title%
)
echo.&echo %line3%&echo.
echo Installing Local KMS Emulator...
)
set "AddExc="
call :qrWD Add
if %winbuild% GEQ 9600 (
  %_qr% %_Nul3% && set "AddExc= and Windows Defender exclusion"
)
if %_verb% EQU 1 (
echo.
echo Adding File%AddExc%...
echo %SystemRoot%\System32\SppExtComObjHook.dll
)
if %_AUR% EQU 1 (
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
)
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do (
  if exist "%SysPath%\%%#" del /f /q "%SysPath%\%%#" %_Nul3%
  if exist "%SystemRoot%\SysWOW64\%%#" del /f /q "%SystemRoot%\SysWOW64\%%#" %_Nul3%
)
setlocal
set "TMP=%SystemRoot%\Temp"
set "TEMP=%SystemRoot%\Temp"
%_Nul3% %_psc% "$d='%_dllPath%';$f=[IO.File]::ReadAllText('!_batp!') -split ':embdbin\:.*';iex ($f[1]);X %_dllNum%"
endlocal
if %Unattend% EQU 0 title %_title%
if %_verb% EQU 1 (
echo.
echo Adding Registry Keys...
)
if %SSppHook% NEQ 0 call :CreateIFEOEntry %SppVer%
if %_AUR% EQU 1 (call :CreateIFEOEntry osppsvc.exe) else (if %OsppHook% NEQ 0 call :CreateIFEOEntry osppsvc.exe)
if %_AUR% EQU 1 if %OSType% EQU Win7 (
call :CreateIFEOEntry SppExtComObj.exe
if %SSppHook% NEQ 0 if not exist %w7inf% (
  if %_verb% EQU 1 (echo.&echo Adding migration fail-safe...&echo %w7inf%)
  if not exist "%SystemRoot%\Migration\WTR" md "%SystemRoot%\Migration\WTR"
  (
  echo [WTR]
  echo Name="KMS_VL_ALL"
  echo.
  echo [WTR.*]
  echo NotifyUser="No"
  echo.
  echo [System.Registry]
  echo "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sppsvc.exe [*]"
  )>%w7inf%
  )
)
if %_AUR% EQU 1 if %OSType% EQU Win8 call :CreateTask
if %_verb% EQU 1 echo.&echo %line3%&echo.
goto :%_rtr%

:RemoveHook
if %_dDbg%==Yes (
set "_para=/d /r"
goto :DoDebug
)
set "RemExc="
call :qrWD Remove
if %winbuild% GEQ 9600 (
  for %%# in (NoGenTicket,NoAcquireGT) do reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v %%# /f %_Null%
  %_qr% %_Nul3% && set "RemExc= and Windows Defender exclusions"
)
if %_verb% EQU 1 (
if %Silent% EQU 0 if %_Debug% EQU 0 (
mode con cols=100 lines=34
%_Nul3% %_psc% "&%_buf%"
if %Unattend% EQU 0 title %_title%
)
echo.&echo %line3%&echo.
echo Uninstalling Local KMS Emulator...
echo.
echo Removing Files%RemExc%...
)
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do if exist "%SysPath%\%%#" (
	if %_verb% EQU 1 echo %SystemRoot%\System32\%%#
	del /f /q "%SysPath%\%%#" %_Nul3%
)
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do if exist "%SystemRoot%\SysWOW64\%%#" (
  if %_verb% EQU 1 echo %SystemRoot%\SysWOW64\%%#
  del /f /q "%SystemRoot%\SysWOW64\%%#" %_Nul3%
)
if exist %w7inf% (
	if %_verb% EQU 1 echo %w7inf%
	del /f /q %w7inf%
)
if %_verb% EQU 1 (
echo.
echo Removing Registry Keys...
)
for %%# in (SppExtComObj.exe,sppsvc.exe,osppsvc.exe) do reg query "%IFEO%\%%#" %_Nul3% && (
  call :RemoveIFEOEntry %%#
)
if %OSType% EQU Win8 schtasks /query /tn "%_TaskEx%" %_Nul3% && (
if %_verb% EQU 1 (
echo.
echo Removing Scheduled Task...
echo %_TaskEx%
)
schtasks /delete /f /tn "%_TaskEx%" %_Nul3%
)
goto :eof

:CreateIFEOEntry
if %_verb% EQU 1 (
echo [%IFEO%\%1]
)
reg delete "%IFEO%\%1" /f /v Debugger %_Null%
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" %_Nul3%
reg add "%IFEO%\%1" /f /v VerifierDebug /t REG_DWORD /d 0x00000000 %_Nul3%
reg add "%IFEO%\%1" /f /v VerifierFlags /t REG_DWORD /d 0x80000000 %_Nul3%
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 0x00000100 %_Nul3%
reg add "%IFEO%\%1" /f /v KMS_Emulation /t REG_DWORD /d %KMS_Emulation% %_Nul3%
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% %_Nul3%
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% %_Nul3%
if /i %1 EQU SppExtComObj.exe if %winbuild% GEQ 9600 (
reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" %_Nul3%
)
goto :eof

:RemoveIFEOEntry
if %_verb% EQU 1 (
echo [%IFEO%\%1]
)
if /i %1 NEQ osppsvc.exe (
reg delete "%IFEO%\%1" /f %_Null%
goto :eof
)
if %OsppHook% EQU 0 (
reg delete "%IFEO%\%1" /f %_Null%
)
if %OsppHook% NEQ 0 for %%A in (Debugger,VerifierDlls,VerifierDebug,VerifierFlags,GlobalFlag,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /v %%A /f %_Null%
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%_uIP%" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" %_Nul3%
goto :eof

:UpdateIFEOEntry
reg add "%IFEO%\%1" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% %_Nul3%
reg add "%IFEO%\%1" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% %_Nul3%
if /i %1 EQU SppExtComObj.exe if %winbuild% GEQ 9600 (
reg add "%IFEO%\%1" /f /v KMS_HWID /t REG_QWORD /d "%KMS_HWID%" %_Nul3%
)
if /i %1 EQU sppsvc.exe (
reg add "%IFEO%\SppExtComObj.exe" /f /v KMS_ActivationInterval /t REG_DWORD /d %KMS_ActivationInterval% %_Nul3%
reg add "%IFEO%\SppExtComObj.exe" /f /v KMS_RenewalInterval /t REG_DWORD /d %KMS_RenewalInterval% %_Nul3%
)

:UpdateOSPPEntry
if /i %1 EQU osppsvc.exe (
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%KMS_IP%" %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "%KMS_Port%" %_Nul3%
)
goto :eof

:CheckFR
if not exist %_Hook% (
echo.
echo %_err%
echo File existence failed.
echo "%SystemRoot%\System32\SppExtComObjHook.dll"
echo.
echo Verify that Antivirus protection is OFF or the file path is excluded.
)

for /f "skip=1 tokens=* delims=" %%# in ('certutil -hashfile %_Hook% SHA1^|findstr /i /v CertUtil') do set "_hash=%%#"
set "_hash=%_hash: =%"
if /i not "%_hash%"=="%_orig%" (
echo.
echo %_err%
echo SHA1 hash verification failed.
echo "%SystemRoot%\System32\SppExtComObjHook.dll"
echo Expected: %_orig%
echo Detected: %_hash%
echo.
echo Verify that Antivirus protection is OFF or the file path is excluded.
)

set E_REG=0
if %SSppHook% NEQ 0 for %%A in (VerifierDlls,VerifierDebug,VerifierFlags,GlobalFlag,KMS_Emulation) do (
reg query "%IFEO%\%SppVer%" /v %%A %_Nul3% || set E_REG=1
)
if %E_REG% EQU 1 (
echo.
echo %_err%
echo Some or all required registry values are missing.
echo [%IFEO%\%SppVer%]
echo VerifierDlls, VerifierDebug, VerifierFlags, GlobalFlag, KMS_Emulation
echo.
echo Verify that Antivirus protection is OFF or the registry path is excluded.
)
set E_REG=0
if %OsppHook% NEQ 0 for %%A in (VerifierDlls,VerifierDebug,VerifierFlags,GlobalFlag,KMS_Emulation) do (
reg query "%IFEO%\osppsvc.exe" /v %%A %_Nul3% || set E_REG=1
)
if %E_REG% EQU 1 (
echo.
echo %_err%
echo Some or all required registry values are missing.
echo [%IFEO%\osppsvc.exe]
echo VerifierDlls, VerifierDebug, VerifierFlags, GlobalFlag, KMS_Emulation
echo.
echo Verify that Antivirus protection is OFF or the registry path is excluded.
)

set WMIe=0
call :CheckWS
if %WMIe% EQU 1 (
echo.
echo %_err%
echo Failed running WMI query check.
echo.
echo Verify that these services are working correctly:
echo Windows Management Instrumentation [WinMgmt]
echo Software Protection [sppsvc]
)
goto :eof

:CheckWS
call :qrCheck Win32_ComputerSystem CreationClassName SoftwareLicensingService Version
%_qrs% %_Nul2% | findstr /r "[0-9]*\.[0-9]*\.[0-9]*\.[0-9]*" %_Nul1% || (
  set WMIe=1
  %_qrw% %_Nul2% | find /i "ComputerSystem" %_Nul1% && (
    echo Error: SPP is not responding
    ) || (
    echo Error: WMI ^& SPP are not responding
  )
)
goto :eof

:cREG
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%_uIP%"
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg add "HKLM\%SPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%_uIP%" /reg:32
reg add "HKLM\%SPPk%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%_uIP%" /reg:32
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688" /reg:32
)
reg delete "HKLM\%SPPk%\%_oApp%" /f
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServiceName /t REG_SZ /d "%_uIP%"
reg add "HKLM\%SPPk%\%_oApp%" /f /v KeyManagementServicePort /t REG_SZ /d "1688"
)
if %winbuild% GEQ 9600 (
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
if %OsppHook% EQU 0 (
goto :eof
)
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d "%_uIP%"
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
goto :eof

:rREG
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%SPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%SPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%SPPk%\%_wApp%" /f
if %winbuild% GEQ 9200 (
if not %xOS%==x86 (
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName /reg:32
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort /reg:32
reg delete "HKLM\%SPPk%\%_oApp%" /f /reg:32
)
reg delete "HKLM\%SPPk%\%_oApp%" /f
)
if %winbuild% GEQ 9600 (
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
)
reg delete "HKLM\%OPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
reg delete "HKLM\%OPPk%" /f /v DisableDnsPublishing
reg delete "HKLM\%OPPk%" /f /v DisableKeyManagementServiceHostCaching
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
goto :eof

:cCache
echo.
echo Clearing KMS Cache...
call :rREG %_Nul3%
set "_C16R="
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath /reg:32" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" set "_C16R=1"
if %winbuild% GEQ 9200 if defined _C16R (
echo.
echo ## Notice ##
echo.
echo To make sure Office programs do not show a non-genuine banner
echo please apply manual or auto-renewal activation, and don't uninstall afterward.
)
if %Unattend% NEQ 0 goto :TheEnd
echo.&echo %line3%&echo.
echo Press any key to continue . . .
pause >nul
goto :MainMenu

:CreateTask
schtasks /query /tn "%_TaskEx%" %_Nul3% || (
  schtasks /query /tn "%_TaskOs%" %_Nul3% && (
    schtasks /query /tn "%_TaskOs%" /xml >"!_temp!\SvcTrigger.xml"
    schtasks /create /tn "%_TaskEx%" /xml "!_temp!\SvcTrigger.xml" /f %_Nul3%
    schtasks /change /tn "%_TaskEx%" /enable %_Nul3%
    del /f /q "!_temp!\SvcTrigger.xml" %_Nul3%
  )
)
schtasks /query /tn "%_TaskEx%" %_Nul3% || (
pushd %_temp%
%_Nul3% %_psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':spptask\:.*'; [IO.File]::WriteAllText('SvcTrigger.xml',$f[1].Trim(),[System.Text.Encoding]::Unicode)"
popd
if %Unattend% EQU 0 title %_title%
if exist "!_temp!\SvcTrigger.xml" (
  schtasks /create /tn "%_TaskEx%" /xml "!_temp!\SvcTrigger.xml" /f %_Nul3%
  del /f /q "!_temp!\SvcTrigger.xml" %_Nul3%
  )
)
schtasks /query /tn "%_TaskEx%" %_Nul3% && if %_verb% EQU 1 (
echo.
echo Adding Scheduled Task...
echo %_TaskEx%
)
goto :eof

:CreateReadMe
if not exist "%PUBLIC%\ReadMeAIO.html" (
pushd %PUBLIC%
%_Nul3% %_psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':readme\:.*'; [IO.File]::WriteAllText('ReadMeAIO.html',$f[1].Trim(),[System.Text.Encoding]::UTF8)"
popd
if %Unattend% EQU 0 title %_title%
)
if exist "%PUBLIC%\ReadMeAIO.html" start "" "%PUBLIC%\ReadMeAIO.html"
timeout /t 2 %_Nul3%
goto :eof

:CreateOEM
cls
if exist "!_oem!\$OEM$\" (
echo.&echo %line3%&echo.
echo $OEM$ Folder already exist...
echo "!_oem!\$OEM$"
echo.
echo Manually remove it if you wish to create a fresh copy.
echo.&echo %line3%&echo.
echo Press any key to continue . . .
pause >nul
goto :eof
)
if not exist "!_oem!\$OEM$\$$\Setup\Scripts\KMS_VL_ALL_AIO.cmd" mkdir "!_oem!\$OEM$\$$\Setup\Scripts"
copy /y "!_batf!" "!_oem!\$OEM$\$$\Setup\Scripts\KMS_VL_ALL_AIO.cmd" %_Nul3%
(
echo @echo off
echo call %%~dp0KMS_VL_ALL_AIO.cmd /s /a
echo cd \
echo ^(goto^) 2^>nul^&rd /s /q "%%~dp0"
)>"!_oem!\$OEM$\$$\Setup\Scripts\setupcomplete.cmd"
echo.&echo %line3%&echo.
echo $OEM$ Folder Created...
echo.
echo "!_oem!\$OEM$"
echo.&echo %line3%&echo.
echo.
echo Press any key to continue . . .
pause >nul
goto :eof

:CreateBIN
cls
if exist "!_oem!\KMS_VL_ALL_AIO-bin\*.dll" if exist "!_oem!\KMS_VL_ALL_AIO-bin\*.cab" (
echo.&echo %line3%&echo.
echo Binaries Folder already exist...
echo "!_oem!\KMS_VL_ALL_AIO-bin"
echo.
echo Manually remove it if you wish to create a fresh copy.
echo.&echo %line3%&echo.
echo Press any key to continue . . .
pause >nul
goto :eof
)
if not exist "!_oem!\KMS_VL_ALL_AIO-bin\*.dll" mkdir "!_oem!\KMS_VL_ALL_AIO-bin"
pushd "!_oem!\KMS_VL_ALL_AIO-bin"
%_Nul3% rmdir /s /q .
setlocal
set "TMP=%SystemRoot%\Temp"
set "TEMP=%SystemRoot%\Temp"
%_Nul3% %_psc% "cd -Lit ($env:__CD__); $f=[IO.File]::ReadAllText('!_batp!') -split ':embdbin\:.*';iex ($f[1]); 2..4|%%{[BAT85]::Decode($_, $f[$_])}; [IO.File]::WriteAllText('CleanOffice.ps1',$f[5].Trim(),[System.Text.Encoding]::ASCII)"
endlocal
if %Unattend% EQU 0 title %_title%
%_Nul3% ren 2 SppExtComObjHook-x86.dll
%_Nul3% ren 3 SppExtComObjHook-x64.dll
%_Nul3% ren 4 SppExtComObjHook-arm64.dll
popd
echo.&echo %line3%&echo.
echo Binaries Folder Created...
echo.
echo "!_oem!\KMS_VL_ALL_AIO-bin"
echo.&echo %line3%&echo.
echo.
echo Press any key to continue . . .
pause >nul
goto :eof

:C2RR2V
set RanR2V=1
set "_SLMGR=%SysPath%\slmgr.vbs"
if %_Debug% EQU 0 (
set "_cscript=cscript.exe //NoLogo //B"
) else (
set "_cscript=cscript.exe //NoLogo"
)
set _LTS19=0
set _LTS21=0
set _LTS24=0
set "_tag="&set "_ons= 2016"
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% EQU 1060 if %error2% EQU 1060 (
echo Error: Office C2R service is not detected
goto :%_fC2R%
)
set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\ProPlus*.xrm-ms" (
  set _Office16=1
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\ProPlus*.xrm-ms" (
  set _Office15=1
)
if %_Office16% EQU 0 if %_Office15% EQU 0 (
echo Error: Office C2R InstallPath is not detected
goto :%_fC2R%
)

:Reg16istry
if %_Office16% EQU 0 goto :Reg15istry
set "_InstallRoot="
set "_ProductIds="
set "_GUID="
set "_Config="
set "_PRIDs="
set "_LicensesPath="
set "_Integrator="
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=%_onat%\ClickToRun\Configuration"
  set "_PRIDs=%_onat%\ClickToRun\ProductReleaseIDs"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do (set "_InstallRoot=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v InstallPath" %_Nul6%') do (set "_OSPPVBS=%%b\Office16\OSPP.VBS")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun /v PackageGUID" %_Nul6%') do (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_ProductIds=%%b")
  set "_Config=%_owow%\ClickToRun\Configuration"
  set "_PRIDs=%_owow%\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
if "%_ProductIds%"=="" (
if %_Office15% EQU 0 (echo Error: Office C2R ProductIDs are not detected&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_LicensesPath%\ProPlus*.xrm-ms" (
if %_Office15% EQU 0 (echo Error: Office C2R Licenses files are not detected&goto :%_fC2R%) else (goto :Reg15istry)
)
if not exist "%_Integrator%" (
if %_Office15% EQU 0 (echo Error: Office C2R Licenses Integrator is not detected&goto :%_fC2R%) else (goto :Reg15istry)
)
if exist "%_LicensesPath%\Word2019VL_KMS_Client_AE*.xrm-ms" (set _LTS19=1&set "_tag=2019"&set "_ons= 2019")
if exist "%_LicensesPath%\Word2021VL_KMS_Client_AE*.xrm-ms" (set _LTS21=1)
if exist "%_LicensesPath%\Word2024VL_KMS_Client_AE*.xrm-ms" (set _LTS24=1)
if %winbuild% LSS 10240 if !_LTS21! EQU 1 (set "_tag=2021"&set "_ons= 2021")
if %_Office15% EQU 0 goto :CheckC2R

:Reg15istry
set "_Install15Root="
set "_Product15Ids="
set "_Con15fig="
set "_PR15IDs="
set "_OSPP15Ready="
set "_Licenses15Path="
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_onat%\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=%_onat%\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=%_onat%\15.0\ClickToRun\Configuration"
) else (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun /v InstallPath" %_Nul6%') do (set "_Install15Root=%%b\root")
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_owow%\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=%_owow%\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=%_owow%\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
reg query %_onat%\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_onat%\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=%_onat%\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
reg query %_owow%\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && (
  for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do (set "_Product15Ids=%%b")
  set "_Con15fig=%_owow%\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=%_owow%\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
  )
)
set "_Licenses15Path=%_Install15Root%\Licenses"
set _OSPP15VBS=
for %%G in (
"%ProgramFiles%"
"%ProgramW6432%"
"%ProgramFiles(x86)%"
) do if exist "%%~G\Microsoft Office\Office15\OSPP.VBS" (
if not defined _OSPP15VBS set "_OSPP15VBS=%%~G\Microsoft Office\Office15\OSPP.VBS"
)
if "%_Product15Ids%"=="" (
if %_Office16% EQU 0 (echo Error: Office 2013 C2R ProductIDs are not detected&goto :%_fC2R%) else (goto :CheckC2R)
)
if not exist "%_Licenses15Path%\ProPlus*.xrm-ms" (
if %_Office16% EQU 0 (echo Error: Office 2013 C2R Licenses files are not detected&goto :%_fC2R%) else (goto :CheckC2R)
)
if %winbuild% LSS 9200 if "%_OSPP15VBS%"=="" (
if %_Office16% EQU 0 (echo Error: Office 2013 C2R Licensing tool OSPP.vbs is not detected&goto :%_fC2R%) else (goto :CheckC2R)
)

:CheckC2R
set _OMSI=0
if %_Office16% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\16.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
)
if %_Office15% EQU 0 (
for /f "skip=2 tokens=2*" %%a in ('"reg query %_onat%\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
for /f "skip=2 tokens=2*" %%a in ('"reg query %_owow%\15.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\EntityPicker.dll" set _OMSI=1
)
if %winbuild% GEQ 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
set "_vbsf=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
set _vbsf="!_OSPPVBS!" /inslic:
)
set "_wmi="
call :qrSingle %_sps% Version
for /f "tokens=2 delims==" %%# in ('%_qr%') do set _wmi=%%#
if "%_wmi%"=="" (
echo Error: %_sps% WMI version is not detected
call :CheckWS
goto :%_fC2R%
)
set _Retail=0
set "_ocq=ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey is not NULL"
call :qrQuery %_spp% "%_ocq%" Description fix
%_qr% %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set rancopp=0
if %_Retail% EQU 0 if %_OMSI% EQU 0 (
set rancopp=1
%_Nul3% %_psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':embdbin\:.*';iex ($f[5])"
if %Unattend% EQU 0 title %_title%
)

:R16V
set _SubID=O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _O16O365=0
set _C16Msg=0
set _C15Msg=0
call :qrQuery %_spp% "%_ocq%" LicenseFamily fix
if %_Retail% EQU 1 %_qr% %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
call :qrQuery %_spp% "ApplicationID='%_oApp%'" LicenseFamily fix
%_qr% %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% EQU 0 goto :R15V

set _S24ID=ProPlus2024,Standard2024
set _S21ID=ProPlus2021,Standard2021
set _S19ID=ProPlus2019,Standard2019
set _S16ID=Mondo,Standard
set _P24ID=ProjectPro2024,ProjectStd2024
set _P21ID=ProjectPro2021,ProjectStd2021
set _P19ID=ProjectPro2019,ProjectStd2019
set _P16ID=ProjectPro,ProjectStd
set _I24ID=VisioPro2024,VisioStd2024
set _I21ID=VisioPro2021,VisioStd2021
set _I19ID=VisioPro2019,VisioStd2019
set _I16ID=VisioPro,VisioStd
set _A24ID=Excel2024,Outlook2024,PowerPoint2024,Word2024
set _A21ID=Excel2021,Outlook2021,PowerPoint2021,Publisher2021,Word2021
set _A19ID=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16ID=Excel,Outlook,PowerPoint,Publisher,Word
set _E24ID=Access2024,SkypeforBusiness2024
set _E21ID=Access2021,SkypeforBusiness2021
set _E19ID=Access2019,SkypeforBusiness2019
set _E16ID=Access,SkypeforBusiness
set _R24ID=Professional2024,HomeBusiness2024,HomeStudent2024,Home2024
set _R21ID=Professional2021,HomeBusiness2021,HomeStudent2021
set _R19ID=Professional2019,HomeBusiness2019,HomeStudent2019
set _R16ID=Professional,HomeBusiness,HomeStudent,%_SubID%
set _V24ID=%_S24ID%,%_A24ID%,%_E24ID%,%_P24ID%,%_I24ID%
set _V21ID=%_S21ID%,%_A21ID%,%_E21ID%,%_P21ID%,%_I21ID%
set _V19ID=%_S19ID%,%_A19ID%,%_E19ID%,%_P19ID%,%_I19ID%
set _V16ID=%_S16ID%,%_A16ID%,%_E16ID%,%_P16ID%,%_I16ID%
set _RetID=%_R24ID%,%_V24ID%,%_R21ID%,%_V21ID%,%_R19ID%,%_V19ID%,%_R16ID%,%_V16ID%
set _Suites=ProPlus,%_S16ID%,%_R16ID%,%_S19ID%,%_R19ID%,%_S21ID%,%_R21ID%,%_S24ID%,%_R24ID%
set _PrjSKU=%_P16ID%,%_P19ID%,%_P21ID%,%_P24ID%
set _VisSKU=%_I16ID%,%_I19ID%,%_I21ID%,%_I24ID%

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetID%,ProPlus,OneNote,Publisher2024,Home,Home2019,Home2021) do (
set _%%a=0
)
for %%a in (%_RetID%,OneNote) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
if !_LTS24! EQU 0 for %%a in (%_V24ID%) do (
set _%%a=0
)
if !_LTS24! EQU 1 for %%a in (%_V24ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office24%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
if !_LTS21! EQU 0 for %%a in (%_V21ID%) do (
set _%%a=0
)
if !_LTS21! EQU 1 for %%a in (%_V21ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office21%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
if !_LTS19! EQU 0 for %%a in (%_V19ID%) do (
set _%%a=0
)
if !_LTS19! EQU 1 for %%a in (%_V19ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V16ID%,OneNote) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PRIDs%\ProPlusVolume.16 %_Nul3% && (
  find /i "Office16ProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_RetID%,OneNote) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R16=1)
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R19=1)
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R19=1)
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R19=1)
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R19=1)
  find /i "Office21%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R21=1)
  find /i "Office21%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R21=1)
  find /i "Office21%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R21=1)
  find /i "Office21%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R21=1)
  find /i "Office24%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R24=1)
  find /i "Office24%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R24=1)
  find /i "Office24%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R24=1)
  find /i "Office24%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R24=1)
  )
)
if %_Retail% EQU 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R16=1)
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R16=1)
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R16=1)
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R16=1)
)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'Office16O365%%%%'" LicenseFamily
find /i "Office16MondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (%_SubID%) do set _%%a=0
  )
)
if %sub_o365% EQU 1 (
for %%a in (%_Suites%) do set _%%a=0
echo.
echo Microsoft Office is activated with a vNext license.
)
if %sub_proj% EQU 1 (
for %%a in (%_PrjSKU%) do set _%%a=0
echo.
echo Microsoft Project is activated with a vNext license.
)
if %sub_vsio% EQU 1 (
for %%a in (%_VisSKU%) do set _%%a=0
echo.
echo Microsoft Visio is activated with a vNext license.
)

for %%a in (%_RetID%,ProPlus,OneNote) do if !_%%a! EQU 1 (
set _C16Msg=1
)
if %_C16Msg% EQU 1 (
echo.
echo Converting Office C2R Retail-to-Volume:
)
if %_C16Msg% EQU 0 goto :endRV16

set "_arr="
for %%# in ("!_LicensesPath!\client-issuance-*.xrm-ms") do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_LicensesPath!\%%~nx#"") else (set "_arr="!_LicensesPath!\%%~nx#"")
  ) else (
  %_cscript% %_vbsf%"!_LicensesPath!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'; InstallLicenseFile '"!_LicensesPath!\pkeyconfig-office.xrm-ms"'"
  ) else (
  %_cscript% %_vbsf%"!_LicensesPath!\pkeyconfig-office.xrm-ms"
  )

set _jump=0
set _DidO365=0
if !_Mondo! EQU 1 (
call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 (
set _DidO365=1
echo O365ProPlus 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365Business! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365Business 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365SmallBusPrem 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365HomePrem! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365HomePrem 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365EduCloud! EQU 1 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365EduCloud 2016 Suite ^<-^> Mondo 2016 Licenses
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_DidO365! EQU 1 set _jump=1&set _O16O365=1
if !_Mondo! EQU 1 if !_DidO365! EQU 0 (
echo Mondo 2016 Suite
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :endRV16
)

for %%a in (%_P16ID%,%_I16ID%) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 SKU&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 SKU&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 SKU -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 SKU -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)

if !_jump! EQU 1 goto :endRV16

for %%a in (ProPlus) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite&call :InsLic ProPlus2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite&call :InsLic ProPlus2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (Professional) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite -^> ProPlus 2024 Licenses&call :InsLic ProPlus2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite -^> ProPlus 2021 Licenses&call :InsLic ProPlus2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> ProPlus%_ons% Licenses&call :InsLic ProPlus%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (SkypeforBusiness) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 App&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 App&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)

for %%a in (Access) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 App&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 App&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)

for %%a in (Standard) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite&call :InsLic Standard2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite&call :InsLic Standard2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (HomeBusiness,HomeStudent,Home) do (
  if !_%%a2024! EQU 1 (set _jump=1&echo %%a 2024 Suite -^> Standard 2024 Licenses&call :InsLic Standard2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (set _jump=1&echo %%a 2021 Suite -^> Standard 2021 Licenses&call :InsLic Standard2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (set _jump=1&echo %%a 2019 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (set _jump=1&echo %%a 2016 Suite -^> Standard%_ons% Licenses&call :InsLic Standard%_tag%)
)
if !_jump! EQU 1 goto :endRV16

for %%a in (%_A16ID%) do (
  if !_%%a2024! EQU 1 (echo %%a 2024 App&call :InsLic %%a2024)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 1 (echo %%a 2021 App&call :InsLic %%a2021)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 1 (echo %%a 2019 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
  if !_%%a2024! EQU 0 if !_%%a2021! EQU 0 if !_%%a2019! EQU 0 if !_%%a! EQU 1 (echo %%a 2016 App -^> %%a%_ons% Licenses&call :InsLic %%a%_tag%)
)
for %%a in (OneNote) do (
  if !_%%a! EQU 1 (echo %%a 2016 App&call :InsLic %%a)
)

:endRV16
if %_Office15% EQU 0 goto :GVLKC2R

:R15V
set _S15ID=Mondo,Standard
set _P15ID=ProjectPro,ProjectStd
set _I15ID=VisioPro,VisioStd
set _A15ID=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _E15ID=Access,Lync
set _V15ID=%_S15ID%,%_A15ID%,%_E15ID%,%_P15ID%,%_I15ID%
set _R15ID=%_V15ID%,SPD,Professional,HomeBusiness,HomeStudent,%_SubID%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15ID%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15ID%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15ID%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
reg query %_PR15IDs%\Active\ProPlusVolume\x-none %_Nul3% && (
  find /i "OfficeProPlusVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _ProPlus=0) || (set _ProPlus=1)
)
if %_Retail% EQU 1 for %%a in (%_R15ID%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _%%a=0&set aC2R15=1)
  )
)
if %_Retail% EQU 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R15=1)
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R15=1)
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R15=1)
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && (set _ProPlus=0&set aC2R15=1)
)
call :qrQuery %_spp% "ApplicationID='%_oApp%' AND LicenseFamily like 'OfficeO365%%%%'" LicenseFamily
find /i "OfficeMondoVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (
%_qr% %_Nul2% | find /i "O365" %_Nul1% && (
  for %%a in (%_SubID%) do set _%%a=0
  )
)

for %%a in (%_R15ID%,ProPlus) do if !_%%a! EQU 1 (
set _C15Msg=1
)
if %_C15Msg% EQU 1 if %_C16Msg% EQU 0 (
echo.
echo Converting Office C2R Retail-to-Volume:
)
if %_C15Msg% EQU 0 goto :endRV15

set "_arr="
for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_Licenses15Path!\%%~nx#"") else (set "_arr="!_Licenses15Path!\%%~nx#"")
  ) else (
  %_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'; InstallLicenseFile '"!_Licenses15Path!\pkeyconfig-office.xrm-ms"'"
  ) else (
  %_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"
  )

set _jump=0
set _DidO365=0
if !_Mondo! EQU 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! EQU 1 if !_O16O365! EQU 0 (
set _DidO365=1
echo O365ProPlus 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365SmallBusPrem 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365HomePrem 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365Business! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
set _DidO365=1
echo O365Business 2013 Suite ^<-^> Mondo 2013 Licenses
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_DidO365! EQU 1 set _jump=1
if !_Mondo! EQU 1 if !_O16O365! EQU 0 if !_DidO365! EQU 0 (
echo Mondo 2013 Suite
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :endRV15
)

for %%a in (%_P15ID%,%_I15ID%) do (
  if !_%%a! EQU 1 (echo %%a 2013 SKU&call :Ins15Lic %%a)
)

if !_Mondo! EQU 0 if !_DidO365! EQU 0 for %%a in (SPD) do (
  if !_%%a! EQU 1 (set _jump=1&echo SharePoint Designer 2013 App -^> Mondo 2013 Licenses&call :Ins15Lic Mondo)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (ProPlus) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite&call :Ins15Lic %%a)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (Professional) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite -^> ProPlus 2013 Licenses&call :Ins15Lic ProPlus)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (Lync) do (
  if !_%%a! EQU 1 (echo SkypeforBusiness 2015 App&call :Ins15Lic %%a)
)

for %%a in (Access) do (
  if !_%%a! EQU 1 (echo %%a 2013 App&call :Ins15Lic %%a)
)

for %%a in (Standard) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite&call :Ins15Lic %%a)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (HomeBusiness,HomeStudent) do (
  if !_%%a! EQU 1 (set _jump=1&echo %%a 2013 Suite -^> Standard 2013 Licenses&call :Ins15Lic Standard)
)
if !_jump! EQU 1 goto :endRV15

for %%a in (%_A15ID%) do (
  if !_%%a! EQU 1 (echo %%a 2013 App&call :Ins15Lic %%a)
)

:endRV15
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
set "_kpey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=PidKey=%2"
set "_kpey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%

set fallback=0
call :qrQuery %_spp% "ApplicationID='%_oApp%'" LicenseFamily fix
%_qr% %_Nul2% | find /i "%_patt%" %_Nul1% || (set fallback=1)
if %fallback% equ 0 goto :IntOK

set "_lsfs="
for %%# in ("!_LicensesPath!\%_patt%*.xrm-ms") do (
set "_lsfs=!_lsfs! %%~nx#"
)
if defined _kpey (
  for %%# in ("!_LicensesPath!\%1DemoR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1E5R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1EDUR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1MSDNR*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1O365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
  for %%# in ("!_LicensesPath!\%1CO365R*.xrm-ms") do (
  set "_lsfs=!_lsfs! %%~nx#"
  )
)
set "_arr="
for %%# in (!_lsfs!) do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_LicensesPath!\%%~nx#"") else (set "_arr="!_LicensesPath!\%%~nx#"")
  ) else (
  %_cscript% %_vbsf%"!_LicensesPath!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'"
  )
call :qrPKey %_sps% %_wmi% %_kpey%
if defined _kpey %_qr% %_Nul3%

:IntOK
reg add %_Config% /f /v %_ID%.OSPPReady /t REG_SZ /d 1 %_Nul1%
reg query %_Config% /v ProductReleaseIds | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Config% /v ProductReleaseIds') do reg add %_Config% /v ProductReleaseIds /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:Ins15Lic
set "_ID=%1Volume"
set "_patt=%1VL_"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_patt=%1R_"
set "_pkey=%2"
)
reg delete %_OSPP15Ready% /f /v %_ID%.OSPPReady %_Nul3%
set "_arr="
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
if %WMI_PS% NEQ 0 (
  if defined _arr (set "_arr=!_arr!;"!_Licenses15Path!\%%~nx#"") else (set "_arr="!_Licenses15Path!\%%~nx#"")
  ) else (
  %_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
  )
)
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); InstallLicenseArr '!_arr!'"
  )
call :qrPKey %_sps% %_wmi% %_pkey%
if defined _pkey %_qr% %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% %_Nul2% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig% %_Nul6%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
set _CtRMsg=0
if %_C16Msg% EQU 1 set _CtRMsg=1
if %_C15Msg% EQU 1 set _CtRMsg=1
if %_Office16% EQU 1 (
for %%a in (%_RetID%,ProPlus,OneNote) do set "_%%a="
for %%A in (19,21,24) do call :officeLoc %%A
)
if %_Office15% EQU 1 (
for %%a in (%_R15ID%,ProPlus,O365EduCloud) do set "_%%a="
)
call :qrMethod %_sps% Version %_wmi% RefreshLicenseStatus
if %winbuild% GEQ 9200 %_qr% %_Nul3%
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" if %rancopp% EQU 1 if %_CtRMsg% EQU 1 (
if %WMI_PS% NEQ 0 (
  %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); ReinstallLicenses"
  if !ERRORLEVEL! NEQ 0 %_Nul3% %_psc% "$sls='%_sps%'; $f=[IO.File]::ReadAllText('!_batp!') -split ':embdxrm\:.*'; iex ($f[1]); ReinstallLicenses"
  ) else (
  %_cscript% %_SLMGR% /rilc
  if !ERRORLEVEL! NEQ 0 %_cscript% %_SLMGR% /rilc
  )
)
goto :%_sC2R%

:casWm
cls
mode con cols=100 lines=34
%_Nul3% %_psc% "&%_buf%"
%_psc% "$f=[IO.File]::ReadAllText('!_batp!') -split ':sppmgr\:.*';iex ($f[1])"
echo.
echo Press any key to continue . . .
pause >nul
goto :eof

:keys
if "%~1"=="" exit /b
goto :%1 %_Nul2%

:: Windows 11 [Ni]
:59eb965c-9150-42b7-a0ec-22151b9897c5
set "_key=KBN8V-HFGQ4-MGXVD-347P6-PDQGT" &:: IoT Enterprise LTSC
exit /b

:: Windows 11 [Co]
:ca7df2e3-5ea0-47b8-9ac1-b1be4d8edd69
set "_key=37D7F-N49CB-WQR8W-TBJ73-FM8RX" &:: SE {Cloud}
exit /b

:d30136fc-cb4b-416e-a23d-87207abc44a9
set "_key=6XN7V-PCBDC-BDBRH-8DQY7-G6R44" &:: SE N {Cloud N}
exit /b

:: Windows 10 [RS5]
:32d2fab3-e4a8-42c2-923b-4bf4fd13e6ee
set "_key=M7XTQ-FN8P6-TTKYV-9D4CC-J462D" &:: Enterprise LTSC 2019
exit /b

:7103a333-b8c8-49cc-93ce-d37c09687f92
set "_key=92NFX-8DJQP-P6BBQ-THF9C-7CG2H" &:: Enterprise LTSC 2019 N
exit /b

:ec868e65-fadf-4759-b23e-93fe37f2cc29
set "_key=CPWHC-NT2C7-VYW78-DHDB2-PG3GK" &:: Enterprise for Virtual Desktops
exit /b

:0df4f814-3f57-4b8b-9a9d-fddadcd69fac
set "_key=NBTWJ-3DR69-3C4V8-C26MC-GQ9M6" &:: Lean
exit /b

:: Windows 10 [RS3]
:82bbc092-bc50-4e16-8e18-b74fc486aec3
set "_key=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J" &:: Pro Workstation
exit /b

:4b1571d3-bafb-4b40-8087-a961be2caf65
set "_key=9FNHH-K3HBT-3W4TD-6383H-6XYWF" &:: Pro Workstation N
exit /b

:e4db50ea-bda1-4566-b047-0ca50abc6f07
set "_key=7NBT4-WGBQX-MP4H7-QXFF8-YP3KX" &:: Enterprise Remote Server
exit /b

:: Windows 10 [RS2]
:e0b2d383-d112-413f-8a80-97f373a5820c
set "_key=YYVX9-NTFWV-6MDM3-9PT4T-4M68B" &:: Enterprise G
exit /b

:e38454fb-41a4-4f59-a5dc-25080e354730
set "_key=44RPN-FTY23-9VTTB-MP9BX-T84FV" &:: Enterprise G N
exit /b

:: Windows 10 [RS1]
:2d5a5a60-3040-48bf-beb0-fcd770c20ce0
set "_key=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ" &:: Enterprise 2016 LTSB
exit /b

:9f776d83-7156-45b2-8a5c-359b9c9f22a3
set "_key=QFFDN-GRT3P-VKWWX-X7T3R-8B639" &:: Enterprise 2016 LTSB N
exit /b

:3f1afc82-f8ac-4f6c-8005-1d233e606eee
set "_key=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y" &:: Pro Education
exit /b

:5300b18c-2e33-4dc2-8291-47ffcec746dd
set "_key=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC" &:: Pro Education N
exit /b

:: Windows 10 [TH]
:58e97c99-f377-4ef1-81d5-4ad5522b5fd8
set "_key=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99" &:: Home
exit /b

:7b9e1751-a8da-4f75-9560-5fadfe3d8e38
set "_key=3KHY7-WNT83-DGQKR-F7HPR-844BM" &:: Home N
exit /b

:cd918a57-a41b-4c82-8dce-1a538e221a83
set "_key=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH" &:: Home Single Language
exit /b

:a9107544-f4a0-4053-a96a-1479abdef912
set "_key=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR" &:: Home China
exit /b

:2de67392-b7a7-462a-b1ca-108dd189f588
set "_key=W269N-WFGWX-YVC9B-4J6C9-T83GX" &:: Pro
exit /b

:a80b5abf-76ad-428b-b05d-a47d2dffeebf
set "_key=MH37W-N47XK-V7XM9-C7227-GCQG9" &:: Pro N
exit /b

:e0c42288-980c-4788-a014-c080d2e1926e
set "_key=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2" &:: Education
exit /b

:3c102355-d027-42c6-ad23-2e7ef8a02585
set "_key=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ" &:: Education N
exit /b

:73111121-5638-40f6-bc11-f1d7b0d64300
set "_key=NPPR9-FWDCX-D2C8J-H872K-2YT43" &:: Enterprise
exit /b

:e272e3e2-732f-4c65-a8f0-484747d0d947
set "_key=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4" &:: Enterprise N
exit /b

:7b51a46c-0c04-4e8f-9af4-8496cca90d5e
set "_key=WNMTR-4C88C-JK8YV-HQ7T2-76DF9" &:: Enterprise 2015 LTSB
exit /b

:87b838b7-41b6-4590-8318-5797951d8529
set "_key=2F77B-TNFGY-69QQF-B8YKP-D69TJ" &:: Enterprise 2015 LTSB N
exit /b

:: Windows Server 2025 [Ge]
:7dc26449-db21-4e09-ba37-28f2958506a6
set "_key=TVRH6-WHNXV-R9WG3-9XRFY-MY832" &:: Standard
exit /b

:c052f164-cdf6-409a-a0cb-853ba0f0f55a
set "_key=D764K-2NDRG-47T6Q-P8T8W-YP6DF" &:: Datacenter
exit /b

:45b5aff2-60a0-42f2-bc4b-ec6e5f7b527e
set "_key=FCNV3-279Q9-BQB46-FTKXX-9HPRH" &:: Azure Core
exit /b

:c2e946d1-cfa2-4523-8c87-30bc696ee584
set "_key=XGN3F-F394H-FD2MY-PP6FD-8MCRC" &:: Turbine
exit /b

:: Windows Server 2022 [Fe]
:9bd77860-9b31-4b7b-96ad-2564017315bf
set "_key=VDYBN-27WPP-V4HQT-9VMD4-VMK7H" &:: Standard
exit /b

:ef6cfc9f-8c5d-44ac-9aad-de6a2ea0ae03
set "_key=WX4NM-KYWYW-QJJR4-XV3QB-6VM33" &:: Datacenter
exit /b

:8c8f0ad3-9a43-4e05-b840-93b8d1475cbc
set "_key=6N379-GGTMK-23C6M-XVVTC-CKFRQ" &:: Azure Core
exit /b

:f5e9429c-f50b-4b98-b15c-ef92eb5cff39
set "_key=67KN8-4FYJW-2487Q-MQ2J7-4C4RG" &:: Standard ACor
exit /b

:39e69c41-42b4-4a0a-abad-8e3c10a797cc
set "_key=QFND9-D3Y9C-J3KKY-6RPVP-2DPYV" &:: Datacenter ACor
exit /b

:: Windows Server 2019 [RS5]
:de32eafd-aaee-4662-9444-c1befb41bde2
set "_key=N69G4-B89J2-4G8F4-WWYCC-J464C" &:: Standard
exit /b

:34e1ae55-27f8-4950-8877-7a03be5fb181
set "_key=WMDGN-G9PQG-XVVXX-R3X43-63DFG" &:: Datacenter
exit /b

:a99cc1f0-7719-4306-9645-294102fbff95
set "_key=FDNH6-VW9RW-BXPJ7-4XTYG-239TB" &:: Azure Core
exit /b

:73e3957c-fc0c-400d-9184-5f7b6f2eb409
set "_key=N2KJX-J94YW-TQVFB-DG9YT-724CC" &:: Standard ACor
exit /b

:90c362e5-0da1-4bfd-b53b-b87d309ade43
set "_key=6NMRW-2C8FM-D24W7-TQWMY-CWH2D" &:: Datacenter ACor
exit /b

:034d3cbb-5d4b-4245-b3f8-f84571314078
set "_key=WVDHN-86M7X-466P6-VHXV7-YY726" &:: Essentials
exit /b

:8de8eb62-bbe0-40ac-ac17-f75595071ea3
set "_key=GRFBW-QNDC4-6QBHG-CCK3B-2PR88" &:: ServerARM64
exit /b

:19b5e0fb-4431-46bc-bac1-2f1873e4ae73
set "_key=NTBV8-9K7Q8-V27C6-M2BTV-KHMXV" &:: Datacenter Azure - Turbine
exit /b

:: Windows Server 2016 [RS4]
:43d9af6e-5e86-4be8-a797-d072a046896c
set "_key=K9FYF-G6NCK-73M32-XMVPY-F9DRR" &:: ServerARM64
exit /b

:: Windows Server 2016 [RS3]
:61c5ef22-f14f-4553-a824-c4b31e84b100
set "_key=PTXN8-JFHJM-4WC78-MPCBR-9W4KR" &:: Standard ACor
exit /b

:e49c08e7-da82-42f8-bde2-b570fbcae76c
set "_key=2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG" &:: Datacenter ACor
exit /b

:: Windows Server 2016 [RS1]
:8c1c5410-9f39-4805-8c9d-63a07706358f
set "_key=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY" &:: Standard
exit /b

:21c56779-b449-4d20-adfc-eece0e1ad74b
set "_key=CB7KF-BWN84-R7R2Y-793K2-8XDDG" &:: Datacenter
exit /b

:3dbf341b-5f6c-4fa7-b936-699dce9e263f
set "_key=VP34G-4NPPG-79JTQ-864T4-R3MQX" &:: Azure Core
exit /b

:2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283
set "_key=JCKRF-N37P4-C2D82-9YXRT-4M63B" &:: Essentials
exit /b

:7b4433f4-b1e7-4788-895a-c45378d38253
set "_key=QN4C6-GBJD2-FB422-GHWJK-GJG2R" &:: Cloud Storage
exit /b

:: Windows 8.1
:fe1c3238-432a-43a1-8e25-97e7d1ef10f3
set "_key=M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK" &:: Core
exit /b

:78558a64-dc19-43fe-a0d0-8075b2a370a3
set "_key=7B9N3-D94CG-YTVHR-QBPX3-RJP64" &:: Core N
exit /b

:c72c6a1d-f252-4e7e-bdd1-3fca342acb35
set "_key=BB6NG-PQ82V-VRDPW-8XVD2-V8P66" &:: Core Single Language
exit /b

:db78b74f-ef1c-4892-abfe-1e66b8231df6
set "_key=NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3" &:: Core China
exit /b

:ffee456a-cd87-4390-8e07-16146c672fd0
set "_key=XYTND-K6QKT-K2MRH-66RTM-43JKP" &:: Core ARM
exit /b

:c06b6981-d7fd-4a35-b7b4-054742b7af67
set "_key=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9" &:: Pro
exit /b

:7476d79f-8e48-49b4-ab63-4d0b813a16e4
set "_key=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY" &:: Pro N
exit /b

:096ce63d-4fac-48a9-82a9-61ae9e800e5f
set "_key=789NJ-TQK6T-6XTH8-J39CJ-J8D3P" &:: Pro with Media Center
exit /b

:81671aaf-79d1-4eb1-b004-8cbbe173afea
set "_key=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7" &:: Enterprise
exit /b

:113e705c-fa49-48a4-beea-7dd879b46b14
set "_key=TT4HM-HN7YT-62K67-RGRQJ-JFFXW" &:: Enterprise N
exit /b

:0ab82d54-47f4-4acb-818c-cc5bf0ecb649
set "_key=NMMPB-38DD4-R2823-62W8D-VXKJB" &:: Embedded Industry Pro
exit /b

:cd4e2d9f-5059-4a50-a92d-05d5bb1267c7
set "_key=FNFKF-PWTVT-9RC8H-32HB2-JB34X" &:: Embedded Industry Enterprise
exit /b

:f7e88590-dfc7-4c78-bccb-6f3865b99d1a
set "_key=VHXM3-NR6FT-RY6RT-CK882-KW2CJ" &:: Embedded Industry Automotive
exit /b

:e9942b32-2e55-4197-b0bd-5ff58cba8860
set "_key=3PY8R-QHNP9-W7XQD-G6DPH-3J2C9" &:: with Bing
exit /b

:c6ddecd6-2354-4c19-909b-306a3058484e
set "_key=Q6HTR-N24GM-PMJFP-69CD8-2GXKR" &:: with Bing N
exit /b

:b8f5e3a3-ed33-4608-81e1-37d6c9dcfd9c
set "_key=KF37N-VDV38-GRRTV-XH8X6-6F3BB" &:: with Bing Single Language
exit /b

:ba998212-460a-44db-bfb5-71bf09d1c68b
set "_key=R962J-37N87-9VVK2-WJ74P-XTMHR" &:: with Bing China
exit /b

:e58d87b5-8126-4580-80fb-861b22f79296
set "_key=MX3RK-9HNGX-K3QKC-6PJ3F-W8D7B" &:: Pro for Students
exit /b

:cab491c7-a918-4f60-b502-dab75e334f40
set "_key=TNFGH-2R6PB-8XM3K-QYHX2-J4296" &:: Pro for Students N
exit /b

:: Windows Server 2012 - 2012 R2 ESU
:f57b5b6b-80c2-46e4-ae9d-9fe98e032cb7
set "_key=GFMWN-WDHVB-4Y4XP-42WKM-RC6CQ" &:: Year1
exit /b

:b1b1ef19-a088-4962-aedb-2a647a891104
set "_key=XN3XP-QGKM4-KT7HM-6HC6T-H8V6F" &:: Year2
exit /b

:1a716f14-0607-425f-a097-5f2f1f091315
set "_key=QCQ4R-N2J93-PWMTK-G2BGF-BY82T" &:: Year3
exit /b

:: Windows Server 2012 R2
:b3ca044e-a358-4d68-9883-aaa2941aca99
set "_key=D2N9P-3P6X9-2R39C-7RTCD-MDVJX" &:: Standard
exit /b

:00091344-1ea4-4f37-b789-01750ba6988c
set "_key=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9" &:: Datacenter
exit /b

:21db6ba4-9a7b-4a14-9e29-64a60c59301d
set "_key=KNC87-3J2TX-XB4WP-VCPJV-M4FWM" &:: Essentials
exit /b

:b743a2be-68d4-4dd3-af32-92425b7bb623
set "_key=3NPTF-33KPT-GGBPR-YX76B-39KDD" &:: Cloud Storage
exit /b

:: Windows 8
:c04ed6bf-55c8-4b47-9f8e-5a1f31ceee60
set "_key=BN3D2-R7TKB-3YPBD-8DRP2-27GG4" &:: Core
exit /b

:197390a0-65f6-4a95-bdc4-55d58a3b0253
set "_key=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY" &:: Core N
exit /b

:8860fcd4-a77b-4a20-9045-a150ff11d609
set "_key=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ" &:: Core Single Language
exit /b

:9d5584a2-2d85-419a-982c-a00888bb9ddf
set "_key=4K36P-JN4VD-GDC6V-KDT89-DYFKP" &:: Core China
exit /b

:af35d7b7-5035-4b63-8972-f0b747b9f4dc
set "_key=DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV" &:: Core ARM
exit /b

:a98bcd6d-5343-4603-8afe-5908e4611112
set "_key=NG4HW-VH26C-733KW-K6F98-J8CK4" &:: Pro
exit /b

:ebf245c1-29a8-4daf-9cb1-38dfc608a8c8
set "_key=XCVCF-2NXM9-723PB-MHCB7-2RYQQ" &:: Pro N
exit /b

:a00018a3-f20f-4632-bf7c-8daa5351c914
set "_key=GNBB8-YVD74-QJHX6-27H4K-8QHDG" &:: Pro with Media Center
exit /b

:458e1bec-837a-45f6-b9d5-925ed5d299de
set "_key=32JNW-9KQ84-P47T8-D8GGY-CWCK7" &:: Enterprise
exit /b

:e14997e7-800a-4cf7-ad10-de4b45b578db
set "_key=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT" &:: Enterprise N
exit /b

:10018baf-ce21-4060-80bd-47fe74ed4dab
set "_key=RYXVT-BNQG7-VD29F-DBMRY-HT73M" &:: Embedded Industry Pro
exit /b

:18db1848-12e0-4167-b9d7-da7fcda507db
set "_key=NKB3R-R2F8T-3XCDP-7Q2KW-XWYQ2" &:: Embedded Industry Enterprise
exit /b

:: Windows Server 2012
:f0f5ec41-0d55-4732-af02-440a44a3cf0f
set "_key=XC9B7-NBPP2-83J2H-RHMBY-92BT4" &:: Standard
exit /b

:d3643d60-0c42-412d-a7d6-52e6635327f6
set "_key=48HP8-DN98B-MYWDG-T2DCC-8W83P" &:: Datacenter
exit /b

:8f365ba6-c1b9-4223-98fc-282a0756a3ed
set "_key=HTDQM-NBMMG-KGYDT-2DTKT-J2MPV" &:: Essentials
exit /b

:7d5486c7-e120-4771-b7f1-7b56c6d3170c
set "_key=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ" &:: MultiPoint Standard
exit /b

:95fd1c83-7df5-494a-be8b-1300e1c9d1cd
set "_key=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G" &:: MultiPoint Premium
exit /b

:: Windows 7
:b92e9980-b9d5-4821-9c94-140f632f6312
set "_key=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4" &:: Professional
exit /b

:54a09a0d-d57b-4c10-8b69-a842d6590ad5
set "_key=MRPKT-YTG23-K7D7T-X2JMM-QY7MG" &:: Professional N
exit /b

:5a041529-fef8-4d07-b06f-b59b573b32d2
set "_key=W82YF-2Q76Y-63HXB-FGJG9-GF7QX" &:: Professional E
exit /b

:ae2ee509-1b34-41c0-acb7-6d4650168915
set "_key=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH" &:: Enterprise
exit /b

:1cb6d605-11b3-4e14-bb30-da91c8e3983a
set "_key=YDRBP-3D83W-TY26F-D46B2-XCKRJ" &:: Enterprise N
exit /b

:46bbed08-9c7b-48fc-a614-95250573f4ea
set "_key=C29WB-22CC8-VJ326-GHFJW-H9DH4" &:: Enterprise E
exit /b

:db537896-376f-48ae-a492-53d0547773d0
set "_key=YBYF6-BHCR3-JPKRB-CDW7B-F9BK4" &:: Embedded POSReady 7
exit /b

:e1a8296a-db37-44d1-8cce-7bc961d59c54
set "_key=XGY72-BRBBT-FF8MH-2GG8H-W7KCW" &:: Embedded Standard
exit /b

:aa6dd3aa-c2b4-40e2-a544-a6bbb3f5c395
set "_key=73KQT-CD9G6-K7TQG-66MRP-CQ22C" &:: Embedded ThinPC
exit /b

:: Windows Server 2008 R2
:a78b8bd9-8017-4df5-b86a-09f756affa7c
set "_key=6TPJF-RBVHG-WBW2R-86QPH-6RTM4" &:: Web
exit /b

:cda18cf3-c196-46ad-b289-60c072869994
set "_key=TT8MH-CG224-D3D7Q-498W2-9QCTX" &:: HPC
exit /b

:68531fb9-5511-4989-97be-d11a0f55633f
set "_key=YC6KT-GKW9T-YTKYR-T4X34-R7VHC" &:: Standard
exit /b

:620e2b3d-09e7-42fd-802a-17a13652fe7a
set "_key=489J6-VHDMP-X63PK-3K798-CPX3Y" &:: Enterprise
exit /b

:7482e61b-c589-4b7f-8ecc-46d455ac3b87
set "_key=74YFP-3QFB3-KQT8W-PMXWJ-7M648" &:: Datacenter
exit /b

:8a26851c-1c7e-48d3-a687-fbca9b9ac16b
set "_key=GT63C-RJFQ3-4GMB6-BRFB9-CB83V" &:: Itanium
exit /b

:f772515c-0e87-48d5-a676-e6962c3e1195
set "_key=736RG-XDKJK-V34PF-BHK87-J6X3K" &:: MultiPoint Server - ServerEmbeddedSolution
exit /b

:: Office 2024
:8d368fc1-9470-4be2-8d66-90e836cbb051
set "_key=NBBBB-BBBBB-BBBBB-BBBJD-VXRPM" &:: Professional Plus
exit /b

:bbac904f-6a7e-418a-bb4b-24c85da06187
set "_key=V28N4-JG22K-W66P8-VTMGK-H6HGR" &:: Standard
exit /b

:f510af75-8ab7-4426-a236-1bfb95c34ff8
set "_key=NBBBB-BBBBB-BBBBB-BBBH4-GX3R4" &:: Project Professional
exit /b

:9f144f27-2ac5-40b9-899d-898c2b8b4f81
set "_key=PD3TT-NTHQQ-VC7CY-MFXK3-G87F8" &:: Project Standard
exit /b

:fa187091-8246-47b1-964f-80a0b1e5d69a
set "_key=NBBBB-BBBBB-BBBBB-BBBCW-6MX6T" &:: Visio Professional
exit /b

:923fa470-aa71-4b8b-b35c-36b79bf9f44b
set "_key=JMMVY-XFNQC-KK4HK-9H7R3-WQQTV" &:: Visio Standard
exit /b

:72e9faa7-ead1-4f3d-9f6e-3abc090a81d7
set "_key=82FTR-NCHR7-W3944-MGRHM-JMCWD" &:: Access
exit /b

:cbbba2c3-0ff5-4558-846a-043ef9d78559
set "_key=F4DYN-89BP2-WQTWJ-GR8YC-CKGJG" &:: Excel
exit /b

:bef3152a-8a04-40f2-a065-340c3f23516d
set "_key=D2F8D-N3Q3B-J28PV-X27HD-RJWB9" &:: Outlook
exit /b

:b63626a4-5f05-4ced-9639-31ba730a127e
set "_key=CW94N-K6GJH-9CTXY-MG2VC-FYCWP" &:: PowerPoint
exit /b

:0002290a-2091-4324-9e53-3cfe28884cde
set "_key=4NKHF-9HBQF-Q3B6C-7YV34-F64P3" &:: Skype for Business
exit /b

:d0eded01-0881-4b37-9738-190400095098
set "_key=MQ84N-7VYDM-FXV7C-6K7CC-VFW9J" &:: Word
exit /b

:fceda083-1203-402a-8ec4-3d7ed9f3648c
set "_key=2TDPW-NDQ7G-FMG99-DXQ7M-TX3T2" &:: Pro Plus Preview
exit /b

:aaea0dc8-78e1-4343-9f25-b69b83dd1bce
set "_key=D9GTG-NP7DV-T6JP3-B6B62-JB89R" &:: Project Pro Preview
exit /b

:4ab4d849-aabc-43fb-87ee-3aed02518891
set "_key=YW66X-NH62M-G6YFP-B7KCT-WXGKQ" &:: Visio Pro Preview
exit /b

:: Office 2021
:fbdb3e18-a8ef-4fb3-9183-dffd60bd0984
set "_key=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH" &:: Professional Plus
exit /b

:080a45c5-9f9f-49eb-b4b0-c3c610a5ebd3
set "_key=KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3" &:: Standard
exit /b

:76881159-155c-43e0-9db7-2d70a9a3a4ca
set "_key=FTNWT-C6WBT-8HMGF-K9PRX-QV9H8" &:: Project Professional
exit /b

:6dd72704-f752-4b71-94c7-11cec6bfc355
set "_key=J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T" &:: Project Standard
exit /b

:fb61ac9a-1688-45d2-8f6b-0674dbffa33c
set "_key=KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4" &:: Visio Professional
exit /b

:72fce797-1884-48dd-a860-b2f6a5efd3ca
set "_key=MJVNY-BYWPY-CWV6J-2RKRT-4M8QG" &:: Visio Standard
exit /b

:1fe429d8-3fa7-4a39-b6f0-03dded42fe14
set "_key=WM8YG-YNGDD-4JHDC-PG3F4-FC4T4" &:: Access
exit /b

:ea71effc-69f1-4925-9991-2f5e319bbc24
set "_key=NWG3X-87C9K-TC7YY-BC2G7-G6RVC" &:: Excel
exit /b

:a5799e4c-f83c-4c6e-9516-dfe9b696150b
set "_key=C9FM6-3N72F-HFJXB-TM3V9-T86R9" &:: Outlook
exit /b

:6e166cc3-495d-438a-89e7-d7c9e6fd4dea
set "_key=TY7XF-NFRBR-KJ44C-G83KF-GX27K" &:: PowerPoint
exit /b

:aa66521f-2370-4ad8-a2bb-c095e3e4338f
set "_key=2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ" &:: Publisher
exit /b

:1f32a9af-1274-48bd-ba1e-1ab7508a23e8
set "_key=HWCXN-K3WBT-WJBKY-R8BD9-XK29P" &:: Skype for Business
exit /b

:abe28aea-625a-43b1-8e30-225eb8fbd9e5
set "_key=TN8H9-M34D3-Y64V9-TR72V-X79KV" &:: Word
exit /b

:f3fb2d68-83dd-4c8b-8f09-08e0d950ac3b
set "_key=HFPBN-RYGG8-HQWCW-26CH6-PDPVF" &:: Pro Plus Preview
exit /b

:76093b1b-7057-49d7-b970-638ebcbfd873
set "_key=WDNBY-PCYFY-9WP6G-BXVXM-92HDV" &:: Project Pro Preview
exit /b

:a3b44174-2451-4cd6-b25f-66638bfb9046
set "_key=2XYX7-NXXBK-9CK7W-K2TKW-JFJ7G" &:: Visio Pro Preview
exit /b

:: Office 2019
:85dd8b5f-eaa4-4af3-a628-cce9e77c9a03
set "_key=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" &:: Professional Plus
exit /b

:6912a74b-a5fb-401a-bfdb-2e3ab46f4b02
set "_key=6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK" &:: Standard
exit /b

:2ca2bf3f-949e-446a-82c7-e25a15ec78c4
set "_key=B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B" &:: Project Professional
exit /b

:1777f0e3-7392-4198-97ea-8ae4de6f6381
set "_key=C4F7P-NCP8C-6CQPT-MQHV9-JXD2M" &:: Project Standard
exit /b

:5b5cf08f-b81a-431d-b080-3450d8620565
set "_key=9BGNQ-K37YR-RQHF2-38RQ3-7VCBB" &:: Visio Professional
exit /b

:e06d7df3-aad0-419d-8dfb-0ac37e2bdf39
set "_key=7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2" &:: Visio Standard
exit /b

:9e9bceeb-e736-4f26-88de-763f87dcc485
set "_key=9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT" &:: Access
exit /b

:237854e9-79fc-4497-a0c1-a70969691c6b
set "_key=TMJWT-YYNMB-3BKTF-644FC-RVXBD" &:: Excel
exit /b

:c8f8a301-19f5-4132-96ce-2de9d4adbd33
set "_key=7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK" &:: Outlook
exit /b

:3131fd61-5e4f-4308-8d6d-62be1987c92c
set "_key=RRNCX-C64HY-W2MM7-MCH9G-TJHMQ" &:: PowerPoint
exit /b

:9d3e4cca-e172-46f1-a2f4-1d2107051444
set "_key=G2KWX-3NW6P-PY93R-JXK2T-C9Y9V" &:: Publisher
exit /b

:734c6c6e-b0ba-4298-a891-671772b2bd1b
set "_key=NCJ33-JHBBY-HTK98-MYCV8-HMKHJ" &:: Skype for Business
exit /b

:059834fe-a8ea-4bff-b67b-4d006b5447d3
set "_key=PBX3G-NWMT6-Q7XBW-PYJGG-WXD33" &:: Word
exit /b

:0bc88885-718c-491d-921f-6f214349e79c
set "_key=VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG" &:: Pro Plus Preview
exit /b

:fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9
set "_key=XM2V9-DN9HH-QB449-XDGKC-W2RMW" &:: Project Pro Preview
exit /b

:500f6619-ef93-4b75-bcb4-82819998a3ca
set "_key=N2CG9-YD3YK-936X4-3WR82-Q3X4H" &:: Visio Pro Preview
exit /b

:: Office 2016
:829b8110-0e6f-4349-bca4-42803577788d
set "_key=WGT24-HCNMF-FQ7XH-6M8K7-DRTW9" &:: Project Professional C2R-P
exit /b

:cbbaca45-556a-4416-ad03-bda598eaa7c8
set "_key=D8NRQ-JTYM3-7J2DX-646CT-6836M" &:: Project Standard C2R-P
exit /b

:b234abe3-0857-4f9c-b05a-4dc314f85557
set "_key=69WXN-MBYV6-22PQG-3WGHK-RM6XC" &:: Visio Professional C2R-P
exit /b

:361fe620-64f4-41b5-ba77-84f8e079b1f7
set "_key=NY48V-PPYYH-3F4PX-XJRKJ-W4423" &:: Visio Standard C2R-P
exit /b

:e914ea6e-a5fa-4439-a394-a9bb3293ca09
set "_key=DMTCJ-KNRKX-26982-JYCKT-P7KB6" &:: MondoR
exit /b

:9caabccb-61b1-4b4b-8bec-d10a3c3ac2ce
set "_key=HFTND-W9MK4-8B7MJ-B6C4G-XQBR2" &:: Mondo
exit /b

:d450596f-894d-49e0-966a-fd39ed4c4c64
set "_key=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" &:: Professional Plus
exit /b

:dedfa23d-6ed1-45a6-85dc-63cae0546de6
set "_key=JNRGM-WHDWX-FJJG3-K47QV-DRTFM" &:: Standard
exit /b

:4f414197-0fc2-4c01-b68a-86cbb9ac254c
set "_key=YG9NW-3K39V-2T3HJ-93F3Q-G83KT" &:: Project Professional
exit /b

:da7ddabc-3fbe-4447-9e01-6ab7440b4cd4
set "_key=GNFHQ-F6YQM-KQDGJ-327XX-KQBVC" &:: Project Standard
exit /b

:6bf301c1-b94a-43e9-ba31-d494598c47fb
set "_key=PD3PC-RHNGV-FXJ29-8JK7D-RJRJK" &:: Visio Professional
exit /b

:aa2a7821-1827-4c2c-8f1d-4513a34dda97
set "_key=7WHWN-4T7MP-G96JF-G33KR-W8GF4" &:: Visio Standard
exit /b

:67c0fc0c-deba-401b-bf8b-9c8ad8395804
set "_key=GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW" &:: Access
exit /b

:c3e65d36-141f-4d2f-a303-a842ee756a29
set "_key=9C2PK-NWTVB-JMPW8-BFT28-7FTBF" &:: Excel
exit /b

:d8cace59-33d2-4ac7-9b1b-9b72339c51c8
set "_key=DR92N-9HTF2-97XKM-XW2WJ-XW3J6" &:: OneNote
exit /b

:ec9d9265-9d1e-4ed0-838a-cdc20f2551a1
set "_key=R69KK-NTPKF-7M3Q4-QYBHW-6MT9B" &:: Outlook
exit /b

:d70b1bba-b893-4544-96e2-b7a318091c33
set "_key=J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6" &:: Powerpoint
exit /b

:041a06cb-c5b8-4772-809f-416d03d16654
set "_key=F47MM-N3XJP-TQXJ9-BP99D-8K837" &:: Publisher
exit /b

:83e04ee1-fa8d-436d-8994-d31a862cab77
set "_key=869NQ-FJ69K-466HW-QYCP2-DDBV6" &:: Skype for Business
exit /b

:bb11badf-d8aa-470e-9311-20eaf80fe5cc
set "_key=WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6" &:: Word
exit /b

:: Office 2013
:1dc00701-03af-4680-b2af-007ffc758a1f
set "_key=CWH2Y-NPYJW-3C7HD-BJQWB-G28JJ" &:: MondoR
exit /b

:dc981c6b-fc8e-420f-aa43-f8f33e5c0923
set "_key=42QTK-RN8M7-J3C4G-BBGYM-88CYV" &:: Mondo
exit /b

:b322da9c-a2e2-4058-9e4e-f59a6970bd69
set "_key=YC7DK-G2NP3-2QQC3-J6H88-GVGXT" &:: Professional Plus
exit /b

:b13afb38-cd79-4ae5-9f7f-eed058d750ca
set "_key=KBKQT-2NMXY-JJWGP-M62JB-92CD4" &:: Standard
exit /b

:4a5d124a-e620-44ba-b6ff-658961b33b9a
set "_key=FN8TT-7WMH6-2D4X9-M337T-2342K" &:: Project Professional
exit /b

:427a28d1-d17c-4abf-b717-32c780ba6f07
set "_key=6NTH3-CW976-3G3Y2-JK3TX-8QHTT" &:: Project Standard
exit /b

:e13ac10e-75d0-4aff-a0cd-764982cf541c
set "_key=C2FG9-N6J68-H8BTJ-BW3QX-RM3B3" &:: Visio Professional
exit /b

:ac4efaf0-f81f-4f61-bdf7-ea32b02ab117
set "_key=J484Y-4NKBF-W2HMG-DBMJC-PGWR7" &:: Visio Standard
exit /b

:6ee7622c-18d8-4005-9fb7-92db644a279b
set "_key=NG2JY-H4JBT-HQXYP-78QH9-4JM2D" &:: Access
exit /b

:f7461d52-7c2b-43b2-8744-ea958e0bd09a
set "_key=VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB" &:: Excel
exit /b

:fb4875ec-0c6b-450f-b82b-ab57d8d1677f
set "_key=H7R7V-WPNXQ-WCYYC-76BGV-VT7GH" &:: Groove
exit /b

:a30b8040-d68a-423f-b0b5-9ce292ea5a8f
set "_key=DKT8B-N7VXH-D963P-Q4PHY-F8894" &:: InfoPath
exit /b

:1b9f11e3-c85c-4e1b-bb29-879ad2c909e3
set "_key=2MG3G-3BNTT-3MFW9-KDQW3-TCK7R" &:: Lync
exit /b

:efe1f3e6-aea2-4144-a208-32aa872b6545
set "_key=TGN6P-8MMBC-37P2F-XHXXK-P34VW" &:: OneNote
exit /b

:771c3afa-50c5-443f-b151-ff2546d863a0
set "_key=QPN8Q-BJBTJ-334K3-93TGY-2PMBT" &:: Outlook
exit /b

:8c762649-97d1-4953-ad27-b7e2c25b972e
set "_key=4NT99-8RJFH-Q2VDH-KYG2C-4RD4F" &:: Powerpoint
exit /b

:00c79ff1-6850-443d-bf61-71cde0de305f
set "_key=PN2WF-29XG2-T9HJ7-JQPJR-FCXK4" &:: Publisher
exit /b

:d9f5b1c6-5386-495a-88f9-9ad6b41ac9b3
set "_key=6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7" &:: Word
exit /b

:: Office 2010
:09ed9640-f020-400a-acd8-d7d867dfd9c2
set "_key=YBJTT-JG6MD-V9Q7P-DBKXJ-38W9R" &:: Mondo
exit /b

:ef3d4e49-a53d-4d81-a2b1-2ca6c2556b2c
set "_key=7TC2V-WXF6P-TD7RT-BQRXR-B8K32" &:: Mondo2
exit /b

:6f327760-8c5c-417c-9b61-836a98287e0c
set "_key=VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB" &:: Professional Plus
exit /b

:9da2a678-fb6b-4e67-ab84-60dd6a9c819a
set "_key=V7QKV-4XVVR-XYV4D-F7DFM-8R6BM" &:: Standard
exit /b

:df133ff7-bf14-4f95-afe3-7b48e7e331ef
set "_key=YGX6F-PGV49-PGW3J-9BTGG-VHKC6" &:: Project Professional
exit /b

:5dc7bf61-5ec9-4996-9ccb-df806a2d0efe
set "_key=4HP3K-88W3F-W2K3D-6677X-F9PGB" &:: Project Standard
exit /b

:92236105-bb67-494f-94c7-7f7a607929bd
set "_key=D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ" &:: Visio Premium
exit /b

:e558389c-83c3-4b29-adfe-5e4d7f46c358
set "_key=7MCW8-VRQVK-G677T-PDJCM-Q8TCP" &:: Visio Professional
exit /b

:9ed833ff-4f92-4f36-b370-8683a4f13275
set "_key=767HD-QGMWX-8QTDB-9G3R2-KHFGJ" &:: Visio Standard
exit /b

:8ce7e872-188c-4b98-9d90-f8f90b7aad02
set "_key=V7Y44-9T38C-R2VJK-666HK-T7DDX" &:: Access
exit /b

:cee5d470-6e3b-4fcc-8c2b-d17428568a9f
set "_key=H62QG-HXVKF-PP4HP-66KMR-CW9BM" &:: Excel
exit /b

:8947d0b8-c33b-43e1-8c56-9b674c052832
set "_key=QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4" &:: Groove - SharePoint Workspace
exit /b

:ca6b6639-4ad6-40ae-a575-14dee07f6430
set "_key=K96W8-67RPQ-62T9Y-J8FQJ-BT37T" &:: InfoPath
exit /b

:ab586f5c-5256-4632-962f-fefd8b49e6f4
set "_key=Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX" &:: OneNote
exit /b

:ecb7c192-73ab-4ded-acf4-2399b095d0cc
set "_key=7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ" &:: Outlook
exit /b

:45593b1d-dfb1-4e91-bbfb-2d5d0ce2227a
set "_key=RC8FX-88JRY-3PF7C-X8P67-P4VTT" &:: Powerpoint
exit /b

:b50c4f75-599b-43e8-8dcd-1081a7967241
set "_key=BFK7F-9MYHM-V68C7-DRQ66-83YTP" &:: Publisher
exit /b

:2d0882e7-a4e7-423b-8ccc-70d91e0158b1
set "_key=HVHB3-C6FV7-KQX9W-YQG79-CRY7T" &:: Word
exit /b

:ea509e87-07a1-4a45-9edc-eba5a39f36af
set "_key=D6QFG-VBYP2-XQHM7-J97RH-VVRCK" &:: Small Business Basics
exit /b

:embdbin:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ [BAT85]::Decode($d+"\\SppExtComObjHook.dll", $f[$r+1]) }; function Y([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $d\$tmp -F:* -R; del $tmp -force }

<#
:: 1st Block (above):
:: Powershell code which decode the embedded files
:: https://github.com/AveYo/Compressed2TXT
::
:: 2nd Block:
:: SppExtComObjHook-x86.dll   SHA-1: 1b6ee9e99b1dbcfc9427bb3a61c65ed53667fc22
:: 3rd Block:
:: SppExtComObjHook-x64.dll   SHA-1: 1a6a3e40f610a6394ef539a039308dbe2f526ac1
:: 4th Block:
:: SppExtComObjHook-arm64.dll SHA-1: 3d158e9cbd6de13f954e8dba356e369b557fe2d9
:: 5th Block:
:: CleanOffice.ps1            SHA-1: eb20e53561980734f678894d29f8ff0783ff769a
#>

:embdbin:
::O;Iru0{{R31ONa4|Nj60xBvhE00000KmY($0000000000000000000000000000000000000000$N(HU4j/M?0JI6sA.Dld&]^51X?&ZOa(KpHVQnB|VQy}3bRc47
::AaZqXAZczOL{C#7ZEs{{E+5L|Bme,a00000v}wA@[CekK[CekK[CekK;YU#E]9a;N[CenL)FoL=XL{Y5^z2XSXL{C}[d)tLQfXso[CekK00000000000000000000
::P)=U$OaTM{NZQs#0000000000.~b{a3jq!s04[Lk01N/C00000&o^jz01yBG06-i$0000G01yBG00IC21][s6000001][s6000000B_]R00aO4m_wlx0RTV(000mG
::01yBG000mG01yBG000005C8xG0000000000girtgC/$Ke0000000000000000000000000000000AK)ByaE6Ka2o(sH~/^u000000000000000000000000000000
::0000000000000000000008jt_ga7~l000000000000000000000000000000E^7vhbN~PVsVx8i01yBG04[Lk00aO4000000000000000AOHYhE[WYJVE^OC00aO4
::06-i$00aO405Sjo000000000000000KmY,1E[[;8bYTDhPy-w}08jt_00aO405$,s000000000000000KmY)hE]=jTZ){&eyaE6K0AK)B00aO406G8w0000000000
::00000KmY)j0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::000000000000000000000000000000asY4uV,qjhbO1B}E(yZzYyfNk00000QgCBabaH8KXF^RiWNB^]LvL-xZ,yf=00000QgCBJX?Md_Zf8bvZ,5a^a&pa7LTPSf
::X?Mm&00000QgCBabaH8KXGU]mWmf;IQgCBJX?Md_Zf8bvWn}/WQgCBIb9ruKNp5L$X;=-?dSysqZe)m^00000QgCBIb9ruKLvL-xY.Mz1Lt$+e00000PGoXHb9ruK
::Lu^efZgfLoY.|7kPGoXJY.wd~bVFfmY&&};Zvb.uZ~$.sZvbKdY5/Qp00000a{zDvZ~$+rVgPCYa{vGUQvh&PZ~#RBcmQ-(LjZ38Z2)UIVgPCY00000000005C9SY
::&n$$(4ge4U/1B?17yudo[DKnH1wDfY_Q_A4?t3d4Z16Y7;nPkf-lX~/B5j4{$ulF4rNb0SK_h3fs64Liicu?ILt-Sp=Aj)Sr/XZE[oPh|dpc/kI9Omm.uZm;d32^r
::YfqyG5OvGFC[rgk^SDyLkDzhUo9vx,i;wr,qqO}@RbVPQ-Q3_u0o8OPicBKvDfr+]e3;o{rdY0UD={Spp@wGKh=tfp]c]kNQbmKO#oc+aWT1ZP?@NkA7(wb[N^^~{
::A@=URMNRQLsc2W7sY,eW/sHY~o6AN71=va;$)-3YE1mz.uvWR)wT|=l)vkiv_3wR0Nm{rs{M1X@o-8VeXD.TPE^8BC)x5q(1u+@,5gsc|KWbS4@aE.365zvo1ODhX
::Je09F)O&J_W8TR{X([nURkV/qgz7=)tzBIj#C@2ek/({W6)g;6[5m_b($U&5UVOO-OJ5Yt.!hc(5Qo9qPWyP?1,B{cpkiK|u/rgY{vPL@_@_yaQVD9-FgB$+zd+m(
::f&Dh;eB)KSn=k+|G?$^=#NO&4RC|/&rotmV@o5?nLi+o^2ri,!DA]?kc3YxJZHv),a_]UShG?_/+TCU[U1heCY/Z^W{q4EhUKK_Hr/VM2kl3pLjJ)qd^vBawxU+qD
::([3L0&0CYR!LPjo0TYUAI,}1UPiNffm.5ff[U.T0maKFl=dCq_/_uk|9ChDrNAVhQ9Vx|$Z@|F(su/c.{8m0o#@pBpn&ltsc-Fb$AKj=khzG|pu[VqjCxGl;U{Qam
::8MR6cE#.QjlgXU#px_[At}6Ag$m^d2gHxGd7b]sQx^8zl/b|0ORUr)0V|/ge[[sF!Fac,P{[1H]&7V##_dLTtt;;8goTPHVxBZhQHb3{wG]OS7ao8~x1ji&87@uT]
::2NHnd?nE~x34;(e8,W/lQajeODdR7MQ^&qJApEggYRkSkN=#VK)C[1ILrpV;Mfn1MP(}WgQKLYQlASp9ytdjQ5dZVi&@uOlUzbD|#HW5eWL-6]V1ZBEA}WxGM))(2
::.d-pa/4)T2Nd^cb!qco_k)K0m=g2p0jnz+6Y,zH[WqPg&x^Bin9Hz9!=.qT5OTCMVa6YwWNCWl_VKrB|hQS[4/rN(lY1xjHn/wVh(Q(PijG?7Qzve;{L76QNuvEJi
::2m$~A4rTxV5C8xG(3;_rDzaV6RsYEEgJi]TWgmZ#(8/p;mAp(!0K9Hk;YPj(k_Km4y!gQ#Uj8Xr_?{!c?hO9=nX6{Xmg&7NX~3UueI?-8w5N3i6w[a|+9-S;1PqBl
::hd]6$I8$0?am!^fj7FnMqc^W($;]wti6-XjsHxXNla0[gpCB1nw+Ka~M$N!Lux,abSEM(Tugt~|4,#xCod_p4cw6_F]Zg./#mnnNgY,6;PG,3ohco4mhcHJ)iG}x3
::G9g/Y;W}J_Z@{rPpON.K.Ic6JBfp@~^0V!ak=fN-]_wEeC-A.1#Azr-nbIL]4-dU17id?Ib$en&o+aQJEuNg0Syr)T1LpBeoFDM+0k{~5z~i5m@4ue;pCv,z1?Un|
::Hr9O7Vj1Z~i&&!E!an;jy0P5Lli.0(q-;|wQjz_nPZQ!/5snv4oU+MyoD~sBSQEwJKK=tjq[p_)Ajxx1fX9!]1uR.gmk[=o|H&Z_kZ5FW1~wW.hO1eNxJu4~ShGMp
::NLjB&k~?q;AIyGvJA1jiq?Ly]mluipy-XvS8B,SV_uj?qg2];|tyAb$--.?tu|qvgqYN,_ogkIQ77Nve2uQC&xI42rCqo#rv7UWlHt(K[hTx_J/Cq)F8}^w[3o^$N
::fl9Y+EBgF_zwxH#K&K+ts.MSuq7_^-M+^LkB_(u|gW;ls?,~f470(SwiK)4OuSW89#y19I00000m/e9)K$,1_KeTKY000000052Yl.;bzr~m+~K;-IPzml2~00000
::0Kl)uBht$O(Hw[e09FqP|F3Zi000000Kfw}FpA9q(Hw[fz#&6Pzk7+i000000KkpX&gW9H(Hw[gzz{@oziOr,000000Kn$YQ0(nG(Hw_iz@at_zr-Y400000005_l
::/#t&I.4O&]Fm)U_FQPF400000004S9_CHckHxmU1AWi[PAA2zY000000KjyvY@s/rHWLL000000Ps0EJ000000KjyvY@s/r(Hw_iI6]A_SCA^J00000005_l/#t&I
::5+TCj00000zv1Kn00000005_l/#t&IU/qLGz;=Bee_[{=0000006;8;uHni7(Hw[hfR6GF|35-x000000D#]8L&q!b(Hw[iKo;.ezsOq~0000006?2o.22c0(Hw_j
::fLgLAe~05J00000003EunN!pO(Hw}lz{Ch5zwtRE00000007C#jjGoH0oMQkau+yq0oMQku]j,aG8F(.[FM]K0T}=QfF&F_91Z{gIXD0S91Z{gV@^V}91Z{gd_|!X
::91Z{g]ko15Lvnd=bU|Zrb!l?CLvL;$Wq5Q~00000K?$PmRscZ(Pyk5+GXOFG00000Lvnd=bV-S-Z,p_@WqAMqLvnd=bW?$@OJ#XbVRB)[00000Lvnd=bVY7sa)Qrc
::00000Lvnd=bVOxybaHQbOJ#WgLvnd=bW(w)Wnpt=LvL;$Wq5P|Lvnd=bVOxia)Qrc00000Lvnd=bVG7wVRU6kVRL8zLvnd=bVy.yXhdOjVE^OCLvnd=bVp[$NMUnm
::P-[XmZ2$lOLvnd=bVOxybaHQbNMUnm00000Lvnd=bW?$@NMUnmP-[XmZ2$lOLvnd=bVp[wQekdnZ,2eoZUAEdVE|)QZUA2ZX#j8lUjTFfV,qdf00000B?.~)IshdA
::a{yZaB?.~)T?t;8F#s|EHvldGFaRz9FaRz9F#rGnM_d)Vd2[7SZA4{eVRdYDOhZXT00000YXD]casX}sWdLjdGXOFGE(yZzYyfNk00000B?,r0H2_&0EdV6|FaR|G
::bpR~[B?,r0GXQk}EdV6|FaS0HbpR~[B?,r0G5~b|EdV6|bpR~[B?/5,E(wn9FaR)BFaRw8B?,r0GXP_(B?,r0Gyr4)00000O8_v)QvhE8MF4F8bpUJtVE}XhX#j5k
::ZU6uPO8_v)QvhE8K?&X^bO31pb]u_jbO31pZvbupNdRsDbO2=lasYM!VE}9Z00000O8_v)QvhE8QUGNDZUAKfcK~4kYye3BZUA&uWdL#jb]u_jYybcNO8_v)QvhE8
::NB~y=NdQCu0000000000O8_v)QvhE8Pyk5+L/zm]B?,r0H~@$^cmOQ_B?,r0GyrG.cmOQ_B?,r0GyrG.cmOQ_B?,r0G5}},cmO2/FaR;DXaINsEdV6|FaR;DXaINs
::B?,r0G5}},cmO2/FaR;DXaINsB?,r0G5}},cmO2/FaR;DXaINsB?,r0G5}},cmMzZ00000NZQs#000004gdfE00000000000000000000NZQs#000005C8xGBme,a
::s2czPs1E=DAOHXW+!{^tM=kmNp2cE]cqn[M#x6Y-#S?93N~4-/NZQs#Rg3J4MGS.J0dypT=mT{q|8+$y[IU|&xiCNg5NZhMBmw{ci$xH}0PsKn5bFv5bqN0z08juB
::Gr(My!VCaai|m8!9Jozd0038u_DQr?4}}K.004^eJb]qoP)=U$4~6#t004_OIDh~E0ENj9gy/YO0E?h/ga7~lg}[Jl,#H0lR,f^{1ICGU=!r$.JMa(K!~g(QQ.gjC
::01t$@0001sUJ#201K$J3iADT^I{,+almGw#zCb^#5XVLA2mk/8iw=d!bOt,MbztioiwK3ucv$~.{EbHf1Hn.L6.ZD35LsD/z/!]4Mfizb[K9;&jYagwMf3y!002/p
::MetB-|Nj,PPyi5&1BnKUMg+lijYarS|8[9{Mf6aOMetB-|Nj,PPyi5&Mf_+t7=vx?0d;Xo!vurC1c]obgF65Zg?V1=|BH3#gT[qzb@7]F;PU/A|NsA6USG-?Rg3Ia
::z7P=r5D(Ko)dbbBb^D.,2?&sKPyi5v#0.VNbqI[14CuA~|Nn!=2!Z}65daW&+Lvb}5CB$JY6!Y8KmZV4i~iB@YsVNlOb_[v4uilL1JHxW|0~A?[aqDL^l5R;.@/EV
::01!LSbOk#~6mtTL)2GJ4UtYsii|m8!7?z~nR##C^i~9c-2v7hJjYa&3^E3#Q[K9D){}l{S01$=1c]HjF{7{WW[K9D){}m8W01$=1c@2_S?n2x@Md)of6&;ea5RFCn
::P,#ma=urQ4{QnggPyi5xz;2|Tb[-@MUtY;?Rg3I{;oJW]G,ebri~5V&i]qw4BoHgZJs1N30DN|fbqH!0S(2,}6ovMFegAa~i^@uvBqUep]Z[^?i]kFZ!T16L[QcSQ
::L@k4cY8YAf_Tunci]z,aBqS[tnfH7o7,/#MeG!XPBq+o]=#T(Z|BGBCB;mB4Oe8Fe,XVx#|No18Bp__=BoK@oE5kh+0{{Shl?c=Mi_R@G=-OWG0AF5Pi{+Ly2mp+s
::EAxxSiF70(Y8aVW^xXuLBoxv3g}{DCYtR[2);{.9d@YAp23h}g42$xKR3sqtrHxD]EJ#uRbR./$)1}DOB#X=Fr2^x}i$o-Ci(P{WY8Y9Cz;;VzbR.~+d@XNqd@XYC
::e2/^2=!twJ5R1l)^nG)kgZKh_J8A}5Y6gvbBrNl#i-m($jYK3YS(Q[YrT=vdi-m($i}z3~]J,Abi]iGv_7^ds&TVY)0{{SO7-LEAEAuP+iF^mwgZ@mmtm!ZR|Nm8s
::@1St$Q/WfgRqTmH[K#ql!ViR/0001uSQv@2=sVgEgpL3J0LMl67ytkOY6e.0]NC(bjeX4XrHg(wi&s}DcocITQBaG-E5U=n7,PKeP,4C6i$)B,#0.m7@2Gv8IE__W
::1MrD/;U4!^{BtLZMc|9pi)T}K_.}36-JovCY6e/Jr8_0ta^x(n[QYRKiCy&IUFeH^]n,qC7?#wzY6gSD7?QlzS[Wek-7E@!|NsAAUW?w8$.+ExRg3I^@6@2_08?^r
::bqI]bSBv^IP4FwniGBE+gU0A/7-LrEiFNeR^.oi0i]eO,nfLhv..&uPgZL;ZPm5g+K#NWI!T1Af[EB|OFpEX^Q/ie_1IJM4)E;Pfi(gxK&2O.OjRY1[=+eL10E;QR
::i^lY0E5|F=i]&A&0ssJuefU$0UHnpw3^j}y|8+$D(sK}m=#(Bg0F6T!?kW&b6gf.$atU^^gX;VO_,QD#O$a.|a}kSE6gf~7atU^^gTNR(!E,A8)2MhlRs4x{3^H;t
::ON)-8Idc@pGj}-Ra2Sbo{5!(QC29s+i|~zo,z=_o23d?vi}H(^{PU&123d_D,o,q}rHf4n?jI0$gTfd+)Q?|vbqtGr^=#2gJNI]1jZWx]P4tUX6gg28axr(5gX;VO
::,?oss23d?qjeWrLrHgg-i~DK[S[Wffee{iWz.k6r]QDV,42ymEi&s.{@ihpV7?RWZiB0r7{(KgARs1{mi&kf1NjuRGgwp]306W5T9cl)yi}Q^rsPm;223d?ni~IAX
::Y6e.2b,PK[]QCGAS(Q@HeUS5|Y6gq[i}G3XrGvm2jdhTV_h(w5?v[C5=!ta;JJEGagX;VO]K?F=23d{0c?n-Z]QDV[42$z?23hl]Y6e.2wRr#k|BL)crD^IQi}Q_W
::VE^OB]QD967?oOA23d?n]QDcoVE^OBi}?p+iFFK)y.5H6|7r$Vi}LfOi,,c)_f3JQ]QDV[42_u(|NsAk!WfBF{AvbS]QDV,42ymEi&kgWdH);Zi&sxfUWplrK@IAz
::Tgk!,09A|ZgX|Dfi_j_+^=_/li]55XMf{0P[X_5-P4re,JJEeSi&tBEGzpDP0,OuhYsfH,L?WXoL?Vk|2#rPoiADVD1B,rcYx+@A(}/A,Y6e,Y_vddyrHf7UJHdR!
::JHd4vi$w]i0d,B@z!/5l=xh2Ii(gMy23Z5}1N.x)IaT;5EQ@JHjY9v@^.pVOiA4wl]NU6lJ3$n46pKX(Yx+@A(}s&,1N.x)IYsz;+QfctYx+?!23Z69]QB,2Tgk!;
::09A|Zi$WBG?=/vv,/ZFOLlko$Y6e.0[{N7y]QCGAS(Q@F_tzk~23d_D=!]UFrHeuoi&keQK[[X3i,,Q&eduZiS[WffeF&(4Y6e/JrD^IQjdkdY_tzlWLKHhg6mt_5
::23d?qi~94WY6e.2b@A&x]QB,2Tgk!,09A|ZgX|20W&vMtW(8l4$U]_SQ/XVGzMvQY5Lb+(g}_-Yx_.G65R1/}i~0ZmqsT+55NH4Z0Pt!UY5.~gjZOGb|ImfN4}^Nh
::005!PLjVwkz/zM2v={(oi^Yj-_Tzf/&tHVWjZOSfY5.~gY8-]o6aWzab[?0#g}_+WGxkP?|8y#$(^e)ag}_-by2Ka(5V[cj01+UL_Tzf/(^e)aY5.SO|8[NT)1pNs
::G[/Z)01$=1bribT7yuBtpcnuU=;[jg|D+7H01#,#0RRC1bqG]x09I.M|8[NT)1pNt0yFlDMf_=pbS8[p,h2sig~[dpy66}H5V[cj01+W3_2YW.,h2sii]z,b{80bV
::i$)Z^z/q~!4ctQj5QWKg8M]ow01(yL7yuCHhxq]hqufIP5R1r,Mfhp}Q2+^ii]g7C$.+Q#Rg3IVQK85~01#7FS2O?M[_.+/yPy~V5Q&/Kg}_-Yx_.G65R1#_L.^yy
::qsT+55NH4Z0Pt!US67Wq^+.7Rg}_+Oq0B=75QV]X5xTS)01&7I=oa|[|D)+901&B${83j|Y8-]o6aWzab[?0#g}_)zq2NOR5QV]X5xNW-01&7I=/rtT|D+hT01,Fm
::2v&1!^C{7$|8[NT)1pNt0,m,HMf_=pbR(xm,h2sig~[dmy66}H5R1#_ulN7|qu4^L5R1r,Mf^0z)Thd-g}_)qiw+dE01$=Abr!n#7yuB9&jkvo|No=hLjVwq$cshz
::S5W_aUyH^GTgk!&09I.QE6QGr1Q9U/jZy[I@g4f4Uc+oe54Hpfih!6C01$_.3POYW5Pa?!Rg3I{?=2Db[K&fQY7kaX{}ohF01$+33]USe!Ucoi4~j$u[Q4I5LWBAc
::eC=Mz!(QsygX|EEMetUO[oErOQ2!NFPyi5v#0+dii]^|^gW(^g/Q}kcgWwN}L;R7O1TsQ{_Vf5NUdh8#R#&JJGyjV/yNkxb=m#YM0RaJP(?M[!C4YZ]{{z4@)2K]4
::&E7=8GtR.r2,Jq-GsrW}Gs=k!yGMin1boL|UR&Rei|kQRQ!~JcP54&cMf|z}0RaJ5i#(mg$HC|aBf[Lg8/i#!e}8}f1Hd!bi]IX^2P493$Qz5pC4YZ]{{z4@$cw@j
::=m#UhYw#P3!6koxfByr)Gw^SnGx0Omi]hw}GsiQ^!N3r~$p|yZ!NLfOMf[|$i(gkD(NIu241z$3Mf{7)x(Z-J0fYDie1?0MTgk(#i}AVW5daWZi_M9Z{{R1K5MPVV
::=#BpW|LC6n|No25=mP+$|Ba965daX2,63gU|Nn!?5P|=o0001d+r.#PV,daC=yU&6|6hyNxrh;~5MJpz{{R0|R,Uhuh!Ox0SBuu[DgOWeyO00@05kuK(gggk|No26
::54Hp_ih!6B01$_.F-qd,5Pa?8i|7#m5R2C6_ThU[i^YlG{{R1j?kx)8f60r]=s]De{|~kVH/RCm5(#g01UE/6{}6obUyIhc=n)+AUR(wH{r~@}i|kR0MF?-@i}8yH
::i]2EMgWwN9SBv^M90.YB{DJ!&0RRAY1T);vOi+mb1UXRu6;AOJ5QD[Fh4yp^i_R@J=zRVE|BKJ+;Np8ugTwH28I4yEi^hq0{r~[i#}JFp=,s]8|AXrgb@1vs2#ZDh
::i2,afKwn/4$.+Q#Rg3IVQBzinbqH6B_imZkUHJCH4|D_G!0Tj;NALr~JI8h&i_a|H=.(SS|BKU(OYrCp{r~[i!|,&bcj${;^?0fz!~XyOgZmJ4Esa;I1JjBEJP?yx
::Gye~]$b_Uj6N~W=wgO5!!F2{lPKEY=[{4r{JNb12E7?!]i5[].UR&k-1OQW3SBvuKwfz78i_Ka+5daYAsr?+{Y6_zN5daWdi^Yk/{Qv,x!2JLJi^YjI{r~[q+{D?R
::h5Y~jgU1kq?j8D#i^Yka{Qv,xoc#a+UyIhc,bx8^UR(v6{Qv)|i|m8!5TVFJ01#7FGuy{S]dJBL0Eu1ri}HzG{8x-lg}_-bx_.G65V[cj01+V$]Z+/($U]_SXaE2J
::[M/+SQ/kjdQUB0|zz?9v0001@;U/[ug}_-bx,Qn,5V[cj01+VC]Z+/(;U/[ujZOSfQ(VURfB,phb[?0#g}_+Pq3A/Z5QV]X6uK.K01(yL7yuCHH}n7hqv&5b5K~rH
::|8[NT)1pNsHKFW701$=1briZh82}KupcnuU=nnJ$|D+^f01#7-P4rR!bqHz#|8[NT)1pNs1vB;W?jI1Oi$)l}z/q]y4e(z&5QWKg8M/sz01(yL7yuCH(-_BOqwqrj
::5R1r,Mf^0z)Thd-g}_)tiw+dE01$=Abs4)(7yuBtpcnuU=&Vuf|D+VP01&7Fi$)ZTQ2+^ii]g7C$.+ExRg3I{?@~GlB51vdXo3H61dH;jG3a3u003$rXzjJSj8ahO
::M.l+41$sv+#,Iz,iADI0Mc7b,)pt@{F=^~jRs34$bN?JTY7mV[=ulSw6/x0F5QD[FUdh8&i|m8!BvV#bGylg$[E_yH0EvD0i}Hzm]of1]iGARSedLLK=!t#o$3[&?
::fB,o6$#fBF5K)9hKmZW_6/x0F5QD[Fp~yo35WAol01$=1brHIV7yuB9&jnVZ|No=NLjVwH0002-Y8Y2njZOGb|ImfN4}|Oh005!PLjVwkz/zM2v={(oi^7Ss[(Es$
::&tHVWjZO4XS66BrXaGO}5dU[f|ImfN4}_b@005!nLjVwkz/zM292o!,i^7S2[(Es$;U/[ujZOSfSO0bNXbFG.0RMIP|ImfN4}]pO005!&LjVwkz/zM2WElVui^7RZ
::[(Es$]g{p/SB,{FQECPMb]QO)g}_+Lq3A/Z5QV]X5xOiH01&7I=nnD!|D+)b01#JJ|8+reb]QO)g}_)(q3lBd5QV]X5xP7X01&7I=/rYM|D+^f01#J=P4H3wbqs0)
::|8[NT)1pNs1vB;W?jI1Si$)N?z/q,v4cJ2f5QWKg7P{yd01&7I=(JDl|D+JL01&7Fi$)NM|Iv#[{Dr_DBa032LjVwk$#oXGP#FLai^7SM[c/j#[IwF+i]z,b{80bV
::i$)Z^z/q/w4ctQj5QWKg7P|Nt01&7I=vMIm|D+VP01&7Fi$)ZXQ2+^ii]g7C$.+Q#Rg3IVp~yo35K~q(-l&svUHn(z_GvrB6uO8Q01(yL7yuCH7x4f8qsT+55NH4Z
::0Pt!UQ(Wvi{89hVg}_)zq4-}p5QV]X6uNjB01(yL7yuCH=kNdjqxeGr5LZ^Jb]QO)g}_)LGxkR70,mvDMf_=pbSH}q-)Q5mg~[dpy7)9X5V[cj01+V~[Bja!-)Q5m
::i]z,b{8Lc?)O.-kUR&k.Rg3I^_2YX~0CX9F_2YY00CWU1!0Q;QbqtGD2;XQ8|NrX@|8+$DRS4-+^W&D~!UzCWi|kQTi_t99R,6OYi,p2n#t2qZQ2!NFPyi4Eb]MKs
::hyVZpY9vus{}ohF01$+37&O&BjYcG1Tgk(!i|m8!BvXsoiB;eojdBQAivx@!i2{vNFz7=L004!-c@pZti^Yko2LJ$8{}o)N01&5-{Ec&!S(alm0ssI2i9!U8x_-S)
::0Hvj-rHeyUi&VFEf|vjR0Evp60001sLr@@2iGrX2000lS1t?rO5Q(1M0000Fw,[Ld01&0Sr~m+}Gr$W(REtYkiAC][Mc_14MevDD]icm598drdUtU|u!(Qsyf$SIp
::002{q!i|-2|NsA1SBuDrMeOK2|NsAul]p/7|Ba38{r~]y[BaV+i|~zo&rn4=b[/1|jqLsZ|BK$MtE+4?jeXp!tE/Pn{t$dE!R_c$#=-nOE7,x$]o!YzRm^Xa=;WIc
::|AWI2gYE&!@u(K!iCyrEUG$4h[P,rc4pUK$Mch#6(kX;ogTxGtRm{dj9o(G8P2A|W_Tzfm9aK/N5RFC5Y7kJ3h3x)R{{zNQ|I=il7ytkOjZP4YbqIsR42celKrm5i
::0RM0p{}nV+01,E-i}/O(@EU}$P,@xcgTxGI?/M1(|8@yDbqxR0gTxGtg(hC@|7ffL002;_brAp4gTxGr$6sDs$.+c)Rg3I^?^7ql08?^r[_-9KS85Q6b@l1;jZOrK
::g?@S]|A|fTUp+W.1psvxiyc&[01&CZc?e$YS5Z,])}Tne?o$wmjZM^);oy5tiyc&[01&Bu+K]fAg@Rq||A~F,Q2,154vj^VgTxF^jfHsr|Nl^vT?bz5i$)N{4uin^
::iACg&MdVP8Mg(mkME)E(jYbrWmBju3|A|HPi]&Ap{r~[smBju3|Ba1]{{R2zQ~m${jkSpW|No2E1MrK]=xhA{|BXfvi_P)#(gjkf|No0k[Qc?yCH4RRi$(~&#xRQj
::Jpcd)0Ci_fAV2]Riwz^|01#LIb@}J?iB1HA#t2sb)}Tneja9]rja?f!{{#2v&=!QSjYXu3+=.Uwc?e$YP&F[lO~mLB_v3n|jYYIjR{zt3#0.r]+C2cdP?qFn{{R0^
::|I??^@1RJ&ivW#{i2ncoEAfeS^(ops1pss#Jpcd)0Cg&-jYas2)NK-zc?e$Y=+)5[|LYb}jYarSjg]T0|No7Rc?e$Y=pgd{|74,U0001sP56s.2!p{CQEC8,4vRoA
::|8N.p6,N!+5dSud^?F~h{{R0^SO3$4#0-Tc00030b[cyr4FA+E#0.sv#Qp#OXsiGL08syR5dYJI#0.naUtU|u!VCaai|m2ypaB2@gZW(25sO6#YFC3G0E;NogCGD{
::?r/U_000C4bS!~7000F5bPRzy000I6bObZW?oARlQ2-n_P?qFP{{R0^|8+rHy#4@G?lTfLQ2-n_P?qFP{{R0^|8+rHjr{.ri]z,b42cLcz{$c009A|ZR,Q8ESB3U/
::b!dng01#IH6;|/R5LsD/z/$Ej^6Ps~gFUc901yCmRe)J#KmZT_bUy!e6l9@o00030br8n|utES30RR91R&oOd01#-]0002Kv^b$7qr5_]5LW,cWKaMQi$xrR#2AS&
::Gr/RBgC)#-01yDbxETNt?jR5[5V[f/01+Um?i^[$bsYb782[z@S62UZ2?/MuUR}Z&09A|ZQG;O509T8B40KS1|8z_+J-MLm5CC,TgFU!I01yCmI.$Hn01#w51sDJT
::0Cg/Av^b$7|8+re6;|/R5LsD/z/zgE1X{VYLjVxy)E0!Wp|nE+5Qzsf!0QmXMKb]p=),|t|5yKY2?/MuSzW?i09A|ZQG.3OLI4l|R,Q8AS9Cpv^H/R-yh8vGWIF{I
::0001WE5Ect01#LIHvbi1Pyi5FS&tuL7ia)h002.|{}otJ01$(F0E5H~Gr/Q(xw$g{5a@9t|NmD1(|h9zUBUzaRg3I{J-MLm5CB$.bqI7ch4yqXiv~Ldv^k-8bta,^
::LjVwDI|Ud3004CuGr,(/LI4n?v^k-8qqsu=5Tm@901+d9xg0bA5a|5q|NmD1(|Y1^1OQcw?_^,]ax@&C=/i4D|NnIii,,QA|ImfNbuKvtutES3bR/;iv^k-8bR0Pa
::xI-LCbrhkzLjVwDI|Ud3004Cci8Z+G01z|4URhnj2mn=!?_^,WbqIy^bUTB63/=XBgFUc901yClEQ39_LjVu}bR(a3v^k-80CXCmyh8vGWIF{I0001W3^CTrLjVwU
::1T);v5V]uM01+U@=?Px!bqrSj(|X;x!UzCWi|m8!41-zeLI4l|Q(x-05Lb(v42xa}iAD5^Mi7Zb^=_pqiADT$UxR&J0CZV{eGC9}Q=zm(01!Jy1a)P]J~=)OLjVwU
::KcT!s01#w51sDJT0Ch2ge;T2OD}#R||8ynjG6nzu|8[L}RrHI@|8[B2ll=exGxk?hcO@H6AW#4hi]l5~xgRwE5a^h$|NmD1bqxP@2?/N7#0-0vTgk(!i|m8!41-ze
::LI4l|Q(x-05Lb(v42ymQi)UwcMfi)G5Q#;nbUTB62mo|7gMADDbT6T^LjVvv26ZN)yh8vGWIF{I0001W80e[40094W{EO3z(/ND!=nDM,|1.er7P.SU01+Um=l}m!
::|8+&ibqN2]gTxG9UR&k.Rg3I{?|9e/S84!^Md,o5;WP)9{}m_u01&5s=xPv(Mfi;I5NNa)01#0B6=-ZZ5QD[FyC6XT5ck,{|AXiHgU1btee{ifFpD,4Pyi5#ef)-=
::SN}8r+_P@hg}_)RgToAg_=CMq5OwE]ee{9,ph5r/bQ-681dUDTiACs)LkNjQ@2SeEQ0Vvp004{lYtI/g!N@W?5Q$Cn54J&9iACrSwm}1pP3Vb5@2SeEQ0UMB003$b
::XrvYZ5dU[ji_a=?@Elt.#0.VNbpng~Y7l6Y761^cb]QO;gTxGlz/y@M_vP^Xi~DL2XoMC35dU[j|JH.V428gT2ZQ@qb]@q0Y7l62761^cb]QO;gTxGlz/y@M_vi6Z
::i~DL2Xk.=u5dU[j|JH.V428gT2ZQ@tb]@q0Y7l626#x-bb]QO;gTxGlz/y@M_vrCai~Ea3]n,?r|7.sk$3@()0001sNC=B@1S_(sMfizDyo34wbqP[a6)mpq5bF$$
::K?;,WO}yw)0{{Sx95A8$LjVwkz/zV5&ozX?x#SrD5a^|?|No=?LjVx})2KX(761]7zuXo85Q#;njYZH;{}n7y01&5!)2GUXf$0FiX+1|Lz?Q6~9mxOy0Eu;DiABVX
::MfgyOeZ1)]0001qMc9o-s8EeX,ojr#Q2!MiPyi4!z)8h-MXF|rMXZfQ$WV;&s8IhEFi.#xGt!H7,o)utKtc}y^wf)EO{gn&$cy@?i&r~&,c)mEXvqKo0B8WfX)~}x
::==cNx0E;oRKv0WC]#3&@Oc)$V|8[Kiw,YDbS5r{_G|fmD01#,Z8UPS#K?u~,{}p6V01$+3FkfC;$._BP@1StKi$Mf~P5l2?i!m$4i&kfLP4ve^^#glP0E77dbqP_b
::6)mpq5bFv-jZO4W=pF)900YKe$.+ExRg3I{??!Il3{#6+2v(;RSBrTBgUSDiMfi)F2x?ryMdX9X|5]{M!2wFq=!s42iAC],z;2L,1&vtjcj#AEiB/)6Mgjl/|8[9^
::[QeCW=s5xa04v6e$Q#!q&7gj/jZNfr2k0RH008R]gTMiU)g0FW=yU+800YJo,HTdE3IYHCUtU|u!UO;Si|m8!3{zHDGr+_ai|~m?{EOI(M-A,X|BKU$#]~Pz004{d
::52W$|1NMnk^&qx8H~$qRPyi4-UkrEias[j^2zTs[P5cY-i$@@l$ctV0iB0]2_9N~MUyH^z)^UN2!(QsygX|=WMGTA4Q/S]+gTw!cRrFSa)Eo{D=rhTQMfg{VP2^_R
::_~ZvbiFNRc_H4/Jg}[JjhX4QocLp=SJJ[!Ai&JCR1B,[YNsGse,62$E004^!]o#h5[).l_0R#X4GyfGNPyi5&Qw+ns2#HPLiB/S=RpfW+as+fUcj1dw{0sOqzzgt/
::Lj/LM[Qq0Si$)0{1p[#8i]z.G52VWh3tij/$qP/3J6.sRUHpq)=!s42a,uZeJIiv3iFNQ)|8+&ibqN1/{88xI0ssJsUEo{+bqtA3-?6.&bqI]{|8@-J=)-,_06W&l
::BzFWm!,U,rUFeHN^={EaY5.X|UEB|a8vp;QIYr;Pg&SV&|NnIi|8+reb@{f{j{,PyUtU|u!UzCWi|kQRQ/S]+R#&Bl{Eb8ZiADG_^A~#5z/rQ.UHpqf1dT}l+8-(A
::|1;v/Bv1elJ3|O}]9&5cLj/XT|I^CK{|kEvJAL@a,o$2VGt!I4^w$R$i~2Lti]li!JHc_XjYI!C^/sv{#xv6Q]Iu-z!duD01OQcw@2A;lxc~qFfQv/4JJELyas.9I
::a|eUK029(c1&v1S8^_|D2mn=!@1StWjYa5I|1|)X01#0B6$nrO5RFCnGxktcjYa5C{}nh/01$=14~Xyp002{0i#;G001&Bu{8nlNTT[U|Xj~cq5dU[f|JQ}UcnUdn
::{124{EI;GdR#S~d{7^S9cp3l[|8[BP,S]50004!-cpi(I{5!]W1dGP(4TJa)a|LP.SnC51x4?qjC^n&ZY5.G?Mf]}xXs8-h5dU[f|JQ}UcpHmF{5!]X4TJa)a|LP.
::Uh4x7xA10)H7Y/=5NZHw2#rPXP.-Wk(?8?_|8[BP,I!$Oz;4.=$]SV_]l}Pn{#bV+IYsby8[Le?5fM2?]mhw#1vy3Va|DaX?jsNO[QY3KqaZ,45Qz/WKmZW_b[=}k
::Ku_b?Gr)TS!(Qsy54J+Nf$WR_0034{54J+Mi}6qown7b154J+KP!G033s4WXLJCk1wn7O|54J+GP!G3422c/TLIhC&GL40F|NsAJ=o$bJYCur_6=YBV5QD[xjYahT
::H2]]X5K#XW2v7hJjYa&WY5;Kz]icm5I8Xo)g}_^uY5/0}R&!u^g?@V^|4{#R{Qng]Pyi7Bb]MF+{}n)_01&7DUdhA&B~)xV5dS4@Pyi7BC2(vx5dS4~Pyi7BC3sK)
::5dS4aPyi7BC45i.5C8xG000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::2m$~A0&iaJ5C8xG0000000000000000000000000b]x{jmINF-cmQB00RR916aW@g01yBW7yuan7!Uvu00000p+vpv6aW@g01yBW8~^~vG!Os~00000MKb]p6aW@g
::01yBW4ge1TR1g3V00000xibI|6aW@g01yBW4ge1TWDo!l0000095etB6aW@g01yBW7yuanbPxa#00000ax@&C6aW@g01yBW6aW;fkPrY600000!ZZL76aW@g01yBW
::5(#nbs1N_U00000A2k3F6aW@g01yBW4ge1Tybu5o00000!!.a9EC2uiph5r/EFAz4000000000000000000000000000000000000000000000q!s_W3jhEB3jhEB
::lokLG3/-NC3/-NCgcbl04FCWD4FCWDbQS/,4gdfE4gdfEWEKDr4,(oF4,(oFR2Bdb0ssI22LJ#7WEB7q0ssI22LJ#7R22Xa0ssI22LJ#7L=]xK0ssI22LJ#7L?2&L
::0ssI22LJ#7G!-040ssI22LJ#7BozP;0ssI22LJ#76cqpv0ssI22LJ#7Bo-V=0ssI22LJ#71Qh[f0ssI22LJ#71Qq}g0{{R32LJ#7]b_OP0{{R32LJ#7;P_uA1ONa4
::2LJ#7#1#M#1ONa42LJ#7v=sml1ONa42LJ#7;P.o91ONa42LJ#7q!j=V1poj52LJ#7+D!?]1poj52LJ#7#1sG!1][s62LJ#7lobFF2LJ#72LJ#7gcSe~2LJ#72LJ#7
::v=jgk2LJ#72LJ#7+D.{^2mk/82mk/8bQJ(+3IG5A3IG5AG!^652?;{92?;{96czvw2?;{92?;{9]c4UQ2?;{92?;{900000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::000000000000000000000000000000Fi_,iQc)Z]Y,7FJgi!zhmQerzq+_9?xKRKA)op~a=urRw^E7+/1X2J1AW{GTG,SQnN?Ts.Tv7l4dQt!Ym{I[$I8y+stWp2~
::wo)89!cqVL&u+aV+=~ff/8Fkp?QVpz^EG=;1XBP24pRUC7,hZMB2xeWEK?jgL{k6(00000tWW?|0000000000qEY|=08jt_0000000000000000000000000Fi_,i
::Qc)Z]Y,7FJgi!zhmQerzq+_9?xKRKA)op~a=urRw^E7+/1X2J1AW{GTG,SQnN?Ts.Tv7l4dQt!Ym{I[$I8y+stWp2~wo)89!cqVL&u+aV+=~ff/8Fkp?QVpz^EG=;
::1XBP24pRUC7,hZMB2xeWEK?jgL{k6(00000ZvaeWaztr!VPb4$RA^Q#VPr#LY/13JbaO]/azt!w007JZPIORmZ,,m2bXI9{bai2DO=WFwa)Ms&Sp.saY+NiubX9I@
::V{c@.Q,@4]Zf5_hdH^shaz|x!L~LwGVQyq?WdMo,Ok{FQZ))FaY.|7kOaxMNY+NiubU|+(X/XA]X?Ml#fdEWoaz|x!P/zf$Wn]_7WkF;Qa&FRK006=TQgm!oX?Dax
::Z(Yb,WkzXbY.Do+Ljq28Q+P5Tc4cmK001}zQgm!mVQyq]ZAEwh]Z_zEQFUc;c~E6@W]ZzBVQyn)LvM9&bY,e@0Rm2RQFUc;c~g0FbY,Q-X?DZyz6DZrY,cA(WkzXb
::Y.Dp)Z(Yb,WdPR#Qgm!VY/131VRU6kWnpjtjQ~t!a!-t(Zb[xnXJtldY.LYybZKvHb4z7/005Q&Ok{FVb!BpSNo_@gWkzXiWlLpwPjGZ/Z,Bkp0s(5RLu^wzWdLq/
::WNd6MWNd5z5CC([a${|9000XBUw313ZfRp}Z~zVfZDnn3Z-2w?4FGLrZDVkG000jFZDnn9Wpn[l5()B(b8Ka9000UAUw313X=810000pHb9ZoZX?N38UvmHe3/=Cq
::ZDVb400000Utw&+WNCH+0svoOY/0|HYyboRUtw&+b7,V/1]{1Sb!=?8X@6er2LNATb!=?8c5.b12moJUb!=?MWo.Ze00000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000001yBGGynhq6fqnzBr+JS;vaB|06c{}tvu5^]E[#;Lp[tPYdw.ZxIM}}.aYX?13uV4
::03ZMW5CH&H{4+hK5i}h&J2XTzP(8jObToZ5f/5aYo/0U4tTeVX-&+Ah?NNB/2{jTm7d0U[J~d4?Sv6&feKm$Pk~N;.tu@ka!ZqYI@KSx|4mKAyBQ{nxb~dLrt2fU#
::],01KCO9=XTR3busW]u[lR3/e.Z}j^0y-#jC]|7ZHab8$Svq1mXF72[f/x,jkUE-=!8,!1(pO#U@mG86{5l3ZZ9A^!wL85#!aK_5,E{n({5uFd6-9;AFg!dwTRdbu
::a6E.Pjy#$@tvt3o!aV6b[/v[L6FnR~C^OzrMm;tJZ9RfLhdq&!sXekiw?_l;/63R,@mhTDS3YV!sy]vH^C6Rt#Xt5x05AXm_~Uy|W.x{[Brz(6=_pS{wKB#s(obmP
::@lSl@{W1hI7Bd^)CNo)xVl!$pcr$[Bk29Gwq&,&W([;jM=QH{;2{a+zG(DOjOEge4RWw?OVKjI(kTjSypfsm6wlvl?.!$kn[HGE43N;G-RyAWaYBiuWsWr;r^cZ_E
::CN[@!W/Sg$bvBDOr8clO+i)Jy6E_h4M?k;NdpEf](o|gN.8bhq@?F[~0ysZ7SU8wCsyMGWx/X7P!#P)vNIS-m&{$(Z@mPWE5Ih^^NIX@Mc|3@b#XQSA+I20TEj?9s
::Ks__BYdv/7fjx(kjXje-nLVF9r9A+u06-i$d/kCdG&!3cL[.P-R4_mHWH4-nbTE7{gfNUSlrWqyq&f?7v[pCd#4yY.+G,vI;S]^o]f34]ATca4I59,qP&(IFXfbp#
::fH90Qm[&X=ura)b$T8G0/4$nm^&Q[B5HcJxC]9rMKr(1-STbZXa58+{h&&Hi00000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
:embdbin:
::O;Iru0{{R31ONa4|Nj60xBvhE00000KmY($0000000000000000000000000000000000000000$N(HU4j/M?0JI6sA.Dld&]^51X?&ZOa(KpHVQnB|VQy}3bRc47
::AaZqXAZczOL{C#7ZEs{{E+5L|Bme,a00000v~j&1[DS3Q[DS3Q[DS3Q;a]Va]AOUT[DS6R?k!hLXJXr$^z=?YXJXKr[etCRQfXso[DS3Q00000000000000000000
::P)=U$WQGI+#,nM]0000000000[Bktp3jz+t06qW!01f~E000001RekY01yBG0001h0RR9101yBG00IC21][s6000001][s6000000Du4h00aO4P(feq0RUhD000mG
::0000001yBG00000000mG0000001yBG00000000005C8xG0000000000,kAwvC/$Ke000000000001yBGumJ!700000000000B_]Rm/e9)s2u;RH~/^u0000000000
::00000000000000000000000000000000000000000AK)B,Z=@k000000000000000000000000000000E^7vhbN~PVAUyy801yBG06qW!00aO4000000000000000
::AOHYhE[WYJVE^OCAO.,c08jt_00sa6073u(000000000000000KmY,1E[[;8bYTDhwgUhF0AK)B00aO407w7/000000000000000KmY)hE]=jTZ){&em/e9)0B_]R
::00IC2089V@000000000000000KmY)j00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::0000000000000000000000000000001RekYP#ypP[JavxP#ypP{~rJV^)}i)SReoZ;sbk66iWaA=pX;9+,}D^AWHxM,dqV{+FuD_KuZ7s+FuD_8Yln(SW5r]yeI$w
::jw&2EY+b$DkSYKG2rK{qgi8Pb2rK{qLo5IQluG~rL[WRRi7WsBluG~r,en15zAXR&oJ#.zz&2j)FE0Q9oJ#.zFfRZA[h;=Xv_YX0]e-GaurL4sKuZ7surL4sw=n;!
::$V(hKxG@|#WHSH(/7b4iWHSH(0X6]t]h,E$1U3Kw/x-(P2uuI~;Td~RH#YzP6ifgBI5z-Q/Wz,QBuoGR;TwBTJURdXJWK!pJURdX8a+62TucA}96bO4lRW@ccuW8Q
::lsy0dh(}+SgiHVch(}+S6h8m}m_nfw6h8m}i9Y}U[Javxj6VPX],/arq+Y$-^(+#uVn6[[[JavxWIzA^(p.eGq+Y$-(^DnHw@O~]tV{p]xIq8^aY6t9z+S!DbV2|C
::UPAx@(_baTU^$[]8$;vA;V,kn97F(B&t!zL^+Gu,&t!zLdr1HQKuZ7sd_SQR=t=-p5KRC8asY4uV,qjhbO1B}E(yZzYyfNk00000QgCBabaH8KXF^RiWNB^]LvL-x
::Z,yf=0000000000QgCBJX?Md_Zf8bvZ,5a^a&pa7LTPSfX?Mm&00000QgCBabaH8KXGU]mWmf;IQgCBJX?Md_Zf8bvWn}/WQgCBIb9ruKNp5L$X;=-?dSysqZe)m^
::0000000000QgCBIb9ruKLvL-xY.Mz1Lt$+e00000PGoXHb9ruKLu^efZgfLoY.|7k00000PGoXJY.wd~bVFfmY&&};PGoX6G)mHDZev4iX=QG7Lt$+e00000PGoXJ
::Y.wd~bVFfmY&?4=Zvb.uZ~$.sZvbKdY5/Qp0000000000a{zDvZ~$+rVgPCYa{vGUQvh&PZ~#RBcmQ-(LjZ38Z2)UIVgPCY000000000000000000005C9SY00000
::uo3_)0RR914ge4U00000$Pxg60RR917yudo00000,b+GM0RR911wDfY_Q_A4?t3d4Z16Y7;nPkf-lX~/B5j4{$ulF4rNb0SK_h3fs64Liicu?ILt-Sp=Aj)Sr/XZE
::[oPh|dpc/kI9Omm.uZm;d32^rYfqyG5OvGFC[rgk^SDyLkDzhUo9vx,i;wr,qqO}@RbVPQ-Q3_u0o8OPicBKvDfr+]e3;o{rdY0UD={Spp@wGKh=tfp]c]kNQbmKO
::#oc+aWT1ZP?@NkA7(wb[N^^~{A@=URMNRQLsc2W7sY,eW/sHY~o6AN71=va;$)-3YE1mz.uvWR)wT|=l)vkiv_3wR0Nm{rs{M1X@o-8VeXD.TPE^8BC)x5q(1u+@,
::5gsc|KWbS4@aE.365zvo1ODhXJe09F)O&J_W8TR{X([nURkV/qgz7=)tzBIj#C@2ek/({W6)g;6[5m_b($U&5UVOO-OJ5Yt.!hc(5Qo9qPWyP?1,B{cpkiK|u/rgY
::{vPL@_@_yaQVD9-FgB$+zd+m(f&Dh;eB)KSn=k+|G?$^=#NO&4RC|/&rotmV@o5?nLi+o^2ri,!DA]?kc3YxJZHv),a_]UShG?_/+TCU[U1heCY/Z^W{q4EhUKK_H
::r/VM2kl3pLjJ)qd^vBawxU+qD([3L0&0CYR!LPjo0TYUAI,}1UPiNffm.5ff[U.T0maKFl=dCq_/_uk|9ChDrNAVhQ9Vx|$Z@|F(su/c.{8m0o#@pBpn&ltsc-Fb$
::AKj=khzG|pu[VqjCxGl;U{Qam8MR6cE#.QjlgXU#px_[At}6Ag$m^d2gHxGd7b]sQx^8zl/b|0ORUr)0V|/ge[[sF!Fac,P{[1H]&7V##_dLTtt;;8goTPHVxBZhQ
::Hb3{wG]OS7ao8~x1ji&87@uT]2NHnd?nE~x34;(e8,W/lQajeODdR7MQ^&qJApEggYRkSkN=#VK)C[1ILrpV;Mfn1MP(}WgQKLYQlASp9ytdjQ5dZVi&@uOlUzbD|
::#HW5eWL-6]V1ZBEA}WxGM))(2.d-pa/4)T2Nd^cb!qco_k)K0m=g2p0jnz+6Y,zH[WqPg&x^Bin9Hz9!=.qT5OTCMVa6YwWNCWl_VKrB|hQS[4/rN(lY1xjHn/wVh
::(Q(PijG?7Qzve;{L76QNuvEJi2m$~A4rTxV5C8xG(3;_rDzaV6RsYEEgJi]T00000WgmZ#(8/p;mAp(!0K9Hk;YPj(k_Km4y!gQ#Uj8Xr_?{!c?hO9=nX6{Xmg&7N
::X~3UueI?-8w5N3i6w[a|+9-S;1PqBlhd]6$I8$0?am!^fj7FnMqc^W($;]wti6-XjsHxXNla0[gpCB1nw+Ka~M$N!Lux,abSEM(Tugt~|4,#xCod_p4cw6_F]Zg./
::#mnnNgY,6;PG,3ohco4mhcHJ)iG}x3G9g/Y;W}J_Z@{rPpON.K.Ic6JBfp@~^0V!ak=fN-]_wEeC-A.1#Azr-nbIL]4-dU17id?Ib$en&o+aQJEuNg0Syr)T1LpBe
::oFDM+0k{~5z~i5m@4ue;pCv,z1?Un|Hr9O7Vj1Z~i&&!E!an;jy0P5Lli.0(q-;|wQjz_nPZQ!/5snv4oU+MyoD~sBSQEwJKK=tjq[p_)Ajxx1fX9!]1uR.gmk[=o
::|H&Z_kZ5FW1~wW.hO1eNxJu4~ShGMpNLjB&k~?q;AIyGvJA1jiq?Ly]mluipy-XvS8B,SV_uj?qg2];|tyAb$--.?tu|qvgqYN,_ogkIQ77Nve2uQC&xI42rCqo#r
::v7UWlHt(K[hTx_J/Cq)F8}^w[3o^$Nfl9Y+EBgF_zwxH#K&K+ts.MSuq7_^-M+^LkB_(u|gW;ls?,~f470(SwiK)4OuSW89#y19Im/e9)K$,1_KeTKY000000052Y
::l.;bzr~m+~K;-IPzml2~000000Kl)uBht$O(Hw[e09FqP|F3Zi000000Kfw}FpA9q(Hw[fz#&6Pzk7+i000000KkpX&gW9H(Hw[gzz{@oziOr,000000Kn$YQ0(nG
::(Hw_iz@at_zr-Y400000005_l/#t&I.4O&]Fm)U_FQPF400000004S9_CHckHxmU1AWi[PAA2zY000000KjyvY@s/rHWLL000000Ps0EJ000000KjyvY@s/r(Hw_i
::I6]A_SCA^J00000005_l/#t&I5+TCj00000zv1Kn00000005_l/#t&IU/qLGz;=Bee_[{=0000006;8;uHni7(Hw[hfR6GF|35-x000000D#]8L&q!b(Hw[iKo;.e
::zsOq~0000006?2o.22c0(Hw_jfLgLAe~05J00000003EunN!pO(Hw}lz{Ch5zwtRE00000007C#jjGoH0oMQkau+yq0oMQku]j,aG8F(.[FM]K0T}=QfF&F_91Z{g
::IXD0S91Z{gV@^V}91Z{gd_|!X91Z{g]ko15Lvnd=bU|Zrb!l?CLvL;$Wq5Q~00000K?$PmRscZ(Pyk5+GXOFG0000000000Lvnd=bV-S-Z,p_@WqAMqLvnd=bW?$@
::OJ#XbVRB)[0000000000Lvnd=bVY7sa)Qrc00000Lvnd=bVOxybaHQbOJ#WgLvnd=bW(w)Wnpt=LvL;$Wq5P|00000Lvnd=bVOxia)Qrc00000Lvnd=bVG7wVRU6k
::VRL8zLvnd=bVy.yXhdOjVE^OCLvnd=bVp[$NMUnmP-[XmZ2$lO00000Lvnd=bVOxybaHQbNMUnm0000000000Lvnd=bW?$@NMUnmP-[XmZ2$lO00000Lvnd=bVp[w
::QekdnZ,2eoZUAEdVE|)QZUA2ZX#j8lUjTFfV,qdf0000000000B?.~)IshdAa{yZaB?.~)T?t;800000F#s|EHvldGFaRz9FaRz9F#rGn00000M_d)Vd2[7SZA4{e
::VRdYDOhZXT00000YXD]casX}sWdLjdGXOFGE(yZzYyfNk0000000000B?,r0H2_&0EdV6|FaR|GbpR~[B?,r0GXQk}EdV6|FaS0HbpR~[B?,r0G5~b|EdV6|bpR~[
::B?/5,E(wn9FaR)BFaRw8B?,r0GXP_(B?,r0Gyr4)0000000000O8_v)QvhE8MF4F8bpUJtVE}XhX#j5kZU6uP00000O8_v)QvhE8K?&X^bO31pb]u_jbO31pZvbup
::NdRsDbO2=lasYM!VE}9Z00000O8_v)QvhE8QUGNDZUAKfcK~4kYye3BZUA&uWdL#jb]u_jYybcNO8_v)QvhE8NB~y=NdQCu0000000000O8_v)QvhE8Pyk5+L/zm]
::B?,r0H~@$^cmOQ_B?,r0GyrG.cmOQ_B?,r0GyrG.cmOQ_B?,r0G5}},cmO2/FaR;DXaINsEdV6|FaR;DXaINsB?,r0G5}},cmO2/FaR;DXaINsB?,r0G5}},cmO2/
::FaR;DXaINsB?,r0G5}},cmMzZ00000#,nM]000004gdfE00000000000000000000#,nM]000005C8xGBme,a/2i)}/1K_.AOHXWOY;sH)omY$UaA8=DcdrPusO&_
::^P;w80p?C8#,nM]KvPJA??x?t,n{c/bS/DW19dJ$icBOpM2$iRNR1WgGXMZcjST_a008K;0ssI=jRn?.002md1M3F=6}v$I0Js4F002mX#2{P4NQ=ZsiC7RwiD)!|
::iEtoDiAV[nK~zCiK~^OmNQ3N9NQ@4FjV,[s|Nlsf!9+-YMF2?P$ViJ.2uacC{}l#8004!-4}~iM006j6S]xlMIR-1f9RUCUNMlAkfjlr!MF0Q~g$w}z0E@6_kN]Mx
::h0-g&{Qv,}i/OUo0001m$q$6@0000/jT{zCjU+[{,Z=@kgL[1B4}{kM004_75JZbJNCW/&1NP~!0000@jXg?-002R~3IG5ANs9-ag}_)LNIM7X8&c_?NQJ;7cS)y]
::^)-XT0!RbM{}o;A004!-bz)[3L@kdsiF^nLLAgKx002mdL@j@gjYK3kNXJAZI0yg$07#8gBtS[m$]ZWqpg/fsNQ)zdjYK3kM2k!$Fi4G5BtS_t2S|g.|4EBS2uO)s
::NR31!Fi43-Bq0A4a6kY6NrT5QNrUYH4~1s_|NlsX#|TM.#t2A,KL8JfNB{r.NQ1+]NjuyRf,=3@|45AmWb];3NQ1=@NIU3Ai.aVA0000Fg_EEX|4fZsBv46;TR2IJ
::Yd}ehb4W?x;3V3RUO_;!TwlY@NR1UKFaQ8SxC#IO08NX@Nzv(_jRZ-Z1Hec,Ou{I06iAH}!bpS2|4A#sNQ1(KOas8{15Jy/O[-XJ(_6C1&rO7}NITGU21q-fD02fy
::i^l1mLLkG/NQqn|2v;mh?^AA1_bdk?NR3P-ApaGeKL7wsjZ7q9Gtx|rL@j@H$]R9RKL7wT-l9b-988T(Bw#brOpQb(AT!DT6[Nbf0ENJL1T);vF.VKoNR3P-F#i?C
::KL7woi&cY7OpQz=Xhk!@NR3n]F#i={KL7xQz)[~CTqI~ni]E8ZTqICPgTz2z!^3UgNQqn|2uO+-BoIi6bR.y8K~zCiK~^OmNP-B.0000/i]4(@Fa_hsNGriXirGnv
::=tzt5NQ@PNJH!u-tN/K2NxiG8s/a80swzl[@GKJ00RR9,ipxliOe8Q(i]fQc,XW7[002yjbR/.Pi^J_nd@YYQi]WL6_2s|W&Sh4vNCW9ei^S@a(Pey]NsG[-i]51N
::(q=}f14+a)NWthoOas74i]WLy=}5u+14xU+NWthoNCVJGJJ3vv#eEP.i_Pht(,)P+|Nlvg{^74&i_Get(gdrp|Nlvg^ehJzNGr!lJH!u.UjP69NIU+ygxCQ907#8r
::Fi1Pz4}{AB002mf?qsl^4.iuzNQ=V{5E~B=V.RK!Z]C8|bJz&GNDqF)4.s4?Fb[&ABtQ=lR3uPHJ69-VgF]uT07#3=OpC=xE6qrY,GPlGF#i?0J]&m[5JV(}4.iZw
::Ko1cI4.iBoP!ADL5DyVYAn8g0002mf@n#TqNGtD1^w7uJ/z^~$14+a.NWthoOatIei~LE]|4oJeeojw~G=E9K;]+Uw^f3s;)1=^lAVCih21+;VOpQz=U^lQMP7qCt
::#7T@S4.rHpU_UJ4K[Si{AP,5tBybNAgd~6f008I.1ONa{h5vuONrU^X4~,LY002mX#@VQN#Yp$;NQ?[B!TJM8EAL2+!brjBKS&[6NQ3$Ud|D3?21$#?OpQz=U=I,Z
::5J.#9h-HHf4.rHpU=I,RAV~M]4.teUfB,mh4.f|r5lkd.NWuC84.p1Ti]2~OPY^AL=s!#Y!ZXqj5l#?f5Jw/n5k@]Bd/;UgNQ3S$NQ3=g4~=/L|NrYvNR3U;NCVGE
::J5)@Db0ZHB0S]&X4.iQZ4.rTZ4.i2h4.r5hNQ=QpgXu6xJ5eZd.478M4.sGx4.rrxNQ1,LNQ1?NNITvSg,,QL|4faXB#/0A07/8mFiDGRI7y3hKuL@]L0?]$L0v(y
::U(GAINQ=ZsiC73oiD)c=iEtQ5iFhDER!D?FAV_b+OpEbJi||d0$V[BBNw}.3s/a80swzo~z+AP&NR4jLNsGWp!TJM8i^J_n&1A5ENx|tqNdwVH)fUY[98yU4=}5uj
::1W1d|NWuC8NCV$Ui]533=s!pU(_5,(D1KH,jY0uP!T1AAjSN6cjTAsh!Qli-!Qur@jT|{ii^=Yw1Ul&U0{{R?i^1.o1Q|(y$w.US=#K,c07#3~NQ=!$E73^S&1n#J
::=z9YI08ER,NsG)t15As?NQ=|xXafKMNR2};LAo#i008R_NQ-A.NI6URate0{Nh{h(JNt6&NIS.J6iACpC_dU.C~]vS2uTCNNIS.I[JNfuNQ=|!LrjZ7C_?s-C~_J/
::JV.mibSw_L0S]!m4.i.o4.jA=4.o)l5J@aZ5l9dZ5J4ah5kMdh5fKj(Sr88qVIWBZ!bt;lNITAQxJWzKc34b}LeNZ$ODIh^NGNhVcST7n-DSXubTUB@5C9Jl6Autu
::5DySzAP,4)K[Si}5DyVc5J3-RKp-nhLm+v95fDKS5m,pG4.sG?Nh{J#E5b.S_f{(GJJ+q^NIT9CgbD!w07yH,bSw_L01pro4.i_r4.jJ[4.o@o5J)UY5lavc5I_Ug
::5knvk5fBd&SP&~pU@2|=0Z9YG4.gX&5L,xr5Mv-[5d#kpNe~YaOArqbK^CwiLm+{5-7A(C4.r_q4.sJ@=^vpI|44)v(_3MZb!bT|-DJRobSw_L0}l_q4.i[q4.jG@
::4.o;n5K9mb5lRpb5JMmj5kepj5fcv+TM!QsV/~O[6G;z=4.i_r4.jJ[4.fzk5d#kpNDvPZOArqbKp-nhLm(@k5DyVp5DyVxAnCgP|NjpV0uK.o4.i=p4.jD?4.o-m
::5K0ga5lIja5JDgi5kVji5fTp)S_ZHrVj$]/{{R0.i+;t~NQ.nNKuC,xBuGh#9!QH]Brr,dMhHoZ#z=$2AVFTkNQ=ZsiC73oiD)c=iEtQ5iFhDMjadIli},/;_&H^_
::NQ@68QcR5m5=+H~Ku819NsC0&SV=jA)ue?607x6aNR3beNdwSJjRadrjSNyq1Ib7Oz)EfX1j;1V5d]|P4.ibs5J3-SOu_UB4.iDkAVCiiM8Y6Ui_qy#)|o3NJ3#iw
::0d-M;jSPQC1Ib8[1X4&@z+X!^|3MEBM9R=X4.o{yK[Sj2&HTl{5lq4mK[Si^&J4xC5k$fuNjvj?N=b_BC_pY=|47mKNR1RpNCVPIIYiPZb1]{=5CqaojRadvi_hs6
::z)EfY1j0cN5KPh#K[Sm3!Vp0Z5Jb_.K[Sl_!XQXH[qD(H4.f&Ji)DiKON|6uNQ.PF5J(]SNQ.nN7+XnJBp]W#5d]|P4.iQZK[Sm3!Vp0Z5J4b84.rJdAj8beNQ,+!
::NIO9(b1+AO0S]!o4.i[q4.jG@4.o;n5J@aZ5lRpb5J4ah5kepj5fKj(Sr88qVIWA0LMTZ(Kqzxv4.f+D4.gPR4.i.oK[SjMAVCii01psK5J3-SNDvPYLLfm85kMdh
::5fTp)S_ZHrVjxJ1K_2N$LMU@~4.f$l5dseoNe~YaN+QhaK^CwiLLd)j5f2er5DyVyAj8beNQ=ZsiC73oiD)c=iEtQ5iFhDES4e~GP+LJd[BmDWJ/)9@|44(i=m1HJ
::[JNdVz$]d(NQ?&7i}][}z/zZ#jTMgZ|Nlvg(,+C~|Nlsd1+nSc071DJ0002TL@j?p008hsGr(lVOe9D&),Mwfzz?Af0000/iv[lx002mZz/zZ#jTL)F|Nlvg(,&#G
::|Nlsd1!F7#07Wy(NR3P-KuC#9Bq(IY6,ukw|455WBuIl~Bp@7qjX[m$)1pMcgoXeB0J{MI002mZ|8zA.iv=z$002mZz/zZ#jTJKS|Nlvg(,.xD|Nlsd1rsa+07#2W
::BtS)o$xMsKi^_zmg}_+2NQ)slEC2vVg}_-dNR1U0[(Erxi^hqN^W&D#i3R2?002mfOe8=?Gs)w9BrqHS002ab!$]sABq&e|NQrDDApg,Xz/yyM-enK{BtS[o$#f}5
::iv^kT002mZz/zo+jTO[H|Nlvg(,(=l|NlsfOe8=]i3OG_0095cNQ-D+NJxdrbSp[U1(b?H07!-vbsI?H6|eCB|4EC^==b(i|455WBuGe!1#2q-05j76(_67HBxsAn
::NQ-z~U_UH}BydQJd@a{CgTzolU(GAINQ=ZsiC7RwiD)!|iAV[nNQ3M!NP}P@07#83_0xM!NP}Pq07/A3NQ)vWDgXdTi~2}~z/zZ#jTOT0|Nlsf(FF(l|Nlsd1=A_3
::071DJ0002TL@j?p008hsGr(lVOe9z{),Mwfzz?9t0000/iv^YO002mZz/zZ#jTNr&|Nlsf(FDV$|Nlsd1)zxS07Wy(NR3P-KuC#9Bq(IY6?IDN|455WBv]xFBp@7q
::jX[m$)1pNsLr9ASkSYKGNQJ;47D$a1#P9$ANQ=$r]z{G#NQnh@DgXdTi&cXyMKj4riCiQoGtx-jd@X.2jadOii]KoWg}_-JGuuduOe8=]g~[a(NQ)tADgXdTg}_-i
::NR1Uu[BjZui^Pez]#A_zi&cXyNQngzDgXfg(_66-Bv@p=$#g47iv;QM002mZz/zo+jTI{I|Nlsf(FFIU|NlsfOe9!Hi3R2;001.6|IkQ^Y$Q/N!$]x;BuGeu#4umO
::KvPJA?[Y}-,-_4gNR3P-AVIhg000306[xSY05j5!1d{,4gWwN}L;As-14#eFgZdDB?PUmcFk8dSKvPJA?[Y}-,-_4gNR3P-AVIhg000306.P7x05j5y&8SB./RFA|
::!Qlc)E5U={4~j$tAczA;|HFg(5PaZBgTydf!^3UgOpDP+Gs#Db1d2h6z_]JTB?[2e0c-43L5sj8e}8}f1Hd!TL5sk^zz~bdK{Lof!N3T@$p|yZK{LoR&0r0]ib@/$
::NrU-We7wWVNQqn|2uO+-BoIi6bR.y8K~^OmNQ3MkOpP_9=Kue^1ONa4Oe]t=+kurkL5l;qF#$,e&}9gl0d@^?dJv2EOpQH)=?Pvni^1uh,AKP;LW&[9hyh6d!AOJZ
::0d@-3i^7TA{{R0.i^7Rm|Ns9;jTAOWi^J_n-d^,38bL7wOat9WgXsZv[JIvqNP-(K0001dvPg[|54Hh9iUc[^0Z9MBNQ3VIb@!+u&jkao|Nlsh4gco;|456~NGr?W
::1Q{]{NCVwSgX#fw[;[wZBtS[uY$QlXi,zJVNQ1/6L0?]$U(GAINQqn|2uO+-BoIi6bR.y8K~^OmNQ3MkNR1V.=Kue^1ONa4NGs7qi]WKb,]2}UF#$,e&}Imk0d@}f
::kN]MxOpP^];]TUoEAvc]J(+&9|BZKmNQ=wpDE|NdNQ=uzi_Eae0YZudIEVpA|G_Lu?H(4_NQ=u(jRZbOi_7Ak1PCz!NCVACgX#fw[kKM-MvDxaL5sn_=m#YM0RaI.
::YtS1)i[^y.e}Df2z)h09L5sq{=m#YM0RaJP$Qwb6!X;xyfByr)Gsug;!RQAi0RaI4L~Fnsi[^y.e}Df2z)g~]L5t8tGr?VK)LsyAi][SW!9g@1!N3r~$p|yZ!NLeL
::&0V/8K{Luii42=b|HDi?.4Bd+|Ns9/EB/7[{|}EX|Ns9/i^7R${r~@-i]~tT0YZudIEVpA|G_Lu@g4e~NQ=!viv&-;14skiNQ3DCb[51xTqHn9i+;uFNQ.nNP+LKs
::AVFV2USGq]NQqn|2uO+]BoIi6d@Xk^R!D?FAVG[(NQ?Hw(_pc,^tHp;_GevQ|H6$l0+hS;0RRAY1T);vazu/5NR12w{}nnh002mZ|8yEii^1tW_@?[G004]w4?18q
::jY$MZgX#fw[kooy=sW&Y|BZg|h5vLfNR3kvLAV3}002mf&1A5Piv$ZX0Z5HW1WAMF0d@^6i][og-UN[X|Nn#U5OvW)i8i;a0000/i)DiyNQ.nNI7o|pBtS[m#2_Ul
::!^3UgNQ=ZsiC73oiD)c=iEtQ5iFhDER!D?FAW35wO]fhIi}Lov4|D_G!0UEMje77&1N.Y9NQ=-tkp2JvGs&lYkMJ=Bk4XQ+NQ3zVeDFwv@-{2k{()F.i^Yk6{r~[i
::_w);BOpQkWOasF~fH+9$C^xXl2s6@|i]-w^bQM7lwg5]0$UDJx3je}E|H)y#$$#rR)RBhV,|.4!002RW9!QH]Brr(eY$P~Hi,zJFNQ.;VNJxXkAVFTkNQqn|2uO+]
::BoIi6d@Xk^R6$ljS4e~GAVIza0000/i_qqt#z.s5iv$rd0!ahbNrUJCb[EJ(B}3x?|456=NsH7@i]fPR)~ATOF#$/f(Pjvl0d@|0jd&f#cMnXBJ(WZ3|456=NQ?4F
::wgEzl1UQHRNdLh|gX#fw@nsNv=-pZD|456==ui9q|BKg5i_(8C1WAj|NGsDx1Jpu_1R6mx15E@oNQ3DCb[2bkgZ~S1ut;x_54Hh9iUc[^0Z9MBNQ3SHb@!+u&jkys
::|Nlsh4Oim.|4ED1NGr?W1Q{]{Ndw-UgX#fw[;[wZBtS[ubR;Yfi-m)dNQ1/6L0?]$L0rSkNQqn|2uO+]BoI|sL03qN?^~(_KuCjS^yA0eJ;sU?|44(n[Bm4R[JNdV
::$Rhv&i_vIUI3NH307#4ZNQJ;47D$a1km(#aNsG^uQ11W#NQniZBLDzFxflQd0LMfmAOHXW[I]DgNR3VSGt(Rig}[Jlxc~qFNQ)urBLDzMg}_-dNR1Wk=?Pvoi^ho{
::@,IQti3OG/002mhP4GoC$w.MzBp]jIz)|Wt^^^?$0095cg}_+qNQ)t|BLDzMg}_-dNR1V}=?Pvoi^hrI@f@Hsi3MIG002mfP4GoC$wZ68NQ?A1)1pNsL_aJTOd|jQ
::NQJ;47D$a1nCSoiNsG^unC;_nNQnh8BLDzMi&sxMjZHX&WF#N}OpC^40ssI2|ImfNbO,Zu0002&0yEo4i&sxIg~[a&NQ)spBLDzMg}_-hNR1U@=?Pvoi^hpu@f@Hs
::i&sxIi3R2(0095cNQ-JQNQKFCDoBe3tResaNQJ;48c2/5)C7dENsG^u814W6NQ-JQNQniFA].q0),Mv&i,zJti]E8ZTqIyfgTz2VUte9rNQqn|5J.u1Bp6j!L03qN
::?^~(_a7cq@$N(#lz)|8.(/T?QNP}g.07/ASNQp+GNsHLWMIaym004{n4.rM!$3[r.fB,mw5k=[og~[ajNQ==#jX+4cjSWg10093L/4A;DNQ)vGApihOjXl!l|NlsZ
::z/zZ#jTNru|Nlvg(,,OK|Nlsd1.~Hx071DJ0002TL@j?p008hsGr(lVO~]CS|ImfN4}@Jh002mf1+m_R07!-vbrwjC6_SV(|4EC^=qv31|44}igdqR_MKj4rjZM&;
::iA,FYNR1V^8vp=Ei&rObWF#N}MU6om|ImfN4}|pq002mf1&n{~07!-vbrwjC6~E]H|4EC^=.=!A|44}iXdwUqOp8U)NR3UvNQq1(AVo9DNQ-I#x)R?+0RPa1zz?A5
::0000/iv@aG002mZz/zZ#jTOq~|Nlvg(,-.#|Nlsd1w$bK07#2Xz+X!r,hMqRL5+!b|ImfNba^aN1rH$r07!-vbrwjC6/tN^|4EC^=ws{u|44}i]dJBLNQ-ItMKj4n
::i]oWd+Bn)gz/r}Niv{8!002mZz/zZ#jTJ8D|Nlvg(,)1e|Nlsd1/.!(07#2Xz+X!zAcJHiAOK8[#;~Ik0095cg}_)Ny8!@I0P6xX-enK|(_5?JbSOxR1#2Jx07!-v
::bs9,G6]G]j|4EC^=.=x9|455X(_5~|OdtRN|IkQ^O~6Qn$#f_4iv@/R002mZz/zl)jTOe^|Nlvg(,.k||NlsfO~6Qr1xp|R0RPZPi&rN#g~[a)NQ)s#AOHYJg}_-h
::NR1UW;]TUli^hqN?i^?pi&rN#i3Rc?001.6|Ikd0TqJOd!&2&;C_pTRFiDH!L0@~8!^3UgNQqn|2v;mh?[Y}!WF$}ki_qzw1)P2D07#4ZNQJ;48c2/5WaR);NR173
::;p2NZAnO1BNQnh/9{?PBxflQd0LMfmAOHXW[I]DgNR3P-P(3m1)1pNsJxGfMtRDaXNQJ;48c2/5DCPhENR16K;p2NZ.0A=SNQniF9{?PIi&cX@OpC{h+Bn)gz/p-[
::0RR91?jE?|NQ-D+P+LQzbT3GY1uGu^07!-vbstEL6/I];|45Au=/QzY=#&OH|455WBv43-1qUAh05j76)2K+Ji)Di@NQ1/MU(GAINQqn|2uO+]BoJ3fgX|!SS^DXo
::_bdMr2uO@ZOpC--6_v]p07wJgNR5|?0000/i^1Z{Bme,a{}qNQ001lANQ.nNIE^OjNQ-z~Fi3/MAYa4GNQqn|7,$qRK~^OmNP-AS0RRAt0Z5H&2#Eqri}6Gcwn08j
::i~2-lwn/q@woy0^wn0Bc54KS}L=U!EI}f(5IuEvaIS/o;HbD=!b~K4ZBxpp5L@mELjRlJ3|NlgZOe9!Ei&u{]iBu#|M2TD]NQqn|K#6=LIEhpwFuFhh004;hBq(IY
::j3kf&004;}Bq0A4xF_SsNQ+I(8UO&DjZ7qP{}nDM002ylgd{+#001.6NR3P-aQ^u~C/$M3z;5DOgJdKq07#9LB#/0A07#3BBtQWG08ER_LAU^_0075CBp_kO0093L
::P$(QZNQ/alKmh/&i_f4aNGJdRh1-=xNsHD;i^PfO3jhE}i^QNPXea/xNR5/vfB,mhjSNKs0000&iBAMU4.iH_]Fa[Phll^G07/8RR7r_4m/e9)L5oIE1HeIvhoAre
::01vkXm=]#5L5YW@0000Fw,_zB002RWho}Gm06~jRSV4;MR7k;$AV?@yNQsOjNC5x;NR5mnNC5x;{}pa0002mfoFq]I002mV#1H_h06||tUSD2a!^3UgNQ=ZsiC73o
::iEt1|iFg=QK~zCiK~^OmNR6nt{{R0.f$Sgx002yj_b~[JNsG[(jZ_E-O]e{[82;nNNR3n]KuC?E[aVSv|Nl(5[JNk+[DEqO4^C=Si}nu[L@kc|5lkc@L4,Dfd@i7P
::.blgg1dGGL.~=o1NR3n]Fi4Bf=tuYe|44)!5J.dT0d@s}W8Q_Keh]HH!$]&xBp]tO,XUjm002mhR3tFQL?wT1NR3P-Am|eJ|NlshP4Gdu5C8xGNdwMEjZ_Es{}u8k
::002RaP7pKFL5U0_Wk_zzB_]R008C[SNQ-2dNQ-4[{}qlV002#61SCj}Rq#lQ$ViLW{}s9?002R^?/M1(NsHG?i^QNPs3rgaNQ/ed0RR9;jZ_E-LAa~{0093LkR|{C
::OpTl,AOZjYNQ=Wsi)4?Bi,q;hi-eyxi{n9GL0(/!L0nzKNQqn|AXQdZK~zCiK~^OmNR6x{p#J~=NP-BN0ssJu0ztk200jVvMSx9;=}n9HNQ@T5,-IMj00scQ5C8xG
::K|98E7EFzD1VoF,NsHD;jZK9A6&![]0P9qX+;}(]Bp~Qv_~Uw/i]hqyr~v=~NR3n]AV_f)g#Q+oB?)^KIX83{L](sQB20~ir~v=~NR3s5=pXg}|LYJ/jfJQI002mh
::RfOoN[BjZyjduirKL7v+0F6WZiG{EM0049rL](sQAWV(gumJ!7NsHF#;n/gl?kmwgg|Gnt07/A1=x]_/|45CMumAu6NR3Yv=m.4(|45CMumAu6NR3UD=&f4p|43uV
::NR3]TNCVl6+;}!X=($;!|4fZT5R2AGi^7R$^W&D#i_PN81ONa4NGsDwi^42d3Is6$NR3GZNrUJCb@}P]NQ1,LK|90_g~0#.07#1kh!g-.NR3MfOpC[yi33TC,8dg#
::Bme-Ni.kx5002ylL@l2n)nz^ziD.c$0000/^wh+&@X|j$AOHXWx)EOO07wt7!AQA!M;~Wx&~ml/1Hnj,ji?;t07!|2r~v=~=+d,;|4fU;NR3n]P+LnTBryLKoFo7M
::NR35=jZ-va(_ga@Bw$R9MTAI=Oe8SqJoo@qOpC[yjZ_FXNR3Mv{}pZ|002yj#z?7-Bp]tQO[v5/!zlk1SR@=diw8+JO^VFyiJhPU002Dz00jVa8$mn7bt-7ag_fcd
::07#8hlt^)Dgy]#J|NrY0OpS&00RR9;ja8ILjZK8/DD40Li/bWG001.6gTWL,i4SE-iv&Sw0000@W5Gy]NMJ~dNihEvtRnybO=Aa0ja7h1i]xce,#8yTBLDzFxa;G_
::07#8j6iJKC{}sj~002mfjlcl^07#9MumAu6LAa~{0093LtRnybNQ=Wri;~650ssI=fy7^}002Q?L0(/!L0n(6UBk[GKvPJ8@9c&K07/A4gZW(24?Q0+i2zB70k~HH
::008Swivm3W00aPZDFpxk_~Ru_|Nj4U1T);vFi4F=Brr]kj3nRy002mdL@j]S)ER_Z?la9kL@kdwjf]DV0000/i9{qI=wJK+|44}gGr(lJ#Lxi&09)V&NQqn|2uO+-
::BoIi6bR.y8NQ3M!NsIAGi},/3-DMD)NQKf5gopqD07#7$;lX=ONQ=w=74sqh0ENJHbm,oF004tMqZR.F0Ci[7Jw^4$004APOpD9M1(bB{009610A+yv1SK#4002mh
::9e+.807#3=NQrbLC_]q7wB7(zNQnh~7XSddfB,mhNQrzTApaFzA].qLi$o.7NQoFT!0SANB|/Ve002mhJ+7PC|45AuNdN!/=(#}b|455WBxp?F(Pa)(Bq(Ua$4HAz
::Bw$F3)[2R.Bp]tO(HvCyi)Di@NQ.PFNJxuxBv43$#4umONQqn|2v;mh??x/s-DMD}NQKFCSV+D]bW@,p,A+N(0CY.7gFW3A0001WKxIga1SK#4002li1=AJ+0Cg[&
::jTN?P0093L??(UEg}_-kjZg@kjSZ$2008J6^y7M$iv]k&002md12e$u6iAH@jQ{_t=?OpV|456|NQ=-](_671Brr(W#2{b8KvPJA??z^ZUljlV07#43bWBKv$#h3$
::NQ)p|FaQ7mNI3/v761TsHAssINR1Ue6#xML6?}i~0ENJHAv4lQjSVUl002R^00961{}oCh002mXBLFkNNQ1/6Tf;0=4V)Y}|LB6@|Nlsf&SeO7AX_ZP/LOa.NQ3Ms
::OpDPo$vF-?6aWBpLQRFzbUjFm1?Y3^07,Flbunc~iv&Sw0000/IR+Ak004C#4^Cm64Yw2k0A?$=1.}(l05iZyi4Cq5002mX#3/i]jSZ(&|NrO}/Q#-gi][oY#3+Gr
::/LOa.KvPJA??x/s,-h&dNR18c|NsB![ZbOcL5tEzi^8DeGs&U)bv_+^ToeERbTmjg4O;le0CX!zISpDB004C(Wk_zzB_]R007#1sNEHA8NI3/b6#xK84[ApIi48Ut
::001-;NQ1/6Tf[xENQ3MsGr(xX-DwblNQKFCK}dztbUZl+02BZKbTmjg1qT&X0CX!zIR,9;0049&Wk_zzB_]R007y9n]b_O9br4894dWC50Cfj5z)|9{D8opN4X6MA
::|L8v7|Nlvg)[BfUNQ1/ENdMr?NQ=ZsiBJ$si9i[kiAW$,RaRF+RzX+tgX~y@J(qFq002ylco;2GA54qVNsE67OpDn/i-2!7i,FcBi-3PRi-@C|Xh@;0bYn;_)sW+)
::IXz-(004DWNI4y06aWBqOl3&m1SK#4002li1y?XR0Ch4,i$gF.|Hw##_2BY!=qCsO08NX}O]e1zi_9$H=)qd.|455VFf.Es70)^305ibrHB5^@B$xmI07#7uO#lD@
::=y&[#|4fN|Bsffq(rFGQBrr]i!$]s2Bq(LX)n,WUiF70(|IkQ.#8]RJL0)]8U0cJ;NQ=ZsiBJ$si9i[kiAW$,RaRF+R!D?FP=h^F5(![IOpS0DNsAs#i^$[heh5s9
::,.49b5J_)}7+]^IAaq7Zg~[b5NQKgLJV.emfD.[!bu)p1iv&Sw0000/IR$;b004Cv=vxN?08NX|OpC=xi_I-J=;E9b|1.erF.)h/B!~b207#7ubpQYV=)pYf|4fN=
::Brrsa(q#[EBq(Ua!&2)MNQrbLAW4hM|IkQ.#85$AUtV2X!^3UgNQ=Zwi9i[kiAW$,RaRF+R6$ljS4e~GV1qsE5dZ+HOpSOTNsAv&i^l4ne-Wd2/z5gd5KN2NNsDh7
::O]bIRO]bghbYn;_$#h/wIX&J?004DWNI4zB5(!]oOl3&m1SK#4002li1.B9a0Ch4,i$gF.|Hw##_2BY!=!XUX08NX}O]e1zi_9$H=ok9_|455VFf.Es6+znC05ibr
::I!ud{B)MMg07#7uQ2-n_=.1r/|4fN|BtT4y&S@&MBsfir!bpj1BrrjX+QNN@C_pUUNQrzTApg+vgT!D#UqN0$Twh,YTf[xENQ=Zwi9i[kiAW$,RaRF+RzX+tgX~y@
::JpvH_002yla3D#G9!.nTL5qF}M2q[Ki,]u9i_hwwZWv9Ab|7?@NQKFCJV.em-z|i)bu)p1iv&Sw0000/IR+Ag004Cv=)hy_08NX|OpC=xi_I-J=tueg|1.erHB5^@
::B&lBQ07#7uc?n-Z=o8&k|4fN=Bsffq&SefABrr{j!bpj9Bq&|P,GY[ZiF^m=|IkQ.#8]RJL0)]8U0cJ;NQq1(2vt,7S3y+kRY6ukS4fR+B/iPb@6@2_07#44NR3T,
::Gt(PR2pj-ai$!.xjY|-gz7PNa0EtC!MT]ErjTK)h|Ns9LU?pDd^t/F0doW0iHR}xk02}{IjV0m^004vM_h(,}Guudu4o!?3NsHJ26,C-F0ENJG2uTC?Njvm+;w3ar
::|Ns9;JM@sBi-(75i,]V?EB.-{^/n6QjZJU=6[wc90E]J/KaF;,O]e^}i]Ge;h5vLANR3Tz{}pB&002abz)I@|Gt(3=i]xQa$}_eIi[]8xJHc_VL^5!Q(NI@Mi[]8x
::LXA8!LAd|^|Nl(lEPp}1KmY($NR0,Ey8r+7iEYP154M!))./5&NR3TzLJzit=-YPf07.-+|44~Ne[KZ;^euZGNzv@0JK&TqNGs4xiFN.?iDd]!gZlq;Gf0c_OpEJD
::jZXha+8;G6^DGFQZ~qm]8UO&_z+3sAckf6G[J$2mNrU}Ba[kCaZO7{mNdLk}jZOFHC;6chNQ?}Hi,;iUjTQdV|NlX}5C8xGNsG[#i]NO=^Wu;D8vp=7i}r=Ubpt^,
::{!NR.NR1W6),OTSi^iZR;Qf0~g}_-ONrV0Z4[dq[i]NEc6^@Wg|4EC^{}sj?004!-bq7g;{sRw3{!NR.NR1U~),OTSi^iZRs2TtOg}_-ONrV0b4[dq[i]NEc6-6=Z
::|4EC^{}qZF004!-bq7g;{ss@6{!NR.NR1UO)f|KRi^iZRY#IOng}_-ONrV0c4[drs$4HGekh&Z=LW{ia,cbo-MvJ_e+EEE(K_Z}EjU-)0|NsC0JNR_DNR3Tz{}rwo
::004{7?q@Dx0ZfbNM2o|V!.fBJ5J.)pZ~qmM82|u8i[.sP#WT|P^KV0wi]@;7L5slm^B-9H1w=c~b;RwSZO1dxL5slm^DGF27_gxdLX8AEiA8rviv=PN002mZz/zl)
::jTQFP|Nlsh4F}c#|L98D|Nlsd1p]NN0RPZVi[fIG7ytlBjZN2yMSuSlEExa,Op8U=NVot&men[[|Nl$^z_sBM002mhMQ=!h(i^b?MfbS?|Ns9/|IbL#@npcEclAiQ
::000000001hNQr(^Ogs5ViFF4^jRX&)jU,3D4[2NdjT8t?JHbhf1PDkEL)oYp+JTK-|8-1.i~C58Pya~M;46PBNR3Tz{}q}T004_@NIS.N@@@/SOauQ)ga1Hs,.VRV
::$LkMB|H4R)P50?e0000/i_YrH000000001hN{tjFNjuO=jSK=xjRXQn4[1F7jaBDJjZM!/iAB&;6,CwB07#8h.]WGi0RRC2NR3VBy8!@I0RI)17ytl7i(b|?xd6av
::D#.u;07MVJMd!GK.Wx-f^u[#4MbAhpMc-h=&1D^-i9{q=NcZ^jjTNf5|NlY3=[LO{$v{F608EVq$kYG.1H,|#Bv3-(gz{7x004;ZBuI.ze@&-MLJzl$?d-Vf0Et8.
::K,vNRI0FCx07QvYBrrsYOe82qi/VJ68UO&|Q$LAJBp]tOO=tfV(=(vzNP+z.0000&UqN0$T|r!5URzzmNQ=ZsiC73oiD)c=iEtQ5iFhDER!D?FAV_byNR4X$OpEa|
::-cW=2h3#}bNsCqpjf)#O002nS;46PBNsA8u6,m^E0E[s$J4O)9@[5bR2uKUqjf)#O002nS=STzpNeg}uNIU(;$1~DMi]oX!^DGA!OpC&Z)nyQLNcZ-gJHc_VOgqJO
::nn/UmBsfTm!$]x;Brr45NQ.;VNJ#hgNQ.nNKuC-iNQ1/6L0.emNQqn|AXQdDR!EENNQ3M]NQ+Jn2mk/_jZN^X6?}B;08EWVC]OPXjZN^X6~h)+0ENI0h$R6407!#m
::Bq#t)jZGj)i&lp;jTOey|NlY01ONa4MT]EniF70({}q}R004!-cngDlAOH_Q1wRG=07!#mBq#t)jZGj)i&lp;jTNra|Nljc#zcv9Bq0A4XchnfyTGUb0KN|Z004!-
::co~aDAUnf$1rN8w?jpc;a{_OTW{Cyy1poj[gJdKq08EWdAV_Z&C_gSJl-yqIMT]EniF70({}mn;004!-co?UCAUnf&3J;s5JH~PY54YH6i5/{B002mXWF#m6OpQ&2
::NQ-G;NR1VJ),OTKxC/OP0LMfmAP4{e0RI++6#xK,z;4W2i&l?{$H4yo|Nlrk$afS.x+Bi)5fKp+5lD.{NIS|&4@[C7i3L(x002mfO)^2rs1,PJGr(lTTqJl&gTz2V
::US3^p|0TK=007L+0R{p91~LLL0U!)jAY?B(AXE|nAT$vGAd)#L8sHev7Qhs60SW{F3N#7/3UUT/0Ur$jA7mN/A5;9tA2b,M9{~~o81NS06wngD5O4qh0T~Ja8FUE&
::8Dt0m8B^.V88ij}88Q{&0TT!S6LbUs4_c&X3seFC2Q(cy0T~Ja8FUW.8DtIs8B_4b88i$48Il$70Tc!R6jTZT6f^9{6jBgy0R{p922uhr0T?DZ7.R|n7,q+W7(Hg~
::7&~,^65tSU0Tl=U6@6yy6=Vkh6,L9^6,3Xv0T~Ja8FUH(8Dt3n88iq088Q{{6W|fR0Tl=U6=V$n6,LS06?;,n3~(oj0Tl=U6=W0u6,Ln7719py3~(oj0SW{F3N#1-
::3Q_7e0S]WM4_c[b4?Se;4?AjI0TT&T6ErFS69FOs4Dbrz2yh2r22cP10VWLqCUi]yCS,$hCNxR^CILhM81NS06wngD5KsUB0UrwhA2e409|24N5bzG,4A2U|2yh2r
::22cP10SN/D2@06+0x$po0Tc+T6l4kj6jTWS6f^6_6jBgy0SW{F3N!_+3Ni-80R#a61VR7-0UHMZ8=[ER72p$a5@~Qf5HJ7$0T~7W8Il#@6L1n?5l|2@0T~DY8L}1d
::6W|fR4{#1)4Nwd+0T&}V7orpJ5#SGS4qy#X3[_uy0UZhe9RU{r5&3S.4bTg~32-Et2QUUu0T2cN5Ht@}5ON9N2Ve$J00000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000002m$~A
::0&iaJ5C8xG0000000000000000000000000b]x{jmINF-cmQB00RR916aW@g00000un^=(0RR917yuan00000$Poa50RR910000000000lsy1|0RR916aW@g00000
::un^=(0RR918~^~v00000=n),b0RR910000000000h(},.0RR916aW@g00000un^=(0RR914ge1T000002oeB,0RR9100000000006h8of0RR916aW@g00000un^=(
::0RR914ge1T000007!m.00RR910000000000j6VQ@0RR916aW@g00000un^=(0RR917yuan00000C=vjG0RR910000000000^(+&E0RR916aW@g00000un^=(0RR91
::6aW;f00000ND=]m0RR910000000000WIzCb0RR916aW@g00000un^=(0RR915(#nb00000U=jd/0RR910000000000(^Doy0RR916aW@g00000un^=(0RR914ge1T
::00000coG1B0RR910000000000xIqAb0RR916aW@g00000un^=(0RR916aW;f00000h!OyR0RR910000000000bV2}t0RR916aW@g00000un^=(0RR914ge1T00000
::pb_Lp0RR910000000000U^$^a0RR91Pyhe_00000U{nBr0RR91P#yq+0RR9100000000000000000000000000000000000000000000000000000000000000000
::00000uowV;0RR913jhEB3jhEBpcnvv0RR913/-NC3/-NCkQe}f0RR914FCWD4FCWDfEWOP0RR914gdfE4gdfEa2No90RR914,(oF4,(oFU?E?]0RR910ssI22LJ#7
::a2Ei80RR910ssI22LJ#7U?5,[0RR910ssI22LJ#7P!|Az0RR910ssI22LJ#7P#6G!0RR910ssI22LJ#7Ko;aj0RR910ssI22LJ#7Fc$!T0RR910ssI22LJ#7AQu3D
::0RR910ssI22LJ#7Fc;+U0RR910ssI22LJ#75ElS|0RR910ssI22LJ#75EuY}0RR910{{R32LJ#702cs(0RR910{{R32LJ#7[D~7p0RR911ONa42LJ#7(=(xJ0RR91
::1ONa42LJ#7z!w030RR911ONa42LJ#7[D?1o0RR911ONa42LJ#7uonP/0RR911poj52LJ#7/1(RY0RR911poj52LJ#7(=vrI0RR911][s62LJ#7pcepu0RR912LJ#7
::2LJ#7kQV[e0RR912LJ#72LJ#7z!m^20RR912LJ#72LJ#7/1?XZ0RR912mk/82mk/8fENIO0RR913IG5A3IG5AKo|gk0RR912?;{92?;{9AQ&9E0RR912?;{92?;{9
::02ly)0RR912?;{92?;{9000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::0000000000000000000000000000000000000000,kJ$w00000_e6V7000006k.4X00000EMfov00000K4Jg?00000Okw~400000U}69O00000dSU;o00000kYWG/
::00000o@.w100000tYQEF00000$YKBh00000--qL#00000[@ro0000001Y.aI00000B4Ypm00000Kw|(]00000o@_$200000RAT[D00000USj|N00000YGVKZ00000
::bYlPj00000eq#Ut00000h-^Z&00000l4Ae?00000tYZKG0000000000000000AT;C0000000000N[D/30AK)B0000000000000000000000000,kJ$w00000_e6V7
::000006k.4X00000EMfov00000K4Jg?00000Okw~400000U}69O00000dSU;o00000kYWG/00000o@.w100000tYQEF00000$YKBh00000--qL#00000[@ro000000
::1Y.aI00000B4Ypm00000Kw|(]00000o@_$200000RAT[D00000USj|N00000YGVKZ00000bYlPj00000eq#Ut00000h-^Z&00000l4Ae?00000tYZKG0000000000
::00000Z2)MUaztr!VPb4$RA^Q#VPr#LY/13JbaO]/azt!w0071TPIORmZ,,m2bXI9{bai2DO=WFwa)Ms&Zv/|wY+NiubX9I@V{c@.Q,@4]Zf5_hcmPafaz|x!L~LwG
::VQyq?WdMl+Ok{FQZ))FaY.|7kVgyojY+NiubU|+(X/XA]X?Ml#fB/Nnaz|x!P/zf$Wn]_7WkF;Qa&FRK007^xQgm!oX?DaxZ(Yb,WkzXbY.Do+JpxX2Q+P5Tc4cmK
::002.0Qgm!mVQyq]ZAEwh@g378QFUc;c~E6@W]ZzBVQyn)LvM9&bY,e@_vFdLQFUc;c~g0FbY,Q-X?DZy-yzo}Y,cA(WkzXbY.Dp)Z(Yb,WdPFxQgm!VY/131VRU6k
::Wnpjti~vkza!-t(Zb[xnXJtldY.LYybZKvHb4z7/005EzOk{FVb!BpSNo_@gWkzXiWlLpwPjGZ/Z,Bkp{QypMLu^wzWdLq/WNd6MWNd5z5CC([a${|9000XBUw313
::ZfRp}Z~zVfZDnn3Z-2w?4FGLrZDVkG000jFZDnn9Wpn[l5()B(b8Ka9000UAUw313X=810000pHb9ZoZX?N38UvmHe3/=CqZDVb40000000000000000000000000
::000000000000000000000000000000000000000000000000000000000000000000000001yBG5C8xG2&{LID5C&X08jt_hyVZpIG{-NSfFU2c&X=(n4qYjxS-^O
::,r4d3^[D[)7[/VkIH5@PSfOa4c&g_+n4zelxS_0Q,rDj5^[M},7[{DeV4_rMfTED1prWv&z[pHi/G,!N0HYA2Afqs(K&.EjV54xOfTNJ3prf#)z[yNk/G]+P0HhG4
::Afzy+K&_Kl0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::0000000000000000000000000000000000000000000000000000000000000000000000000000000000000
:embdbin:
::O;Iru0{{R31ONa4|Nj60xBvhE00000KmY($0000000000000000000000000000000000000000$N(HU4j/M?0JI6sA.Dld&]^51X?&ZOa(KpHVQnB|VQy}3bRc47
::AaZqXAZczOL{C#7ZEs{{E+5L|Bme,a00000v~j&1[DS3Q[DS3Q[DS3Q;a]Va]AOUT[DS6R?k!hLXJXr$^z=?YXJXKr[etCRQfXso[DS3Q00000000000000000000
::P)=U$WU2&J)37W^0000000000[Bktp3jz+t073u(01f~E00000SRDWW01yBG0001h0RR9101yBG00IC21][y8000001][y8000000FVFx00aO4dOHCC0RUhD000mG
::0000001yBG00000000mG0000001yBG00000000005C8xG0000000000,l-,;C/$Ke000000000001yBG7y$qP00000000000Du4hm/e9)7##or8~]|S0000000000
::00000000000000000000000000000000000000000B_]R,Z=@k000000000000000000000000000000E^7vhbN~PV3^$;[01yBG073u(00aO4000000000000000
::AOHYhE[WYJVE^OCAO.,c0AK)B00sa607d_-000000000000000KmY,1E[[;8bYTDhwgUhF0B_]R00aO4089V@000000000000000KmY)hE]=jTZ){&em/e9)0Du4h
::00IC208jt_000000000000000KmY)j00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::000000000000000000000000000000SRDWWWdPv/z#RYpy#Z-j=pO)8wE,D+pdbJMsRC(QNFx9M5d_52SSJ7gDFK22fG7X}5dnY!kSPEF$pPg8SStVkc?!bs([2D|
::SpeYyC[la0VF2L+AT9s]1p#FOC[&m2]#NuCC[}y4/Q.^T2r?Ww.2mhP=rRBRDFNgG5HtV+y#Zwd([}+6tper,fH)jE[c@83a5)@~fdOL([Hzkhj8FgokURhYoKOG(
::=sy4ec?rM!U^bx?2@6H;Xh8q~r2yjr5JCU|l?lM]s6qe$bpYT1AVUBEaR6ZfkV60fbpYT12t+t@uuuR1Xhi[3$WQ;Pct!vK-+w}j,hc]W]iTi,=tuwn[darH=uQ9t
::nE~Mf0000000000asY4uV,qjhbO1B}E(yZzYyfNk00000QgCBabaH8KXF^RiWNB^]LvL-xZ,yf=0000000000QgCBJX?Md_Zf8bvZ,5a^a&pa7LTPSfX?Mm&00000
::QgCBabaH8KXGU]mWmf;IQgCBJX?Md_Zf8bvWn}/WQgCBIb9ruKNp5L$X;=-?dSysqZe)m^0000000000QgCBIb9ruKLvL-xY.Mz1Lt$+e00000PGoXHb9ruKLu^ef
::ZgfLoY.|7k00000PGoXJY.wd~bVFfmY&&};PGoX6G)mHDZev4iX=QG7Lt$+e00000PGoXJY.wd~bVFfmY&?4=Zvb.uZ~$.sZvbKdY5/Qp0000000000a{zDvZ~$+r
::VgPCYa{vGUQvh&PZ~#RBcmQ-(LjZ38Z2)UIVgPCY000000000000000000005C9SY00000AQAw80RR914ge4U00000I1(JW0RR917yudo00000ND=]m0RR911wDfY
::_Q_A4?t3d4Z16Y7;nPkf-lX~/B5j4{$ulF4rNb0SK_h3fs64Liicu?ILt-Sp=Aj)Sr/XZE[oPh|dpc/kI9Omm.uZm;d32^rYfqyG5OvGFC[rgk^SDyLkDzhUo9vx,
::i;wr,qqO}@RbVPQ-Q3_u0o8OPicBKvDfr+]e3;o{rdY0UD={Spp@wGKh=tfp]c]kNQbmKO#oc+aWT1ZP?@NkA7(wb[N^^~{A@=URMNRQLsc2W7sY,eW/sHY~o6AN7
::1=va;$)-3YE1mz.uvWR)wT|=l)vkiv_3wR0Nm{rs{M1X@o-8VeXD.TPE^8BC)x5q(1u+@,5gsc|KWbS4@aE.365zvo1ODhXJe09F)O&J_W8TR{X([nURkV/qgz7=)
::tzBIj#C@2ek/({W6)g;6[5m_b($U&5UVOO-OJ5Yt.!hc(5Qo9qPWyP?1,B{cpkiK|u/rgY{vPL@_@_yaQVD9-FgB$+zd+m(f&Dh;eB)KSn=k+|G?$^=#NO&4RC|/&
::rotmV@o5?nLi+o^2ri,!DA]?kc3YxJZHv),a_]UShG?_/+TCU[U1heCY/Z^W{q4EhUKK_Hr/VM2kl3pLjJ)qd^vBawxU+qD([3L0&0CYR!LPjo0TYUAI,}1UPiNff
::m.5ff[U.T0maKFl=dCq_/_uk|9ChDrNAVhQ9Vx|$Z@|F(su/c.{8m0o#@pBpn&ltsc-Fb$AKj=khzG|pu[VqjCxGl;U{Qam8MR6cE#.QjlgXU#px_[At}6Ag$m^d2
::gHxGd7b]sQx^8zl/b|0ORUr)0V|/ge[[sF!Fac,P{[1H]&7V##_dLTtt;;8goTPHVxBZhQHb3{wG]OS7ao8~x1ji&87@uT]2NHnd?nE~x34;(e8,W/lQajeODdR7M
::Q^&qJApEggYRkSkN=#VK)C[1ILrpV;Mfn1MP(}WgQKLYQlASp9ytdjQ5dZVi&@uOlUzbD|#HW5eWL-6]V1ZBEA}WxGM))(2.d-pa/4)T2Nd^cb!qco_k)K0m=g2p0
::jnz+6Y,zH[WqPg&x^Bin9Hz9!=.qT5OTCMVa6YwWNCWl_VKrB|hQS[4/rN(lY1xjHn/wVh(Q(PijG?7Qzve;{L76QNuvEJi2m$~A4rTxV5C8xG(3;_rDzaV6RsYEE
::gJi]T00000WgmZ#(8/p;mAp(!0K9Hk;YPj(k_Km4y!gQ#Uj8Xr_?{!c?hO9=nX6{Xmg&7NX~3UueI?-8w5N3i6w[a|+9-S;1PqBlhd]6$I8$0?am!^fj7FnMqc^W(
::$;]wti6-XjsHxXNla0[gpCB1nw+Ka~M$N!Lux,abSEM(Tugt~|4,#xCod_p4cw6_F]Zg./#mnnNgY,6;PG,3ohco4mhcHJ)iG}x3G9g/Y;W}J_Z@{rPpON.K.Ic6J
::Bfp@~^0V!ak=fN-]_wEeC-A.1#Azr-nbIL]4-dU17id?Ib$en&o+aQJEuNg0Syr)T1LpBeoFDM+0k{~5z~i5m@4ue;pCv,z1?Un|Hr9O7Vj1Z~i&&!E!an;jy0P5L
::li.0(q-;|wQjz_nPZQ!/5snv4oU+MyoD~sBSQEwJKK=tjq[p_)Ajxx1fX9!]1uR.gmk[=o|H&Z_kZ5FW1~wW.hO1eNxJu4~ShGMpNLjB&k~?q;AIyGvJA1jiq?Ly]
::mluipy-XvS8B,SV_uj?qg2];|tyAb$--.?tu|qvgqYN,_ogkIQ77Nve2uQC&xI42rCqo#rv7UWlHt(K[hTx_J/Cq)F8}^w[3o^$Nfl9Y+EBgF_zwxH#K&K+ts.MSu
::q7_^-M+^LkB_(u|gW;ls?,~f470(SwiK)4OuSW89#y19Im/e9)K$,1_KeTKY000000052Yl.;bzr~m+~K;-IPzml2~000000Kl)uBht$O(Hw[e09FqP|F3Zi00000
::0Kfw}FpA9q(Hw[fz#&6Pzk7+i000000KkpX&gW9H(Hw[gzz{@oziOr,000000Kn$YQ0(nG(Hw_iz@at_zr-Y400000005_l/#t&I.4O&]Fm)U_FQPF400000004S9
::_CHckHxmU1AWi[PAA2zY000000KjyvY@s/rHWLL000000Ps0EJ000000KjyvY@s/r(Hw_iI6]A_SCA^J00000005_l/#t&I5+TCj00000zv1Kn00000005_l/#t&I
::U/qLGz;=Bee_[{=0000006;8;uHni7(Hw[hfR6GF|35-x000000D#]8L&q!b(Hw[iKo;.ezsOq~0000006?2o.22c0(Hw_jfLgLAe~05J00000003EunN!pO(Hw}l
::z{Ch5zwtRE00000007C#jjGoH0oMQkau+yq0oMQku]j,aG8F(.[FM]K0T}=QfF&F_91Z{gIXD0S91Z{gV@^V}91Z{gd_|!X91Z{g]ko15Lvnd=bU|Zrb!l?CLvL;$
::Wq5Q~00000K?$PmRscZ(Pyk5+GXOFG0000000000Lvnd=bV-S-Z,p_@WqAMqLvnd=bW?$@OJ#XbVRB)[0000000000Lvnd=bVY7sa)Qrc00000Lvnd=bVOxybaHQb
::OJ#WgLvnd=bW(w)Wnpt=LvL;$Wq5P|00000Lvnd=bVOxia)Qrc00000Lvnd=bVG7wVRU6kVRL8zLvnd=bVy.yXhdOjVE^OCLvnd=bVp[$NMUnmP-[XmZ2$lO00000
::Lvnd=bVOxybaHQbNMUnm0000000000Lvnd=bW?$@NMUnmP-[XmZ2$lO00000Lvnd=bVp[wQekdnZ,2eoZUAEdVE|)QZUA2ZX#j8lUjTFfV,qdf0000000000B?.~)
::IshdAa{yZaB?.~)T?t;800000F#s|EHvldGFaRz9FaRz9F#rGn00000M_d)Vd2[7SZA4{eVRdYDOhZXT00000YXD]casX}sWdLjdGXOFGE(yZzYyfNk0000000000
::B?,r0H2_&0EdV6|FaR|GbpR~[B?,r0GXQk}EdV6|FaS0HbpR~[B?,r0G5~b|EdV6|bpR~[B?/5,E(wn9FaR)BFaRw8B?,r0GXP_(B?,r0Gyr4)0000000000O8_v)
::QvhE8MF4F8bpUJtVE}XhX#j5kZU6uP00000O8_v)QvhE8K?&X^bO31pb]u_jbO31pZvbupNdRsDbO2=lasYM!VE}9Z00000O8_v)QvhE8QUGNDZUAKfcK~4kYye3B
::ZUA&uWdL#jb]u_jYybcNO8_v)QvhE8NB~y=NdQCu0000000000O8_v)QvhE8Pyk5+L/zm]B?,r0H~@$^cmOQ_B?,r0GyrG.cmOQ_B?,r0GyrG.cmOQ_B?,r0G5}},
::cmO2/FaR;DXaINsEdV6|FaR;DXaINsB?,r0G5}},cmO2/FaR;DXaINsB?,r0G5}},cmO2/FaR;DXaINsB?,r0G5}},cmMzZ00000)37W^000005C8xGBme,aG#vl{
::G!Xy,AOHXWhSsx)/uABg$&EuEZ7=2~Rx[)G(/D{rn7il,)37W^]A8{R{d?Nt{R04z]8,5]KLh}ApaB3?KM)-M!2tkNC/$Mk0Kou};3m6?0e}aQLIHr&#Q,[5C/$Mk
::2tf#uXaWHF1ONaOC/$M]2mwI)00BSNAOL^/{d?Zw]9Morzyn{^00000]HaO2]/.d{^hSO7_D-8I_y(AP{d?Eq{R04z2mk;)8]H/Y=?q^(_zHjc^5(NL]aBB]$Ob_p
::.~$P(0{ubL!Gd4.^8$QGC/$M]2u)ow00BSN/0XXVhyp.+sY#1c9{~w#VF?^Kh)18M3#y1xioz)1NdZ8+KLHDCp$Gs}NRdFfXb1o^NtHmkDF]]GlR^wqdO|6S_WpcG
::e,zlof)HOpHUI#yXbwQR2nPT)Xc9oVX-}Y~l|m@s]A_a5mqICvr~,Lw=mh|[,-K!476E|LSOI|2I{,OCIsgFBGXMb4DF/LNsQ?_9r~,LwKLH5qp#uO]2?;{T=mJ3b
::NCWt{2mus}A&k4^00{t,XiGr)00BSNfC2zDNC!aq;U/^F^SXTa0|;ap/$r}j/e!B@004lJ00BSNr~,Lw;U/^FDT7_3/$r}j/}bx-/e!B@004lJC;7h&Xa-#}sR97_
::00BSN7zY5-,!&yrsE$DR^aXq1HUI#yi2DDv]Xo#Xe,zlo1Nr|{^U}WfXu|.J=^f$?.vS8h!Sw&B{d?iz_y+X4_D/U|^hUk.]/;!y]HasBzyn{^00000C/$Mk3c(!8
::=?rO@O96n=4hDeIZ2dvgjU]77s1.o[9{~XCq5uF@XaNk&3k3ktslfn|0ssIM?jMm_e,pmTtp5L0NP!2DKLH5qfB,ngC?22Y9{?pJLI40&Nr4BEAQ3@Mzyn{^]A8{R
::{d?Hr{R04zC/$M]2nj(?]8,2[/R67w/DZ2?00BSNC/$Mk2n|5^;3j-E/+4K[0RVu~004l}00BSNU/-3xC/$Mk2oXT};3j-E/+4K[0RVu~004l}00BSNU/y|w004l}
::5C8xaC/$Mk2o,s2/R6$[/KKls00BSNC/$Mk2pK]6/0r-c;6{7k0sw$g/llut00BSN=np{o9{?Px0HL3n{d?fy]9Morzyn{^00000]HaO2]/.d{^hSO7_D-8I_y(AP
::{d?Eq{R04z]aB8[]#cK^r~)wrAHf,$^5&W{]8,7aD-B/k83usT8U}#U.vR,fEeHTq/{y{a/sX^_/R6)]/6nhBlmGyf^-LS)$o[f.773P&/{y{a/sX|{r~))u3Juws
::2m=)$2[TqsKLHBs$]ZaV/R6^|.~$w[.2eZV]aB]F1pojP/R6-^.~$)_,Z=?Q]#d5Hr~)wrAHf,$YW+9Hp8]&[fDQmulfnRze,zWjAPxXjb]/X3mm(bsn8E;j_hx)G
::EdT)p_.1@H_GWwFXeL0Z?/n^3YA!,kNGAZPXeL6bN.qJaiWWfmNGAfR@k7O_.vJ8i!UzCVXaW|@0Kou}s8K.q/sX|{/R6)]00BSNXeU6aYA.?lh$aB3XeUCciY[_E
::h$aH5O#lECwgME)2nK.C!VbuqSNuVf{{jH;Edl]k^+.X+_GWwF2q,oi^XYsb;O35b;AVT]/R6@{C@]1]3NJya.~$w[2q!|RDlY.4C@]7_NdW-q{{jH;Z2tdLUkCv4
::O9uc{wZZ^=7Y2aR^=5nE83usT$PU-;9|.{Qivj?ts3t)E;O35bsxCpP;AVT]0RVu~.~$w[2qyrks3t;G3NHbv2qyxmh$cX)s3riZiY_H]sxASkssa@th$cd,LJirP
::s3robD,,tMwZZ^=r~)wr7Qq0K.-}[0&KZOS8UO$k=xTQO4FeX73IG5Us3kzDh$R52sx3jOiY+=Ds3k)Fh$RB4?/ny|t.&1&s1.o[.vJ2g!~XwNC@_OvDlb8+h$R52
::C@_UxiY+=Dh$RB4bHV^Te,zWj,1_ahEdT)p0rUS;{d?iz_y+X4_D/U|^hUk.]/;!y]HasBzyn{^]A8{R{d=mZ{R04z;]ut$;pTn$2@l_Dr~n4b2o1[a3I?4EufPD(
::]8+~@.~$G#3H@EnE,T1(=m7[H2@l_D2[T1bKcN8eZ2|yPC;OqK4gEut2MmDH1O|Z8q8SI9p(105H30yWpt&H^/R6n/qB#VcF#!OSp}ho~/sXz=puGp1Edc;O/sXJy
::pcw@40ssIM/sXz=/R6n/CjkJI(A|YX.vAElLID6(${^(JNd/Z^NHIY9KLH5q03k{Gg8&@j(cOiD.vAElBme)YzX1j7A]_wY$rV8Ps1.o[9{~yLAR$Qlg8&@jt.&11
::zX1j7L/wF(@,k30=K~I]{{aQ.A^M@b?^Y(N;U/^F$Q3~O=[mfv9{~yL0|Nk5KLH5qBLe^bzX1?HU/-SCs3kzDsx3jOh$R52s3k)FiY+=Dh$RB4=p{g[s3icY?McR3
::sx1Mj=p{m^iY!5@s3iiah$KL&=p^KDh$KR)?Ma4O=p^QFt.&11[4,0){{aQ.WBmVA{{RN.sRRI2@7#rg/{ySa/sXJZh!sHj$rV8P9{~yLBLe^bKLH5qV,?zG#K8d3
::@gIp?p8yQ(U/-SCh$KL&iY!5@Xe0osh$KR)YAgY&Xe0uuh$TR(h$H~1iY.B[iYx+Ch$TX+N.ROCh$I53NF-e1h$R52NF-k3iY+=Dh$RB4=fD8b.v9]ejKKiWBmDnV
::{{RN.0R{k6{{aQ.pbh|3zX1?HpaK9@$R$9j@85;)$}K]uh$R52$R$FliY+=Dh$RB4s3kzD@85;)$Rz.(sx3jO$}It[s3k)FiY.B[$Rz[+h$TR(s3icYh$TX+sx1Mj
::s3iia$R$9j@1KW4$}K]ut.&11h$R52$R$FliY+=Dh$RB4=p{g[@1KW4$Rz.(?McR3$}It[=p{m^iY.B[$Rz[+h$TR((cOhY=p^KDh$TX+?Ma4O=p^QFMgRa5{{aQ.
::0R{k6=fD8b(cOiD{{Rl^paK9@=p/a@?^Y?Q?MTL2h$R52=p/g]iY+=Dh$RB4$R$9j?^Y?Q=p-EC$}K]u?MQ}N$R$FliY.B[=p-KEh$TR($Rz.(h$TX+$}It[$Rz[+
::$R$9j;O2ke$}K]utib[$t.&1&h$H~1$R$FliYx+Ch$I53h$TR(;O2ke$Rz.(iY.B[$}It[h$TX+iY!5@$Rz[+h$KL&h$R52h$KR)iY+=Dh$RB4SpWYQ=p{g[@85|-
::?McR3h$R52=p{m^iY+=Dh$RB4=p^BA@85|-=p^KD?MTL2?Ma4O=p/g]iY.B[=p^QFh$TR(=p-ECh$TX+?MQ}N=p-KE[4,0)LjV64ZZ.g].~$t@{d@A]]9Morzyn{^
::]Haa6PXqwbKLn5K=K}$&@gIg/3IhOC1^prAMF4=)BmjWY69$0N6b69O&K3lOEdUgoNC5^$2_xbR2t_2o9{~yLsUU=!E((RQ&mEXd/R6n/.vy8Bh$TR(s3icYiY.B[
::sx1Mjh$TX+s3iia?/3/!.vy8Bp#cC@f(l;G2nK.CO#ld-2@l_DEC30c/R6q;s3m==h$R52sx5x0iY+=Ds3m_@h$RB4{{R8(iUI(s4-enJ1^prAC;Fk}X&s/D4,fxs
::&?fUas1.o[9{~yLVgUeDs3kzDEC2@Z{{Rl^/R6n/h$R52sx3jOiY+=Ds3k)Fh$RB43/zF92nK.CEC2|bXe2;Xh$R52YAiviiY+=DXe2^Zh$RB4]Hag7zyn{^|HA/$
::DHK5Y2oym1KLH5q!U6zPC@r6s?/nLiDl9?&h$R52C@rCuiY+=Dh$RB42qZwM?/nLiC@o+?3M[gXDl7r12qZ$OiY.B[C@o=[h$TR(2qXZhh$TX+3M?Js2qXfjDHK5Y
::NEAT&9{~yL!UO;RNF-e1?/nLiN.ROCh$R52NF-k3iY+=Dh$RB4C@r6s?/nLiNF+HMDl9?&N.P1XC@rCuiY.B[NF+NOh$TR(C@o+?h$TX+Dl7r1C@o=[3lu?4DilEZ
::UjYm2!T|tO2qZwM?/nLiC@o+?3M[gXDl7r12qZ$OiY.B[C@o=[h$TR(2qXZhh$TX+3M?Js2qXfj|HA/0zyn{^]HaU4]/.d{^Y)m5{d?Nt{R04zHUI#S$l5~r|9=6g
::]8+~@]aBB]]#cN_^y7O!=l}q;?Hq+mApt0n/9[9|hW.DS=mP-&$l3z=1OUEL0|S6k0sw(00RVu~/9~&h00BSN/06FRHUI#S$lgNv=l}q;?Hq+mA/Bn./9[9|cK!dC
::=mP-&$le0]/159g?Hq+mAwd|C;wF3G1OR|i0|0?1f(-k300BSNpacLk69NFVHUI#S$l]lz=l}q;?Hq+mApt3o/9[9|WBvb]=mP-&$l@O|/0r-c0|0;h/sX;]Apn3;
::00BSNpaK9iGXMaPXy!us=l}q;?Hq+mAz?-z/9[9|RQ?/#=mP-&XyyX?/0r-ch9iJd;pUL};O39{0|0;hA]@C=0RVu~00BSNU/qF#GXQ{60ssIM699lx/0r-cfB]us
::GynjQi0VT52mt_K?Hq+mA+zUe/9[9|J]lZe=mP-&/0r-ci0T6Q00BSN/159gpaB51GynjQi0)r92mt_K?Hq+mA?k?J/9[9|F#Z3R=mP-&/159gi0&UU0RVtf00BSN
::.~$sX{d?Zw^Y,-,]/;!y]Ham9zyn{^]HaX5]/.d{{d?Nt{R04zH2@sRsM;pL|9=6g]8,2[]aBE^^W&Fz=l}q;?Hq+mApt0nz-xzo7XAO1=mP-&sM.Sg1OUEL0|S6k
::0sw(00RVu~/9~&h00BSNzyts]H2@sRsNO?P=l}q;?Hq+mA/Bn.z-xzo2L1n-=mP-&sNMqk/159g?Hq+mAwd|C;wF3G1OR|i0|0?1f(-k300BSNfC2zCH2@sRsPaPj
::=l}q;?Hq+mAwepUz-xzo]!+#q=mP-&sPY2(/0r-c;pUI|;O36_0|0;hA]@C=0RVu~00BSNU/qF#GXQ{60ssIM699lx/0r-cfB]usGynjQi0VT52mt_K?Hq+mA+zUe
::z-xzo.u)ZU=mP-&/0r-ci0T6Q00BSN/159gpaB51GynjQi0)r92mt_K?Hq+mA?k?Jz-xzo)ft3H=mP-&/159gi0&UU0RVtf00BSN.~$sX{d?Zw]/;!y]Haj8zyn{^
::]A8{R{d?Nt{R04zC/$Mk2vtD(]8+~@0s@]2/R6$[/6nhB00BSN3jlyp?^Y(NXbB4oYXtxie@b6o2[OD!DrsyuY8C+EOaK2={d?Zw]9Morzyn{^]A8{R{d?Nt{R04z
::C/$Mk2vtD(]8+~@0s@]2/R6$[/6nhB00BSN3/=,q@Lz?Oi3J{0h;!lQ2|-2#j0FG[pFsd|Dh+uAOKEL5YZd[F3/-LA{d?Zw]9Morzyn{^hy)x^9C$!X5Df&Q=+)Xq
::G6Vom6&7PVb^W1Y1^uC71qA@48U^GQ5gt5FnMb8=u)-UZH&78;[(o_-n[75CQ[WsTlt/5}l]!+tww|^5#~wFs,SMf={~SDSm_As6[kg-39v_^,S.YTann$]A{70s4
::T^3wnJ084Fd?=h.ogY4Kz8[!U9)Vvuzyn{^e}8}f00000]HaU4]/.d{^Y)m5{d?Qu{R04z=?Pxl6oCzq]8+~@ivWO9i~;wO?H_z1iD^!MYXtyNNJT+nDFFydNx?hu
::YybZ?mO=oLH35K9m&/{.=?rq03Ic#qC?20BN)BH?2x+gXDDfXSivRyL.~$t@kpKUe.~$t@y#N1~?H_z1ivWO9Nku[oYXtyN$VNc8DFFydNx?huYybZ?wFUrDmHq!U
::=?rq03Ic#qi]2wxC?20BN)BH?2x+6LDDfXSivRyL.~$t@djJ2Ih=Kx;3jq^$iU5F8X=!t~N)BH?XhuM|DFFydX~G}4YXAQ={d?Wv^Y,-,]/;!y]Ham9zyn{^00000
::]HaU4]/.d{^hSO7{d?Eq{R04z=?Pxl6oCzq]8+~@h=Kx;?H_z13/_3&ivWO9iD^!MYXtyNNJT+nDFFydNx?huYybZ?5eEQIcMSj.[B{!+^_@7+l|llMF}k2_HUWTA
::6uO{p5Cs5F]v3|L5e5KH5W1jlF}k2_[xuYF.~$t@U/qD@=?rq03Ic#qi]2ktC?20BN)BH?2x+6LDDfXSivRyL?H_z1?/o05ivWO9Nku[oYXtyNh)$oSDFFydNx?hu
::YybZ?_VIt6{s-K4wL$?Ve|kVn+(?Ak8xI6dIRpSt[khRHP#.[|G9Eil6h]sja0dWSQ=YI-WF9nbl|/U7R39WxwjMi9m_1s7,PgIW{T@_OSRXx397nlsxktWkIv-bu
::Tc5B^[C)2^d?=e-o,zGMg(#d_.AAx+5Cs5FzZ]eq#~(na(^}Rt1|Gdm[DIQ}[;,^45C#BG[kg-3[Dsp2J|418]F,-25C/HH[;gz1Qy#NUbRIr#l]!N/wjL#J,B(Hp
::cX|L!cK81].~$t@8UO#6=?rq03Ic#qi]2ktC?20BN)BH?2x+6LDDfXSivRyLivknNiU5F8X=!t~N)BH?XhuM|DFFydX~G}4YXAQ={d?iz^hUk.]/;!y]Ham9zyn{^
::A0PwOe}8}f00000]HaX5]/.d{{d?Qu{R04z^5&W{$]t/S]8,2[]aB8[=mRP$2[L=eAq4/tH2@|=zj6d|X#fCJ004keB?)]vC/$ME2w6b-U^vU3/sXJy00BSNQ~@0A
::?H_z1i~;wOivWO9iD^!MYXtyNNJT+nDFFydNx?huYybZ?.~$t@9{?NBv^b$/6aoM=YC.]!?jMg]Z2}6,i~xXAscCDtj0FHuXhlG{DFFydX~7[3Z2$i@]8,U1.~$t@
::5C8v{ltKVeRQ~[p+dB#yAOL^/{d?Wv]/;!y]Haj8zyn{^]HaX5]/.d{{d?Qu{R04z^5&W{),i+b]#cK^e-~e0U/qGA004keDF6TzsKPUg6hQ#d3/-NW.~$w[IsgBc
::?H_$2ivWO9NdaHDYXtyNNJT+nDFFydNx?huYybZ?ltKW}p8]&[i2nan.~$z]EdT$Pe@kCpU/-SCsKPUg3GrVzKS2O.=m7v!3IKpo?jMcYDFFa93;Utui1lAM9{~w#
::p#T6?YXtyNe,pk.N)BHBO#lB?UjYegFae)$a{?[cAOL^/),gjw{d?Wv]/;!y]Haj8zyn{^00000]HaX5]/.d{{d?Qu{R04z]8,2[?H_z13/-|$ivWO9iD^!MYXtyN
::NJT+nDFFydNx?huYybZ?=?PxF6[dzotO66u?jM-2iU5F8iD^&NN)BH?XhlG{DFFydX~7[3YXAQ=Gys57w!#UK=?rq03Ic#qC?20BN)BH?2x+dWDDfXSivRyL.~$t@
::SN{K).~$t@gZ}[QiEbQItU[V]?H_z1ivWO9Nku[oYXtyNh)$oSDFFydNx?huYybZ?lm.A1pDqA#BmMtW=?rq03Ic#qtHKG9C?20BN)BH?2x+6LDDfXSivRyL.~$t@
::KK}ogsKNq~3jq^$iU5F8X=!t~N)BH?XhuM|DFFydX~G}4YXAQ={d?Wv]/;!y]Haj8zyn{^]HaU4]/.d{^Y)m5{d?Ks{R04z2n2vq|NjB0=o0|B761V7$l5~r]8+~@
::]aBAZ]#cN_^y7OU=l}q;?Hq+GApt0n/9[9|;of[Y=mP-&$l3z=1OUEL0|S6k0sw(00RVu~/DZ2?00BSNKn4Ib761V7$o4|{=l}q;?Hq+GA&QB9/9[9|+cXII=mP-&
::$o2yH/1fXk;YNGl0|0;h0sw(0fdP;G00BSNKm.6Z761V7$ofM0=l}q;?Hq+GA/Bt;/9[9|#QOi2=mP-&$oc~L/159g0|0;h/sX?a/R6$[00BSNU/-R&69544X#PU]
::=l}q;?Hq+GApt9q/9[9|wfg]/=mP-&X#N8E/159g1OR|i;3j-E/sX^_K?(bK00BSNU/qF#GXQ{60ssIM699lx/159gfB]us6aWD5hyp|T2mt_K?Hq+GAz?@#/9[9|
::p!+xp=mP-&/159ghynxo00BSN/1fXkpaB516aWD5i0)r92mt_K?Hq+GA?k?J/9[9|lluRc=mP-&/1fXki0&UU0RVtf00BSN.~$sX{d?cx^Y,-,]/;!y]Ham9zyn{^
::00000]HaU4]/.d{^hSO7{d?Bp{R04z=+)Y!|9=9hAAJC-i2/yOAAJF.]aBAZ9}xig2n2vq=o0|B2mk=]69E8_{|]B9]#cN_=_#Si^5&Z|.~a&$C/$ME2vtD(/R67w
::0s@]2U[_!a00BSN7ytn92._yW^y7OU=l}q;?Hq+GApt0n/9[9|WcvS@=mP-&2.]br1OUEL0|S6k0sw(00RVu~/DZ2?00BSNAPN997ytn92/V~a=l}q;?Hq+GA/Bn.
::/9[9|RQmsy=mP-&2/Tzv/1fXk;+Z-R1OR|i0|0?1f(-k3/R6$[00BSN00/my69544X!b)-=l}q;?Hq+GA&QB9/9[9|L/C.h=mP-&X!Zj6/0r?j;YNGl0|0;hApww5
::00BSNAO.-569544Xa-;1=l}q;?Hq+GA/Bw=/9[9|H2VLS=mP-&Xa+oM/159g0|0;h/==&up#XqV00BSNKm.6Z69544X!=6==l}q;?Hq+GA/Bt;/9[9|CHnuD=mP-&
::X!.,A/159g0|0;h/sX?a/R6-^00BSNU/-R&69544X#PU]=l}q;?Hq+GApt9q/9[9|7W+5}=mP-&X#N8E/159g1OR|i;3j-E/sX|{K?(bK00BSNU/qF#GXQ{60ssIM
::699lx/1[vofB]us6aWD5i0VT52mt_K?Hq+GA+zUe/9[9|0s8.!=mP-&/1[voi0T6Q00BSN/159gfB]us6aWD5hyp|T2mt_K?Hq+GAz?@#/9[9|]!fjn=mP-&/159g
::hynxo00BSN/1fXkpaB516aWD5i0)r92mt_K?Hq+GA?k?J/9[9|=lTDa=mP-&/1fXki0&UU0RVtf00BSN.~$sX{d?l!^hUk.]/;!y]Ham9zyn{^00000]HaX5]/.d{
::{d?Nt{R04z6#xM6sM;pL{|f/5]8+~@]aBAZ^W&FT=l}q;?Hq+GApt0nz-xzo&=!P9=mP-&sM.Sg1OUEL0|S6k0sw(00RVu~/6nhB00BSNAOZk16#xM6s0u],=l}q;
::?Hq+GAt5Z0z-xzoy!ro]=mP-&s0su5/0r-c/sX;]/R6(Z00BSNU/qF#GXQ{60ssIM699lx/0r-cpaB516aWD5i0)r92mt_K?Hq+GA?k?Jz-xzosrmnx=mP-&/0r-c
::i0&UU0RVtf00BSN.~$sX{d?Zw]/;!y]Haj8zyn{^]HaX5]$P(_{d=]j{R04z|HA/$]#cH]r~,K^]aBB]0SJK7/KKothynn+sOmsDt]PnctolGXtM++S=mP-_=?PxF
::0zop7C/$ME2t_2os_5ZNsqR2I@JEGer{-L8??~iVrs6;3?l,/MrEWlZ?JtFDq.sEU=@eh4qcT9b00BSN2mk=]0U1I0C/$ME2nj(?/6nkC00BSNC/$ME2suFc/sXJZ
::0RVtf/6nkC00BSN00Q^oC/$ME2t7dg/3Gi!1pt83#1DW{gCYQtA]@C=/llxu00BSNC/$ME2th#k]8+}X/3Gi!00BSNlK}WO/R6-^fFb~qp927tC/$ME2wgz=fFb~q
::00BSN/e!E[2@PKU/3EN&DtRAMiUt6=s4hgQh]_2!sX|5giB16ds8T@=2?;}^DS.fy3V9z?C=oz/DHT9[iXs##iK-m)sH#dS34sc#C/$ME2pvHA=^dgB00BSN|HA/0
::{d?&+]$S4x]Haj8zyn{^]HaR3]/.d{^hSO7_D-8I|3e7T{d+kZ{R04z^X7c{.~$)_/llut^yYo}_2z#0_U3?2lmGvh=r=(Q/llut/DZB]6CnVRC/$ME2vtD(/sX;]
::00BSN=z{~1a{?s9C/$ME2vtD(f(^rl/o}04.~$t@00BSN=z{~1X#xmKHjw}k=|cdK=z{=}KYakH]n)MDAAJC-]#c|v.$DR!D,,sh)|!a~+e/j-/e!B@.~$w[)f$9I
::ltKWJa|QrWbN~M}zXAYptpEU2qJ98V/R6)]/6nhBE)HLT=|cdK2oQi$/e!B@D9JTA/6nhB!u|i3=z{~10KqnkC/$ME2vtD(0s@]2/e!B@00BSN&0d7U=mQd}3IhPS
::2{AzVC/$ME2sJ@YLVZA!0RVtfAQ@dU00BSNC/$ME2vtD(0t0}#/e!K^]8+~@00BSNC/$ME2vtD(f,pX//R6@{.~$;|00BSNC/$ME2vtD(f+#-$/llut.~$@}00BSN
::.~$t@{d-,E|3e6o_D/U|^hUk.]/;!y]HapAzyn{^]HaR3]/.d{^hSO7_5OTF|APt9{d+kZ{R04z]aBB]hyp/l]8+}X^5&W{^X7i}^yYv0=tBUxA3/HJAprnXC/$ME
::2vtD(l[b7v0s@]2/R6-^/1dCn00BSN82|tj0Rn)h/DZ2?/{N}a2m,jo=o;jJC/$ME2vtD(0s@]2/e!B@/1dCn00BSNhyp/lA3/HJ.~a$rAAvz}0RaG1/$r}j/S(Lo
::H2wdV1ONaO/$r}j/S(Log!}+Ol[b7vi2]{mXc7QX=obLFKS4op.~a$rKY?AU0RaG1/!]/T/R6-^CH@=G1ONaO/!]/T/R6-^b]HI9/ll=zfKmXF_2PQw=)j;-/ll=z
::/8OvS6CnVRC/$ME2vtD(/sX;]00BSN=u.iaa{?s9C/$ME2vtD(f(^rl/o}IA.~$t@00BSN=u.iaX#xmKGm!uh_BMRrAj30[0Rn)hrT-hy=#v4FAj30[0?Lwj0Rn)h
::g#G_QD#J62?/o05ivWO9Nku[oYXtyNh)$oSDFFydNx?huYybZ?ivmEo=o12w6Tvf!e}O[90R{k62mk=]2oXT}0s@]2/R6-^U@KpKXaWHFC/$ME2vtD(00BSN=qEw?
::X$t]Yiwgi+stW,E/==_z2norW0Rezg/9~&hb7BCI2q^Dj=nnw.Vg3J@C/$ME2vtD(0s@]2/R6Pd/KKls00BSN=o0~vVFCzC;3k3K/u8Up/KKls#r].6C/$ME2vtD(
::0s@]2/e!T|.~$t@00BSNC/$ME2vtD(0s@]2/e!B@/1dCnb3y=.00BSNivmEo=u.iae@dWUY61vL?JtFD00970e}O[9VF3VC/zIzD/Zp(T/1dCng8cuN1pojP/zIzD
::/Zp(T/1dCnm.^#g?Jvb[N?Kn2=mQd}$]rnn2{AzVC/$ME2sJ@YLVZA!0RVtfAQ@dU00BSNC/$ME2vtD(!UBM~/R6AY]8+~@00BSNC/$ME2vtD(f,pX/fl?gG.~$)_
::00BSNC/$ME2vtD(f+#-$/ll=z.~$-{00BSN.~$t@{d-,E|APsU_5Qp^^hUk.]/;!y]HapAzyn{^A0PwOy[^anA].pY@X|j$AOHXWdPgY6TFq85]A8{R{d=XU{R04z
::]8,8^A8.M2ssI2~UjP8P/0l0Je,ysc5(![cC/(jY9|1veKmh;$2th$nA9+XQU/qGA004l}2mk/S;U/^F/{yYc7XSa31ONaO;U/^F/{yYcs{a3&U/-U7004ke{d@P}
::]9Morzyn{^]HaR3]/.d{^hSO7_5OTF{d?Qu{R04z^yYi{]aBB]]#cN_^5&Z|_2z(1^X7p0v/-XO=?Pw+0bwkW2?;}^C}BYP.~$w[00BSNzykm]jspOc2mk=k69E#D
::XaYdFC;6dB2mk=k2)dspXaWE;C/+(_XaWGa=mQd}XpR8-=?Pw+0]ux@2mk=]2t_2o;pUO~;O3C|/{z0_0T6+FU=je400BSNXc7RC=mG&w004ke4,(oZ=?Pw+6-tbL
::e,yrx2mk;)0D&FKmG}Rb@,jm/.~$-{;pUS0;O3P1/{z6|/sX^_/R6)].~m6[{d?Wv_5Qp^^hUk.]/;!y]HapAzyn{^]HaX5]$P(_{d?Qu{R04z]#cH]]aBB]6$1dY
::]a2312mk=k69E#DXaYdFXaWE;Xo]7jC/|YrXpTVn=?rm~9{~yLp#cC@2mk=]2w^0]VG/n500BSN0096s0RezgU@K#Ot]NO)Xof+f004kehynol2mk/S2mk;)0HFnu
::X7~S?@,jm/.~$z]/R6)].~m6[{d?Wv]$S4x]Haj8zyn{^00000]Haa6{d?Qu{R04z2mk=k69E#DXaYdF]aB8[r~({qlmY/?XpTVn=?rm~9{~yL0RjM22mk=]2w^0]
::fC51IVG/n500BSNKmh;X2mk=]2w6b-0w93W0RVtfU=je400BSNp8]2-004ke2LJ#R2mk;)06^@mKKK8Z@,jm/.~$w[.~m6[{d?Wv]Hag7zyn{^00000{d?Qu{R04z
::2mk=k3IP)4NC7~)=K}z$=m7vU#{mGeNrgc9=m0@Z9{~yLK?-|&NR2[G=?rm~9{~yL!2keMUjYEQ004keU/PlNUyT6y2LJ#R2mk;)0HF$zANT,4@,jm/.~$J$.~m6[
::{d?Wvzyn{^00000]A8{R{d?Qu{R04z]8+~@2mk;)0D&mV68Ha@@,jm/.~$t@.~m6[.~j-N2mk=k3IP)4NC7~)NC5yeNQFT82mt_JONl_F&K!kiNR2[G=?rm~9{~yL
::X#$IyNr@dY004ke{d?Wv]9Morzyn{^{d?Qu{R04z=K}z$M,/w}Ap!uj2mk=k3IP)4NC7~)r~v?pONl_FYXJbXNQFT8C/;SpNR2[G=?rm~9{~yLp#T6?{{Rc@VE^PB
::004ke2LJ#R2mk;)06_9s;[W!V@,jm/.~$J$.~m6[{d?Wvzyn{^00000]HaL1]/.d{^hSO7/tvC;;QD{~;{t(A{d?Qu{R04z?H_6h=mQGN2@/=wDgg@MOCbP}=mQJO
::bAey[2@/=wDgg^NOd$Y~=mQMPgMnZ82@aosDgg|OOCbP}=mQJOb&9]F2@/=wDgg^NOd$Y~=mQMPmVsaR2@/=wDgg|OOCbP}=mQJOcY$B{2@/=w2mk=k3IP)4IB9G6
::NC7~)]aB8[$O8a0v/zRNfdc[vNQFT89{~gFAp.zZNQprC9|05V!2$qONR2[G=?rm~9{~yLK?_3(Xc|EI2@YSrKMeq}$N?OUegXiL/{zC~/sY0|/R6;_.~$yZx(Hr_
::2mk=]2q8fEU?ZRA0RVu~00BSN004l}3/-NW2mk;)03i]OmG=La@,jm/.~$w[=K~n3;]vb1;pUO~;O3Bd/sX;].~m6[{d?Wv^hUk.]/;!y]HavCzyn{^00000]HaO2
::]/.d{/tv9/;QD]};{t#9{d?Qu{R04z?caq$=mQGN2@/=wDgg@MOCbP}=mQJOVu4[y2@aosDgg^NOd$Y~=mQMPlYw8j2@/=wDgg|OOCbP}=mQJObb),]2@/=wDgg^N
::Od$Y~=mQMPm4RRQ2@/=w2mk=k3IP)4Hfe15NC7~)=K}z$NCE(fCjtPp0RjNDNQFT89{~dEp#cC@NR2[G=?rm~9{~yLAprnXGXemV/{z6|/sX^_/R6)].~$sXZvOw5
::004l}3jhEV2mk;)0HGC/Pxk-p@,jm/.~$J$;]vY0;pUL};O39{/{y{a.~m6[{d?Wv]/;!y]HasBzyn{^00000]HaL1]/.d{^hSO7;C6oa;)mYl=Pv/H{d?Qu{R04z
::|3d+L?SF;s=mQGN2@/=wDgg@MOCbP}=mQJObAey[2@/=wDgg^NOd$Y~=mQMPgn@i92@aosDgg|OOCbP}=mQJOb&9]F2@/=wDgg^NOd$Y~=mQMPmVsaR2@/=wDgg|O
::OCbP}=mQJOcY$B{2@/=w2mk=k3IP)4IB9G6NC7~)]aB8[r~@2rlmh]@NQFT89{~jGAp.zZNQprC9|05V!2$qONR2[G=?rm~9{~yLK?_3(Xc|EI2@YSrKMeq}$N?OU
::/sF4Z/{zC~/sY0|/R6;_.~$yZ9sd892mk=]2q8fEU?ZRA0RVu~00BSN004l}3/-NW2mk;)0O1,t_St(o@,jm/^yYj?.~$w[=K~k2;]vY0;pUKe/{y|^.~m6[|3d)g
::{d?Wv^hUk.]/;!y]HavCzyn{^]HaO2]/.d{^Y)m5;C6lZ;)mVk=O-O9{d?Qu{R04z?f.?B=mQGN2@/=wDgg@MOCbP}=mQJOV}W1z2@aosDgg^NOd$Y~=mQMPl!0Hk
::2@/=wDgg|OOCbP}=mQJOb&9]^2@/=wDgg^NOd$Y~=mQMPmVsaR2@/=w2mk=k3IP)4H+)A6NC7~)]8+~@C/|X969NFVNQFT89{~gFp#cC@NR2[G=?rm~9{~yLAprnX
::mjM8j/{z9}/sX|{/R6-^.~$vY),6IJ004l}3jhEV2mk;)0AU}Iv.SU&@,jm/.~$t@=K~k2;]vY0;pUL};O38c.~m6[{d?Wv^Y,-,]/;!y]HasBzyn{^]HaO2]/.d{
::^hSO7_D-8I_y(AP{d?8o{R04z|APS02mk=]2q{4M/6niU^yYl|0RVu~.~$1X^#XiI00BSN=tDrc0s@]207C$g=raJh=?Pw+0l][V2?;}^C}lwT00BSNsR4je=m3CH
::9{?PxIRS^o2?;}lG=T|_]8,yBl?!Kn761V7U//q.0s@]2$YwzK0mA]100BSNfB,nAb3y=.e,zWj0ssG0hyn;ae,y]WzyknObs|A26af_Wp9TOi!2keM2mk=]2qi&I
::/KKls00BSN6aWAelLi10761V7004ke2mpXmv/Y7!$R;Gf/KKls00BSN2mpW,Qvd+pfB.)!2@]60Dxnh^2nf?}9{~w#L@KjqLH^@#2nf?}0D&+582}Xv?Hq),I{^h+
::D1SiH3Il.B76]dS8f].j699mc?VH78NDDwY_XdCXNe[6dbN+foUjY/A+(?C4qyPU[Gyw@9p9TQ2Ap!tY2mpZ66aawI,9HL57ytn96)QG}2q![K/KKls00BSN2z+[)
::9|05VX#f9IlmZru2n7|Yp8]5#!~XwN9{?Opl@DKj3jl!97XX0L6aoOW,aiU66#xM6dLh_Fs3t+9/KKls00BSN2z+[)9|05VX#f9I3k4dg=pxvfcLoTm{{j]2qyGO@
::l@DKjRR93BfB.)!=@c}FXbIJtD&}/E9{~yLL@KlA/r#zpXbIJt_XfO32)1/H=^f$?0s@]2^#Z(|6#+s0Xafe32n_jB0K+,0=sy7Y=?Pw+7Qq{l2?;}^C}lwT!NLTQ
::00BSN2mk=]2xUO|9{?Px0s@]2!GZ,l0K+,0vjUkK00BSN2mk=]2xUO|9{?PxqzXd&0s@]2!9oO+0K+,0vx1Wv00BSN2mk=]2xUO|9{?PxqzXg(0s@]2!2$&40K+,0
::vx1Wv00BSN2mk=]2xUO|9{?PxqzXj)0s@]2!NLQP0K+,0vx1Wv00BSN2mk=]2xUO|9{?PxqzXp,0s@]2/R6;_0K+,0vx1Wv00BSN3Il.B9{?PxqzXm+sDhIkRR=+1
::7JUkvNDDx[6af_Wp9TOi!2keM2mk=]2qi&I/KKls00BSN6aWAelLi107XSe8004ke2mpXmv/Y7!=q5n;/KKls00BSN2mpW,Qvd+pfB.)!2@]60Dxnh^2nf?}9{~w#
::L@KjqLH^@#2nf?}0D&+57yuOu6953vU=Bn1^+7q}.~a&&=?Pw+0ii9C?Hq),AOS9s)DeV8U=9QM002MMXhQ(z+M{w?2mk=]2rWSQ/5z_h/G-SN00BSN=r=)5DG(fy
::76]dS699mc3IPer=zl=6bN+fo9|05V+(?C4qyPU[2mtWX3/]+b3jpxai2@}Ahyn|Xp8]c+Iw6Rf2mtWXNdXAUUjYp3NC69rC@SZN7Xcc}{{{fDAp!tY2mpZ66aawI
::,9HL582|wA6)QG}C@_Pq/KKls00BSN2z+[)9|05VX#f9IlmZru2n7|Yp8]5#!~XwN9smFo]acQt8vp@C3jl!96##)J,8u?u,aiU6c^G.EXeU7V/KKls00BSN2z+[)
::9|05VX#f9I3k4afs3O?!bp{BkzXBKQqyGO@Q~(^AfB.)!sS4DZXbIGsD&BL59{~yLL@KuDA]rbUXbIGs^+9?!2)1,G3IQ663jpxa2mtWX2?|fYNC60oNdXDVUjYm2
::DItiO2mk=]2pvHA=^?(F/bQ[j/6nkC00BSNC/+(_=)hlQ/llxu=)^.U2mk=]2r+qU/9~+i00BSN/6p)9X8@dw0SJK7?$AXlhy)x^2!E+X3;LmJ3H[Nx=sQ69C4CZ8
::=^]3^?l,/MDgg-~=p#V-N+61KOalNIrr.=41OY(KNd,8A?OVmF=[S6C={rEVLm[yZ=?q^{=nDY3=?Pw+0+Z})2mk=]2t_2o00BSN|APRL{d?o#_y+X4_D/U|^hUk.
::]/;!y]HasBzyn{^0KjP~$p8QVgWelMKtc}y]A8{R{d?Hr{R04z2mk=k0U1I02mk=]2nj(?/G-PM00BSN2mk=]2suFc/$r}j0RVtf/G-PM00BSNfC~6G=?Pw+0Rb@P
::/159g2mk=]2t7dg1pt83gaCk2;3j-Ef(hS000BSN2mk=k6M-DcpaA$c=o3J?9{~Vy=@9.0X+,vg=?Pw+0f8_+/159g2mk=]2t7dg1pt83gaCk2;3j-Ef(hS000BSN
::KmqtS=o3J?2]f_9KLH49LI40&1ONaOA3XqZ=?dRJDKUr|X&YZ==?Pw+0YNd5/159g2mk=]2t7dg1pt83gaCk2;3j-Ef(hS000BSNKmqtS=o3J?2]f_9KLH49LI40&
::1ONaOA3XqZ=?dRJDKUr|X&-x]=?Pwa0pT&_/159g2mk=k2t7dg1pt831Ob3j;AVT]VgZ0s00BSNU/-3y=o3Ks9|._lX#fCJDFA@y1pojP?f.?i9{~#M?Ei(hDKUte
::X#xQG2mk=k2th#k/159g00BSN004ke{d?fy]9Morzyn{^5C8zs5LQ6?00JM[XaHas/XuG4$&e[U$bu/3+(M{l$N+e9/XuG9)T2$c$bu/3R{.ED/eq4h;H.cbf.K~L
::$ppxPEac;kLjYhR/eq4h;H(-4;blY7D(,tiSODNE/eq4h;H.cbf.K|#fyo5Of.L0YL/(C^/eq4h;Ix1jf.2/J)FDkXD(,ti000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000002m$~A0&iaJ5C8xG0000000000000000000000000b]x{jmINF-cmQB00RR916aW@g00000AQ1q70RR917yuan00000I1vDV0RR91
::0000000000U^bzX0RR916aW@g00000AQ1q70RR918~^~v00000SP=k#0RR910000000000Xh8sg0RR916aW@g00000AQ1q70RR914ge1T00000co6_A0RR9100000
::000005JCWe0RR916aW@g00000AQ1q70RR914ge1T00000h!FsQ0RR910000000000s6qgM0RR916aW@g00000AQ1q70RR917yuan00000m=OSg0RR910000000000
::AVUCv0RR916aW@g00000AQ1q70RR916aW;f00000xDfz=0RR910000000000kV61~0RR916aW@g00000AQ1q70RR915(#nb00000(=CND0RR9100000000002t+vY
::0RR916aW@g00000AQ1q70RR914ge1T00000=n),b0RR910000000000Xhi]k0RR916aW@g00000AQ1q70RR916aW;f00000^z@hr0RR910000000000ct!w#0RR91
::6aW@g00000AQ1q70RR914ge1T000005E1}[0RR910000000000,hc^?0RR91Pyhe_00000U}OM,0RR91z#Ra90RR9100000000000000000000000000000000000
::00000000000000000000000000000000000AQ&9E0RR913jhEB3jhEB5EuY}0RR913/-NC3/-NC02ly)0RR914FCWD4FCWD[D~7p0RR914gdfE4gdfE/1?XZ0RR91
::4,(oF4,(oF(=(xJ0RR910ssI22LJ#7/1(RY0RR910ssI22LJ#7(=vrI0RR910ssI22LJ#7z!m^20RR910ssI22LJ#7z!w030RR910ssI22LJ#7uoeJ.0RR910ssI2
::2LJ#7pcVjt0RR910ssI22LJ#7kQM.d0RR910ssI22LJ#7pcepu0RR910ssI22LJ#7fEECN0RR910ssI22LJ#7fENIO0RR910{{R32LJ#7a25c70RR910{{R32LJ#7
::U?5,[0RR911ONa42LJ#7Ko;aj0RR911ONa42LJ#7Fc$!T0RR911ONa42LJ#7U={#@0RR911ONa42LJ#7AQu3D0RR911poj52LJ#7P!;4y0RR911poj52LJ#7Ko$Ui
::0RR911][s62LJ#75ElS|0RR912LJ#72LJ#702cs(0RR912LJ#72LJ#7FctuS0RR912LJ#72LJ#7P!|Az0RR912mk/82mk/8[D?1o0RR913IG5A3IG5AuonP/0RR91
::2?;{92?;{9kQV[e0RR912?;{92?;{9a2Ei80RR912?;{92?;{9000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::0000000000000000000000000000000000000000000000000000000000000000000000,l^?=00000_f(gN000006mkFn00000EOGz;00000K5^s600000OmYAK
::00000U~(Ke00000dU5~(00000ka7S300000o]k,H00000ta1PV00000$Z_Mx00000-/RW^00000[]SzG000001aklY00000B69!$00000Kyv]900000o]t?I00000
::RC53T00000UUL8d00000YI6Vp00000baMaz00000escf.00000h/sk{00000l5-q600000taAVW0000000000000000C4~S0000000000N];}J0B_]R0000000000
::000000000000000,l^?=00000_f(gN000006mkFn00000EOGz;00000K5^s600000OmYAK00000U~(Ke00000dU5~(00000ka7S300000o]k,H00000ta1PV00000
::$Z_Mx00000-/RW^00000[]SzG000001aklY00000B69!$00000Kyv]900000o]t?I00000RC53T00000UUL8d00000YI6Vp00000baMaz00000escf.00000h/sk{
::00000l5-q600000taAVW000000000000000ZU9VVaztr!VPb4$RA^Q#VPr#LY/13JbaO]/azt!w0074UPIORmZ,,m2bXI9{bai2DO=WFwa)Ms&U;6WhY+NiubX9I@
::V{c@.Q,@4]Zf5_hc?qjgaz|x!L~LwGVQyq?WdMo,Ok{FQZ))FaY.|7kQv^0UY+NiubU|+(X/XA]X?Ml#fdEWoaz|x!P/zf$Wn]_7WkF;Qa&FRK007MeQgm!oX?Dax
::Z(Yb,WkzXbY.Do+J^1g3Q+P5Tc4cmK002G)Qgm!mVQyq]ZAEwh@,UG9QFUc;c~E6@W]ZzBVQyn)LvM9&bY,e@_~gmMQFUc;c~g0FbY,Q-X?DZy$pun$Y,cA(WkzXb
::Y.Dp)Z(Yb,WdPIyQgm!VY/131VRU6kWnpjtjQ~t!a!-t(Zb[xnXJtldY.LYybZKvHb4z7/005H!Ok{FVb!BpSNo_@gWkzXiWlLpwPjGZ/Z,Bkp{s2yNLu^wzWdLq/
::WNd6MWNd5z5CC([a${|9000XBUw313ZfRp}Z~zVfZDnn3Z-2w?4FGLrZDVkG000jFZDnn9Wpn[l5()B(b8Ka9000UAUw313X=810000pHb9ZoZX?N38UvmHe3/=Cq
::ZDVb40000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000001yBG5C8xGc&q1-n4$mx0AK)B
::hyVZpIG{-NSfFU2c&X=(n4qYjxS-^O,r4d3^[D[)7[/VkIH5@PSfOa4c&g_+n4zelxS_0Q,rDj5^[M},7[{DeV4_rMfTED1prWv&z[pHi/G,!N0HYA2Afqs(K&.Ej
::V54xOfTNJ3prf#)z[yNk/G]+P0HhG4Afzy+K&_Kl0000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
::0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
:embdbin:
function UninstallLicenses($DllPath) {
    $TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2).DefineType(0)
    
    [void]$TB.DefinePInvokeMethod('SLOpen', $DllPath, 22, 1, [int], @([IntPtr].MakeByRefType()), 1, 3)
    [void]$TB.DefinePInvokeMethod('SLGetSLIDList', $DllPath, 22, 1, [int],
        @([IntPtr], [int], [Guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    [void]$TB.DefinePInvokeMethod('SLUninstallLicense', $DllPath, 22, 1, [int], @([IntPtr], [IntPtr]), 1, 3)

    $SPPC = $TB.CreateType()
    $Handle = 0
    [void]$SPPC::SLOpen([ref]$Handle)
    $pnReturnIds = 0
    $ppReturnIds = 0

    if (!$SPPC::SLGetSLIDList($Handle, 0, [ref][Guid]"0ff1ce15-a989-479d-af46-f275c6370663", 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
        foreach ($i in 0..($pnReturnIds - 1)) {
            [void]$SPPC::SLUninstallLicense($Handle, [Int64]$ppReturnIds + [Int64]16 * $i)
        }    
    }
}

$OSPP = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -ErrorAction SilentlyContinue).Path
if ($OSPP) {
    UninstallLicenses ($OSPP + "osppc.dll")
}
UninstallLicenses "sppc.dll"
:embdbin:

:embdxrm:
$wmi = Get-WmiObject $sls
function InstallLicenseFile($Lsc) {
    try {
        $null = $wmi.InstallLicense([IO.File]::ReadAllText($Lsc))
    } catch {
        $host.SetShouldExit($_.Exception.HResult)
    }
}
function InstallLicenseArr($Str) {
    $a = $Str -split ';'
    ForEach ($x in $a) {InstallLicenseFile "$x"}
}
function InstallLicenseDir($Loc) {
    dir $Loc *.xrm-ms -af -s | select -expand FullName | % {InstallLicenseFile "$_"}
}
function ReinstallLicenses() {
    $Oem = "$env:SystemRoot\system32\oem"
    $Spp = "$env:SystemRoot\system32\spp\tokens"
    InstallLicenseDir "$Spp"
    If (Test-Path $Oem) {InstallLicenseDir "$Oem"}
}
:embdxrm:

:sppmgr:
function ExitScript($ExitCode = 0)
{
	Exit $ExitCode
}

if (-Not $PSVersionTable) {
	Write-Host "==== ERROR ====`r`n"
	Write-Host 'Windows PowerShell 1.0 is not supported by this script.'
	ExitScript 1
}

if ($ExecutionContext.SessionState.LanguageMode.value__ -NE 0) {
	Write-Host "==== ERROR ====`r`n"
	Write-Host 'Windows PowerShell is not running in Full Language Mode.'
	ExitScript 1
}

$winbuild = 1
try {
	$winbuild = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$env:SystemRoot\System32\kernel32.dll").FileBuildPart
} catch {
	$winbuild = [int](Get-WmiObject Win32_OperatingSystem).BuildNumber
}

if ($winbuild -EQ 1) {
	Write-Host "==== ERROR ====`r`n"
	Write-Host 'Could not detect Windows build.'
	ExitScript 1
}

if ($winbuild -LT 2600) {
	Write-Host "==== ERROR ====`r`n"
	Write-Host 'This build of Windows is not supported by this script.'
	ExitScript 1
}

$NT6 = $winbuild -GE 6000
$NT7 = $winbuild -GE 7600
$NT9 = $winbuild -GE 9600

$Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$line2 = "============================================================"
$line3 = "____________________________________________________________"

function echoWindows
{
	Write-Host "$line2"
	Write-Host "===                   Windows Status                     ==="
	Write-Host "$line2"
	if (!$All.IsPresent) {Write-Host}
}

function echoOffice
{
	if ($doMSG -EQ 0) {
		return
	}

	if ($All.IsPresent) {Write-Host}
	Write-Host "$line2"
	Write-Host "===                   Office Status                      ==="
	Write-Host "$line2"
	if (!$All.IsPresent) {Write-Host}

	$script:doMSG = 0
}

function strGetRegistry($strKey, $strName)
{
Get-ItemProperty -EA 0 $strKey | select -EA 0 -Expand $strName
}

function CheckOhook
{
	$ohook = 0
	$paths = "${env:ProgramFiles}", "${env:ProgramW6432}", "${env:ProgramFiles(x86)}"

	15, 16 | foreach `
	{
		$A = $_; $paths | foreach `
		{
			if (Test-Path "$($_)$('\Microsoft Office\Office')$($A)$('\sppc*dll')") {$ohook = 1}
		}
	}

	"System", "SystemX86" | foreach `
	{
		$A = $_; "Office 15", "Office" | foreach `
		{
			$B = $_; $paths | foreach `
			{
				if (Test-Path "$($_)$('\Microsoft ')$($B)$('\root\vfs\')$($A)$('\sppc*dll')") {$ohook = 1}
			}
		}
	}

	if ($ohook -EQ 0) {
		return
	}

	if ($All.IsPresent) {Write-Host}
	Write-Host "$line2"
	Write-Host "===                Office Ohook Status                   ==="
	Write-Host "$line2"
	Write-Host
	Write-Host -back 'Black' -fore 'Yellow' 'Ohook for permanent Office activation is installed.'
	Write-Host -back 'Black' -fore 'Yellow' 'You can ignore the below mentioned Office activation status.'
	if (!$All.IsPresent) {Write-Host}
}

#region WMI
function DetectID($strSLP, $strAppId, [ref]$strAppVar)
{
	$fltr = "ApplicationID='$strAppId'"
	if (!$All.IsPresent) {
		$fltr = $fltr + " AND PartialProductKey <> NULL"
	}
	Get-WmiObject $strSLP ID -Filter $fltr -EA 0 | select ID -EA 0 | foreach {
		$strAppVar.Value = 1
	}
}

function GetID($strSLP, $strAppId, $strProperty = "ID")
{
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)
	$IDs = [Collections.ArrayList]@()

	if ($All.IsPresent) {
		$fltr = "ApplicationID='$strAppId' AND PartialProductKey IS NULL"
		$clause = $fltr
		if (-Not $NT5) {
		$clause = $fltr + " AND LicenseDependsOn <> NULL"
		}
		Get-WmiObject $strSLP $strProperty -Filter $clause -EA 0 | select -Expand $strProperty -EA 0 | foreach {$IDs += $_}
		if (-Not $NT5) {
		$clause = $fltr + " AND LicenseDependsOn IS NULL"
		Get-WmiObject $strSLP $strProperty -Filter $clause -EA 0 | select -Expand $strProperty -EA 0 | foreach {$IDs += $_}
		}
	}

	$fltr = "ApplicationID='$strAppId' AND PartialProductKey <> NULL"
	$clause = $fltr
	if (-Not $NT5) {
	$clause = $fltr + " AND LicenseDependsOn <> NULL"
	}
	Get-WmiObject $strSLP $strProperty -Filter $clause -EA 0 | select -Expand $strProperty -EA 0 | foreach {$IDs += $_}
	if (-Not $NT5) {
	$clause = $fltr + " AND LicenseDependsOn IS NULL"
	Get-WmiObject $strSLP $strProperty -Filter $clause -EA 0 | select -Expand $strProperty -EA 0 | foreach {$IDs += $_}
	}

	return $IDs
}

function DetectSubscription {
	if ($null -EQ $objSvc.SubscriptionType -Or $objSvc.SubscriptionType -EQ 120) {
		return
	}

	if ($objSvc.SubscriptionType -EQ 1) {
		$SubMsgType = "Device based"
	} else {
		$SubMsgType = "User based"
	}

	if ($objSvc.SubscriptionStatus -EQ 120) {
		$SubMsgStatus = "Expired"
	} elseif ($objSvc.SubscriptionStatus -EQ 100) {
		$SubMsgStatus = "Disabled"
	} elseif ($objSvc.SubscriptionStatus -EQ 1) {
		$SubMsgStatus = "Active"
	} else {
		$SubMsgStatus = "Not active"
	}

	$SubMsgExpiry = "Unknown"
	if ($objSvc.SubscriptionExpiry) {
		if ($objSvc.SubscriptionExpiry.Contains("unspecified") -EQ $false) {$SubMsgExpiry = $objSvc.SubscriptionExpiry}
	}

	$SubMsgEdition = "Unknown"
	if ($objSvc.SubscriptionEdition) {
		if ($objSvc.SubscriptionEdition.Contains("UNKNOWN") -EQ $false) {$SubMsgEdition = $objSvc.SubscriptionEdition}
	}

	Write-Host
	Write-Host "Subscription information:"
	Write-Host "    Edition: $SubMsgEdition"
	Write-Host "    Type   : $SubMsgType"
	Write-Host "    Status : $SubMsgStatus"
	Write-Host "    Expiry : $SubMsgExpiry"
}

function DetectAvmClient
{
	Write-Host
	Write-Host "Automatic VM Activation client information:"
	if (-Not [String]::IsNullOrEmpty($IAID)) {
		Write-Host "    Guest IAID: $IAID"
	} else {
		Write-Host "    Guest IAID: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($AutomaticVMActivationHostMachineName)) {
		Write-Host "    Host machine name: $AutomaticVMActivationHostMachineName"
	} else {
		Write-Host "    Host machine name: Not Available"
	}
	if ($AutomaticVMActivationLastActivationTime.Substring(0,4) -NE "1601") {
		$EED = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($AutomaticVMActivationLastActivationTime),$null,48).ToString('yyyy-MM-dd hh:mm:ss tt')
		Write-Host "    Activation time: $EED UTC"
	} else {
		Write-Host "    Activation time: Not Available"
	}
	if (-Not [String]::IsNullOrEmpty($AutomaticVMActivationHostDigitalPid2)) {
		Write-Host "    Host Digital PID2: $AutomaticVMActivationHostDigitalPid2"
	} else {
		Write-Host "    Host Digital PID2: Not Available"
	}
}

function DetectKmsHost
{
	if ($Vista -Or $NT5) {
		$KeyManagementServiceListeningPort = strGetRegistry $SLKeyPath "KeyManagementServiceListeningPort"
		$KeyManagementServiceDnsPublishing = strGetRegistry $SLKeyPath "DisableDnsPublishing"
		$KeyManagementServiceLowPriority = strGetRegistry $SLKeyPath "EnableKmsLowPriority"
		if (-Not $KeyManagementServiceDnsPublishing) {$KeyManagementServiceDnsPublishing = "TRUE"}
		if (-Not $KeyManagementServiceLowPriority) {$KeyManagementServiceLowPriority = "FALSE"}
	} else {
		$KeyManagementServiceListeningPort = $objSvc.KeyManagementServiceListeningPort
		$KeyManagementServiceDnsPublishing = $objSvc.KeyManagementServiceDnsPublishing
		$KeyManagementServiceLowPriority = $objSvc.KeyManagementServiceLowPriority
	}

	if (-Not $KeyManagementServiceListeningPort) {$KeyManagementServiceListeningPort = 1688}
	if ($KeyManagementServiceDnsPublishing -EQ "TRUE") {
		$KeyManagementServiceDnsPublishing = "Enabled"
	} else {
		$KeyManagementServiceDnsPublishing = "Disabled"
	}
	if ($KeyManagementServiceLowPriority -EQ "TRUE") {
		$KeyManagementServiceLowPriority = "Low"
	} else {
		$KeyManagementServiceLowPriority = "Normal"
	}

	Write-Host
	Write-Host "Key Management Service host information:"
	Write-Host "    Current count: $KeyManagementServiceCurrentCount"
	Write-Host "    Listening on Port: $KeyManagementServiceListeningPort"
	Write-Host "    DNS publishing: $KeyManagementServiceDnsPublishing"
	Write-Host "    KMS priority: $KeyManagementServiceLowPriority"
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceTotalRequests)) {
		Write-Host
		Write-Host "Key Management Service cumulative requests received from clients:"
		Write-Host "    Total: $KeyManagementServiceTotalRequests"
		Write-Host "    Failed: $KeyManagementServiceFailedRequests"
		Write-Host "    Unlicensed: $KeyManagementServiceUnlicensedRequests"
		Write-Host "    Licensed: $KeyManagementServiceLicensedRequests"
		Write-Host "    Initial grace period: $KeyManagementServiceOOBGraceRequests"
		Write-Host "    Expired or Hardware out of tolerance: $KeyManagementServiceOOTGraceRequests"
		Write-Host "    Non-genuine grace period: $KeyManagementServiceNonGenuineGraceRequests"
		Write-Host "    Notification: $KeyManagementServiceNotificationRequests"
	}
}

function DetectKmsClient
{
	if ($null -NE $VLActivationTypeEnabled) {Write-Host "Configured Activation Type: $($VLActTypes[$VLActivationTypeEnabled])"}
	Write-Host
	if ($LicenseStatus -NE 1) {
		Write-Host "Please activate the product in order to update KMS client information values."
		return
	}

	if ($Vista) {
		$KeyManagementServicePort = strGetRegistry $SLKeyPath "KeyManagementServicePort"
		$DiscoveredKeyManagementServiceMachineName = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServiceName"
		$DiscoveredKeyManagementServiceMachinePort = strGetRegistry $NSKeyPath "DiscoveredKeyManagementServicePort"
	}

	if ([String]::IsNullOrEmpty($KeyManagementServiceMachine)) {
		$KmsReg = $null
	} else {
		if (-Not $KeyManagementServicePort) {$KeyManagementServicePort = 1688}
		$KmsReg = "Registered KMS machine name: ${KeyManagementServiceMachine}:${KeyManagementServicePort}"
	}

	if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceMachineName)) {
		$KmsDns = "DNS auto-discovery: KMS name not available"
		if ($Vista -And -Not $Admin) {$KmsDns = "DNS auto-discovery: Run the script as administrator to retrieve info"}
	} else {
		if (-Not $DiscoveredKeyManagementServiceMachinePort) {$DiscoveredKeyManagementServiceMachinePort = 1688}
		$KmsDns = "KMS machine name from DNS: ${DiscoveredKeyManagementServiceMachineName}:${DiscoveredKeyManagementServiceMachinePort}"
	}

	if ($null -NE $objSvc.KeyManagementServiceHostCaching) {
		if ($objSvc.KeyManagementServiceHostCaching -EQ "TRUE") {
			$KeyManagementServiceHostCaching = "Enabled"
		} else {
			$KeyManagementServiceHostCaching = "Disabled"
		}
	}

	Write-Host "Key Management Service client information:"
	Write-Host "    Client Machine ID (CMID): $($objSvc.ClientMachineID)"
	if ($null -EQ $KmsReg) {
		Write-Host "    $KmsDns"
		Write-Host "    Registered KMS machine name: KMS name not available"
	} else {
		Write-Host "    $KmsReg"
	}
	if ($null -NE $DiscoveredKeyManagementServiceMachineIpAddress) {Write-Host "    KMS machine IP address: $DiscoveredKeyManagementServiceMachineIpAddress"}
	Write-Host "    KMS machine extended PID: $KeyManagementServiceProductKeyID"
	Write-Host "    Activation interval: $VLActivationInterval minutes"
	Write-Host "    Renewal interval: $VLRenewalInterval minutes"
	if ($null -NE $KeyManagementServiceHostCaching) {Write-Host "    KMS host caching: $KeyManagementServiceHostCaching"}
	if (-Not [String]::IsNullOrEmpty($KeyManagementServiceLookupDomain)) {Write-Host "    KMS SRV record lookup domain: $KeyManagementServiceLookupDomain"}
}

function GetResult($strSLP, $strSLS, $strID)
{
	try {$objPrd = Get-WmiObject $strSLP -Filter "ID='$strID'" -EA 1} catch {return}
	$objPrd | select -Expand Properties -EA 0 | foreach {
		if (-Not [String]::IsNullOrEmpty($_.Value)) {set $_.Name $_.Value}
	}

	$winID = ($ApplicationID -EQ $winApp)
	$winPR = ($winID -And -Not $LicenseIsAddon)
	$Vista = ($winID -And $NT6 -And -Not $NT7)
	$NT5 = ($strSLP -EQ $wslp -And $winbuild -LT 6001)

	if ($Description | Select-String "VOLUME_KMSCLIENT") {$cKmsClient = 1; $_mTag = "Volume"}
	if ($Description | Select-String "TIMEBASED_") {$cTblClient = 1; $_mTag = "Timebased"}
	if ($Description | Select-String "VIRTUAL_MACHINE_ACTIVATION") {$cAvmClient = 1; $_mTag = "Automatic VM"}
	if ($null -EQ $cKmsClient) {
		if ($Description | Select-String "VOLUME_KMS") {$cKmsHost = 1}
	}

	$_gpr = [Math]::Round($GracePeriodRemaining/1440)
	if ($_gpr -GT 0) {
		$_xpr = [DateTime]::Now.addMinutes($GracePeriodRemaining).ToString('yyyy-MM-dd hh:mm:ss tt')
	}

	if ($null -EQ $LicenseStatusReason) {$LicenseStatusReason = -1}
	$LicenseReason = '0x{0:X}' -f $LicenseStatusReason
	$LicenseMsg = "Time remaining: $GracePeriodRemaining minute(s) ($_gpr day(s))"
	if ($LicenseStatus -EQ 0) {
		$LicenseInf = "Unlicensed"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 1) {
		$LicenseInf = "Licensed"
		$LicenseMsg = $null
		if ($GracePeriodRemaining -EQ 0) {
			if ($winPR) {$ExpireMsg = "The machine is permanently activated."} else {$ExpireMsg = "The product is permanently activated."}
		} else {
			$LicenseMsg = "$_mTag activation expiration: $GracePeriodRemaining minute(s) ($_gpr day(s))"
			if ($null -NE $_xpr) {$ExpireMsg = "$_mTag activation will expire $_xpr"}
		}
	}
	if ($LicenseStatus -EQ 2) {
		$LicenseInf = "Initial grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Initial grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 3) {
		$LicenseInf = "Additional grace period (KMS license expired or hardware out of tolerance)"
		if ($null -NE $_xpr) {$ExpireMsg = "Additional grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 4) {
		$LicenseInf = "Non-genuine grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Non-genuine grace period ends $_xpr"}
	}
	if ($LicenseStatus -EQ 5 -And -Not $NT5) {
		$LicenseInf = "Notification"
		$LicenseMsg = "Notification Reason: $LicenseReason"
		if ($LicenseReason -EQ "0xC004F200") {$LicenseMsg = $LicenseMsg + " (non-genuine)."}
		if ($LicenseReason -EQ "0xC004F009") {$LicenseMsg = $LicenseMsg + " (grace time expired)."}
	}
	if ($LicenseStatus -GT 5 -Or ($LicenseStatus -GT 4 -And $NT5)) {
		$LicenseInf = "Unknown"
		$LicenseMsg = $null
	}
	if ($LicenseStatus -EQ 6 -And -Not $Vista -And -Not $NT5) {
		$LicenseInf = "Extended grace period"
		if ($null -NE $_xpr) {$ExpireMsg = "Extended grace period ends $_xpr"}
	}

	if ($winPR -And $PartialProductKey -And -Not $NT9) {
		$dp4 = Get-ItemProperty -EA 0 "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" | select -EA 0 -Expand DigitalProductId4
		if ($null -NE $dp4) {
			$ProductKeyChannel = ([System.Text.Encoding]::Unicode.GetString($dp4, 1016, 128)).Trim([char]$null)
		}
	}

	if ($All.IsPresent) {Write-Host}
	Write-Host "Name: $Name"
	Write-Host "Description: $Description"
	Write-Host "Activation ID: $ID"
	if ($null -NE $ProductKeyID) {Write-Host "Extended PID: $ProductKeyID"}
	if ($null -NE $OfflineInstallationId -And $IID.IsPresent) {Write-Host "Installation ID: $OfflineInstallationId"}
	if ($null -NE $ProductKeyChannel) {Write-Host "Product Key Channel: $ProductKeyChannel"}
	if ($null -NE $PartialProductKey) {Write-Host "Partial Product Key: $PartialProductKey"} else {Write-Host "Product Key: Not installed"}
	Write-Host "License Status: $LicenseInf"
	if ($null -NE $LicenseMsg) {Write-Host "$LicenseMsg"}
	if ($LicenseStatus -NE 0 -And $EvaluationEndDate.Substring(0,4) -NE "1601") {
		$EED = [DateTime]::Parse([Management.ManagementDateTimeConverter]::ToDateTime($EvaluationEndDate),$null,48).ToString('yyyy-MM-dd hh:mm:ss tt')
		Write-Host "Evaluation End Date: $EED UTC"
	}

	if ($winID -And $null -NE $cAvmClient -And $null -NE $PartialProductKey) {
		DetectAvmClient
	}

	$chkSub = ($winPR -And $cSub)

	$chkSLS = ($null -NE $PartialProductKey) -And ($null -NE $cKmsClient -Or $null -NE $cKmsHost -Or $chkSub)

	if (!$chkSLS) {
		if ($null -NE $ExpireMsg) {Write-Host; Write-Host "    $ExpireMsg"}
		return
	}

	$objSvc = Get-WmiObject $strSLS -EA 0

	if ($Vista) {
		$objSvc | select -Expand Properties -EA 0 | foreach {
			if (-Not [String]::IsNullOrEmpty($_.Value)) {set $_.Name $_.Value}
		}
	}

	if ($strSLS -EQ $wsls -And $NT9) {
		if ([String]::IsNullOrEmpty($DiscoveredKeyManagementServiceMachineIpAddress)) {
			$DiscoveredKeyManagementServiceMachineIpAddress = "not available"
		}
	}

	if ($null -NE $cKmsHost -And $IsKeyManagementServiceMachine -GT 0) {
		DetectKmsHost
	}

	if ($null -NE $cKmsClient) {
		DetectKmsClient
	}

	if ($null -NE $ExpireMsg) {Write-Host; Write-Host "    $ExpireMsg"}

	if ($chkSub) {
		DetectSubscription
	}

}
#endregion

#region vNextDiag
if ($PSVersionTable.PSVersion.Major -Lt 3)
{
	function ConvertFrom-Json
	{
		[CmdletBinding()]
		Param(
			[Parameter(ValueFromPipeline=$true)][Object]$item
		)
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
		$psjs = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		Return ,$psjs.DeserializeObject($item)
	}
	function ConvertTo-Json
	{
		[CmdletBinding()]
		Param(
			[Parameter(ValueFromPipeline=$true)][Object]$item
		)
		[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
		$psjs = New-Object System.Web.Script.Serialization.JavaScriptSerializer
		Return $psjs.Serialize($item)
	}
}

function PrintModePerPridFromRegistry
{
	$vNextRegkey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext"
	$vNextPrids = Get-Item -Path $vNextRegkey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty 'property' -ErrorAction SilentlyContinue | Where-Object -FilterScript {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"}
	If ($null -Eq $vNextPrids)
	{
		Write-Host
		Write-Host "No registry keys found."
		Return
	}
	Write-Host
	$vNextPrids | ForEach `
	{
		$mode = (Get-ItemProperty -Path $vNextRegkey -Name $_).$_
		Switch ($mode)
		{
			2 { $mode = "vNext"; Break }
			3 { $mode = "Device"; Break }
			Default { $mode = "Legacy"; Break }
		}
		Write-Host $_ = $mode
	}
}

function PrintSharedComputerLicensing
{
	$scaRegKey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$scaValue = Get-ItemProperty -Path $scaRegKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	$scaRegKey2 = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing"
	$scaValue2 = Get-ItemProperty -Path $scaRegKey2 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	$scaPolicyKey = "HKLM:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing"
	$scaPolicyValue = Get-ItemProperty -Path $scaPolicyKey -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SharedComputerLicensing" -ErrorAction SilentlyContinue
	If ($null -Eq $scaValue -And $null -Eq $scaValue2 -And $null -Eq $scaPolicyValue)
	{
		Write-Host
		Write-Host "No registry keys found."
		Return
	}
	$scaModeValue = $scaValue -Or $scaValue2 -Or $scaPolicyValue
	If ($scaModeValue -Eq 0)
	{
		$scaMode = "Disabled"
	}
	If ($scaModeValue -Eq 1)
	{
		$scaMode = "Enabled"
	}
	Write-Host
	Write-Host "Status:" $scaMode
	Write-Host
	$tokenFiles = $null
	$tokenPath = "${env:LOCALAPPDATA}\Microsoft\Office\16.0\Licensing"
	If (Test-Path $tokenPath)
	{
		$tokenFiles = Get-ChildItem -Path $tokenPath -Filter "*authString*" -Recurse | Where-Object { !$_.PSIsContainer }
	}
	If ($null -Eq $tokenFiles)
	{
		Write-Host "No tokens found."
		Return
	}
	If ($tokenFiles.Length -Eq 0)
	{
		Write-Host "No tokens found."
		Return
	}
	$tokenFiles | ForEach `
	{
		$tokenParts = (Get-Content -Encoding Unicode -Path $_.FullName).Split('_')
		$output = New-Object PSObject
		$output | Add-Member 8 'ACID' $tokenParts[0];
		$output | Add-Member 8 'User' $tokenParts[3];
		$output | Add-Member 8 'NotBefore' $tokenParts[4];
		$output | Add-Member 8 'NotAfter' $tokenParts[5];
		Write-Output $output
	}
}

function PrintLicensesInformation
{
	Param(
		[ValidateSet("NUL", "Device")]
		[String]$mode
	)
	If ($mode -Eq "NUL")
	{
		$licensePath = "${env:LOCALAPPDATA}\Microsoft\Office\Licenses"
	}
	ElseIf ($mode -Eq "Device")
	{
		$licensePath = "${env:PROGRAMDATA}\Microsoft\Office\Licenses"
	}
	$licenseFiles = $null
	If (Test-Path $licensePath)
	{
		$licenseFiles = Get-ChildItem -Path $licensePath -Recurse | Where-Object { !$_.PSIsContainer }
	}
	If ($null -Eq $licenseFiles)
	{
		Write-Host
		Write-Host "No licenses found."
		Return
	}
	If ($licenseFiles.Length -Eq 0)
	{
		Write-Host
		Write-Host "No licenses found."
		Return
	}
	$licenseFiles | ForEach `
	{
		$license = (Get-Content -Encoding Unicode $_.FullName | ConvertFrom-Json).License
		$decodedLicense = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($license)) | ConvertFrom-Json
		$licenseType = $decodedLicense.LicenseType
		If ($null -Ne $decodedLicense.ExpiresOn)
		{
			$expiry = [System.DateTime]::Parse($decodedLicense.ExpiresOn, $null, 'AdjustToUniversal')
		}
		Else
		{
			$expiry = New-Object System.DateTime
		}
		$licenseState = "Grace"
		If ((Get-Date) -Gt (Get-Date $decodedLicense.Metadata.NotAfter))
		{
			$licenseState = "RFM"
		}
		ElseIf ((Get-Date) -Lt (Get-Date $expiry))
		{
			$licenseState = "Licensed"
		}
		$output = New-Object PSObject
		$output | Add-Member 8 'File' $_.PSChildName;
		$output | Add-Member 8 'Version' $_.Directory.Name;
		$output | Add-Member 8 'Type' "User|${licenseType}";
		$output | Add-Member 8 'Product' $decodedLicense.ProductReleaseId;
		$output | Add-Member 8 'Acid' $decodedLicense.Acid;
		If ($mode -Eq "Device") { $output | Add-Member 8 'DeviceId' $decodedLicense.Metadata.DeviceId; }
		$output | Add-Member 8 'LicenseState' $licenseState;
		$output | Add-Member 8 'EntitlementStatus' $decodedLicense.Status;
		$output | Add-Member 8 'EntitlementExpiration' ("N/A", $decodedLicense.ExpiresOn)[!($null -eq $decodedLicense.ExpiresOn)];
		$output | Add-Member 8 'ReasonCode' ("N/A", $decodedLicense.ReasonCode)[!($null -eq $decodedLicense.ReasonCode)];
		$output | Add-Member 8 'NotBefore' $decodedLicense.Metadata.NotBefore;
		$output | Add-Member 8 'NotAfter' $decodedLicense.Metadata.NotAfter;
		$output | Add-Member 8 'NextRenewal' $decodedLicense.Metadata.RenewAfter;
		$output | Add-Member 8 'TenantId' ("N/A", $decodedLicense.Metadata.TenantId)[!($null -eq $decodedLicense.Metadata.TenantId)];
		#$output.PSObject.Properties | foreach { $ht = @{} } { $ht[$_.Name] = $_.Value } { $output = $ht | ConvertTo-Json }
		Write-Output $output
	}
}

function vNextDiagRun
{
	$fNUL = ([IO.Directory]::Exists("${env:LOCALAPPDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:LOCALAPPDATA}\Microsoft\Office\Licenses", "*", 1).Length -GE 0)
	$fDev = ([IO.Directory]::Exists("${env:PROGRAMDATA}\Microsoft\Office\Licenses")) -and ([IO.Directory]::GetFiles("${env:PROGRAMDATA}\Microsoft\Office\Licenses", "*", 1).Length -GE 0)
	$rPID = $null -NE (GP "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LicensingNext" -EA 0 | select -Expand 'property' -EA 0 | where -Filter {$_.ToLower() -like "*retail" -or $_.ToLower() -like "*volume"})
	$rSCA = $null -NE (GP "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -EA 0 | select -Expand "SharedComputerLicensing" -EA 0)
	$rSCL = $null -NE (GP "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\Licensing" -EA 0 | select -Expand "SharedComputerLicensing" -EA 0)

	if (($fNUL -Or $fDev -Or $rPID -Or $rSCA -Or $rSCL) -EQ $false) {
		Return
	}

	if ($All.IsPresent) {Write-Host}
	Write-Host "$line2"
	Write-Host "===                  Office vNext Status                 ==="
	Write-Host "$line2"
	Write-Host
	Write-Host "========== Mode per ProductReleaseId =========="
	PrintModePerPridFromRegistry
	Write-Host
	Write-Host "========== Shared Computer Licensing =========="
	PrintSharedComputerLicensing
	Write-Host
	Write-Host "========== vNext licenses ==========="
	PrintLicensesInformation -Mode "NUL"
	Write-Host
	Write-Host "========== Device licenses =========="
	PrintLicensesInformation -Mode "Device"
	Write-Host "$line3"
	Write-Host
}
#endregion

#region clic

<#
;;; Source: https://github.com/asdcorp/clic
;;; Powershell port: abbodi1406

Copyright 2023 asdcorp

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

function BoolToWStr($bVal) {
	("TRUE", "FALSE")[!$bVal]
}

function InitializePInvoke {
	$Marshal = [System.Runtime.InteropServices.Marshal]
	$Module = [AppDomain]::CurrentDomain.DefineDynamicAssembly((Get-Random), 'Run').DefineDynamicModule((Get-Random))

	$Class = $Module.DefineType('NativeMethods', 'Public, Abstract, Sealed, BeforeFieldInit', [Object], 0)
	$Class.DefinePInvokeMethod('SLIsWindowsGenuineLocal', 'slc.dll', 'Public, Static', 'Standard', [Int32], @([UInt32].MakeByRefType()), 'Winapi', 'Unicode').SetImplementationFlags('PreserveSig')
	$Class.DefinePInvokeMethod('SLGetWindowsInformationDWORD', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType()), 1, 3).SetImplementationFlags(128)
	$Class.DefinePInvokeMethod('SLGetWindowsInformation', 'slc.dll', 22, 1, [Int32], @([String], [UInt32].MakeByRefType(), [UInt32].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)

	if ($DllSubscription) {
		$Class.DefinePInvokeMethod('ClipGetSubscriptionStatus', 'Clipc.dll', 22, 1, [Int32], @([IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
		$Struct = $Class.DefineNestedType('SubStatus', 'NestedPublic, SequentialLayout, Sealed, BeforeFieldInit', [ValueType], 0)
		[void]$Struct.DefineField('dwEnabled', [UInt32], 'Public')
		[void]$Struct.DefineField('dwSku', [UInt32], 6)
		[void]$Struct.DefineField('dwState', [UInt32], 6)
		$SubStatus = $Struct.CreateType()
	}

	$Win32 = $Class.CreateType()
}

function InitializeDigitalLicenseCheck {
	$CAB = [System.Reflection.Emit.CustomAttributeBuilder]

	$ICom = $Module.DefineType('EUM.IEUM', 'Public, Interface, Abstract, Import')
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.ComImportAttribute].GetConstructor(@()), @()))
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.GuidAttribute].GetConstructor(@([String])), @('F2DCB80D-0670-44BC-9002-CD18688730AF')))
	$ICom.SetCustomAttribute($CAB::new([System.Runtime.InteropServices.InterfaceTypeAttribute].GetConstructor(@([Int16])), @([Int16]1)))

	1..4 | % { [void]$ICom.DefineMethod('VF'+$_, 'Public, Virtual, HideBySig, NewSlot, Abstract', 'Standard, HasThis', [Void], @()) }
	[void]$ICom.DefineMethod('AcquireModernLicenseForWindows', 1478, 33, [Int32], @([Int32], [Int32].MakeByRefType()))

	$IEUM = $ICom.CreateType()
}

function PrintStateData {
	$pwszStateData = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-Action-StateData",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pwszStateData
	)) {
		return $FALSE
	}

	[string[]]$pwszStateString = $Marshal::PtrToStringUni($pwszStateData) -replace ";", "`n    "
	Write-Host "    $pwszStateString"

	$Marshal::FreeHGlobal($pwszStateData)
	return $TRUE
}

function PrintLastActivationHRresult {
	$pdwLastHResult = 0
	$cbSize = 0

	if ($Win32::SLGetWindowsInformation(
		"Security-SPP-LastWindowsActivationHResult",
		[ref]$null,
		[ref]$cbSize,
		[ref]$pdwLastHResult
	)) {
		return $FALSE
	}

	Write-Host ("    LastActivationHResult=0x{0:x8}" -f $Marshal::ReadInt32($pdwLastHResult))

	$Marshal::FreeHGlobal($pdwLastHResult)
	return $TRUE
}

function PrintIsWindowsGenuine {
	$dwGenuine = 0
	$ppwszGenuineStates = @(
		"SL_GEN_STATE_IS_GENUINE",
		"SL_GEN_STATE_INVALID_LICENSE",
		"SL_GEN_STATE_TAMPERED",
		"SL_GEN_STATE_OFFLINE",
		"SL_GEN_STATE_LAST"
	)

	if ($Win32::SLIsWindowsGenuineLocal([ref]$dwGenuine)) {
		return $FALSE
	}

	if ($dwGenuine -lt 5) {
		Write-Host ("    IsWindowsGenuine={0}" -f $ppwszGenuineStates[$dwGenuine])
	} else {
		Write-Host ("    IsWindowsGenuine={0}" -f $dwGenuine)
	}

	return $TRUE
}

function PrintDigitalLicenseStatus {
	try {
		. InitializeDigitalLicenseCheck
		$ComObj = New-Object -Com EditionUpgradeManagerObj.EditionUpgradeManager
	} catch {
		return $FALSE
	}

	$parameters = 1, $null

	if ([EUM.IEUM].GetMethod("AcquireModernLicenseForWindows").Invoke($ComObj, $parameters)) {
		return $FALSE
	}

	$dwReturnCode = $parameters[1]
	[bool]$bDigitalLicense = $FALSE

	$bDigitalLicense = (($dwReturnCode -ge 0) -and ($dwReturnCode -ne 1))
	Write-Host ("    IsDigitalLicense={0}" -f (BoolToWStr $bDigitalLicense))

	return $TRUE
}

function PrintSubscriptionStatus {
	$dwSupported = 0

	if ($winbuild -ge 15063) {
		$pwszPolicy = "ConsumeAddonPolicySet"
	} else {
		$pwszPolicy = "Allow-WindowsSubscription"
	}

	if ($Win32::SLGetWindowsInformationDWORD($pwszPolicy, [ref]$dwSupported)) {
		return $FALSE
	}

	Write-Host ("    SubscriptionSupportedEdition={0}" -f (BoolToWStr $dwSupported))

	$pStatus = $Marshal::AllocHGlobal($Marshal::SizeOf([Type]$SubStatus))
	if ($Win32::ClipGetSubscriptionStatus([ref]$pStatus)) {
		return $FALSE
	}

	$sStatus = [Activator]::CreateInstance($SubStatus)
	$sStatus = $Marshal::PtrToStructure($pStatus, [Type]$SubStatus)
	$Marshal::FreeHGlobal($pStatus)

	Write-Host ("    SubscriptionEnabled={0}" -f (BoolToWStr $sStatus.dwEnabled))

	if ($sStatus.dwEnabled -eq 0) {
		return $TRUE
	}

	Write-Host ("    SubscriptionSku={0}" -f $sStatus.dwSku)
	Write-Host ("    SubscriptionState={0}" -f $sStatus.dwState)

	return $TRUE
}

function ClicRun
{
	if ($All.IsPresent) {Write-Host}
	Write-Host "Client Licensing Check information:"

	$null = PrintStateData
	$null = PrintLastActivationHRresult
	$null = PrintIsWindowsGenuine

	if ($DllDigital) {
		$null = PrintDigitalLicenseStatus
	}

	if ($DllSubscription) {
		$null = PrintSubscriptionStatus
	}

	Write-Host "$line3"
	if (!$All.IsPresent) {Write-Host}
}
#endregion

$Host.UI.RawUI.WindowTitle = "Check Activation Status"

$SysPath = "$env:SystemRoot\System32"
if (Test-Path "$env:SystemRoot\Sysnative\reg.exe") {
	$SysPath = "$env:SystemRoot\Sysnative"
}

$wslp = "SoftwareLicensingProduct"
$wsls = "SoftwareLicensingService"
$oslp = "OfficeSoftwareProtectionProduct"
$osls = "OfficeSoftwareProtectionService"
$winApp = "55c92734-d682-4d71-983e-d6ec3f16059f"
$o14App = "59a52881-a989-479d-af46-f275c6370663"
$o15App = "0ff1ce15-a989-479d-af46-f275c6370663"
$cSub = ($winbuild -GE 19041) -And (Select-String -Path "$SysPath\wbem\sppwmi.mof" -Encoding unicode -Pattern "SubscriptionType")
$DllDigital = ($winbuild -GE 14393) -And (Test-Path "$SysPath\EditionUpgradeManagerObj.dll")
$DllSubscription = ($winbuild -GE 14393) -And (Test-Path "$SysPath\Clipc.dll")
$VLActTypes = @("All", "AD", "KMS", "Token")
$SLKeyPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"
$NSKeyPath = "Registry::HKEY_USERS\S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SL"

'cW1nd0ws', 'c0ff1ce15', 'c0ff1ce14', 'ospp14', 'ospp15' | foreach {set $_ $null}

$OsppHook = 1
try {gsv osppsvc -EA 1 | Out-Null} catch {$OsppHook = 0}

if ($NT7 -Or -Not $NT6) {
	try {sasv sppsvc -EA 1} catch {}
}
else
{
	try {sasv slsvc -EA 1} catch {}
}

DetectID $wslp $winApp ([ref]$cW1nd0ws)
DetectID $wslp $o15App ([ref]$c0ff1ce15)
DetectID $wslp $o14App ([ref]$c0ff1ce14)

if ($OsppHook -NE 0) {
	try {sasv osppsvc -EA 1} catch {}
	DetectID $oslp $o15App ([ref]$ospp15)
	DetectID $oslp $o14App ([ref]$ospp14)
}

if ($null -NE $cW1nd0ws)
{
	echoWindows
	GetID $wslp $winApp | foreach -EA 1 {
	GetResult $wslp $wsls $_
	Write-Host "$line3"
	if (!$All.IsPresent) {Write-Host}
	}
}
elseif ($NT6)
{
	echoWindows
	Write-Host
	Write-Host "Error: product key not found."
}

if ($winbuild -GE 9200) {
	. InitializePInvoke
	ClicRun
}

if ($c0ff1ce15 -Or $ospp15) {
	CheckOhook
}

$doMSG = 1

if ($null -NE $c0ff1ce15) {
	echoOffice
	GetID $wslp $o15App | foreach -EA 1 {
	GetResult $wslp $wsls $_
	Write-Host "$line3"
	if (!$All.IsPresent) {Write-Host}
	}
}

if ($null -NE $c0ff1ce14) {
	echoOffice
	GetID $wslp $o14App | foreach -EA 1 {
	GetResult $wslp $wsls $_
	Write-Host "$line3"
	if (!$All.IsPresent) {Write-Host}
	}
}

if ($null -NE $ospp15) {
	echoOffice
	GetID $oslp $o15App | foreach -EA 1 {
	GetResult $oslp $osls $_
	Write-Host "$line3"
	if (!$All.IsPresent) {Write-Host}
	}
}

if ($null -NE $ospp14) {
	echoOffice
	GetID $oslp $o14App | foreach -EA 1 {
	GetResult $oslp $osls $_
	Write-Host "$line3"
	if (!$All.IsPresent) {Write-Host}
	}
}

if ($NT7) {
	vNextDiagRun
}
:sppmgr:

:spptask:
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.3" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Source>Microsoft Corporation</Source>
    <Author>Microsoft Corporation</Author>
    <Version>1.0</Version>
    <Description>This task restarts the Software Protection Platform service when user logon occurs</Description>
    <URI>\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger</URI>
    <SecurityDescriptor>D:P(A;;FA;;;SY)(A;;FA;;;BA)(A;;FRFW;;;S-1-5-80-123231216-2592883651-3715271367-3753151631-4175906628)(A;;FR;;;S-1-5-4)</SecurityDescriptor>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="InteractiveUser">
      <GroupId>S-1-5-4</GroupId>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>false</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>true</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="InteractiveUser">
    <ComHandler>
      <ClassId>{B1AEBB5D-EAD9-4476-B375-9C3ED9F32AFC}</ClassId>
      <Data>logon</Data>
    </ComHandler>
  </Actions>
</Task>
:spptask:

:readme:
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>KMS_VL_ALL_AIO</title>
    <style>
        #nav {
            position: absolute;
            top: 0;
            left: 0;
            bottom: 0;
            width: 220px;
            overflow: auto;
        }

        main {
            position: fixed;
            top: 0;
            left: 220px;
            right: 0;
            bottom: 0;
            overflow: auto;
        }

        .innertube {
            margin: 15px;
        }

        * html main {
            height: 100%;
            width: 100%;
        }

        td, h1, h2, h3, h4, h5, p, ul, ol, li {
            page-break-inside: avoid; 
        }
    </style>
  </head>
  <body>
    <main>
        <div class="innertube">

            <h1 id="Overview">KMS_VL_ALL_AIO - Smart Activation Script</h1>
    <ul>
      <li>A standalone batch script to automate the activation of supported Windows and Office products using local KMS server emulator or an external server.</li>
    </ul>
    <ul>
      <li>Designed to be unattended and smart enough not to override the permanent activation of products (Windows or Office),<br />
      only non-activated products will be KMS-activated (if supported).</li>
    </ul>
    <ul>
      <li>The ultimate feature of this solution when installed, will provide 24/7 activation, whenever the system itself requests it (renewal, reactivation, hardware change, Edition upgrade, new Office...), without needing interaction from the user.</li>
    </ul>
    <ul>
      <li>Some security programs will report infected files due to KMS emulating (see source code near the end),<br />
      this is false-positive, as long as you download the file from the trusted Home Page.</li>
    </ul>
    <ul>
      <li>Home Page:<br />
      <a href="https://forums.mydigitallife.net/posts/838808/" target="_blank">https://forums.mydigitallife.net/posts/838808/</a><br />
      Backup links:<br />
      <a href="https://github.com/abbodi1406/KMS_VL_ALL_AIO" target="_blank">https://github.com/abbodi1406/KMS_VL_ALL_AIO</a><br />
      <a href="https://pastebin.com/cpdmr6HZ" target="_blank">https://pastebin.com/cpdmr6HZ</a><br />
      <a href="https://rentry.co/KMS_VL_ALL" target="_blank">https://rentry.co/KMS_VL_ALL</a></li>
    </ul>
            <hr />
            <br />

            <h2 id="AIO">AIO vs. Traditional</h2>
    <p>The KMS_VL_ALL_AIO fork has these differences and extra features compared to the traditional KMS_VL_ALL:</p>
    <ul>
      <li>Portable all-in-one script, easier to move and distribute alone.</li>
    </ul>
    <ul>
      <li>All options and configurations are accessed via easy-to-use menu.</li>
    </ul>
    <ul>
      <li>Combine all the functions of the traditional scripts (Activate, AutoRenewal-Setup, Check-Activation-Status, setupcomplete).</li>
    </ul>
    <ul>
      <li>Required binary files are embedded in the script (including ReadMeAIO.html itself), using ascii encoder by AveYo.</li>
    </ul>
    <ul>
      <li>The needed files get extracted (decoded) later on-demand, via Windows PowerShell.</li>
    </ul>
    <ul>
      <li>Simple text colorization for some menu options (for easier differentiation).</li>
    </ul>
    <ul>
      <li>Auto administrator elevation request.</li>
    </ul>
            <hr />
            <br />

            <h2 id="How">How does it work?</h2>
    <ul>
      <li>Key Management Service (KMS) is a genuine activation method provided by Microsoft for volume licensing customers (organizations, schools or governments).<br />
      The machines in those environments (called KMS clients) activate via the environment KMS host server (authorized Microsoft's licensing key), not via Microsoft activation servers.
      <div>For more info, see <a href="https://www.microsoft.com/Licensing/servicecenter/Help/FAQDetails.aspx?id=201#215" target="_blank">here</a> and <a href="https://technet.microsoft.com/en-us/library/ee939272(v=ws.10).aspx#kms-overview" target="_blank">here</a>.</div></li>
    </ul>
    <ul>
      <li>By design, the KMS activation period lasts up to <strong>180 Days</strong> (6 Months) at max, with the ability to renew and reinstate the period at any time.<br />
      With the proper auto renewal configuration, it will be a continuous activation (essentially permanent).</li>
    </ul>
    <ul>
      <li>KMS Emulators (server and client) are sophisticated tools based on the reversed engineered KMS protocol.<br />
      It mimics the KMS server/client communications, and provide a clean activation for the supported KMS clients, without altering or hacking any system files integrity.</li>
    </ul>
    <ul>
      <li>Updates for Windows or Office do not affect or block KMS activation, only a new KMS protocol will not work with the local emulator.</li>
    </ul>
    <ul>
      <li>The mechanism of <strong>SppExtComObjPatcher</strong> makes it act as a ready-on-request KMS server, providing instant activation without external scheduled tasks or manual intervention.<br />
      Including auto renewal, auto activation of volume Office afterward, reactivation because of hardware change, date change, windows or office edition change... etc.
      <div>On Windows 7, later installed Office may require initiating the first activation vis OSPP.vbs or the script, or opening Office program.</div></li>
    </ul>
    <ul>
      <li>That feature makes use of the "Image File Execution Options" technique to work, programmed as an Application Verifier custom provider for the system file responsible for the KMS process.<br />
      Hence, OS itself handle the DLL injection, allowing the hook to intercept the KMS activation request and write the response on the fly.
      <div>On Windows 8.1/10, it also handles the localhost restriction for KMS activation and redirects any local/private IP address as it were external (different stack).</div></li>
    </ul>
    <ul>
      <li>KMS_VL_ALL scripts make use of Windows Management Instrumentation <strong>WMI</strong> utilities, which query the properties and executes the methods of Windows and Office licensing classes,<br />
      providing a native activation processing, which is almost identical to the official VBScript tools slmgr.vbs and ospp.vbs, but in an automated way.</li>
    </ul>
    <ul>
      <li>The script make these changes to the system (if the emulator is used):
      <div>copy or link the file <code>"C:\Windows\System32\SppExtComObjHook.dll"</code><br />
      add the hook registry keys to <code>"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"</code><br />
      add osppsvc.exe keys to <code>"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"</code><br />
      create scheduled task <code>"\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"</code> (on Windows 8 and later)</div></li>
    </ul>
            <hr />
            <br />

            <h2 id="Supported">Supported Products</h2>
    <p>Volume-capable:</p>
    <ul>
      <li>Windows 11:<br />
      Enterprise, Enterprise LTSC, IoT Enterprise LTSC, Enterprise G, Education, Pro, Pro Workstation, Pro Education, Home, Home Single Language, Home China, SE (CloudEdition)</li><br />
      <li>Windows 10:<br />
      Enterprise, Enterprise LTSC/LTSB, IoT Enterprise LTSC, Enterprise G, Education, Pro, Pro Workstation, Pro Education, Home, Home Single Language, Home China</li><br />
      <li>Windows 8.1:<br />
      Enterprise, Pro, Pro with Media Center, Core, Core Single Language, Core China, Pro for Students, Bing, Bing Single Language, Bing China, Embedded Industry Enterprise/Pro/Automotive</li><br />
      <li>Windows 8:<br />
      Enterprise, Pro, Pro with Media Center, Core, Core Single Language, Core China, Embedded Industry Enterprise/Pro</li><br />
      <li>Windows 10/11 on <strong>ARM64</strong> is supported. Windows 8/8.1/10/11 <strong>N editions</strong> variants are also supported (e.g. Pro N)</li><br />
      <li>Windows 7:<br />
      Enterprise /N/E, Professional /N/E, Embedded POSReady/ThinPC</li><br />
      <li>Windows Server 2025/2022/2019/2016:<br />
      LTSC editions (Standard, Datacenter, Essentials, Cloud Storage, Azure Core, Datacenter Azure Edition, Server ARM64), SAC editions (Standard ACor, Datacenter ACor)</li><br />
      <li>Windows Server 2012 R2:<br />
      Standard, Datacenter, Essentials, Cloud Storage</li><br />
      <li>Windows Server 2012:<br />
      Standard, Datacenter, Essentials, MultiPoint Standard, MultiPoint Premium</li><br />
      <li>Windows Server 2008 R2:<br />
      Standard, Datacenter, Enterprise, MultiPoint, Web, HPC Cluster</li><br />
      <li>Office Volume 2010 / 2013 / 2016 / 2019 / 2021 / 2024</li>
    </ul>
    <p>______________________________</p>
    <p>These editions are only KMS-activatable for <em>45</em> days at max:</p>
    <ul>
      <li>Windows 10/11 Home edition variants</li>
      <li>Windows 8.1 Core edition variants, Pro with Media Center, Pro Student</li>
    </ul>
    <p>These editions are only KMS-activatable for <em>30</em> days at max:</p>
    <ul>
      <li>Windows 8 Core edition variants, Pro with Media Center</li>
    </ul>
    <p>Windows 10/11 Enterprise multi-session:</p>
    <ul>
      <li>This edition is officially supported for Azure Virtual Desktop service</li>
      <li>The edition KMS activation may not work without AVD license</li>
      <li>For more info, see <a href="https://learn.microsoft.com/en-us/azure/virtual-desktop/windows-multisession-faq" target="_blank">here</a></li>
    </ul>
    <p>Notes:</p>
    <ul>
      <li>supported <u>Windows</u> products do not need volume conversion, only the GVLK (KMS key) is needed, which the script will install accordingly.</li>
      <li>KMS activation on Windows 7 has a limitation related to OEM Activation 2.0 and Windows marker. For more info, see <a href="https://support.microsoft.com/en-us/help/942962" target="_blank">here</a> and <a href="https://technet.microsoft.com/en-us/library/ff793426(v=ws.10).aspx#activation-of-windows-oem-computers" target="_blank">here</a>. To verify the activation possibility before attempting, see <a href="https://forums.mydigitallife.net/posts/1553139/" target="_blank">this</a>.</li>
    </ul>
    <p>______________________________</p>
            <h3>Unsupported Products</h3>
    <ul>
      <li>Office MSI Retail 2010/2013, Office 2010 C2R Retail</li>
      <li>Office UWP (Windows 10/11 Apps)</li>
      <li>Windows editions which do not support KMS activation by design:<br />
      Windows Evaluation Editions<br />
      Windows 7 (Starter, HomeBasic, HomePremium, Ultimate)<br />
      Windows 10 (Cloud "S", IoT Enterprise, Professional SingleLanguage, Professional China... etc)<br />
      Windows 11 (IoT Enterprise, Professional SingleLanguage, Professional China... etc)<br />
      Windows Server (Azure Stack HCI, Server Foundation, Storage Server, Home Server 2011... etc)</li>
    </ul>
    <p>______________________________</p>
            <h3>Office C2R 'Your license isn't genuine' notification banner</h3>
    <ul>
      <li>Office Click-to-Run builds (since February 2021) that are activated with KMS checks the existence of the KMS server name in the registry.</li>
      <li>If KMS server is not present, a banner is shown in Office programs notifying that "Office isn't licensed properly", see <a href="https://i.imgur.com/gLFxssD.png" target="_blank">here</a>.</li>
      <li>Therefore in manual mode, <code>KeyManagementServiceName</code> value containing an internal private-network IP address <strong>172.16.0.2</strong> will be kept in the below registry keys:
      <div><code>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform</code><br />
      <code>HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform</code></div></li>
      <li>This is perfectly fine to keep, and it does not affect Windows or Office activation.</li>
      <li>For more explanation, see <a href="https://massgrave.dev/office-license-is-not-genuine" target="_blank">here</a>.</li>
    </ul>
            <hr />
            <br />

            <h2 id="OfficeR2V">Office Retail to Volume</h2>
    <p>Office Retail must be converted to Volume first before it can be activated with KMS</p>
    <p>specifically, Office Click-to-Run products, whether installed from ISO (e.g. ProPlus2019Retail.img) or using Office Deployment Tool.</p>
    <p><b>Starting version 36, the activation script implements automatic license conversion for Office C2R.</b></p>
    <p>Notes:</p>
    <ul>
      <li>Supported Click-to-Run products: Microsoft 365 Apps (Office 365), Office 2013 / 2016 / 2019 / 2021 / 2024</li>
      <li>Activated Office Retail or Subscription products will be skipped from conversion</li>
      <li>Office 365 itself does not have volume licenses, therefore it will be converted to Office Mondo licenses</li>
      <li>Windows 10/11: Office 2016 products will be converted with corresponding Office 2019 licenses (if RTM detected)</li>
      <li>Windows 8.1: Office 2016/2019 products will be converted with corresponding Office 2021 licenses (if RTM detected)</li>
      <li>Office Professional suite will be converted with Office Professional Plus licenses</li>
      <li>Office HomeBusiness/HomeStudent suites will be converted with Office Standard licenses</li>
      <li>Office 2013 products follow the same logic, but handled separately</li>
    </ul>
    <p>Alternatively, if the automatic conversion did not work, or if you prefer to use the standalone converter script:<br />
    <a href="https://forums.mydigitallife.net/posts/1150042/" target="_blank">Office-C2R-Retail2Volume</a></p>
    <p>You can also use other tools that can convert licensing:</p>
    <ul>
      <li><a href="https://forums.mydigitallife.net/threads/78950/" target="_blank">Office Tool Plus</a></li>
      <li><a href="https://forums.mydigitallife.net/posts/1125229/" target="_blank">OfficeRTool</a></li>
    </ul>
            <hr />
            <br />

            <h1 id="Using">How To Use</h1>
    <ul>
      <li>Built-in Windows PowerShell is required for certain functions, make sure it is not disabled or removed from the system.</li>
    </ul>
    <ul>
      <li>Remove any other KMS solutions.</li>
    </ul>
    <ul>
      <li>Temporary suspend Antivirus realtime protection, or exclude the downloaded file and the extracted folder from scanning to avoid quarantine.</li>
    </ul>
    <ul>
      <li>If you are using <strong>Windows Defender</strong> on Windows 11, 10 or 8.1, the script automatically adds an exclusion for <code>C:\Windows\System32\SppExtComObjHook.dll</code><br />
      therefore, <u>it's best not to disable Windows Defender</u>, and instead exclude the downloaded file and the extracted folder before running the script(s).</li>
    </ul>
    <ul>
      <li>Extract the downloaded file contents to a simple path without special characters or long spaces.</li>
    </ul>
    <ul>
      <li>Administrator rights are required to run the script.</li>
    </ul>
    <ul>
      <li>KMS_VL_ALL_AIO offer 3 flavors of activation modes.</li>
    </ul>
            <hr />
            <br />

            <h2 id="Modes">Activation Modes</h2>
            <br />
            <h3 id="ModesAut">Auto Renewal</h3>
    <p>Recommended mode, where you only need to install the activation emulator once. Afterward, the system itself handles and renew activation per schedule.</p>
    <p>To run this mode:</p>
    <ul>
      <li>from the menu, press <b>2</b> to <strong>Install Activation Auto-Renewal</strong></li>
    </ul>
    <p>If you use Antivirus software, make sure to exclude this file from real-time protection:<br /><code>C:\Windows\System32\SppExtComObjHook.dll</code></p>
    <p>If you later installed Volume Office product(s), it will be auto activated in this mode.</p>
    <p>Additionally, If you want to convert and activate Office C2R, renew the activation, or activate new products:</p>
    <ul>
      <li>from the menu, press <b>1</b> to <strong>Activate [Auto Renewal Mode]</strong></li>
    </ul>
    <p>On Windows 8 and later, the script <em>duplicate</em> inbox system scheduled task <code>SvcRestartTaskLogon</code> to <code>SvcTrigger</code><br />
    this is just a precaution step to insure that the auto renewal period is evaluated and respected, it's not directly related to activation itself, and you can manually remove it.</p>
    <p>To remove this mode:</p>
    <ul>
      <li>from the menu, press <b>3</b> to <strong>Uninstall Completely</strong></li>
    </ul>
            <p>____________________________________________________________</p>
            <br />

            <h3 id="ModesMan">Manual</h3>
    <p>No remnants mode, where the activation is executed, and then any KMS emulator traces will be cleared from the system.</p>
    <p>To run this mode:</p>
    <ul>
      <li>make sure that auto renewal mode is not installed, or remove it</li>
      <li>from the menu, press <b>1</b> to <strong>Activate [Manual Mode]</strong></li>
    </ul>
    <p>You will have to run the script again to activate newly installed products (e.g. Office) or if Windows edition is switched.</p>
    <p>You will have to run the script again to activate before the KMS activation period expires.</p>
    <p>You can run and activate anytime during that period to renew the period to the max interval.</p>
    <p>If the script is accidentally terminated before it completes the process, run the script again, then:</p>
    <ul>
      <li>from the menu, press <b>3</b> to <strong>Uninstall Completely</strong></li>
    </ul>
            <p>____________________________________________________________</p>
            <br />

            <h3 id="ModesExt">External</h3>
    <p>Standalone mode, where you activate against trusted external KMS server, without using the local KMS emulator.</p>
    <p>The external server can be a web address, or a network IP address (local LAN or VM).</p>
    <p>To run this mode:</p>
    <ul>
      <li>from the menu, press letter <b>E</b> to <strong>Activate [External Mode]</strong></li>
      <li>input or paste the server address, then press Enter</li>
    </ul>
    <p>If you later installed Volume Office product(s), it will be auto activated if the external server is still connected.</p>
    <p>The used server address will be left registered in the system to allow activated products to auto renew against it,<br />
    if the server is no longer available, you will need to run the mode again with a new available server.</p>
    <p>If you want to clear the server registration and traces:</p>
    <ul>
      <li>from the menu, press <b>3</b> to <strong>Uninstall Completely</strong> (this will also clear KMS cache)</li>
    </ul>
            <hr />
            <br />

            <h2 id="OptConf">Configuration Options</h2>
            <br />
            <h3 id="ConfDbg">Enable Debug Mode</h3>
    <p>Debug Mode is turned OFF by default.</p>
    <p>This option only works with activation functions (menu options [1], [2], [3], [E]).</p>
    <p>If you need to enable this function for troubleshooting or to detect any activation errors:</p>
    <ul>
      <li>from the menu, press <b>4</b> to change the state to <strong>Enable Debug Mode</strong> <b>[Yes]</b></li>
      <li>then, run the desired activation option.</li>
    </ul>
    <p>______________________________</p>

            <h3 id="ConfAct">Process Windows / Process Office</h3>
    <p>The script is set by default to process and try to activate both Windows and Office.</p>
    <p>However, if you want to turn OFF processing Windows <b>or</b> Office, for whatever reason:</p>
    <ul>
      <li>you afraid it may override permanent activation</li>
      <li>you want to speed up the operation (you have Windows or Office already permanently activated)</li>
      <li>you want to activate Windows or Office later on your terms</li>
    </ul>
    <p>To do that:</p>
    <ul>
      <li>from the menu, press <b>5</b> to change the state to <strong>Process Windows</strong> <b>[No]</b></li>
      <li>from the menu, press <b>6</b> to change the state to <strong>Process Office</strong> <b>[No]</b></li>
    </ul>
    <p>Notice:<br />
    this turn OFF is not very effective if Windows or Office installation is already Volume (GVLK installed),<br />
    because the system itself may try to reach and KMS activate the products, especially on Windows 8 and later.</p>
    <p>______________________________</p>

            <h3 id="ConfC2R">Convert Office C2R-R2V</h3>
    <p>The script is set by default to auto convert detected Office C2R Retail to Volume (except activated Retail products).</p>
    <p>However, if you prefer to turn OFF this function:</p>
    <ul>
      <li>from the menu, press <b>7</b> to change the state to <strong>Convert Office C2R-R2V</strong> <b>[No]</b></li>
    </ul>
    <p>______________________________</p>

            <h3 id="ConfOVR">Override Office C2R vNext</h3>
    <p>The script is set by default to override Office C2R vNext license (subscription or lifetime) or its residue.</p>
    <p>However, if you prefer to turn OFF this function:</p>
    <ul>
      <li>from the menu, press letter <b>V</b> to change the state to <strong>Override Office C2R vNext</strong> <b>[No]</b></li>
    </ul>
    <p>Notice:<br />
    If Office vNext license is detected, the option and state will be highlighted, to draw the user attention</p>
    <p>______________________________</p>

            <h3 id="ConfW10">Skip Windows 10/11 KMS 2038</h3>
    <p>The script is set by default to check and skip Windows activation if KMS 2038 is detected.</p>
    <p>However, if you want to revert to normal KMS activation:</p>
    <ul>
      <li>from the menu, press letter <b>X</b> to change the state to <strong>Skip Windows KMS38</strong> <b>[No]</b></li>
    </ul>
    <p>Notice:<br />
    On Windows 10/11, if <code>SkipKMS38</code> is ON (default), Windows will be processed and only checked, even if <code>Process Windows</code> is No</p>
            <hr />
            <br />

            <h2 id="OptMisc">Miscellaneous Options</h2>
            <br />
            <h3 id="MiscChk">Check Activation Status</h3>
    <p>Embedded Windows Powershell script to display the licensing status of Microsoft Windows and Office.</p>
    <ul>
      <li>Robust replacement for the legacy [vbs]/[wmi] options</li>
      <li>For features and more info, check <a href="https://gravesoft.dev/cas" target="_blank">Gravesoft</a></li>
    </ul>
    <p>You can download the legacy scripts here if needed:</p>
    <ul>
      <li><a href="https://pastebin.com/VcT04VRZ" target="_blank">Check-Activation-Status-vbs.bat</a> | <a href="https://gist.github.com/abbodi1406/acba83a99c717aab0be7cd50504d3d99" target="_blank">Mirror</a></li>
      <li><a href="https://pastebin.com/Y7Y5HmkF" target="_blank">Check-Activation-Status-wmi.bat</a> | <a href="https://gist.github.com/abbodi1406/f3cbb251e15ce64f9325ff646e241f58" target="_blank">Mirror</a></li>
    </ul>
    <p>______________________________</p>

            <h3 id="MiscOEM">Create $OEM$ Folder</h3>
    <p>Create needed folder structure and scripts to use during Windows installation to preactivates the system.</p>
    <p>Afterwards, copy <code>$oem$</code> folder to <code>sources</code> folder in the installation media (ISO/USB).</p>
    <p>If you already use another <strong>setupcomplete.cmd</strong>, copy this command line and paste it properly in your setupcomplete.cmd<br />
    <code>call %~dp0KMS_VL_ALL_AIO.cmd /s /a</code></p>
    <p>Notes:</p>
    <ul>
      <li>Created <strong>setupcomplete.cmd</strong> is set by default to run <strong>KMS_VL_ALL_AIO.cmd</strong> in <em>Auto Renewal</em> mode.</li>
      <li>You can change the command line switches to other modes, and add any configuration switches too.</li>
      <li>Later, if you want to uninstall the project, use the menu option <strong>[3] Uninstall Completely</strong>.</li>
      <li>On Windows 8 and later, running setupcomplete.cmd is disabled if the default installed key for the edition is OEM Channel.</li>
    </ul>
    <p>______________________________</p>

            <h3 id="MiscRed">Read Me</h3>
    <p>Extract and start this ReadMeAIO.html.</p>
            <hr />
            <br />

            <h2 id="OptKMS">Advanced KMS Options</h2>
    <p>You can manually modify these KMS-related options by editing the script with Notepad before running.</p>
    <ul>
      <li>
        <strong>KMS_RenewalInterval</strong>
        <br />
        Set the interval for KMS auto renewal schedule for activated clients (default is 10080 = 7 days)<br />
        this only have much effect on Auto Renewal or External modes<br />
        allowed values in minutes: from 15 to 43200</li>
    </ul>
    <ul>
      <li>
        <strong>KMS_ActivationInterval</strong>
        <br />
        Set the interval for KMS reattempt schedule for unactivated clients (default is 120 = 2 hours)<br />
        this does not affect the overall KMS period (180 Days), or the renewal schedule<br />
        allowed values in minutes: from 15 to 43200</li>
    </ul>
    <ul>
      <li>
        <strong>KMS_HWID</strong>
        <br />
        Set the Hardware Hash for local KMS emulator server (only affect Windows 8.1/10)<br />
        <b>0x</b> prefix is mandatory</li>
    </ul>
    <ul>
      <li>
        <strong>KMS_Port</strong>
        <br />
        Set TCP port for KMS communications</li>
    </ul>
    <p>Tip:<br />
    Advanced users can also edit the script and change the default state of configuration options, or activation modes.
    However, command line switches take precedence over inner options.</p>
            <hr />
            <br />

            <h2 id="Switch">Command line Switches</h2>
    <p>
      <strong>Activation switches:</strong>
    </p>
    <ul>
      <li>Auto Renewal mode:<br /><code>/a</code></li>
    </ul>
    <ul>
      <li>Manual mode:<br /><code>/m</code></li>
    </ul>
    <ul>
      <li>External mode:<br /><code>/e pseudo.kms.server</code></li>
    </ul>
    <ul>
      <li>Uninstall and remove all:<br /><code>/r</code></li>
    </ul>
    <p>
      <strong>Configuration switches:</strong>
    </p>
    <ul>
      <li>Silent run:<br /><code>/s</code></li>
    </ul>
    <ul>
      <li>Silent and create simple log:<br /><code>/s /L</code></li>
    </ul>
    <ul>
      <li>Debug mode run:<br /><code>/d</code></li>
    </ul>
    <ul>
      <li>Silent Debug mode:<br /><code>/s /d</code></li>
    </ul>
    <ul>
      <li>Process Windows only:<br /><code>/w</code></li>
    </ul>
    <ul>
      <li>Process Office only:<br /><code>/o</code></li>
    </ul>
    <ul>
      <li>Turn OFF Office C2R-R2V conversion:<br /><code>/c</code></li>
    </ul>
    <ul>
      <li>Do not override Office C2R vNext:<br /><code>/v</code></li>
    </ul>
    <ul>
      <li>Do not skip Windows 10/11 KMS38:<br /><code>/x</code></li>
    </ul>
    <p>
      <strong>Rules:</strong>
    </p>
    <ul>
      <li>All switches are case-insensitive, works in any order, but must be separated with spaces.</li>
    </ul>
    <ul>
      <li>You can specify Configuration switches along with Activation switches.</li>
    </ul>
    <ul>
      <li>If External mode switch <code>/e</code> is specified without server address, it will be changed to Manual or Auto (depending on SppExtComObjHook.dll presence).</li>
    </ul>
    <ul>
      <li>If multiple Activation switches <code>/a /m /e</code> are specified together, the last one takes precedence.</li>
    </ul>
    <ul>
      <li>Uninstall switch <code>/r</code> always takes precedence over Activation switches</li>
    </ul>
    <ul>
      <li>If these Configuration switches <code>/w /o /c /v /x</code> are specified without other switches, they only change the corresponding state in Menu.</li>
    </ul>
    <ul>
      <li>If Process Windows/Office switches <code>/o /w</code> are specified together, the last one takes precedence.</li>
    </ul>
    <ul>
      <li>Log switch <code>/L</code> only works with Silent switch <code>/s</code></li>
    </ul>
    <ul>
      <li>If Silent switch <code>/s</code> and/or Debug switch <code>/d</code> are specified without Activation switches, the script will just run activation in Manual or Auto Renewal mode (depending on SppExtComObjHook.dll presence).</li>
    </ul>
    <p>
      <strong>Examples:</strong>
    </p>
    <pre>
<code>
Silent External activation:
KMS_VL_ALL_AIO.cmd /s /e pseudo.kms.server

Auto Renewal activation for Windows only:
KMS_VL_ALL_AIO.cmd /o /w /a

Manual activation in silent debug mode, do not skip W10 KMS38:
KMS_VL_ALL_AIO.cmd /m /x /d /s

Change config options in menu, Process Office only, do not convert C2R-R2V: 
KMS_VL_ALL_AIO.cmd /o /c

Silent activation (Auto Renewal mode if already installed, otherwise Manual mode):
KMS_VL_ALL_AIO.cmd /s
</code>
    </pre>
    <p>
      <strong>Remarks:</strong>
    </p>
    <ul>
      <li>In general, Windows batch scripts do not work well with unusual folder paths and files name, which contain non-ascii and unicode characters, long paths and spaces, or some of these special characters <code>` ~ ; ' , ! @ % ^ &amp; ( ) [ ] { } + =</code></li>
    </ul>
    <ul>
      <li>KMS_VL_ALL_AIO script is coded to correctly handle those limitations, as much as possible.</li>
    </ul>
    <ul>
      <li>If you changed the script file name and added some unusual characters or spaces, make sure to enclose the script name (or full path) in qoutes marks "" when you run it from command line prompt or another script.</li>
    </ul>
    <ul>
      <li>By default, even explorer context menu option "Run as administrator" will fail to execute on some of those paths.<br />
      In order to fix that, open command prompt as administrator, then copy/paste and execute these commands:</li>
    </ul>
    <pre>
<code>
set _r=^%SystemRoot^%
reg add HKLM\SOFTWARE\Classes\batfile\shell\runas\command /f /v "" /t REG_EXPAND_SZ /d "%_r%\System32\cmd.exe /C \"\"%1\" %*\""
reg add HKLM\SOFTWARE\Classes\cmdfile\shell\runas\command /f /v "" /t REG_EXPAND_SZ /d "%_r%\System32\cmd.exe /C \"\"%1\" %*\""
</code>
    </pre>
            <hr />
            <br />

            <h2 id="Debug">Troubleshooting</h2>
    <p>If the activation failed at first attempt:</p>
    <ul>
      <li>Run the script one more time.</li>
      <li>Reboot the system and try again.</li>
      <li>Verify that Antivirus software is not blocking <code>C:\Windows\SppExtComObjHook.dll</code></li>
      <li>Check System integrity, open command prompt as administrator, and execute these command respectively:<br />
      for Windows 11, 10 or 8.1 only: <code>Dism /online /Cleanup-Image /RestoreHealth</code><br />
      then, for any OS: <code>sfc /scannow</code></li>
    </ul>
    <p>if Auto-Renewal is installed already, but the activation started to fail, run the installation again (option <b>2</b>), or Uninstall Completely then run the installation again.</p>
    <p>For Windows 7, if you have the errors described in <a href="https://support.microsoft.com/en-us/help/4487266" target="_blank">KB4487266</a>, execute the suggested fix.</p>
    <p>If you got Error <strong>0xC004F035</strong> on Windows 7, it means your Machine is not qualified for KMS activation. For more info, see <a href="https://support.microsoft.com/en-us/help/942962" target="_blank">here</a> and <a href="https://technet.microsoft.com/en-us/library/ff793426(v=ws.10).aspx#activation-of-windows-oem-computers" target="_blank">here</a>.</p>
    <p>If you got Error <strong>0x80040154</strong>, it is mostly related to misconfigured Windows 10/11 KMS38 activation, rearm the system and start over, or revert to Normal KMS.</p>
    <p>If you got Error <strong>0xC004E015</strong>, it is mostly related to misconfigured Office retail to volume conversion, try to reinstall system licenses:<br /><code>cscript //Nologo %SystemRoot%\System32\slmgr.vbs /rilc</code></p>
    <p>If you got one of these Errors on Windows Server, verify that the system is properly converted from Evaluation to Retail/Volume:<br /><strong>0xC004E016</strong> - <strong>0xC004F014</strong> - <strong>0xC004F034</strong></p>
    <p>If the activation still failed after the above tips, you may enable the debug mode to help determine the reason:</p>
    <ul>
      <li>from the menu, press <b>5</b> to change the state to <strong>Enable Debug Mode</strong> <b>[Yes]</b></li>
      <li>then, run the desired activation option.</li>
      <li><strong>OR</strong></li>
      <li>run the script with debug command line switch accompanied with an activation mode switch: <code>KMS_VL_ALL_AIO.cmd /d /m</code></li>
      <li>wait until the operation is finished and Debug.log is created</li>
      <li>upload or post the log file on the home page (MDL forums) for inspection</li>
    </ul>
    <p>If you have issues with Office activation, or got undesired or duplicate licenses (e.g. Office 2016 and 2019):</p>
    <ul>
      <li>Download Office Scrubber pack from <a href="https://forums.mydigitallife.net/posts/1466365/" target="_blank">here</a>.</li>
      <li>To get rid of any conflicted licenses, run <strong>Uninstall_Licenses.cmd</strong>, then you must start any Office program to repair the licensing.</li>
      <li>You may also try <strong>Uninstall_Keys.cmd</strong> for similar manner.</li>
      <li>If you wish to remove Office and leftovers completely and start clean:<br />
      uninstall Office normally from Control Panel / Programs and Feature<br />
      then run <strong>Full_Scrub.cmd</strong><br />
      afterward, install new Office.</li>
    </ul>
    <p>Final tip, you may try to rebuild licensing Tokens.dat as suggested in <a href="https://support.microsoft.com/en-us/help/2736303" target="_blank">KB2736303</a> (this will require to repair Office afterward).</p>
            <hr />
            <br />

            <h2 id="Source">Source Code</h2>
            <br />
            <h3 id="srcAvrf">SppExtComObjHookAvrf</h3>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1508167/" target="_blank">https://forums.mydigitallife.net/posts/1508167/</a>
      <br />
      <a href="https://app.box.com/s/mztbabp2n21vvjmk57cl1puel0t088bs" target="_blank">https://app.box.com/s/mztbabp2n21vvjmk57cl1puel0t088bs</a>
    </p>
    <h4 id="visual-studio">Visual Studio:</h4>
    <p>launch shortcut Developer Command Prompt for VS 2017 (or 2019)<br />
    execute:<br />
    <code>MSBuild SppExtComObjHook.sln /p:configuration="Release" /p:platform="Win32"</code><br />
    <code>MSBuild SppExtComObjHook.sln /p:configuration="Release" /p:platform="x64"</code></p>
    <h4 id="mingw-gcc">MinGW GCC:</h4>
    <p>download mingw-w64<br />
    <a href="https://sourceforge.net/projects/mingw-w64/files/i686-8.1.0-release-win32-sjlj-rt_v6-rev0.7z" target="_blank">Windows x86</a><br />
    <a href="https://sourceforge.net/projects/mingw-w64/files/x86_64-8.1.0-release-win32-sjlj-rt_v6-rev0.7z" target="_blank">Windows x64</a><br />
    both can compile 32-bit and 64-bit binaries<br />
    extract and place SppExtComObjHook folder inside mingw32 or mingw64 folder<br />
    run <code>_compile.cmd</code></p>
    <p>______________________________</p>

            <h3 id="srcDebg">SppExtComObjPatcher</h3>
    <h4 id="visual-studio-1">Visual Studio:</h4>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1457558/" target="_blank">https://forums.mydigitallife.net/posts/1457558/</a>
      <br />
      <a href="https://app.box.com/s/mztbabp2n21vvjmk57cl1puel0t088bs" target="_blank">https://app.box.com/s/mztbabp2n21vvjmk57cl1puel0t088bs</a>
    </p>
    <h4 id="mingw-gcc-1">MinGW GCC:</h4>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1462101/" target="_blank">https://forums.mydigitallife.net/posts/1462101/</a>
    </p>
            <hr />
            <br />

            <h2 id="Credits">Credits</h2>
    <p>
      <a href="https://forums.mydigitallife.net/posts/862774" target="_blank">qad</a> - SppExtComObjPatcher, IFEO Debugger.<br />
      <a href="https://forums.mydigitallife.net/posts/1508167/" target="_blank">namazso</a> - SppExtComObjHook, IFEO AVrf custom provider.<br />
      <a href="https://forums.mydigitallife.net/posts/1448556/" target="_blank">Mouri_Naruto</a> - SppExtComObjPatcher-DLL<br />
      <a href="https://forums.mydigitallife.net/posts/1462101/" target="_blank">os51</a> - SppExtComObjPatcher ported to MinGW GCC, Retail/MAK checks examples.<br />
      <a href="https://forums.mydigitallife.net/posts/309737/" target="_blank">MasterDisaster</a> - Original script, WMI methods.<br />
      <a href="https://forums.mydigitallife.net/members/1108726/" target="_blank">Windows_Addict</a> - Features suggestion, ideas, testing, and co-enhancing.<br />
      <a href="https://gist.github.com/ave9858/9fff6af726ba3ddc646285d1bbf37e71" target="_blank">ave9858</a> - CleanOffice.ps1<br />
      <a href="https://github.com/asdcorp/clic" target="_blank">asdcorp</a> - clic tool.<br />
      <a href="https://github.com/AveYo/Compressed2TXT" target="_blank">AveYo</a> - Compressed2TXT ascii encoder.<br />
      <a href="https://stackoverflow.com/a/10407642" target="_blank">dbenham, jeb</a> - Color text in batch script.<br />
      <a href="https://stackoverflow.com/a/13351373" target="_blank">dbenham</a> - Set buffer height independently of window height.<br />
      <a href="https://forums.mydigitallife.net/threads/74769/" target="_blank">hearywarlot</a> - Auto Elevate as admin.<br />
      <a href="https://forums.mydigitallife.net/posts/1296482/" target="_blank">qewpal</a> - KMS-VL-ALL script.<br />
      <a href="https://forums.mydigitallife.net/members/846864/" target="_blank">NormieLyfe</a> - GVLK categorize, Office checks help.<br />
      <a href="https://forums.mydigitallife.net/members/120394/" target="_blank">rpo</a>, <a href="https://forums.mydigitallife.net/members/2574/" target="_blank">mxman2k</a>, <a href="https://forums.mydigitallife.net/members/58504/" target="_blank">BAU</a>, <a href="https://forums.mydigitallife.net/members/presto1234.647219/" target="_blank">presto1234</a> - scripting suggestions.<br />
      <a href="https://forums.mydigitallife.net/members/80361/" target="_blank">Nucleus</a>, <a href="https://forums.mydigitallife.net/members/104688/" target="_blank">Enthousiast</a>, <a href="https://forums.mydigitallife.net/members/293479/" target="_blank">s1ave77</a>, <a href="https://forums.mydigitallife.net/members/325887/" target="_blank">l33tisw00t</a>, <a href="https://forums.mydigitallife.net/members/77147/" target="_blank">LostED</a>, <a href="https://forums.mydigitallife.net/members/1023044/" target="_blank">Sajjo</a> and MDL Community for interest, feedback, and assistance.</p>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1343297/" target="_blank">abbodi1406</a> - KMS_VL_ALL author</p>

            <h2 id="acknow">Acknowledgements</h2>
    <p>
      <a href="https://forums.mydigitallife.net/forums/51/" target="_blank">MDL forums</a> - the home of the latest and current emulators.<br />
      <a href="https://forums.mydigitallife.net/posts/838505" target="_blank">mikmik38</a> - fixed reversed source of KMSv5 and KMSv6.<br />
      <a href="https://forums.mydigitallife.net/threads/41010/" target="_blank">CODYQX4</a> - easy to use KMSEmulator source.<br />
      <a href="https://forums.mydigitallife.net/threads/50234/" target="_blank">Hotbird64</a> - the resourceful vlmcsd tool, and KMSEmulator source development.<br />
      <a href="https://forums.mydigitallife.net/threads/50949/" target="_blank">cynecx</a> - SECO Injector bypass, SppExtComObj KMS functions.<br />
      <a href="https://forums.mydigitallife.net/posts/856978" target="_blank">deagles</a> - SppExtComObjHook Injector.<br />
      <a href="https://forums.mydigitallife.net/posts/839363" target="_blank">deagles</a> - KMSServerService.<br />
      <a href="https://forums.mydigitallife.net/posts/1475544/" target="_blank">ColdZero</a> - CZ VM System.<br />
      <a href="https://forums.mydigitallife.net/posts/1476097/" target="_blank">ColdZero</a> - KMS ePID Generator.<br />
      <a href="https://forums.mydigitallife.net/posts/838023" target="_blank">kelorgo</a>, <a href="http://forums.mydigitallife.net/posts/838114" target="_blank">bedrock</a> - TAP adapter TunMirror bypass.<br />
      <a href="https://forums.mydigitallife.net/posts/1259604/" target="_blank">mishamosherg</a> - WinDivert FakeClient bypass.<br />
      <a href="https://forums.mydigitallife.net/posts/860489" target="_blank">Duser</a> - KMS Emulator fork.<br />
      <a href="https://forums.mydigitallife.net/threads/67038/" target="_blank">Boops</a> - Tool Ghost KMS (TGK).<br />
      ZWT, nosferati87, crony12, FreeStyler, Phazor - KMS Emulator development.</p>
        </div>
    </main>

    <nav id="nav">
        <div class="innertube">
            <a href="#Overview">Overview</a><br />
            <a href="#AIO">AIO vs. Traditional</a><br />
            <a href="#How">How does it work?</a><br />
            <a href="#Supported">Supported Products</a><br />
            <a href="#OfficeR2V">Office Retail to Volume</a><br />
            <a href="#Using">How To Use</a><br /><br />
            <a href="#Modes">Activation Modes</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ModesAut">Auto Renewal</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ModesMan">Manual</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ModesExt">External</a><br /><br />
            <a href="#OptConf">Configuration Options</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ConfDbg">Debug Mode</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ConfAct">Activation Choice</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ConfC2R">Office C2R-R2V</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ConfOVR">Office C2R vNext</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ConfW10">KMS38 Win 10/11</a><br /><br />
            <a href="#OptMisc">Miscellaneous Options</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#MiscChk">Activation Status</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#MiscOEM">$OEM$ Folder</a><br /><br />
            <a href="#OptKMS">Advanced KMS Options</a><br />
            <a href="#Switch">Command line Switches</a><br />
            <a href="#Debug">Troubleshooting</a><br /><br />
            <a href="#Source">Source Code</a><br />
            <a href="#Credits">Credits</a><br />
        </div>
    </nav>
  </body>
</html>
:readme:

:DoDebug
set _dDbg=No
cmd.exe /c ""!_batf!" !_para!"
set _dDbg=Yes
echo.
echo Done.
echo Press any key to continue . . .
pause >nul
goto :MainMenu

:E_Admin
echo %_err%
echo This script requires administrator privileges.
echo To do so, right-click on this script and select 'Run as administrator'
goto :E_Exit

:E_PTH
echo.
echo === WARNING ===
echo Disallowed special characters are detected in the file path or name.
echo Make sure either do not contain any of the following characters:
echo ^` ^~ ^! ^@ %% ^^ ^& ^( ^) [ ] { } ^+ ^= ^; ^' ^,
goto :E_Exit

:E_PWS
echo %_err%
echo Windows PowerShell is not installed.
echo It is required for this script to work.
goto :E_Exit

:E_VBS
echo %_err%
echo VBScript engine is not installed.
echo It is required for this script to work.
goto :E_Exit

:E_WSH
echo %_err%
echo Windows Script Host is disabled.
echo It is required for this script to work.
goto :E_Exit

:E_WMS
echo %_err%
echo Windows Management Instrumentation [WinMgmt] service is disabled.
echo It is required for this script to work.
goto :E_Exit

:E_PLM
echo %_err%
echo Windows PowerShell is not properly responding.
echo check if it is working, and not locked in Constrained Language Mode.
goto :E_Exit

:E_WMI
echo %_err%
echo This script require one of these to work:
echo wmic.exe tool
echo VBScript engine
echo Windows PowerShell
goto :E_Exit

:E_Exit
if %_Debug% EQU 1 goto :eof
if %Unattend% EQU 1 goto :eof
echo.
echo Press any key to exit.
pause >nul
goto :eof

:UnsupportedVersion
echo %_err%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10/11 and their Server equivalent.
:TheEnd
if exist "%PUBLIC%\ReadMeAIO.html" del /f /q "%PUBLIC%\ReadMeAIO.html"
if exist "%_temp%\'" del /f /q "%_temp%\'"
if exist "%_temp%\`.txt" del /f /q "%_temp%\`.txt"
if defined _quit goto :eof
echo.
if %Unattend% EQU 0 echo Press any key to exit.
%_Pause%
goto :eof

:qrPKey
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where Version='%2' call InstallProductKey ProductKey="%3""
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csp% %1 "%3""
exit /b
)
set _qr=%_psc% "try {$null=([WMI]'%1=''%2''').InstallProductKey('%3')} catch {$host.SetShouldExit($_.Exception.HResult)}"
exit /b

:qrMethod
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where %2='%3' call %4"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csm% "%1.%2='%3'" %4"
exit /b
)
set _qr=%_psc% "try {$null=([WMI]'%1.%2=''%3''').%4()} catch {$host.SetShouldExit($_.Exception.HResult)}"
exit /b

:qrSingle
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 get %2 /value"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csq% %1 %2"
exit /b
)
set _qr=%_psc% "(([WMISEARCHER]'SELECT %2 FROM %1').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

:qrQuery
set "_quxt="
set "_quxt=%~4"
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where "%~2" get %3 /value"
if defined _quxt set "_qr=wmic path %1 where "%~2" get %3"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csq% %1 "%~2" %3"
exit /b
)
set "_rq=%~2"
set "_rq=%_rq:'=''%"
set _qr=%_psc% "(([WMISEARCHER]'SELECT %3 FROM %1 WHERE %_rq%').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

:qrWD
if %_cwmi% EQU 1 (
set "_qr=WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call %1 ExclusionPath=%_Hook% Force=True"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csd% %1 %_Hook%"
exit /b
)
set _qr=%_psc% "try {$null = icim MSFT_MpPreference @{ExclusionPath = @('%_Hops%'); Force = $True} %1 -Namespace root/Microsoft/Windows/Defender -EA 1} catch {$host.SetShouldExit($_.Exception.HResult)}"
exit /b

:qrCheck
if %_cwmi% EQU 1 (
set "_qrw=wmic path %1 get %2 /value"
set "_qrs=wmic path %3 get %4 /value"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qrw=%_csq% %1 %2"
set "_qrs=%_csq% %3 %4"
exit /b
)
set _qrw=%_psc% "(([WMISEARCHER]'SELECT %2 FROM %1').Get()).Properties | %% {$_.Name+'='+$_.Value}"
set _qrs=%_psc% "(([WMISEARCHER]'SELECT %4 FROM %3').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

:casWqr
if %_cwmi% EQU 1 (
set "_qr=wmic path %1 where "ApplicationID='%2' and PartialProductKey is not null" get %3 /value"
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csq% %1 "ApplicationID='%2' and PartialProductKey is not null" %3"
exit /b
)
set _qr=%_psc% "(([WMISEARCHER]'SELECT %3 FROM %1 WHERE ApplicationID=''%2'' AND PartialProductKey IS NOT NULL').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

:casWall
if %_cwmi% EQU 1 (
set "_qr="wmic path %~1 get %~2 /value" ^| findstr ^="
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csg% %~1 "%~2""
exit /b
)
set _qr=%_psc% "(([WMISEARCHER]'SELECT %~2 FROM %~1').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

:casWsng
if %_cwmi% EQU 1 (
set "_qr="wmic path %~1 where ID='%~2' get %~3 /value" ^| findstr ^="
exit /b
)
if %WMI_VBS% NEQ 0 (
set "_qr=%_csg% %~1 "ID='%~2'" "%~3""
exit /b
)
set _qr=%_psc% "(([WMISEARCHER]'SELECT %~3 FROM %~1 WHERE ID=''%~2''').Get()).Properties | %% {$_.Name+'='+$_.Value}"
exit /b

----- Begin wsf script --->
<package>
   <job id="WmiQuery">
      <script language="VBScript">
         If WScript.Arguments.Count = 3 Then
            wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
            wGet = WScript.Arguments.Item(2)
         Else
            wExc = "Select " & WScript.Arguments.Item(1) & " from " & WScript.Arguments.Item(0)
            wGet = WScript.Arguments.Item(1)
         End If
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               If LCase(Prop.Name) = LCase(wGet) Then
                  WScript.Echo Prop.Name & "=" & Prop.Value
                  Exit For
               End If
            Next
         Next
      </script>
   </job>
   <job id="WmiMethod">
      <script language="VBScript">
         On Error Resume Next
         wPath = WScript.Arguments.Item(0)
         wMethod = WScript.Arguments.Item(1)
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2:" & wPath)
         objCol.ExecMethod_(wMethod)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="WmiPKey">
      <script language="VBScript">
         On Error Resume Next
         wExc = "SELECT Version FROM " & WScript.Arguments.Item(0)
         wKey = WScript.Arguments.Item(1)
         Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For each colService in objWMIService
            Exit For
         Next
         set objService = colService
         objService.InstallProductKey(wKey)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="XPDT">
      <script language="VBScript">
         WScript.Echo DateAdd("n", WScript.Arguments.Item(0), Now)
      </script>
   </job>
   <job id="MPS">
      <script language="VBScript">
         On Error Resume Next
         wMethod = WScript.Arguments.Item(0)
         wValue = WScript.Arguments.Item(1)
         Set objID = GetObject("winmgmts:\\.\root\Microsoft\Windows\Defender").ExecQuery("Select ComputerID from MSFT_MpPreference")
         For Each objItm in objID
            cid = objItm.ComputerID
         Next
         Set objCol = GetObject("winmgmts:\\.\root\Microsoft\Windows\Defender:MSFT_MpPreference.ComputerID='" & cid & "'")
         Set objInp = objCol.Methods_(wMethod).inParameters.SpawnInstance_()
         objInp.Properties_.Item("ExclusionPath") = Split(wValue, ";")
         objInp.Properties_.Item("Force") = True
         Set objOut = objCol.ExecMethod_(wMethod, objInp)
         WScript.Quit Err.Number
      </script>
   </job>
   <job id="WmiMulti">
      <script language="VBScript">
         If WScript.Arguments.Count = 3 Then
            wExc = "Select " & WScript.Arguments.Item(2) & " from " & WScript.Arguments.Item(0) & " where " & WScript.Arguments.Item(1)
         Else
            wExc = "Select " & WScript.Arguments.Item(1) & " from " & WScript.Arguments.Item(0)
         End If
         Set objCol = GetObject("winmgmts:\\.\root\CIMV2").ExecQuery(wExc,,48)
         For Each objItm in objCol
            For each Prop in objItm.Properties_
               WScript.Echo Prop.Name & "=" & Prop.Value
            Next
         Next
      </script>
   </job>
   <job id="ELAV">
      <script language="VBScript">
         Set strArg=WScript.Arguments.Named
         Set strRdlproc = CreateObject("WScript.Shell").Exec("rundll32 kernel32,Sleep")
         With GetObject("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" & strRdlproc.ProcessId & "'")
            With GetObject("winmgmts:\\.\root\CIMV2:Win32_Process.Handle='" & .ParentProcessId & "'")
               If InStr (.CommandLine, WScript.ScriptName) <> 0 Then
                  strLine = Mid(.CommandLine, InStr(.CommandLine , "/File:") + Len(strArg("File")) + 8)
               End If
            End With
            .Terminate
         End With
         CreateObject("Shell.Application").ShellExecute "cmd.exe", "/c " & chr(34) & chr(34) & strArg("File") & chr(34) & strLine & chr(34), "", "runas", 1
      </script>
   </job>
</package>