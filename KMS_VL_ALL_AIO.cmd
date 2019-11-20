<!-- : Begin batch script
@set uivr=v36
@echo off
:: ### Configuration Options ###

:: change to 1 to enable debug mode (can be used with unattend options)
set _Debug=0

:: change to 0 to turn OFF Windows or Office activation processing via the script
set ActWindows=1
set ActOffice=1

:: change to 0 to turn OFF auto conversion for Office C2R Retail to Volume
set AutoR2V=1

:: change to 0 to revert Windows 10 KMS38 to normal KMS
set SkipKMS38=1

:: ### Unattend Options ###

:: change to 1 and set KMS_IP address to activate via external KMS server unattended
set External=0
set KMS_IP=172.16.0.2

:: change to 1 to run Manual activation mode unattended
set uManual=0

:: change to 1 to run AutoRenewal activation mode unattended
set uAutoRenewal=0

:: change to 1 to suppress any output
set Silent=0

:: change to 1 to redirect output to text file, works only with Silent=1
set Logger=0

:: ### Advanced Options ###

:: change KMS auto renewal schedule, range in minutes: from 15 to 43200
:: example: 10080 = weekly, 1440 = daily, 43200 = monthly
set KMS_RenewalInterval=10080

:: change KMS reattempt schedule for failed activation or unactivated, range in minutes: from 15 to 43200
set KMS_ActivationInterval=120

:: change Hardware Hash for KMS emulator server (only affect Windows 8.1 and 10)
set KMS_HWID=0x3A1C049600B60076

:: change KMS TCP port
set KMS_Port=1688

:: ##################################################################
:: # NORMALY THERE IS NO NEED TO CHANGE ANYTHING BELOW THIS COMMENT #
:: ##################################################################

set KMS_Emulation=1
set Unattend=0

set fAUR=
set rAUR=
set "_args=%*"
if not defined _args goto :NoProgArgs
if "%~1"=="" set "_args="&goto :NoProgArgs

set _args=%_args:"=%
for %%A in (%_args%) do (
if /i "%%A"=="/d" (set _Debug=1
) else if /i "%%A"=="/u" (set Unattend=1
) else if /i "%%A"=="/s" (set Silent=1
) else if /i "%%A"=="/l" (set Logger=1
) else if /i "%%A"=="/o" (set ActOffice=1&set ActWindows=0
) else if /i "%%A"=="/w" (set ActOffice=0&set ActWindows=1
) else if /i "%%A"=="/c" (set AutoR2V=0
) else if /i "%%A"=="/x" (set SkipKMS38=0
) else if /i "%%A"=="/e" (set fAUR=0&set External=1
) else if /i "%%A"=="/m" (set fAUR=0&set External=0
) else if /i "%%A"=="/a" (set fAUR=1&set External=0
) else if /i "%%A"=="/r" (set rAUR=1
) else (set "KMS_IP=%%A")
)

:NoProgArgs
if %External% EQU 1 (if "%KMS_IP%"=="172.16.0.2" (set fAUR=0&set External=0) else (set fAUR=0))
if %uManual% EQU 1 (set fAUR=0&set External=0)
if %uAutoRenewal% EQU 1 (set fAUR=1&set External=0)
if defined fAUR set Unattend=1
if defined rAUR set Unattend=1
if %Silent% EQU 1 set Unattend=1
set "_run=nul"
if %Logger% EQU 1 set _run="%~dpn0_Silent.log"

set "SysPath=%SystemRoot%\System32"
if exist "%SystemRoot%\Sysnative\reg.exe" (set "SysPath=%SystemRoot%\Sysnative")
set "Path=%SysPath%;%SystemRoot%;%SysPath%\Wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
set "_err===== ERROR ===="
set "xOS=amd64"
set "xBit=amd64"
set "_bit=64"
set "_wow=1"
if /i %PROCESSOR_ARCHITECTURE%==x86 (if not defined PROCESSOR_ARCHITEW6432 (
  set "xOS=x86"
  set "xBit=x86"
  set "_bit=32"
  set "_wow=0"
  )
)

setlocal DisableDelayedExpansion
set "param=%~f0"
cmd /v:on /c echo(^^!param^^!| findstr /R "[| ` ~ ! @ %% \^ & ( ) \[ \] { } + = ; ' , |]*^"
endlocal
if %errorlevel% EQU 0 (
echo.
echo %_err%
echo Disallowed special characters detected in file path name.
echo Make sure the path do not contain the following special characters,
echo ^` ^~ ^! ^@ %% ^^ ^& ^( ^) [ ] { } ^+ ^= ^; ^' ^,
echo.
echo Press any key to exit.
pause >nul
goto :eof
)

if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" goto :E_PS

1>nul 2>nul reg query HKU\S-1-5-19 && goto :Passed

set "_PSarg="""%~f0""" "
if defined _args (
set "_PSarg="""%~f0""" %_args:"="""%"
)

(1>nul 2>nul cscript //NoLogo "%~f0?.wsf" //job:ELAV /File:"%~f0" %*) && (
  exit /b
  ) || (
  call setlocal EnableDelayedExpansion
  1>nul 2>nul %SysPath%\WindowsPowerShell\v1.0\powershell -noprofile -exec bypass -c "start cmd -ArgumentList '/c \"!_PSarg!\"' -verb runas" && (
    exit /b
    ) || (
    goto :E_Admin
  )
)

:Passed
set "_temp=%SystemRoot%\Temp"
set "_log=%~dpn0"
set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"
for /f "skip=2 tokens=2*" %%a in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders" /v Desktop') do call set "_dsk=%%b"
if exist "%SystemDrive%\Users\Public\Desktop\desktop.ini" set "_dsk=%SystemDrive%\Users\Public\Desktop"
set "_suf="
setlocal EnableDelayedExpansion

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
  )
  copy /y nul "!_work!\#.rw" 1>nul 2>nul && (if exist "!_work!\#.rw" del /f /q "!_work!\#.rw") || (set "_log=!_dsk!\%~n0")
  if exist "!_log!_Debug.log" (
  for /f "tokens=2 delims==." %%# in ('wmic os get localdatetime /value') do set "_date=%%#"
  set "_suf=_!_date:~8,6!"
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
if %_Debug% EQU 1 if defined _args echo %_args%
set "_wApp=55c92734-d682-4d71-983e-d6ec3f16059f"
set "_oApp=0ff1ce15-a989-479d-af46-f275c6370663"
set "_oA14=59a52881-a989-479d-af46-f275c6370663"
set "IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"
set "OPPk=SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"
set "SPPk=SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
set _Hook="%SysPath%\SppExtComObjHook.dll"
set w7inf=%SystemRoot%\Migration\WTR\KMS_VL_ALL.inf
set "_TaskEx=\Microsoft\Windows\SoftwareProtectionPlatform\SvcTrigger"
set "_TaskOs=\Microsoft\Windows\SoftwareProtectionPlatform\SvcRestartTaskLogon"
set "line1============================================================="
set "line2=************************************************************"
set "line3=____________________________________________________________"
set "line4=__________________________________________________"
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G
set SSppHook=0
for /f %%A in ('dir /b /ad %SysPath%\spp\tokens\skus') do (
  if %winbuild% GEQ 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*GVLK*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VLKMS*.xrm-ms" set SSppHook=1
  if %winbuild% LSS 9200 if exist "%SysPath%\spp\tokens\skus\%%A\*VL-BYPASS*.xrm-ms" set SSppHook=1
)
set OsppHook=1
sc query osppsvc %_Nul3%
if %errorlevel% EQU 1060 set OsppHook=0

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
set _uRI=
set _uAI=
for /f "tokens=2 delims==" %%# in ('findstr /i /b /c:"set KMS_RenewalInterval" "%~f0"') do if not defined _uRI set _uRI=%%#
for /f "tokens=2 delims==" %%# in ('findstr /i /b /c:"set KMS_ActivationInterval" "%~f0"') do if not defined _uAI set _uAI=%%#
set _dDbg=No
if %ActWindows% EQU 0 if %ActOffice% EQU 0 set ActWindows=1
if %_Debug% EQU 1 if not defined fAUR set fAUR=0&set External=0
if %Unattend% EQU 1 if not defined fAUR set fAUR=0&set External=0
if not defined fAUR if not defined rAUR goto :MainMenu
if defined rAUR (set _verb=1&cls&call :RemoveHook&goto :cCache)
set Unattend=1
set AUR=0
if exist %_Hook% dir /b /al %_Hook% %_Nul3% || (
  reg query "%IFEO%\%SppVer%" /v KMS_Emulation %_Nul3% && set AUR=1
  reg query "%IFEO%\osppsvc.exe" /v KMS_Emulation %_Nul3% && set AUR=1
)
if %fAUR% EQU 1 (if %AUR% EQU 0 (set AUR=1&set _verb=1&set _rtr=DoActivate&cls&goto :InstallHook) else (set _verb=0&set _rtr=DoActivate&cls&goto :InstallHook))
if %External% EQU 0 (set AUR=0&cls&goto :DoActivate)
cls&goto :DoActivate

:MainMenu
cls
mode con cols=80 lines=34
title KMS_VL_ALL AIO %uivr%
color 07
set _dMode=Manual
set AUR=0
if exist %_Hook% dir /b /al %_Hook% %_Nul3% || (
  reg query "%IFEO%\%SppVer%" /v KMS_Emulation %_Nul3% && (set AUR=1&set _dMode=Auto Renewal)
  reg query "%IFEO%\osppsvc.exe" /v KMS_Emulation %_Nul3% && (set AUR=1&set _dMode=Auto Renewal)
)
set _ckc=
if %AUR% EQU 0 (
set _dHook=Not Installed
set _ckc=1
) else (
set _dHook=Already Installed
reg query "HKLM\%SPPk%" /v KeyManagementServiceName /s %_Nul2% | findstr 172.16.0.2 %_Nul1% || set _ckc=1
if %OSType% EQU Win8 reg query "HKU\S-1-5-20\%SPPk%" /v DiscoveredKeyManagementServiceIpAddress /s %_Nul2% | findstr 172.16.0.2 %_Nul1% || set _ckc=1
if %OSType% EQU Win7 reg query "HKLM\%OPPk%" /v KeyManagementServiceName /s %_Nul2% | findstr 172.16.0.2 %_Nul1% || set _ckc=1
)
if %ActWindows% EQU 0 (set _dAwin=No) else (set _dAwin=Yes)
if %ActOffice% EQU 0 (set _dAoff=No) else (set _dAoff=Yes)
if %AutoR2V% EQU 0 (set _dArtv=No) else (set _dArtv=Yes)
if %SkipKMS38% EQU 0 (set _dWXKMS=No) else (set _dWXKMS=Yes)
if %_Debug% EQU 0 (set _dDbg=No) else (set _dDbg=Yes)
set _el=
echo.        
echo           %line3%
echo.        
echo                [1] Activate                [%_dMode% Mode]
echo.              
echo                [2] Install Auto Renewal    [%_dHook%]
echo                [3] Uninstall Completely
echo.              
echo                [4] Activate                [External Mode]
echo                %line4%
echo.              
echo                    Configuration:
echo.              
echo                [5] Enable Debug Mode       [%_dDbg%]
echo                [6] Process Windows         [%_dAwin%]
echo                [7] Process Office          [%_dAoff%]
echo                [C] Convert Office C2R-R2V  [%_dArtv%]
if %winbuild% GEQ 10240 echo                [X] Skip Windows 10 KMS38   [%_dWXKMS%]
echo                %line4%
echo.              
echo                    Miscellaneous:
echo.              
echo                [8] Check Activation Status [vbs]
echo                [9] Check Activation Status [wmic]
if defined _ckc echo                [K] Clear KMS Cache
echo                [S] Create $OEM$ Folder
echo                [R] Read Me
echo           %line3%
echo.
choice /c 1234567890CRSKX /n /m ">           Choose a menu option, or press 0 to Exit: "
set _el=%errorlevel%
if %_el%==15 if %winbuild% GEQ 10240 (if %SkipKMS38% EQU 0 (set SkipKMS38=1) else (set SkipKMS38=0))&goto :MainMenu
if %_el%==14 if defined _ckc (set _verb=0&cls&goto :cCache)
if %_el%==13 (call :CreateOEM)&goto :MainMenu
if %_el%==12 (call :CreateReadMe)&goto :MainMenu
if %_el%==10 goto :eof
if %_el%==11 (if %AutoR2V% EQU 0 (set AutoR2V=1) else (set AutoR2V=0))&goto :MainMenu
if %_el%==9 (call :casWm)&goto :MainMenu
if %_el%==8 (call :casVm)&goto :MainMenu
if %_el%==7 (if %ActOffice% EQU 0 (set ActOffice=1) else (set ActWindows=1&set ActOffice=0))&goto :MainMenu
if %_el%==6 (if %ActWindows% EQU 0 (set ActWindows=1) else (set ActWindows=0&set ActOffice=1))&goto :MainMenu
if %_el%==5 (if %_Debug% EQU 0 (set _Debug=1) else (set _Debug=0))&goto :MainMenu
if %_el%==4 goto :E_IP
if %_el%==3 (if %_dDbg%==No (set _verb=1&cls&call :RemoveHook&goto :cCache) else (set _verb=1&cls&goto :RemoveHook))
if %_el%==2 (if %AUR% EQU 0 (set AUR=1&set _verb=1&set _rtr=DoActivate&cls&goto :InstallHook) else (set _verb=0&set _rtr=DoActivate&cls&goto :InstallHook))
if %_el%==1 (cls&goto :DoActivate)
goto :MainMenu

:E_IP
cls
set kip=
echo.
echo Enter or Paste the external KMS Server address:
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
if %SkipKMS38% EQU 0 set "_para=!_para! /x"
if %External% EQU 1 set "_para=!_para! /e %KMS_IP%"
if %External% EQU 0 if %AUR% EQU 0 set "_para=!_para! /m"
if %External% EQU 0 if %AUR% EQU 1 set "_para=!_para! /a"
goto :DoDebug
)
if %External% EQU 1 (
if "%KMS_IP%"=="172.16.0.2" set External=0
)
if %External% EQU 1 (
set AUR=1
)
if %External% EQU 0 (
set KMS_IP=172.16.0.2
)
if %AUR% EQU 0 (
set KMS_RenewalInterval=43200
set KMS_ActivationInterval=43200
) else (
set KMS_RenewalInterval=%_uRI%
set KMS_ActivationInterval=%_uAI%
)
if %External% EQU 1 (
color 8F&set "mode=External ^(%KMS_IP%^)"
) else (
if %AUR% EQU 0 (color 1F&set "mode=Manual") else (color 07&set "mode=Auto Renewal")
)
if %Unattend% EQU 0 (
if %_Debug% EQU 0 (title KMS_VL_ALL AIO %uivr%) else (title KMS_VL_ALL %uivr% : %mode%)
) else (
echo.
echo Running KMS_VL_ALL AIO %uivr%
)
if %Silent% EQU 0 if %_Debug% EQU 0 (
%_Nul3% powershell -noprofile -exec bypass -c "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.height=300;$W.buffersize=$B;}"
)
if %winbuild% GEQ 9600 (
  reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /f /v NoGenTicket /t REG_DWORD /d 1 %_Nul3%
)
echo.
echo Activation Mode: %mode%
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
if %External% EQU 0 if %AUR% EQU 0 (set _verb=0&set _rtr=ReturnHook&goto :InstallHook)

:ReturnHook
if %External% EQU 0 if %AUR% EQU 1 (
call :UpdateIFEOEntry %SppVer%
call :UpdateIFEOEntry osppsvc.exe
)
if %External% EQU 1 if %AUR% EQU 1 (
call :UpdateOSPPEntry osppsvc.exe
)

SET Win10Gov=0
IF %winbuild% LSS 14393 if %SSppHook% NEQ 0 GOTO :Main
SET "EditionWMI="
SET "EditionID="
SET "RegKey=HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
SET "Pattern=Microsoft-Windows-*Edition~31bf3856ad364e35"
SET "EditionPKG=NUL"
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
FOR /F "TOKENS=2 DELIMS==" %%A IN ('"WMIC PATH SoftwareLicensingProduct WHERE (ApplicationID='%_wApp%' AND PartialProductKey is not NULL) GET LicenseFamily /VALUE" %_Nul6%') DO IF NOT ERRORLEVEL 1 SET "EditionWMI=%%A"
IF NOT DEFINED EditionWMI (
IF %winbuild% GEQ 17063 FOR /F "SKIP=2 TOKENS=3 DELIMS= " %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%A"
IF %winbuild% LSS 14393 FOR /F "SKIP=2 TOKENS=3 DELIMS= " %%A IN ('REG QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v EditionId') DO SET "EditionID=%%A"
GOTO :Main
)
FOR %%A IN (Cloud,CloudN,IoTEnterprise,IoTEnterpriseS,ProfessionalSingleLanguage,ProfessionalCountrySpecific) DO (IF /I "%EditionWMI%"=="%%A" GOTO :Main)
SET "EditionID=%EditionWMI%"

:Main
IF DEFINED EditionID FOR %%A IN (EnterpriseG,EnterpriseGN) DO (IF /I "%EditionID%"=="%%A" SET Win10Gov=1)
if defined EditionID (set "_winos=Windows %EditionID% edition") else (set "_winos=Detected Windows")
if %winbuild% LSS 10240 for /f "skip=2 tokens=2*" %%a in ('reg query "hklm\software\microsoft\Windows NT\currentversion" /v productname %_Nul6%') do if not errorlevel 1 set "_winos=%%b"
if %winbuild% GEQ 10240 for /f "tokens=2* delims== " %%a in ('"wmic os get caption /value" %_Nul6%') do if not errorlevel 1 set "_winos=%%b"
set "nKMS=does not support KMS activation..."
set "nEval=Evaluation Editions cannot be activated. Please install full Windows OS."
if defined EditionID echo %EditionID%| findstr /I /E Eval %_Nul1% && (
set _eval=1
echo %EditionID%| findstr /I /B Server %_Nul1% && (set "nEval=Server Evaluation cannot be activated. Please convert to full Server OS.")
)
set "_C16R="
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if not defined _C16R reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C16R=HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
set "_C15R="
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds %_Nul3% && set "_C15R=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
if not defined _C15R reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid %_Nul3% && set "_C15R=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag"
set _V16Ids=Mondo,ProPlus,ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness,OneNote,Excel,Outlook,PowerPoint,Publisher,Word
set _R16Ids=%_V16Ids%,Professional,HomeBusiness,HomeStudent,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
for %%A in (14,15,16,19) do call :officeLoc %%A

call :RunSPP
if %ActOffice% NEQ 0 call :RunOSPP
if %ActOffice% EQU 0 (echo.&echo Office activation is OFF...)

if exist "!_temp!\crv*.txt" del /f /q "!_temp!\crv*.txt"
if exist "!_temp!\*chk.txt" del /f /q "!_temp!\*chk.txt"
if exist "!_temp!\slmgr.vbs" del /f /q "!_temp!\slmgr.vbs"
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc

if %AUR% EQU 0 call :RemoveHook

sc start sppsvc trigger=timer;sessionid=0 %_Nul3%
set External=0
set KMS_IP=172.16.0.2
if %Unattend% NEQ 0 goto :TheEnd
echo.
echo Press any key to continue...
pause >nul
goto :MainMenu

:RunSPP
set spp=SoftwareLicensingProduct
set sps=SoftwareLicensingService
set W1nd0ws=1
set WinPerm=0
set WinVL=0
set Off1ce=0
set RunR2V=0
if %winbuild% GEQ 9200 if %ActOffice% NEQ 0 call :sppoff
wmic path %spp% where (Description like '%%KMSCLIENT%%') get Name %_Nul2% | findstr /i Windows %_Nul1% && (
set WinVL=1
) || (
if %ActWindows% EQU 0 (
  echo.&echo Windows activation is OFF...
  ) else (
  echo.&echo %_winos% %nKMS%
  if defined _eval echo %nEval%
  )
)
if %Off1ce% EQU 0 if %WinVL% EQU 0 exit /b
if %AUR% EQU 0 (
reg delete "HKLM\%SPPk%\%_wApp%" /f %_Nul3%
reg delete "HKLM\%SPPk%\%_oApp%" /f %_Nul3%
)
wmic path %spp% where (Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get Name %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1) || (set _gvlk=0)
set gpr=0
if %winbuild% GEQ 10240 if %SkipKMS38% NEQ 0 if %_gvlk% EQU 1 for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get GracePeriodRemaining /VALUE" %_Nul6%') do set "gpr=%%A"
if %gpr% NEQ 0 if %gpr% GTR 259200 (
set W1nd0ws=0
wmic path %spp% where "ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL" get LicenseFamily %_Nul2% | findstr /i EnterpriseG %_Nul1% && (call set W1nd0ws=1)
)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE"') do set ver=%%A
wmic path %sps% where version='%ver%' call SetKeyManagementServiceMachine MachineName="%KMS_IP%" %_Nul3%
wmic path %sps% where version='%ver%' call SetKeyManagementServicePort %KMS_Port% %_Nul3%
if %W1nd0ws% EQU 0 for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%') get ID /VALUE"') do (set app=%%G&call :sppchkwin)
if %W1nd0ws% EQU 1 if %ActWindows% NEQ 0 for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (ApplicationID='%_wApp%' and Description like '%%KMSCLIENT%%') get ID /VALUE"') do (set app=%%G&call :sppchkwin)
if %W1nd0ws% EQU 1 if %ActWindows% EQU 0 (echo.&echo Windows activation is OFF...)
if %Off1ce% EQU 1 if %ActOffice% NEQ 0 for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (ApplicationID='%_oApp%' and Description like '%%KMSCLIENT%%') get ID /VALUE"') do (set app=%%G&call :sppchkoff)
if %AUR% EQU 0 (
call :cKMS %_Nul3%
call :cREG %_Nul3%
) else (
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 0 %_Nul3%
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 0 %_Nul3%
)
exit /b

:sppoff
set _sC2R=sppoff
set _fC2R=ReturnSPP
set vol_off15=0&set vol_off16=0&set vol_off19=0
wmic path %spp% where (Description like '%%KMSCLIENT%%') get Name > "!_temp!\sppchk.txt" 2>&1
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off19=1)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off16=1)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off15=1)
for %%A in (15,16,19) do if !loc_off%%A! EQU 0 set vol_off%%A=0
if %vol_off16% EQU 1 find /i "Office16MondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
wmic path %spp% where 'ApplicationID="%_oApp%" AND LicenseFamily like "Office16O365%%"' get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set vol_off16=0)
)
if %vol_off15% EQU 1 find /i "OfficeMondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
wmic path %spp% where 'ApplicationID="%_oApp%" AND LicenseFamily like "OfficeO365%%"' get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set vol_off15=0)
)
set loc_offgl=1
if %loc_off19% EQU 0 if %loc_off16% EQU 0 if %loc_off15% EQU 0 (set loc_offgl=0)
if %loc_offgl% EQU 1 set Off1ce=1
set vol_offgl=1
if %vol_off19% EQU 0 if %vol_off16% EQU 0 if %vol_off15% EQU 0 (set vol_offgl=0)
:: mixed Volume + Retail scenario
if %loc_off19% EQU 1 if %vol_off19% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
if %winbuild% GTR 9600 if %loc_off16% EQU 1 if %vol_off16% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
if %winbuild% LEQ 9600 if %loc_off16% EQU 1 if %vol_off16% EQU 0 if %vol_off19% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
if %loc_off15% EQU 1 if %vol_off15% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
:: all Volume scenario
if %vol_offgl% EQU 1 exit /b
set Off1ce=0
:: nothing installed scenario
if %loc_offgl% EQU 0 (echo.&echo No Installed Office 2013/2016/2019 Product Detected...&exit /b)
:: Retail C2R scenario
if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
:ReturnSPP
:: Retail MSI scenario or failed C2R-R2V scenario
echo.
if %loc_off15% EQU 1 if %vol_off15% EQU 0 echo Detected Office 2013 %nKMS%
if %loc_off16% EQU 1 if %vol_off16% EQU 0 echo Detected Office 2016 %nKMS%
if %loc_off19% EQU 1 if %vol_off19% EQU 0 echo Detected Office 2019 %nKMS%
echo Retail Products need to be converted to Volume first.
exit /b

:sppchkoff
wmic path %spp% where ID='%app%' get Name > "!_temp!\sppchk.txt"
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off15% EQU 0 exit /b)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off16% EQU 0 exit /b)
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off19% EQU 0 exit /b)
set _office=1
wmic path %spp% where (PartialProductKey is not NULL) get ID %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
for /f "tokens=3 delims==, " %%G in ('"wmic path %spp% where ID='%app%' get Name /value"') do set OffVer=%%G
call :offchk%OffVer%
exit /b

:sppchkwin
set _office=0
if %winbuild% GEQ 14393 if %_gvlk% EQU 0 wmic path %spp% where (Description like '%%KMSCLIENT%%' and PartialProductKey is not NULL) get Name %_Nul2% | findstr /i Windows %_Nul1% && (set _gvlk=1)
wmic path %spp% where ID='%app%' get LicenseStatus %_Nul2% | findstr "1" %_Nul1% && (echo.&call :activate&exit /b)
wmic path %spp% where (PartialProductKey is not NULL) get ID %_Nul2% | findstr /i "%app%" %_Nul1% && (echo.&call :activate&exit /b)
if %_gvlk% EQU 1 exit /b
if %WinPerm% EQU 1 exit /b
if %winbuild% LSS 10240 (call :winchk&exit /b)
for %%A in (
b71515d9-89a2-4c60-88c8-656fbcca7f3a,af43f7f0-3b1e-4266-a123-1fdb53f4323b,075aca1f-05d7-42e5-a3ce-e349e7be7078
11a37f09-fb7f-4002-bd84-f3ae71d11e90,43f2ab05-7c87-4d56-b27c-44d0f9a3dabd,2cf5af84-abab-4ff0-83f8-f040fb2576eb
6ae51eeb-c268-4a21-9aae-df74c38b586d,ff808201-fec6-4fd4-ae16-abbddade5706,34260150-69ac-49a3-8a0d-4a403ab55763
4dfd543d-caa6-4f69-a95f-5ddfe2b89567,5fe40dd6-cf1f-4cf2-8729-92121ac2e997,903663f7-d2ab-49c9-8942-14aa9e0a9c72
2cc171ef-db48-4adc-af09-7c574b37f139,5b2add49-b8f4-42e0-a77c-adad4efeeeb1
) do (
if /i '%app%' EQU '%%A' exit /b
)
if not defined EditionID (call :winchk&exit /b)
if %winbuild% LSS 14393 (call :winchk&exit /b)
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
if /i '%app%' EQU 'e4db50ea-bda1-4566-b047-0ca50abc6f07' (
wmic path %spp% where 'Description like "%%KMSCLIENT%%"' get ID | findstr /i "ec868e65-fadf-4759-b23e-93fe37f2cc29" %_Nul3% && (exit /b)
)
call :winchk
exit /b

:winchk
if not defined tok (if %winbuild% GEQ 9200 (set "tok=4") else (set "tok=7"))
wmic path %spp% where (LicenseStatus='1' and Description like '%%KMSCLIENT%%') get Name %_Nul2% | findstr /i "Windows" %_Nul3% && (exit /b)
echo.
wmic path %spp% where (LicenseStatus='1' and GracePeriodRemaining='0' and PartialProductKey is not NULL) get Name %_Nul2% | findstr /i "Windows" %_Nul3% && (
set WinPerm=1
)
if %WinPerm% EQU 0 (
wmic path %spp% where "ApplicationID='%_wApp%' and LicenseStatus='1'" get Name %_Nul2% | findstr /i "Windows" %_Nul3% && (
for /f "tokens=%tok% delims=, " %%G in ('"wmic path %spp% where (ApplicationID='%_wApp%' and LicenseStatus='1') get Description /VALUE"') do set "channel=%%G"
  for %%A in (VOLUME_MAK, RETAIL, OEM_DM, OEM_SLP, OEM_COA, OEM_COA_SLP, OEM_COA_NSLP, OEM_NONSLP, OEM) do if /i "%%A"=="!channel!" set WinPerm=1
  )
)
if %WinPerm% EQU 0 (
copy /y %SysPath%\slmgr.vbs "!_temp!\slmgr.vbs" %_Nul3%
cscript //nologo "!_temp!\slmgr.vbs" /xpr %_Nul2% | findstr /i "permanently" %_Nul3% && set WinPerm=1
)
if %WinPerm% EQU 1 (
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where (ApplicationID='%_wApp%' and LicenseStatus='1') get Name /VALUE"') do echo Checking: %%x
echo Product is Permanently Activated.
exit /b
)
call :insKey
exit /b

:RunOSPP
set spp=OfficeSoftwareProtectionProduct
set sps=OfficeSoftwareProtectionService
set Off1ce=0
set RunR2V=0
if %winbuild% LSS 9200 (set "aword=2010/2013/2016/2019") else (set "aword=2010")
if %OsppHook% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
if %winbuild% GEQ 9200 if %loc_off14% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
if %winbuild% GEQ 9200 wmic path %spp% get Description %_Nul2% | findstr /i KMSCLIENT %_Nul1% || (echo.&echo Detected Office %aword% %nKMS%&echo Retail Products need to be converted to Volume first.&exit /b)
if %winbuild% GEQ 9200 set Off1ce=1
if %winbuild% LSS 9200 call :win7off
if %Off1ce% EQU 0 exit /b
if %AUR% EQU 0 (
reg delete "HKLM\%OPPk%\%_oA14%" /f %_Nul3%
reg delete "HKLM\%OPPk%\%_oApp%" /f %_Nul3%
)
for /f "tokens=2 delims==" %%A in ('"wmic path %sps% get Version /VALUE" %_Nul6%') do set ver=%%A
wmic path %sps% where version='%ver%' call SetKeyManagementServiceMachine MachineName="%KMS_IP%" %_Nul3%
wmic path %sps% where version='%ver%' call SetKeyManagementServicePort %KMS_Port% %_Nul3%
for /f "tokens=2 delims==" %%G in ('"wmic path %spp% where (Description like '%%KMSCLIENT%%') get ID /VALUE"') do (set app=%%G&call :osppchk)
if %AUR% EQU 0 (
call :cKMS %_Nul3%
call :cREG %_Nul3%
) else (
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 0 %_Nul3%
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 0 %_Nul3%
)
exit /b

:win7off
set _sC2R=win7off
set _fC2R=ReturnOSPP
set vol_off14=0&set vol_off15=0&set vol_off16=0&set vol_off19=0
wmic path %spp% where (Description like '%%KMSCLIENT%%') get Name > "!_temp!\sppchk.txt" 2>&1
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off19=1)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off16=1)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off15=1)
find /i "Office 14" "!_temp!\sppchk.txt" %_Nul1% && (set vol_off14=1)
for %%A in (14,15,16,19) do if !loc_off%%A! EQU 0 set vol_off%%A=0
if %vol_off16% EQU 1 find /i "Office16MondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
wmic path %spp% where 'ApplicationID="%_oApp%" AND LicenseFamily like "Office16O365%%"' get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set vol_off16=0)
)
if %vol_off15% EQU 1 find /i "OfficeMondoVL_KMS_Client" "!_temp!\sppchk.txt" %_Nul1% && (
wmic path %spp% where 'ApplicationID="%_oApp%" AND LicenseFamily like "OfficeO365%%"' get LicenseFamily %_Nul2% | find /i "O365" %_Nul1% || (set vol_off15=0)
)
set loc_offgl=1
if %loc_off19% EQU 0 if %loc_off16% EQU 0 if %loc_off15% EQU 0 if %loc_off14% EQU 0 (set loc_offgl=0)
if %loc_offgl% EQU 1 set Off1ce=1
set vol_offgl=1
:: mixed Volume + Retail scenario
if %vol_off19% EQU 0 if %vol_off16% EQU 0 if %vol_off15% EQU 0 if %vol_off14% EQU 0 (set vol_offgl=0)
if %loc_off19% EQU 1 if %vol_off19% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
if %loc_off16% EQU 1 if %vol_off16% EQU 0 if %vol_off19% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
if %loc_off15% EQU 1 if %vol_off15% EQU 0 if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
:: all Volume scenario
if %vol_offgl% EQU 1 exit /b
set Off1ce=0
:: nothing installed scenario
if %loc_offgl% EQU 0 (echo.&echo No Installed Office %aword% Product Detected...&exit /b)
:: Retail C2R scenario
if %RunR2V% EQU 0 if %AutoR2V% EQU 1 goto :C2RR2V
:ReturnOSPP
:: Retail MSI scenario or failed C2R-R2V scenario
echo.
if %loc_off14% EQU 1 if %vol_off14% EQU 0 echo Detected Office 2010 %nKMS%
if %loc_off15% EQU 1 if %vol_off15% EQU 0 echo Detected Office 2013 %nKMS%
if %loc_off16% EQU 1 if %vol_off16% EQU 0 echo Detected Office 2016 %nKMS%
if %loc_off19% EQU 1 if %vol_off19% EQU 0 echo Detected Office 2019 %nKMS%
echo Retail Products need to be converted to Volume first.
exit /b

:osppchk
wmic path %spp% where ID='%app%' get Name > "!_temp!\sppchk.txt"
find /i "Office 14" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off14% EQU 0 exit /b)
find /i "Office 15" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off15% EQU 0 exit /b)
find /i "Office 16" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off16% EQU 0 exit /b)
find /i "Office 19" "!_temp!\sppchk.txt" %_Nul1% && (if %loc_off19% EQU 0 exit /b)
set _office=0
wmic path %spp% where (PartialProductKey is not NULL) get ID | findstr /i "%app%" %_Nul3% && (echo.&call :activate&exit /b)
for /f "tokens=3 delims==, " %%G in ('"wmic path %spp% where ID='%app%' get Name /value"') do set OffVer=%%G
call :offchk%OffVer%
exit /b

:offchk
set ls=0
set ls2=0
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (LicenseFamily='Office%~2') get LicenseStatus /VALUE" %_Nul6%') do set /a ls=%%A
if "%~4" NEQ "" (
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (LicenseFamily='Office%~4') get LicenseStatus /VALUE" %_Nul6%') do set /a ls2=%%A
)
if "%ls2%" EQU "1" (
echo Checking: %~5
echo Product is Permanently Activated.
exit /b
)
if "%ls%" EQU "1" (
echo Checking: %~3
echo Product is Permanently Activated.
exit /b
)
call :insKey
exit /b

:offchk19
if /i '%app%' EQU '0bc88885-718c-491d-921f-6f214349e79c' exit /b
if /i '%app%' EQU 'fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9' exit /b
if /i '%app%' EQU '500f6619-ef93-4b75-bcb4-82819998a3ca' exit /b
if /i '%app%' EQU '85dd8b5f-eaa4-4af3-a628-cce9e77c9a03' (
call :offchk "%app%" "19ProPlus2019VL_MAK_AE" "Office ProPlus 2019" "19ProPlus2019XC2RVL_MAKC2R" "Office ProPlus 2019 C2R"
exit /b
)
if /i '%app%' EQU '6912a74b-a5fb-401a-bfdb-2e3ab46f4b02' (
call :offchk "%app%" "19Standard2019VL_MAK_AE" "Office Standard 2019"
exit /b
)
if /i '%app%' EQU '2ca2bf3f-949e-446a-82c7-e25a15ec78c4' (
call :offchk "%app%" "19ProjectPro2019VL_MAK_AE" "Project Pro 2019" "19ProjectPro2019XC2RVL_MAKC2R" "Project Pro 2019 C2R"
exit /b
)
if /i '%app%' EQU '1777f0e3-7392-4198-97ea-8ae4de6f6381' (
call :offchk "%app%" "19ProjectStd2019VL_MAK_AE" "Project Standard 2019"
exit /b
)
if /i '%app%' EQU '5b5cf08f-b81a-431d-b080-3450d8620565' (
call :offchk "%app%" "19VisioPro2019VL_MAK_AE" "Visio Pro 2019" "19VisioPro2019XC2RVL_MAKC2R" "Visio Pro 2019 C2R"
exit /b
)
if /i '%app%' EQU 'e06d7df3-aad0-419d-8dfb-0ac37e2bdf39' (
call :offchk "%app%" "19VisioStd2019VL_MAK_AE" "Visio Standard 2019"
exit /b
)
call :insKey
exit /b

:offchk16
if /i '%app%' EQU 'd450596f-894d-49e0-966a-fd39ed4c4c64' (
call :offchk "%app%" "16ProPlusVL_MAK" "Office ProPlus 2016"
exit /b
)
if /i '%app%' EQU 'dedfa23d-6ed1-45a6-85dc-63cae0546de6' (
call :offchk "%app%" "16StandardVL_MAK" "Office Standard 2016"
exit /b
)
if /i '%app%' EQU '4f414197-0fc2-4c01-b68a-86cbb9ac254c' (
call :offchk "%app%" "16ProjectProVL_MAK" "Project Pro 2016"
exit /b
)
if /i '%app%' EQU 'da7ddabc-3fbe-4447-9e01-6ab7440b4cd4' (
call :offchk "%app%" "16ProjectStdVL_MAK" "Project Standard 2016"
exit /b
)
if /i '%app%' EQU '6bf301c1-b94a-43e9-ba31-d494598c47fb' (
call :offchk "%app%" "16VisioProVL_MAK" "Visio Pro 2016"
exit /b
)
if /i '%app%' EQU 'aa2a7821-1827-4c2c-8f1d-4513a34dda97' (
call :offchk "%app%" "16VisioStdVL_MAK" "Visio Standard 2016"
exit /b
)
if /i '%app%' EQU '829b8110-0e6f-4349-bca4-42803577788d' (
call :offchk "%app%" "16ProjectProXC2RVL_MAKC2R" "Project Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU 'cbbaca45-556a-4416-ad03-bda598eaa7c8' (
call :offchk "%app%" "16ProjectStdXC2RVL_MAKC2R" "Project Standard 2016 C2R"
exit /b
)
if /i '%app%' EQU 'b234abe3-0857-4f9c-b05a-4dc314f85557' (
call :offchk "%app%" "16VisioProXC2RVL_MAKC2R" "Visio Pro 2016 C2R"
exit /b
)
if /i '%app%' EQU '361fe620-64f4-41b5-ba77-84f8e079b1f7' (
call :offchk "%app%" "16VisioStdXC2RVL_MAKC2R" "Visio Standard 2016 C2R"
exit /b
)
call :insKey
exit /b

:offchk15
if /i '%app%' EQU 'b322da9c-a2e2-4058-9e4e-f59a6970bd69' (
call :offchk "%app%" "ProPlusVL_MAK" "Office ProPlus 2013"
exit /b
)
if /i '%app%' EQU 'b13afb38-cd79-4ae5-9f7f-eed058d750ca' (
call :offchk "%app%" "StandardVL_MAK" "Office Standard 2013"
exit /b
)
if /i '%app%' EQU '4a5d124a-e620-44ba-b6ff-658961b33b9a' (
call :offchk "%app%" "ProjectProVL_MAK" "Project Pro 2013"
exit /b
)
if /i '%app%' EQU '427a28d1-d17c-4abf-b717-32c780ba6f07' (
call :offchk "%app%" "ProjectStdVL_MAK" "Project Standard 2013"
exit /b
)
if /i '%app%' EQU 'e13ac10e-75d0-4aff-a0cd-764982cf541c' (
call :offchk "%app%" "VisioProVL_MAK" "Visio Pro 2013"
exit /b
)
if /i '%app%' EQU 'ac4efaf0-f81f-4f61-bdf7-ea32b02ab117' (
call :offchk "%app%" "VisioStdVL_MAK" "Visio Standard 2013"
exit /b
)
call :insKey
exit /b

:offchk14
set "vPrem="&set "vPro="
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (LicenseFamily='OfficeVisioPrem-MAK') get LicenseStatus /VALUE" %_Nul6%') do set vPrem=%%A
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where (LicenseFamily='OfficeVisioPro-MAK') get LicenseStatus /VALUE" %_Nul6%') do set vPro=%%A
if /i '%app%' EQU '6f327760-8c5c-417c-9b61-836a98287e0c' (
call :offchk "%app%" "ProPlus-MAK" "Office ProPlus 2010" "ProPlusAcad-MAK" "Office Professional Academic 2010"
exit /b
)
if /i '%app%' EQU '9da2a678-fb6b-4e67-ab84-60dd6a9c819a' (
call :offchk "%app%" "Standard-MAK" "Office Standard 2010" "StandardAcad-MAK"  "Office Standard Academic 2010"
exit /b
)
if /i '%app%' EQU 'ea509e87-07a1-4a45-9edc-eba5a39f36af' (
call :offchk "%app%" "SmallBusBasics-MAK" "Office Small Business Basics 2010"
exit /b
)
if /i '%app%' EQU 'df133ff7-bf14-4f95-afe3-7b48e7e331ef' (
call :offchk "%app%" "ProjectPro-MAK" "Project Pro 2010"
exit /b
)
if /i '%app%' EQU '5dc7bf61-5ec9-4996-9ccb-df806a2d0efe' (
call :offchk "%app%" "ProjectStd-MAK" "Project Standard 2010" "ProjectStd-MAK2" "Project Standard 2010"
exit /b
)
if /i '%app%' EQU '92236105-bb67-494f-94c7-7f7a607929bd' (
call :offchk "%app%" "VisioPrem-MAK" "Visio Premium 2010" "VisioPro-MAK" "Visio Pro 2010"
exit /b
)
if defined vPrem exit /b
if /i '%app%' EQU 'e558389c-83c3-4b29-adfe-5e4d7f46c358' (
call :offchk "%app%" "VisioPro-MAK" "Visio Pro 2010" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
if defined vPro exit /b
if /i '%app%' EQU '9ed833ff-4f92-4f36-b370-8683a4f13275' (
call :offchk "%app%" "VisioStd-MAK" "Visio Standard 2010"
exit /b
)
call :insKey
exit /b

:officeLoc
set loc_off%1=0
if %1 EQU 19 (
if defined _C16R reg query %_C16R% /v ProductReleaseIds %_Nul2% | findstr 2019 %_Nul1% && set loc_off%1=1
exit /b
)

for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set loc_off%1=1
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot /v Path" %_Nul6%') do if exist "%%b\OSPP.VBS" set loc_off%1=1

if %1 EQU 16 if defined _C16R (
for /f "skip=2 tokens=2*" %%a in ('reg query %_C16R% /v ProductReleaseIds') do echo %%b> "!_temp!\c2rchk.txt"
for %%a in (%_V16Ids%,ProjectProX,ProjectStdX,VisioProX,VisioStdX) do (
  findstr /I /C:"%%aVolume" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
for %%a in (%_R16Ids%) do (
  findstr /I /C:"%%aRetail" "!_temp!\c2rchk.txt" %_Nul1% && set loc_off%1=1
  )
exit /b
)

if %1 EQU 15 if defined _C15R (
set loc_off%1=1
exit /b
)

if exist "%ProgramFiles%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if exist "%ProgramW6432%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
if exist "%ProgramFiles(x86)%\Microsoft Office\Office%1\OSPP.VBS" set loc_off%1=1
exit /b

:insKey
echo.
set "_key="
for /f "tokens=2 delims==" %%A in ('"wmic path %spp% where ID='%app%' get Name /VALUE"') do echo Installing Key for: %%A
call :keys %app%
if "%_key%"=="" (echo Could not find matching KMS Client key&exit /b)
wmic path %sps% where version='%ver%' call InstallProductKey ProductKey="%_key%" %_Nul3%
set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
cmd /c exit /b %ERRORCODE%
echo Failed: 0x!=ExitCode!
exit /b
)

:activate
wmic path %spp% where ID='%app%' call ClearKeyManagementServiceMachine %_Nul3%
wmic path %spp% where ID='%app%' call ClearKeyManagementServicePort %_Nul3%
if %W1nd0ws% EQU 0 if %_office% EQU 0 if %sps% EQU SoftwareLicensingService (
wmic path %spp% where ID='%app%' call SetKeyManagementServiceMachine MachineName="127.0.0.2" %_Nul3%
wmic path %spp% where ID='%app%' call SetKeyManagementServicePort %KMS_Port% %_Nul3%
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%app%' get Name /VALUE"') do echo Checking: %%x
echo Product is KMS 2038 Activated.
exit /b
)
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%app%' get Name /VALUE"') do echo Activating: %%x
wmic path %spp% where ID='%app%' call Activate %_Nul3%
call set ERRORCODE=%ERRORLEVEL%
if %ERRORCODE% NEQ 0 (
if %sps% EQU SoftwareLicensingService (call :StopService sppsvc) else (call :StopService osppsvc)
wmic path %spp% where ID='%app%' call Activate %_Nul3%
call set ERRORCODE=!ERRORLEVEL!
)
if %sps% EQU SoftwareLicensingService wmic path %sps% where version='%ver%' call RefreshLicenseStatus %_Nul3%
set gpr=0
set gpr2=0
for /f "tokens=2 delims==" %%x in ('"wmic path %spp% where ID='%app%' get GracePeriodRemaining /VALUE"') do (set gpr=%%x&set /a gpr2=%%x/1440)
if %gpr% EQU 43200 if %_office% EQU 0 if %winbuild% GEQ 9200 (
echo Product Activation Successful
echo Remaining Period: 30 days ^(%gpr% minutes^)
exit /b
)
if %gpr% EQU 64800 (
echo Product Activation Successful
echo Remaining Period: 45 days ^(%gpr% minutes^)
exit /b
)
if %gpr% GTR 259200 if %Win10Gov% EQU 1 (
echo Product Activation Successful
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
exit /b
)
if %gpr% EQU 259200 (
echo Product Activation Successful
) else (
cmd /c exit /b %ERRORCODE%
echo Product Activation Failed: 0x!=ExitCode!
)
echo Remaining Period: %gpr2% days ^(%gpr% minutes^)
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
if %SkipKMS38% EQU 0 set "_para=!_para! /x"
goto :DoDebug
)
if %_verb% EQU 1 (
if %Silent% EQU 0 if %_Debug% EQU 0 (
mode con cols=100 lines=34
%_Nul3% powershell -noprofile -exec bypass -c "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.height=300;$W.buffersize=$B;}"
)
echo.&echo %line3%&echo.
echo Installing Local KMS Emulator...
)
if %winbuild% GEQ 9600 (
  WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Add ExclusionPath="%SystemRoot%\System32\SppExtComObjHook.dll" %_Nul3% && set "AddExc= and Windows Defender exclusion"
)
if %_verb% EQU 1 (
echo.
echo Adding File%AddExc%...
echo %SystemRoot%\System32\SppExtComObjHook.dll
)
if %AUR% EQU 1 (
call :StopService sppsvc
if %OsppHook% NEQ 0 call :StopService osppsvc
)
for %%# in (SppExtComObjHookAvrf.dll,SppExtComObjHook.dll,SppExtComObjPatcher.dll,SppExtComObjPatcher.exe) do if exist "%SysPath%\%%#" (
	del /f /q "%SysPath%\%%#" %_Nul3%
)
pushd %SysPath%
%_Nul3% powershell -noprofile -exec bypass -c "$f=[io.file]::ReadAllText('%~f0') -split ':%xOS%dll\:.*';iex ($f[1]);X 1;"
popd
if not exist "%SysPath%\SppExtComObjHook.dll" goto :E_DLL
if %_verb% EQU 1 (
echo.
echo Adding Registry Keys...
)
if %SSppHook% NEQ 0 call :CreateIFEOEntry %SppVer%
if %AUR% EQU 1 (call :CreateIFEOEntry osppsvc.exe) else (if %OsppHook% NEQ 0 call :CreateIFEOEntry osppsvc.exe)
if %AUR% EQU 1 if %OSType% EQU Win7 (
call :CreateIFEOEntry SppExtComObj.exe
if %SSppHook% NEQ 0 if not exist %w7inf% (
  if %_verb% EQU 1 (echo.&echo Adding migration fail-safe...&echo %w7inf%)
  if not exist "%SystemRoot%\Migration\WTR" md "%SystemRoot%\Migration\WTR"
  (
  echo [WTR]
  echo Name="KMS_VL_ALL"
  echo.
  echo [WTR.W8]
  echo NotifyUser="No"
  echo.
  echo [System.Registry]
  echo "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\sppsvc.exe [*]"
  )>%w7inf%
  )
)
if %AUR% EQU 1 if %OSType% EQU Win8 call :CreateTask
if %_verb% EQU 1 echo.&echo %line3%&echo.
goto :%_rtr%

:RemoveHook
if %_dDbg%==Yes (
set "_para=/d /r"
goto :DoDebug
)
if %winbuild% GEQ 9600 (
  reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform" /v "NoGenTicket" /f %_Nul3%
  WMIC /NAMESPACE:\\root\Microsoft\Windows\Defender PATH MSFT_MpPreference call Remove ExclusionPath="%SystemRoot%\System32\SppExtComObjHook.dll" %_Nul3% && set "RemExc= and Windows Defender exclusions"
)
if %_verb% EQU 1 (
if %Silent% EQU 0 if %_Debug% EQU 0 (
mode con cols=100 lines=34
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
echo Removing Schedule Task...
echo %_TaskEx%
)
schtasks /delete /f /tn "%_TaskEx%" %_Nul3%
)
goto :eof

:CreateIFEOEntry
if %_verb% EQU 1 (
echo [%IFEO%\%1]
)
reg delete "%IFEO%\%1" /f /v Debugger %_Nul3%
reg add "%IFEO%\%1" /f /v VerifierDlls /t REG_SZ /d "SppExtComObjHook.dll" %_Nul3%
reg add "%IFEO%\%1" /f /v GlobalFlag /t REG_DWORD /d 256 %_Nul3%
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
reg delete "%IFEO%\%1" /f %_Nul3%
goto :eof
)
if %OsppHook% EQU 0 (
reg delete "%IFEO%\%1" /f %_Nul3%
)
if %OsppHook% NEQ 0 for %%A in (Debugger,VerifierDlls,GlobalFlag,KMS_Emulation,KMS_ActivationInterval,KMS_RenewalInterval,Office2010,Office2013,Office2016,Office2019) do reg delete "%IFEO%\%1" /v %%A /f %_Nul3%
reg delete "HKLM\%OPPk%" /f /v KeyManagementServiceName %_Nul3%
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort %_Nul3%
goto :eof

:UpdateIFEOEntry
reg query "%IFEO%\%1" /v KMS_Emulation %_Nul3% || goto :eof
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
reg add "HKLM\%OPPk%" /f /v KeyManagementServiceName /t REG_SZ /d %KMS_IP% %_Nul3%
reg add "HKLM\%OPPk%" /f /v KeyManagementServicePort /t REG_SZ /d %KMS_Port% %_Nul3%
)
goto :eof

:cKMS
if not "%1"=="" (
set spp=%1
set sps=%2
for /f "tokens=2 delims==" %%A in ('"wmic path %2 get Version /VALUE"') do set ver=%%A
)
wmic path %sps% where version='%ver%' call ClearKeyManagementServiceMachine
wmic path %sps% where version='%ver%' call ClearKeyManagementServicePort
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceDnsPublishing 1
wmic path %sps% where version='%ver%' call DisableKeyManagementServiceHostCaching 1
goto :eof

:cREG
reg delete "HKLM\%SPPk%\%_wApp%" /f
reg delete "HKLM\%SPPk%\%_oApp%" /f
reg delete "HKLM\%SPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%SPPk%" /f /v KeyManagementServicePort
reg delete "HKU\S-1-5-20\%SPPk%\%_wApp%" /f
reg delete "HKU\S-1-5-20\%SPPk%\%_oApp%" /f
reg delete "HKLM\%OPPk%\%_oA14%" /f
reg delete "HKLM\%OPPk%\%_oApp%" /f
reg delete "HKLM\%OPPk%" /f /v KeyManagementServiceName
reg delete "HKLM\%OPPk%" /f /v KeyManagementServicePort
if %OsppHook% EQU 0 (
reg delete "HKLM\%OPPk%" /f
reg delete "HKU\S-1-5-20\%OPPk%" /f
)
goto :eof

:cCache
echo.
if %_verb% EQU 0 echo %line3%&echo.
echo Clearing KMS Cache...
call :cKMS SoftwareLicensingProduct SoftwareLicensingService %_Nul3%
if %OsppHook% NEQ 0 call :cKMS OfficeSoftwareProtectionProduct OfficeSoftwareProtectionService %_Nul3%
call :cREG %_Nul3%
if %Unattend% NEQ 0 goto :TheEnd
echo.&echo %line3%&echo.
echo Press any key to continue...
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
%_Nul3% powershell -noprofile -exec bypass -c "$f=[io.file]::ReadAllText('%~f0') -split ':spptask\:.*';iex ($f[1]);"
popd
if exist "!_temp!\SvcTrigger.xml" (
  schtasks /create /tn "%_TaskEx%" /xml "!_temp!\SvcTrigger.xml" /f %_Nul3%
  del /f /q "!_temp!\SvcTrigger.xml" %_Nul3%
  )
)
schtasks /query /tn "%_TaskEx%" %_Nul3% && if %_verb% EQU 1 (
echo.
echo Adding Schedule Task...
echo %_TaskEx%
)
goto :eof

:CreateReadMe
if exist "!_temp!\ReadMe.html" del /f /q "!_temp!\ReadMe.html"
pushd %_temp%
%_Nul3% powershell -noprofile -exec bypass -c "$f=[io.file]::ReadAllText('%~f0') -split ':readme\:.*';iex ($f[1]);"
popd
if exist "!_temp!\ReadMe.html" start "" "!_temp!\ReadMe.html"
goto :eof

:CreateOEM
cls
set "_oem=!_work!"
copy /y nul "!_work!\#.rw" 1>nul 2>nul && (if exist "!_work!\#.rw" del /f /q "!_work!\#.rw") || (set "_oem=!_dsk!")
if exist "!_oem!\$OEM$\$$\Setup\Scripts\setupcomplete.cmd" (
echo.&echo %line3%&echo.
echo $OEM$ Folder already exist...
echo "!_oem!\$OEM$"
echo.
echo manually remove it if you wish to create a fresh copy.
echo.&echo %line3%&echo.
echo Press any key to continue...
pause >nul
goto :eof
)
echo.&echo %line3%&echo.
echo $OEM$ Folder Created...
echo.
echo "!_oem!\$OEM$"
echo.&echo %line3%&echo.
if not exist "!_oem!\$OEM$\$$\Setup\Scripts\KMS_VL_ALL_AIO.cmd" mkdir ""!_oem!\$OEM$\$$\Setup\Scripts"
copy /y "%~f0" "!_oem!\$OEM$\$$\Setup\Scripts\KMS_VL_ALL_AIO.cmd" %_Nul3%
(
echo @echo off
echo call %%~dp0KMS_VL_ALL_AIO.cmd /s /a
echo cd \
echo ^(goto^) 2^>nul^&rd /s /q "%%~dp0"
)>"!_oem!\$OEM$\$$\Setup\Scripts\setupcomplete.cmd"
echo.
echo Press any key to continue...
pause >nul
goto :eof

:C2RR2V
set RunR2V=1
set "_SLMGR=%SysPath%\slmgr.vbs"
if %_Debug% EQU 0 (
set "_cscript=cscript //Nologo //B"
) else (
set "_cscript=cscript //Nologo"
)
sc query ClickToRunSvc %_Nul3%
set error1=%errorlevel%
sc query OfficeSvc %_Nul3%
set error2=%errorlevel%
if %error1% EQU 1060 if %error2% EQU 1060 (
goto :%_fC2R%
)
set _Office16=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses16\*.xrm-ms" (
  set _Office16=1&set "_OSPPVBS=%%b\Office16\OSPP.VBS"
)
if exist "%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS" (
  set _Office16=1&set "_OSPPVBS=%ProgramFiles%\Microsoft Office\Office16\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office16\OSPP.VBS" (
  set _Office16=1&set "_OSPPVBS=%ProgramW6432%\Microsoft Office\Office16\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" (
  set _Office16=1&set "_OSPPVBS=%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS"
)
set _Office15=0
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if exist "%%b\root\Licenses\*.xrm-ms" (
  set _Office15=1
)
if exist "%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP15VBS=%ProgramFiles%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP15VBS=%ProgramW6432%\Microsoft Office\Office15\OSPP.VBS"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS" (
  set _Office15=1&set "_OSPP15VBS=%ProgramFiles(x86)%\Microsoft Office\Office15\OSPP.VBS"
)
if %_Office16% EQU 0 if %_Office15% EQU 0 (
goto :%_fC2R%
)
if %_Office16% EQU 1 call :Reg16istry
if %_Office15% EQU 1 call :Reg15istry
goto :CheckC2R

:Reg16istry
set "_InstallRoot="
set "_GUID="
set "_ProductIds="
set "_Config="
set "_PRIDs="
set "_LicensesPath="
set "_Integrator="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_InstallRoot=%%b\root")
if not "%_InstallRoot%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v PackageGUID" %_Nul6%') do if not errorlevel 1 (set "_GUID=%%b")
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_ProductIds=%%b")
  set "_Config=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
  set "_PRIDs=HKLM\SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"
)
set "_LicensesPath=%_InstallRoot%\Licenses16"
set "_Integrator=%_InstallRoot%\integration\integrator.exe"
for /f "skip=2 tokens=2*" %%a in ('"reg query %_PRIDs% /v ActiveConfiguration" %_Nul6%') do set "_PRIDs=%_PRIDs%\%%b"
exit /b

:Reg15istry
set "_Install15Root="
set "_Product15Ids="
set "_Con15fig="
set "_PR15IDs="
set "_OSPP15Ready="
set "_Licenses15Path="
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath" %_Nul6%') do if not errorlevel 1 (set "_Install15Root=%%b\root")
if not "%_Install15Root%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration /v ProductReleaseIds"
  set "_PR15IDs=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration"
)
set "_OSPP15ReadT=REG_SZ"
if "%_Product15Ids%"=="" (
  for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid" %_Nul6%') do if not errorlevel 1 (set "_Product15Ids=%%b")
  set "_Con15fig=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun\propertyBag /v productreleaseid"
  set "_OSPP15Ready=HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun"
  set "_OSPP15ReadT=REG_DWORD"
)
set "_Licenses15Path=%_Install15Root%\Licenses"
exit /b

:CheckC2R
if %_Office16% EQU 1 (
if "%_ProductIds%"=="" (
if %_Office15% EQU 0 goto :%_fC2R%
)
if not exist "!_LicensesPath!\*.xrm-ms" (
if %_Office15% EQU 0 goto :%_fC2R%
)
if not exist "!_Integrator!" (
if %_Office15% EQU 0 goto :%_fC2R%
)
if exist "!_LicensesPath!\Word2019VL_KMS_Client_AE*.xrm-ms" (set "_tag=2019") else (set "_tag=")
)
if %_Office15% EQU 1 (
if "%_Product15Ids%"=="" (
if %_Office16% EQU 0 goto :%_fC2R%
)
if not exist "!_Licenses15Path!\*.xrm-ms" (
if %_Office16% EQU 0 goto :%_fC2R%
)
if %winbuild% LSS 9200 if not exist "!_OSPP15VBS!" (
if %_Office16% EQU 0 goto :%_fC2R%
)
)

:PassedC2R
if %winbuild% GEQ 9200 (
set _spp=SoftwareLicensingProduct
set _sps=SoftwareLicensingService
set "_vbsi=%_SLMGR% /ilc "
) else (
set _spp=OfficeSoftwareProtectionProduct
set _sps=OfficeSoftwareProtectionService
set _vbsi="!_OSPP15VBS!" /inslic:
)
set "_wmi="
for /f "tokens=2 delims==" %%# in ('"wmic path %_sps% get version /value" %_Nul6%') do if not errorlevel 1 set "_wmi=%%#"
if not defined _wmi (
goto :%_fC2R%
)
echo.
echo Checking Office Click-to-Run...
set _Retail=0
wmic path %_spp% where (ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey IS NOT NULL) get Description %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
find /i "RETAIL channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "RETAIL(MAK) channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
find /i "TIMEBASED_SUB channel" "!_temp!\crvRetail.txt" %_Nul1% && set _Retail=1
set "_copp="
if exist "%SysPath%\msvcr100.dll" (
set _copp=%_temp%
) else if exist "!_InstallRoot!\vfs\System\msvcr100.dll" (
set _copp="!_InstallRoot!\vfs\System"
) else if exist "!_Install15Root!\vfs\System\msvcr100.dll" (
set _copp="!_Install15Root!\vfs\System"
) else if exist "%SystemRoot%\SysWOW64\msvcr100.dll" (
set _copp=%_temp%
set xBit=x86
) else if exist "!_InstallRoot!\vfs\SystemX86\msvcr100.dll" (
set _copp="!_InstallRoot!\vfs\SystemX86"
set xBit=x86
) else if exist "!_Install15Root!\vfs\SystemX86\msvcr100.dll" (
set _copp="!_Install15Root!\vfs\SystemX86"
set xBit=x86
)
if %_Retail% EQU 0 if defined _copp (
pushd %_copp%
%_Nul3% powershell -noprofile -exec bypass -c "$d='!cd!';$f=[io.file]::ReadAllText('%~f0') -split ':%xBit%exe\:.*';iex ($f[1]);X 1;"
%_Nul3% cleanospp.exe -Licenses
%_Nul3% del /f /q cleanospp.exe
popd
)
set _C2rMsg=0
set _O16O365=0
if %_Retail% EQU 1 wmic path %_spp% where (ApplicationID='%_oApp%' AND LicenseStatus='1' AND PartialProductKey IS NOT NULL) get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvRetail.txt"
wmic path %_spp% where (ApplicationID='%_oApp%') get LicenseFamily %_Nul2% |findstr /V /R "^$" >"!_temp!\crvVolume.txt" 2>&1

if %_Office16% EQU 0 goto :R15V

set _O19Ids=ProPlus2019,ProjectPro2019,VisioPro2019,Standard2019,ProjectStd2019,VisioStd2019,Access2019,SkypeforBusiness2019
set _O16Ids=ProjectPro,VisioPro,Standard,ProjectStd,VisioStd,Access,SkypeforBusiness
set _A19Ids=Excel2019,Outlook2019,PowerPoint2019,Publisher2019,Word2019
set _A16Ids=Excel,Outlook,PowerPoint,Publisher,Word
set _V19Ids=%_O19Ids%,%_A19Ids%
set _V16Ids=Mondo,%_O16Ids%,%_A16Ids%,OneNote
set _R16Ids=%_V16Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem,O365EduCloud
set _RetIds=%_V19Ids%,Professional2019,HomeBusiness2019,HomeStudent2019,%_R16Ids%

echo %_ProductIds%>"!_temp!\crvProductIds.txt"
for %%a in (%_RetIds%,ProPlus) do (
set _%%a=0
)
for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V19Ids%) do (
findstr /I /C:"%%aVolume" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office19%%aVL_KMS_Client" "!_temp!\crvVolume.txt" %_Nul1% && (set _%%a=0) || (set _%%a=1)
  )
)
for %%a in (%_V16Ids%) do (
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
if %_Retail% EQU 1 for %%a in (%_RetIds%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProductIds.txt" %_Nul1% && (
  find /i "Office16%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aE5R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aEDUR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office16%%aXC2RVL_MAKC2R" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office19%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  )
)
if %_Retail% EQU 1 reg query %_PRIDs%\ProPlusRetail.16 %_Nul3% && (
  find /i "Office16ProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "Office16ProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
)

if %_C2rMsg% NEQ 2 for %%a in (%_RetIds%,ProPlus) do if !_%%a! EQU 1 (
set _C2rMsg=1
)
if %_C2rMsg% EQU 1 (
echo Converting Retail-to-Volume:
set _C2rMsg=2
)

if %_C2rMsg% NEQ 2 (if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R))

if !_Mondo! EQU 1 (
call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 (
echo O365ProPlus 2016 Suite -^> Mondo 2016 Licenses
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365Business! EQU 1 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365Business 2016 Suite -^> Mondo 2016 Licenses
call :InsLic O365Business NCHRJ-3VPGW-X73DM-6B36K-3RQ6B
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2016 Suite -^> Mondo 2016 Licenses
call :InsLic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2016 Suite -^> Mondo 2016 Licenses
call :InsLic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365EduCloud! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365Business! EQU 0 if !_O365ProPlus! EQU 0 (
set _O365ProPlus=1
echo O365EduCloud 2016 Suite -^> Mondo 2016 Licenses
call :InsLic O365EduCloud 8843N-BCXXD-Q84H8-R4Q37-T3CPT
if !_Mondo! EQU 0 call :InsLic Mondo
)
if !_O365ProPlus! EQU 1 set _O16O365=1
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 (
echo Mondo 2016 Suite
call :InsLic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)
)
if !_ProPlus2019! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2019 Suite -^> ProPlus%_tag% Licenses
call :InsLic ProPlus%_tag%
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 (
echo ProPlus 2016 Suite -^> ProPlus%_tag% Licenses
call :InsLic ProPlus%_tag%
)
if !_Professional2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2019 Suite -^> ProPlus%_tag% Licenses
call :InsLic ProPlus%_tag%
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 (
echo Professional 2016 Suite -^> ProPlus%_tag% Licenses
call :InsLic ProPlus%_tag%
)
if !_Standard2019! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
echo Standard 2019 Suite -^> Standard2019 Licenses
call :InsLic Standard2019
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 (
echo Standard 2016 Suite -^> Standard%_tag% Licenses
call :InsLic Standard%_tag%
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a2019! EQU 1 (
echo %%a 2019 SKU -^> %%a%_tag% Licenses
if defined _tag (call :InsLic %%a2019) else (call :InsLic %%a)
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
if !_%%a2019! EQU 0 (
  echo %%a 2016 SKU -^> %%a%_tag% Licenses
  call :InsLic %%a%_tag%
  )
)
for %%a in (HomeBusiness2019,HomeStudent2019) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  set _Standard2019=1
  echo %%a Suite -^> Standard2019 Licenses
  call :InsLic Standard2019
  )
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2019! EQU 0 (
  set _Standard2019=1
  echo %%a 2016 Suite -^> Standard%_tag% Licenses
  call :InsLic Standard%_tag%
  )
)
for %%a in (%_A19Ids%,OneNote) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 (
  echo %%a App
  call :InsLic %%a
  )
)
for %%a in (%_A16Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_Standard2019! EQU 0 if !_Standard! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App
  call :InsLic %%a%_tag%
  )
)
for %%a in (Access2019) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 (
  echo %%a App
  call :InsLic %%a
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_Professional2019! EQU 0 if !_Professional! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App
  call :InsLic %%a%_tag%
  )
)
for %%a in (SkypeforBusiness2019) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 (
  echo %%a App
  call :InsLic %%a
  )
)
for %%a in (SkypeforBusiness) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus2019! EQU 0 if !_ProPlus! EQU 0 if !_%%a2019! EQU 0 (
  echo %%a 2016 App
  call :InsLic %%a%_tag%
  )
)
if %_Office15% EQU 1 (goto :R15V) else (goto :GVLKC2R)

:R15V
for %%# in ("!_Licenses15Path!\client-issuance-*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
%_cscript% %_vbsi%"!_Licenses15Path!\pkeyconfig-office.xrm-ms"

set _O15Ids=Standard,ProjectPro,VisioPro,ProjectStd,VisioStd,Access,Lync
set _A15Ids=Excel,Groove,InfoPath,OneNote,Outlook,PowerPoint,Publisher,Word
set _R15Ids=SPD,Mondo,%_O15Ids%,%_A15Ids%,Professional,HomeBusiness,HomeStudent,O365ProPlus,O365Business,O365SmallBusPrem,O365HomePrem
set _V15Ids=Mondo,%_O15Ids%,%_A15Ids%

echo %_Product15Ids%>"!_temp!\crvProduct15s.txt"
for %%a in (%_R15Ids%,ProPlus) do (
set _%%a=0
)
for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && set _%%a=1
)
for %%a in (%_V15Ids%) do (
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
if %_Retail% EQU 1 for %%a in (%_R15Ids%) do (
findstr /I /C:"%%aRetail" "!_temp!\crvProduct15s.txt" %_Nul1% && (
  find /i "Office%%aR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_Sub" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aR_PIN" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aCO365R_" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  find /i "Office%%aVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _%%a=0
  )
)
if %_Retail% EQU 1 reg query %_PR15IDs%\Active\ProPlusRetail\x-none %_Nul3% && (
  find /i "OfficeProPlusR_Retail" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusR_OEM" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusMSDNR_" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
  find /i "OfficeProPlusVL_MAK" "!_temp!\crvRetail.txt" %_Nul1% && set _ProPlus=0
)

if %_C2rMsg% NEQ 2 for %%a in (%_R15Ids%,ProPlus) do if !_%%a! EQU 1 (
set _C2rMsg=1
)
if %_C2rMsg% EQU 1 (
echo Converting Retail-to-Volume:
set _C2rMsg=2
)

if %_C2rMsg% NEQ 2 goto :GVLKC2R

if !_Mondo! EQU 1 (
call :Ins15Lic Mondo
)
if !_O365ProPlus! EQU 1 if !_O16O365! EQU 0 (
echo O365ProPlus 2013 Suite -^> Mondo 2013 Licenses
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365SmallBusPrem! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365SmallBusPrem 2013 Suite -^> Mondo 2013 Licenses
call :Ins15Lic O365SmallBusPrem 3FBRX-NFP7C-6JWVK-F2YGK-H499R
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365HomePrem! EQU 1 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365HomePrem 2013 Suite -^> Mondo 2013 Licenses
call :Ins15Lic O365HomePrem 9FNY8-PWWTY-8RY4F-GJMTV-KHGM9
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_O365Business! EQU 1 if !_O365HomePrem! EQU 0 if !_O365SmallBusPrem! EQU 0 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
set _O365ProPlus=1
echo O365Business 2013 Suite -^> Mondo 2013 Licenses
call :Ins15Lic O365Business MCPBN-CPY7X-3PK9R-P6GTT-H8P8Y
if !_Mondo! EQU 0 call :Ins15Lic Mondo
)
if !_Mondo! EQU 1 if !_O365ProPlus! EQU 0 if !_O16O365! EQU 0 (
echo Mondo 2013 Suite
call :Ins15Lic O365ProPlus DRNV7-VGMM2-B3G9T-4BF84-VMFTK
goto :GVLKC2R
)
if !_SPD! EQU 1 if !_Mondo! EQU 0 if !_O365ProPlus! EQU 0 (
echo SharePointDesigner 2013 App -^> Mondo 2013 Licenses
call :Ins15Lic Mondo
goto :GVLKC2R
)
if !_ProPlus! EQU 1 if !_O365ProPlus! EQU 0 (
echo ProPlus 2013 Suite -^> ProPlus 2013 Licenses
call :Ins15Lic ProPlus
)
if !_Professional! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
echo Professional 2013 Suite -^> ProPlus 2013 Licenses
call :Ins15Lic ProPlus
)
if !_Standard! EQU 1 if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
echo Standard 2013 Suite -^> Standard 2013 Licenses
call :Ins15Lic Standard
)
for %%a in (ProjectPro,VisioPro,ProjectStd,VisioStd) do if !_%%a! EQU 1 (
echo %%a 2013 SKU -^> %%a 2013 Licenses
call :Ins15Lic %%a
)
for %%a in (HomeBusiness,HomeStudent) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  set _Standard=1
  echo %%a 2013 Suite -^> Standard 2013 Licenses
  call :Ins15Lic Standard
  )
)
for %%a in (%_A15Ids%) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 if !_Standard! EQU 0 (
  echo %%a 2013 App
  call :Ins15Lic %%a
  )
)
for %%a in (Access) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 if !_Professional! EQU 0 (
  echo %%a 2013 App
  call :Ins15Lic %%a
  )
)
for %%a in (Lync) do if !_%%a! EQU 1 (
if !_O365ProPlus! EQU 0 if !_ProPlus! EQU 0 (
  echo SkypeforBusiness 2015 App
  call :Ins15Lic %%a
  )
)
goto :GVLKC2R

:InsLic
set "_ID=%1Volume"
set "_pkey="
if not "%2"=="" (
set "_ID=%1Retail"
set "_pkey=PidKey=%2"
)
reg delete %_Config% /f /v %_ID%.OSPPReady %_Nul3%
"!_Integrator!" /I /License PRIDName=%_ID%.16 %_pkey% PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!" %_Nul1%
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
for %%# in ("!_Licenses15Path!\%_patt%*.xrm-ms") do (
%_cscript% %_vbsi%"!_Licenses15Path!\%%~nx#"
)
if defined _pkey wmic path %_sps% where version='%_wmi%' call InstallProductKey ProductKey="%_pkey%" %_Nul3%
reg add %_OSPP15Ready% /f /v %_ID%.OSPPReady /t %_OSPP15ReadT% /d 1 %_Nul1%
reg query %_Con15fig% | findstr /I "%_ID%" %_Nul1%
if %errorlevel% NEQ 0 (
for /f "skip=2 tokens=2*" %%a in ('reg query %_Con15fig%') do reg add %_Con15fig% /t REG_SZ /d "%%b,%_ID%" /f %_Nul1%
)
exit /b

:GVLKC2R
if %_Office16% EQU 1 (
for %%a in (%_RetIds%,ProPlus) do set "_%%a="
)
if %_Office15% EQU 1 (
for %%a in (%_R15Ids%,ProPlus) do set "_%%a="
)
if exist "%SysPath%\spp\store_test\2.0\tokens.dat" if defined _copp (
%_cscript% %_SLMGR% /rilc
)
goto :%_sC2R%

:casVm
cls
mode con cols=100 lines=34
title Check Activation Status [vbs]
%_Nul3% powershell -noprofile -exec bypass -c "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.height=300;$W.buffersize=$B;}"
setlocal EnableDelayedExpansion
echo %line2%
echo ***                   Windows Status                     ***
echo %line2%
copy /y %Windir%\System32\slmgr.vbs "!_temp!\slmgr.vbs" >nul 2>&1
cscript //nologo "!_temp!\slmgr.vbs" /dli || (echo Error executing slmgr.vbs&del /f /q "!_temp!\slmgr.vbs"&goto :casVend)
cscript //nologo "!_temp!\slmgr.vbs" /xpr
del /f /q "!_temp!\slmgr.vbs" >nul 2>&1
echo %line3%

:casVo16
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2016 %_bit%-bit Status               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVo13
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2016 32-bit Status               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVo13
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2013 %_bit%-bit Status               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\15.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2013 32-bit Status               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVo10
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2010 %_bit%-bit Status               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)
if %_wow%==0 goto :casVc16
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\14.0\Common\InstallRoot /v Path" 2^>nul') do (set "office=%%b")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2010 32-bit Status               ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVc16
reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath >nul 2>&1 || goto :casVc13
set office=
for /f "skip=2 tokens=2*" %%a in ('"reg query HKLM\SOFTWARE\Microsoft\Office\ClickToRun /v InstallPath" 2^>nul') do (set "office=%%b\Office16")
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***              Office 2016/2019 C2R Status             ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVc13
reg query HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun /v InstallPath >nul 2>&1 || goto :casVc10
set office=
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" (
  set "office=%ProgramFiles%\Microsoft Office\Office15"
) else if exist "%ProgramW6432%\Microsoft Office\Office15\ospp.vbs" (
  set "office=%ProgramW6432%\Microsoft Office\Office15"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office15"
)
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***                Office 2013 C2R Status                ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVc10
::reg query HKLM\SOFTWARE\Microsoft\Office\14.0\Common\InstallRoot\Virtual >nul 2>&1 || goto :casVend
reg query HKLM\SOFTWARE\Microsoft\Office\14.0\CVH /f Click2run /k >nul 2>&1 || goto :casVend
set office=
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" (
  set "office=%ProgramFiles%\Microsoft Office\Office14"
) else if exist "%ProgramW6432%\Microsoft Office\Office14\ospp.vbs" (
  set "office=%ProgramW6432%\Microsoft Office\Office14"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" (
  set "office=%ProgramFiles(x86)%\Microsoft Office\Office14"
)
if exist "!office!\ospp.vbs" (
echo.
echo %line2%
echo ***                Office 2010 C2R Status                ***
echo %line2%
cscript //nologo "!office!\ospp.vbs" /dstatus
)

:casVend
endlocal
echo.
echo Press any key to continue...
pause >nul
goto :eof

:casWm
cls
mode con cols=100 lines=34
title Check Activation Status [wmic]
%_Nul3% powershell -noprofile -exec bypass -c "&{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.height=300;$W.buffersize=$B;}"
setlocal
set wspp=SoftwareLicensingProduct
set wsps=SoftwareLicensingService
set ospp=OfficeSoftwareProtectionProduct
set osps=OfficeSoftwareProtectionService
set winApp=55c92734-d682-4d71-983e-d6ec3f16059f
set o14App=59a52881-a989-479d-af46-f275c6370663
set o15App=0ff1ce15-a989-479d-af46-f275c6370663
for %%# in (spp_get,ospp_get,cW1nd0ws,sppw,c0ff1ce15,sppo,osppsvc,ospp14,ospp15) do set "%%#="
set "spp_get=Description, DiscoveredKeyManagementServiceMachineName, DiscoveredKeyManagementServiceMachinePort, EvaluationEndDate, GracePeriodRemaining, ID, KeyManagementServiceMachine, KeyManagementServicePort, KeyManagementServiceProductKeyID, LicenseStatus, LicenseStatusReason, Name, PartialProductKey, ProductKeyID, VLActivationInterval, VLRenewalInterval"
set "ospp_get=%spp_get%"
if %winbuild% GEQ 9200 set "spp_get=%spp_get%, DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled"

call :casWpkey %wspp% %winApp% cW1nd0ws sppw
if %winbuild% GEQ 9200 call :casWpkey %wspp% %o15App% c0ff1ce15 sppo
wmic path %osps% get Version 1>nul 2>nul && (
call :casWpkey %ospp% %o14App% osppsvc ospp14
if %winbuild% LSS 9200 call :casWpkey %ospp% %o15App% osppsvc ospp15
)

echo %line2%
echo ***                   Windows Status                     ***
echo %line2%
if not defined cW1nd0ws (
echo.
echo Error: product key not found.
goto :casWcon
)
for /f "tokens=2 delims==" %%# in ('"wmic path %wspp% where (ApplicationID='%winApp%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :casWdet "%wspp%" "%wsps%" "%spp_get%"
  call :casWout
  echo %line3%
  echo.
)

:casWcon
set verbose=1
if not defined c0ff1ce15 (
if defined osppsvc goto :casWospp
goto :casWend
)
echo %line2%
echo ***                   Office Status                      ***
echo %line2%
for /f "tokens=2 delims==" %%# in ('"wmic path %wspp% where (ApplicationID='%o15App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :casWdet "%wspp%" "%wsps%" "%spp_get%"
  call :casWout
  echo %line3%
  echo.
)
set verbose=0
if defined osppsvc goto :casWospp
goto :casWend

:casWospp
if %verbose%==1 (
echo %line2%
echo ***                   Office Status                      ***
echo %line2%
)
if defined ospp15 for /f "tokens=2 delims==" %%# in ('"wmic path %ospp% where (ApplicationID='%o15App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :casWdet "%ospp%" "%osps%" "%ospp_get%"
  call :casWout
  echo %line3%
  echo.
)
if defined ospp14 for /f "tokens=2 delims==" %%# in ('"wmic path %ospp% where (ApplicationID='%o14App%' and PartialProductKey is not null) get ID /value"') do (
  set "chkID=%%#"
  call :casWdet "%ospp%" "%osps%" "%ospp_get%"
  call :casWout
  echo %line3%
  echo.
)
goto :casWend

:casWpkey
wmic path %1 where (ApplicationID='%2' and PartialProductKey is not null) get ID /value 2>nul | findstr /i ID 1>nul && (set %3=1&set %4=1)
exit /b

:casWdet
for %%# in (%~3) do set "%%#="
if %~1 EQU %ospp% for %%# in (DiscoveredKeyManagementServiceMachineIpAddress, KeyManagementServiceLookupDomain, ProductKeyChannel, VLActivationTypeEnabled) do set "%%#="
set "cKmsClient="
for /f "tokens=* delims=" %%# in ('"wmic path %~1 where (ID='%chkID%') get %~3 /value" ^| findstr ^=') do set "%%#"

set /a gprDays=%GracePeriodRemaining%/1440
echo %Description%| findstr /i VOLUME_KMSCLIENT 1>nul && (set cKmsClient=1)
cmd /c exit /b %LicenseStatusReason%
set "LicenseReason=%=ExitCode%"
set "LicenseMsg=Time remaining: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"

if %LicenseStatus% EQU 0 (
set "License=Unlicensed"
set "LicenseMsg="
)
if %LicenseStatus% EQU 1 (
set "License=Licensed"
set "LicenseMsg="
if not %GracePeriodRemaining%==0 set "LicenseMsg=Volume activation expiration: %GracePeriodRemaining% minute(s) (%gprDays% day(s))"
)
if %LicenseStatus% EQU 2 (
set "License=Initial grace period"
)
if %LicenseStatus% EQU 3 (
set "License=Additional grace period (KMS license expired or hardware out of tolerance)"
)
if %LicenseStatus% EQU 4 (
set "License=Non-genuine grace period."
)
if %LicenseStatus% EQU 6 (
set "License=Extended grace period"
)
if %LicenseStatus% EQU 5 (
set "License=Notification"
  if "%LicenseReason%"=="C004F200" (set "LicenseMsg=Notification Reason: 0xC004F200 (non-genuine)."
  ) else if "%LicenseReason%"=="C004F009" (set "LicenseMsg=Notification Reason: 0xC004F009 (grace time expired)."
  ) else (set "LicenseMsg=Notification Reason: 0x%LicenseReason%"
  )
)
if %LicenseStatus% GTR 6 (
set "License=Unknown"
set "LicenseMsg="
)
if not defined cKmsClient exit /b

if %KeyManagementServicePort%==0 set KeyManagementServicePort=1688
set "KmsReg=Registered KMS machine name: %KeyManagementServiceMachine%:%KeyManagementServicePort%"
if "%KeyManagementServiceMachine%"=="" set "KmsReg=Registered KMS machine name: KMS name not available"

if %DiscoveredKeyManagementServiceMachinePort%==0 set DiscoveredKeyManagementServiceMachinePort=1688
set "KmsDns=KMS machine name from DNS: %DiscoveredKeyManagementServiceMachineName%:%DiscoveredKeyManagementServiceMachinePort%"
if "%DiscoveredKeyManagementServiceMachineName%"=="" set "KmsDns=DNS auto-discovery: KMS name not available"

for /f "tokens=* delims=" %%# in ('"wmic path %~2 get ClientMachineID, KeyManagementServiceHostCaching /value" ^| findstr ^=') do set "%%#"
if /i %KeyManagementServiceHostCaching%==True (set KeyManagementServiceHostCaching=Enabled) else (set KeyManagementServiceHostCaching=Disabled)

if %winbuild% LSS 9200 exit /b
if %~1 EQU %ospp% exit /b

if "%DiscoveredKeyManagementServiceMachineIpAddress%"=="" set "DiscoveredKeyManagementServiceMachineIpAddress=not available"

if "%KeyManagementServiceLookupDomain%"=="" set "KeyManagementServiceLookupDomain="

if %VLActivationTypeEnabled% EQU 3 (
set VLActivationType=Token
) else if %VLActivationTypeEnabled% EQU 2 (
set VLActivationType=KMS
) else if %VLActivationTypeEnabled% EQU 1 (
set VLActivationType=AD
) else (
set VLActivationType=All
)
exit /b

:casWout
echo.
echo Name: %Name%
echo Description: %Description%
echo Activation ID: %ID%
echo Extended PID: %ProductKeyID%
if defined ProductKeyChannel echo Product Key Channel: %ProductKeyChannel%
echo Partial Product Key: %PartialProductKey%
echo License Status: %License%
if defined LicenseMsg echo %LicenseMsg%
if not %LicenseStatus%==0 if not %EvaluationEndDate:~0,8%==16010101 echo Evaluation End Date: %EvaluationEndDate:~0,4%-%EvaluationEndDate:~4,2%-%EvaluationEndDate:~6,2% %EvaluationEndDate:~8,2%:%EvaluationEndDate:~10,2% UTC
if not defined cKmsClient exit /b
if defined VLActivationTypeEnabled echo Configured Activation Type: %VLActivationType%
echo.
if not %LicenseStatus%==1 (
echo Please activate the product in order to update KMS client information values.
exit /b
)
echo Most recent activation information:
echo Key Management Service client information
echo.    Client Machine ID (CMID): %ClientMachineID%
echo.    %KmsDns%
echo.    %KmsReg%
if defined DiscoveredKeyManagementServiceMachineIpAddress echo.    KMS machine IP address: %DiscoveredKeyManagementServiceMachineIpAddress%
echo.    KMS machine extended PID: %KeyManagementServiceProductKeyID%
echo.    Activation interval: %VLActivationInterval% minutes
echo.    Renewal interval: %VLRenewalInterval% minutes
echo.    KMS host caching: %KeyManagementServiceHostCaching%
if defined KeyManagementServiceLookupDomain echo.    KMS SRV record lookup domain: %KeyManagementServiceLookupDomain%
exit /b

:casWend
endlocal
echo.
echo Press any key to continue...
pause >nul
goto :eof

:keys
if "%~1"=="" exit /b
goto :%1 %_Nul2% || exit /b

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

:: Windows Server 2019 [RS5]
:de32eafd-aaee-4662-9444-c1befb41bde2
set "_key=N69G4-B89J2-4G8F4-WWYCC-J464C" &:: Standard
exit /b

:34e1ae55-27f8-4950-8877-7a03be5fb181
set "_key=WMDGN-G9PQG-XVVXX-R3X43-63DFG" &:: Datacenter
exit /b

:034d3cbb-5d4b-4245-b3f8-f84571314078
set "_key=WVDHN-86M7X-466P6-VHXV7-YY726" &:: Essentials
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

:8de8eb62-bbe0-40ac-ac17-f75595071ea3
set "_key=GRFBW-QNDC4-6QBHG-CCK3B-2PR88" &:: ServerARM64
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

:2b5a1b0f-a5ab-4c54-ac2f-a6d94824a283
set "_key=JCKRF-N37P4-C2D82-9YXRT-4M63B" &:: Essentials
exit /b

:7b4433f4-b1e7-4788-895a-c45378d38253
set "_key=QN4C6-GBJD2-FB422-GHWJK-GJG2R" &:: Cloud Storage
exit /b

:3dbf341b-5f6c-4fa7-b936-699dce9e263f
set "_key=VP34G-4NPPG-79JTQ-864T4-R3MQX" &:: Azure Core
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

:7482e61b-c589-4b7f-8ecc-46d455ac3b87
set "_key=74YFP-3QFB3-KQT8W-PMXWJ-7M648" &:: Datacenter
exit /b

:620e2b3d-09e7-42fd-802a-17a13652fe7a
set "_key=489J6-VHDMP-X63PK-3K798-CPX3Y" &:: Enterprise
exit /b

:8a26851c-1c7e-48d3-a687-fbca9b9ac16b
set "_key=GT63C-RJFQ3-4GMB6-BRFB9-CB83V" &:: Itanium
exit /b

:f772515c-0e87-48d5-a676-e6962c3e1195
set "_key=736RG-XDKJK-V34PF-BHK87-J6X3K" &:: MultiPoint Server
exit /b

:: Office 2019
:0bc88885-718c-491d-921f-6f214349e79c
set "_key=VQ9DP-NVHPH-T9HJC-J9PDT-KTQRG" &:: Professional Plus C2R-P
exit /b

:fc7c4d0c-2e85-4bb9-afd4-01ed1476b5e9
set "_key=XM2V9-DN9HH-QB449-XDGKC-W2RMW" &:: Project Professional C2R-P
exit /b

:500f6619-ef93-4b75-bcb4-82819998a3ca
set "_key=N2CG9-YD3YK-936X4-3WR82-Q3X4H" &:: Visio Professional C2R-P
exit /b

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
set "_key=QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4" &:: Groove (SharePoint Workspace)
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

:x86dll:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $env:SystemRoot\System32\$tmp -F:* -R; del $tmp -force }

:x86dll:
::O/bZg00000.y#4200000EC2ui000000|5a50RR9100000Q2-n{0RRIP06-i$00000004SU/-.G?Q,dxacyvQ=ZBJrqNN/azE[W)M0MM?+kA[.uKwE&.jRX)@0000n
::24Exr0A2vq;?GKz/X).MWvaWbyQQRfQ_K3e&0-KcHM?(0iOVUwA2+Nh,^(/bWgEMKVP!LUD_.Qpn9O3-ZJez]8DN}.0A[x2pbP..!8B{P]/m}Y4H66[T{KBKtaf+z
::!+UGTK8z(OHhY3}Qb[bIxF=g!8-#JSN}#swIR}|&v|AA?BDlA0aaBlhWUQqXfn0-jDbU;P0_XxdA].pX|7iRGKmY,I3/;b!5#Q}Sx3=w./e-T+6.9AUxSXiaDNdB/
::G!#z!|H2FZzc{h1SbDp71i&RNPPtqFhYep5L|UwK].lVJDhmt9JDgJ20orv7_$_i#D?p/Y7_0{DBCPDJtkJ{FY]@|x1XxkZyLhE7a-#[ueLksApS&$blQC5WT[G-?
::IpUY||KO+eRrF.+91D=uhu!e{#bkn;KL-0O+6psKgFfI{@#ceH/@TW,Gr|myFYjHi_F!W~/15LmJY.LWF/|h!cOms1rizt4sHlOOc{ln_,V~8P]_4_$4Hs3T@7Hm6
::_SyLEzFPdrhEYs7vKZAjwvP7X8#B3DWnmD4q0Q3!XIV?vKWVCYsWkoQ,3oNedbnXt8}/fy3lu_lgg@SNLOQ}?LONkh4f(&h3qxzbh!P]v/+#-_S386wD&&B[B$#x,
::N?z;A039$9cq2SsZin3014L}dzxE;33gtlK/Oz-|x1VmlK,m5+m?[YhP=/u-1Bf83)P&,R&&rk)298by1|SFI9TbWK4Zv(o]ZL-5SP{|1daz[GxhJEii@)fC-VzAS
::z$Z2XC6A#fL_VU3@S^zr(|_7i3J^6]Ivoil9W/W^JSBcQ/f;?2Z9Z|8GNtB_5=3OmE4CBz0]nrci^gHNftIK2d(Xm)^os)vdCb110?gENnM8;Dkn]5+k=A#Xfryfi
::^#Y|,RgBqadnH8!BKK(!4umu#hyduwKmCnZD)EE]o)XJ~BK=~e]!(lOpIu5T[u,JokfpXJuSOkIL55~Q,;KQ7M9M+cImOY?esP]X+z,1ISz2gdFl;h}?hKvYHl=y8
::jD2{9v@XY1!N5p|nGc${gcQr|!B/|-qC4hIB?NFAt?[DLN(cjiz$f^7yeJccWR#ZCuR9Z,$OtIeXa[R|w8Jx=YAO4ZbHTh$!i;Zc3~(zdD_qI6Q8NmIGow18qVwS(
::uuvo~{,(S^Q!208e,pnHoGAoQC1Avlfvy!Ggb;[,gZ]f#IF=(P1~A9{QiWGrN}&O/IHBxpdKx,.K]=A#,ea4hl_[V3;U,Hjy(_L8NU[JJ1}!jyueUGvAuz^{,hFP]
::K{IrVckwa]jrt3if=mcbfkf!deJ!o1hlK(#ur7qBAbT)}8+6addDt,Ku4[cvmLBvKN1)7kb5GvhcpW7L+S^3fYMG7PCD4MTk[n&EFj-v9$+0n_NtAIR96FgibT/WJ
::2n(2)BG5KU94W(?n9xK4K}Q}gXZ^SQv.zk^4qL$?J6UFQG(x}[Nkd~SiApE#;W3YzfDzDFr=&l-I[8zN2ekBnNTY4YhOqW&Ad}+p-7]_@,8p{x2ulRal2~Z.Uj-Qj
::4|/svn1mZbmSLby!vL@Y0v8?HcTUs}_m{qBjN8H)[#gX.Zkt1!_GXQWGUWCId(LWUc4t#02Q?ek{qq]6b2wF3=vQUbO[|&}0l~Yt!r}fG?625|47XuEM(N!#85_l9
::QX)PYraAKl;wPM{TWVp4Q,T+8JxJkUfBfqvU-J=EGr!kkmY74_uNN$tdK]{a$M13}^3fXZ,2]SJOXD{$PAsgUM_PXg#z.^t^FROdgDiT{3qP+Ir}ANTKn~klqjuP1
::/?gT42G|{PXUL[M1A(Zcb@[9WH$eRbRaHaOXXNB#_/9$$&o?mIY6EwELd9vd{XKW~/WKeN$e_J2i2U9@sOL#NB[0sO9Q.?u?[qhQHwc7$82Y$&o9f@W1ou2(&q+;8
::uEsvZ#O3&Jy.yYHG4b{AF^4Z=c40_)?&KTGvF?ZJkoW-DP3cAM_M11CE15pB0|&]^sQ6n,w05SAmfqc1)]dl^?7|#syW@(!y#9vJ@t?72_h!cRut5vEk6Zfan~=L[
::[JWq7Xn/&5LJ,J?pY2/]0Ksq$E8)5i(kmKnI&/A^Q1CTLm=D.lF8b5jY^svcHRN1ljHO9HP/z?HDB0P&MTBomoB;YT,6Y=EIAatGa52PyI?9;IFsOrdasdik2q!2d
::N-#9H8-VE8a=8MBII&WL4kf[8U-KsYf@R9F]t!&~O;5JH&MEH,6KJmIh82J@_476^aNf7Z.[8s2_gV{h(C_Qcp;I4pd_nosj(=xYYlsK&q#Diwj-2_V4Xza8Qq4Ck
::K#fwsuN!mKDQ^TV103d7rtFhIA]}4g5g7k!4?=GrI3X[=0PS?t/S#Gd6u51pxN]S361hcDw[Fpw7xeW5,Vw2Si=uT)oRtJa8[^n-@fMWj]7&#]2Z;!(lW^^k6p}@q
::ERc5-F)RRT-0CyUh=0&S|4Q-FR4BW7J~meS[u.Pp(;gf6?2]Sa1MZH+6in{.hS;lQXFKTe$VpDv7n]PdA;-))P)qU3;lKDyn]2g#ukH73PRfk~V7HP)7[3t{Ovip[
::qi/?LGT=1)2cV3]C1V]2aa0hc^qc=Z(1|0._!rwA#gt}aihsm@0pPcLWw_&+2Z^QXGQ{pck6arc_5,L4ES9EiG4O-ZoJr6JD?jxh0KkWCpWGcubWZ~KRuq}IMB)px
::3DMu84g7/!2[m_u8ge52_=Mt|?Ho^lg642^6/)rnUl]N?(0F.)$n)L2qImUa;7$m??S$64Zj-fiymlvdzijBw(Ey{)]P]JlEwXcaQZ5Y$6L(2y#3=,uhP5j}4rK2O
::{AdUTtS_jjpbPzhb8F@g8CuCn+_DNHU5a]&m~c;]w|0]8KUTTyBFi7@{=CfkorJ8ma^l/q9XP+R0-pjPK(9=8!|;x|Oq7pB4-Xk-.Y(WTA]CSZD&p4oy-s+EuDHU&
::t4afV7t1yVw)P0U-[tUP,4;evo.^|lsT^t0dT,oR&BQ4t8d!xe0tiT03Ps!f{#mmc(Ya7f+U[D,cO/L6]/-kos&D}-B=_d@$&g1-IR;E&Xqpxs0ktWk/jYU!(@TYI
::$|4,q=3nEr^iq6GD6x=/LIS#;]XBgHOAW/H&#WH@sFuOzrXj0=^nIyAZ4Twu7DW)P0q{B[3sr!33TY_{X/)h(7mGpO.#KLkQ{,vcu6gVmwIWwqLwUr=hSDps0xx&l
::r-q&8a?D3-]yK7g]/!(04p.lSd~$C#zz!M-+32]t@$J3bcZl_IGGJFs^&q7abs+$x-2&7}oR&ub.J;&X_}H(-lExhW5Ak;^ld,rxOy80FwAF]!gN8mYun?N]x5CLn
::)f1YuG8#-GUOxK09)m76nqUr&^5sbe0R[u2=#$M.xC{&nf&0&1)xz-57Kb0^.j=N_umG)6vL(f-3vz36r6iJ2_e)Gs$]yrqm_$8(YQF)4fjF4y_RzKrAb{ef,K7~=
::KSeKcY3Q6,IAMv@{SLG?Wn=_XX+Tp=PZyJI&{E)K00zr0e)xsx?}w9L;$x1$6Sme}vp@tBefB^mO5@x2{~CP44mn]lfH]P}9#]}SCQdwMImWgv4-gs50pEeQ=z~EK
::jb[;wMnh(@ih}zoXh9$aa55w4qchU~TXq4|tMPrWT+JHmHA1^yJ4KWiJpzTAf,e/rR5SupJxztSC@A)3RWNuO]+PB9L4[/xn1fOzM(~#oty?9A;3D/iOPsQx7A&xv
::Ib}Y?9?GnaN;[1yt-rT@G8hI;bb];5ABLoU_Gv3vlwh^EWo]5M5hsLX{)a176Yu,0a,+5hRPRE](QZIsZ}dHrcEQKuFsL~~sd0Yw@=Gu$1W4Oa=JmV{7g/mi8xqE$
::OAS80g|YxIa($Cf6xU129(j92oH(dfmCIZMoWZvu.Nms2qy9=7(Sk=89]#0=l#.J($/&HTkt=$eBAMAW6XtNV,v,dwuhm/sldjDM,mdrk&Wc!te7M_8Vx0bbU.KBq
::Es]l{zJFN_=HSzv@y/_Fy1TF9D&Iq5j[YOf!3T~0$PocO^(dztD+|}E8lUO1-~VrmfEn9k[kKJ?&sbH)!zqIrJI+@Lq!Gc5)L,(G0[pMq@x$MeH63=cg]@y;|Ac}4
::[(}09@pZkHc?ht4n}JjGkxD{IVF)Uab7#]4KJaA$4d^Gw?$2N$Sc#!=-!$z@b6m#BD[;ch#!)PYlK/Fe3DTK1D7ZNU!.AI^n@DstLR@HaeTz)gxhZ~hAoa13cAM.[
::W5=s[{xPW8(J^Vl/G+mjiUz1WS^55Ha{~MYdOahNN,ANO&RxuLypfWVdqVIlShTv5MmyPD;4HD/!z6~{W}.AnNTE1Oa+2(6OnRBvbQc&U)?zLAr/]z/6]-X2XUPS$
::ym~jt?yCZNU6}Wefs;M$zf[/-]jMZ2?]AUIXOz+o3;u6KMA$|nX/FJrHk[Vap5Q6z7j8IJEQtkI9.9nIWQs}$l2yH9+SxfYF1Zj642.GS[EtN4z7PX0,KFA3s9x}F
::Z;8-tceEMc[&T.2ER6p3=f$~y|HY]m?&G}4htU,,g@M6EPrAe/Nm]pSKI$9As8u3rgc!.,E/8EdY!6(Psr+dv0)h,rh4B?Y&6kji;L=]S[Qv8?DUpGJGg?L,E_;w!
::JE;u5JsIgT_EPHUog[fAO_ocPU+J;B[1Uuv6yjU[_1ukoB,@s)TAxAE2go+[M,xJihkwRG{)hGO.^2pCRFR)##s)$d8&5Iw-t~LI;K@zTINFz(cScGcse&lv5rB_y
::qqwPuY|DM&iMhUnh,Y=9O-n/]L1r^|5XwF##)8k5tR0vxzX(4PpX,PAc(-_tJ?]R};EWI}&v)e|x&AQe2uH^g1S=t6SWp(v$?t]hPJ#4U{AS@oe,bH;n}geOHXBy2
::H+}Q;if8MNrhn{suy0q1;{IZNhv3V(/IWYSoj~;-9#&TCOg9-jG#ivwXqDW4&0vcqu9}+Z9Ze#-T.9FJKaapCo}e@Hv~0NIB^T[V=C,zfE6VT4cHj$uFCY=dVJx7z
::cO3i._W!S4^6K=_J7Wqo5DfakA_Ee=FVMTM75!6rw0D0|P}wQ@)qJlW4,yhq&Dr6j{&8#wzDHuWc37GIi6{vjo#dZ[KR5Z]xERuuVdlQ~rK!NGo;bm#m/,1pTHHuL
::4HDLbZQWsfEVBc-,i7tp_-TW_;Etx?__AE_x_OlMvcSX7P{!t6z^c9HMT;Y0O&#j}bcUO4{=1c]vWZ?$e_eZFfB$3aH+gA{nF#ptcaA$jCPdyTHF{^S9DG5n[CG=N
::WMHs$eMmpLM9Bqe2C!f}M^@VfYMMm0Prr^6PS@+mDFor)9eTc]^)50C@u!pyLladR.y_EA/IJR?3~y3;^TqvNb9{krfU/lHfP}NR{zivykda4&(|O0iBM2cz|Mv$}
::;3PXmCwRJERRn8UqmeN/C;VSs[cWt{^1|aZy|+4Uh92)N)5#pcyO{#W0Y?o;t7v~jLmfCTqt{#9HU,OZ+eXN_cM~]&8M;iJ2Z93@/2j2mnmorcZicxN1+s4e{69QH
::B)^~PhY[qps{BSyWRYAK9FO5n1S?uzOo;aoNYkp;E!u@GbgURqP3Rm-X|n9X(pT(Zed{.f#FIXLSvyztu||!9U;7u$Z[+[11i/7pb;o_Ok/Y{nC@rvZRRJ;^]vx@{
::5Rp)fw&?l|Am-kB-$C;^WgNsI7,H4Mj|qDP,{S|+/8K^zHyN72v;(JZesk!eH@URwl6F9tT6cDks|#H_dEO=XLEM?Lxx^RO?2RZ0mE^0[^tCV4G@a5tA-@cg/6o=b
::r}&S9/U.3-)o9n~d2}Pk]g-~sxDcc)o?,xT^pfQUk]!!pgaJLiirdQ[Q;oTdNg4Bf4_K4R4/=TW=K7!/PmCirO1ZiHn~e@hckj06XI[|z^b8?PL([pi6O=s)?Vt,.
::s;yeJTL#xcNJqcqH$RFZTn!}NkD_F8CHDpWpza&ZtBl/.UMpwqzxXF,1)nb,8ZiaU/jc;bqug+5lx93XVVJZv4TlGH.ng6K/?yqD(cc~Q+M]q(4GC&-[P&=4ztPj0
::AagCj63US=(7+qANjG#GF.i4Mg4)pm+w2gdfn9D+/Wub&Xs_!BmM0FW@lzzjop[gbP9KLu{fLrKXPC+soPt0qqsip9M5vs2;!&2qZ@-DL{_HW8J7ORCnEM&6.~!n~
::JFLGxwj=BP;Yo[&/Aj~2u)1YYw!VGP8IniRMvhP}A=XWiiNR=s5xk8VjV4;{+88uX_iq~7v7RW3[wSBT6NS#|JB]=EmZjyIJU&XDu]q7S7Xyy]9n_QBc-3EKXfx1;
::G79-QX!JqP4k_M^6QR(I;(tpx9=[w#=ScwXwmMndsR!t)bNn$TwjFgkFp[|z=/]}Za/XV8gV1{xH?5{ZG9KDyoPxLx)#f)zm#{o}y;|iVEfUQ8l8c;;ifZ+mxDsTM
::cU1c-k?xttqs&cYg~#EJPw~-NkO^sMkByW)$qSdN#/YVqf1L?t8=_i7.r2@ic|Ee7$mnIsVJI$E|Bgg6SU8FlWvmQ|jxUXO{dGMwKAy(|pp.xmcPH}d7?&XSfj)jc
::,a.eUg7)-P!oi8Y!=gVN)|V{n$+)JDDeB^Tv4K.?f?[!@na04|&WQh?(Wr)C_O$k67j.oqhBSl=rC4~sZ!,O6RGp+4I^OrPSDyR,Wp{a$lC{^uTa,~Y^ptWPd)CSY
::9+faWJoIsqaN[Xu;wm(05fI8HU({{2QDn]r=8-JZ&=(-Z{[@@~Jm1|rN;kef8$T|[d+#]tK_uSwtkhMhjX2VbP2P[lO2+{^!x/HHBxIl6{,oOhMni8VFHO5eN|KBQ
::D28u1(gq]?3{GPsotdU|?b5|Pqg(9558=[fgjS/z$45W!jX~BkW[O43jl/I3+pDiDxsjvkS0g=ZAxcu!9-i=t2^c~+$ms@(-H)YH$Y{$DEj2@K4Z=HFs4280!8EAf
::0y@.2e(~$&{th(j.?vlHzkTdE.&Q_hBDvvq2)CJiK1}1p|59g2Z6tF6&Ji7d5PM2}R5MA@^NN/_x&dJfMg|Rq8,Iz{#1~&x_TvZqE!WVqkU2u5+x_G4FW1nQisuB0
::nrdO.TJJ6(G^dsQJ8z^dAp+X9!p~Z1b}9N55EdL0I#2]AG@3sF0enTE+Ec7+x({Bk@=mlZMBJ/,URf9w2;RtsB!WhS3MvWV68KJW^w7HXME!Zzb]&-#VS#cJ5rJAM
::S;JvJ[y7$2Td-r!S}hn.7LgX!^0l[2rW24R08D^yu-oYoL8;~[0.M,;YY0J?1b^(hDv7qBCa6)^NOJxWsgb|g]!b0IWF$XwOxO$j$6_X5EO6FdRvPMpu?[ia(xjA!
::9md?Akix(Iz3s53Om5L1XR}wx-B7jKFfy~/{#y_NF)@Jpx.?9r6^/?Ti(LvJJNiHc;H5U^BM?$)-iUuQKUvlPZ22wJxm0H}w6;Fda25//mQO2xmbMl3WM$K96WJ!g
::hXGq6QhZr=5P/vP{D]rld$wa]O^g2BV!p}{(81mqbP)]D+gBWCzS{Gt9R,dgF)]$MrRdcay$l.R,_!G6m5Uqp499U|0]@=16!]taW[NdM62(d+zlODqh+lBdw=d#s
::FnEu[D|R)/n.$qsrtBeGRu[Yw8_wmATJ|uA$Ic])$Oh-]/yl_J9xCPYAJ4DR16,~j1m!AGj#Q~hh.LL5l~#(h2I?s1zFRCrHeD.z=wj;@EM(Gj7EiC-vbGkgRwc/5
::xcO8K;_wYMk4)nqj8m]#Y7fJ2{UDNQ{@y[fYXkrOG$9kWa#S/xFmfDs$@;(m6@R4!9urpbS6Yqq^^EHc@=wEWtj}ue&ndrmU+VfrwJrL[XKU8xN|Vf[xoG&Sl?F;p
::FfgS;)GD_h5t!qPFN8WXN9O+U6KJv.wevjIo=JQK1]-Xz29f?[PYtK-T+EDU.Q&wyEz3[}]y#]7kg7C&aJ,#-3e&Z(h)1b?F{GXbnB#cm&5hPFj(]B^P2XI&XFWD)
::K+W_A(Ru,.K~[8!cnw3wMXx+(QEOttU9MAZcOS+=SW&^71lM{fD4z&2s5[iW4EHaIhpYg9lX$J{)Wnxxpw,CL5pyx=)siq(n]wZ3M]PNHYF)sFDPpRUIEOupQ=KcI
::bu)l4{G),K&=D;-T+)wKoN/Psqs{Qtv(N3uxWrpjGP.o8dzuf~xMxL_t~u^W_-i/Wx37.&t5yn44O&[L{(k2];~Mm}TKzKZZ;@NY6uTZ?t)+yI-t&q&{NGNa$NAP6
::T3#ASp#sy?xz#8)Yn2Y-R}Lpq$[g)_D-}f9s[f{WgZ@HtaTSrL/jfCbP@Sr}3tA0@_&mPkuL[M[(3=V]+Qc-PSQfBO0=]aKuI;+i$ZC)ZOH{6JJ6&d9/8s.OQnQjD
::Q3=)lSA5u8ZQD&8YjH4r.2Y/~TgQtJb8a01cjLse(1u|~t}o0$J^Yrh@xv]K!MU#7)razNmL6J1dH4I{Z1dQpeEtPmi|OV;s(kJs9f1W0Ydk=-@A!Gu@Etg{H[.V.
::ZAq}iW?ze|_A0|VS3LP;4U^EToWd[ZI2&A!U[6kCOCB1P.s,V8u@SUl===zwRiCfVdtEICLPi0i@9jY9n1gl4rJ&aYPnGm!@wcx7I}h}SX|C.x(x4GPiq#6pdABW4
::i.!Zkc]!0B4w6gOCwX34wC5s+.neINTL$_WMnI[CKB&MM&8xA!Gdk_ETU!#zqn.2hEup//$MV}aTnMFl96MLfoZs4]1a;]Drt-b[hyA(+1y)eKojV~@BhjK+nagk4
::37SsWpBM4jQC$!(q6VMGyy}itMyu/&oqN{PqN]#kB;ZSGC$&weURR+~zaR3M3--sJv9+|TC=PufONAO]0afL0DE,v0$.{nWL]iUYS9NLC3$fm~-U6irEW6x7]73qa
::w8f2H;2LX|Zq1Xs-i]Fdc5i,q6/|3=v;}v!H/N0UY7y1!IjJkWP9}ZSxx5n9GuE+lijL.Ge^KoO#8l2}euA+2zPcFA?zMCdr23zfwa;Lq9{U8[889jN+tYJw2GhUG
::[U|1QKCVDBb+y,Oe?9t2Im0y?vF3DwIAr)AtD~A?rXBycU)),}IWcXuKk}~Cb1N-rpQsjS.RnoNk/RsB@rOD?CA6k!LHO4j;^-g?lkwsf^CE}@!q(6+MyCBOCqiN]
::tmN~X_RT0y]N$O@#URgXbFPOn9rLM5+yR?kJWe3GoS;GE[bC)o$k533DlzUm^vp(bI;~B_7v,UqFn#GppWBwwnkWlHh|4nP#J1=@x~e~$0cH{OpM6LhO|@$0R](R!
::t.id-UDZ&skW$h9dM1H?@kDU+2Vf61r;;KuWc0aAG/TFheR0m/yzwN[eV;?IlS/a6hj+V0D[fJlj;l@ZI-jb[D(dp{_x^-9IyETu#ECqU{#rOz@$5?AFbhk~cCzP8
::S=^xXgVZkB-ax@M+9XA9ny-egQg[i@RZBa+!g[w#rxFX1-ww7Gk7tJz,???O]imw;VndOnF=+8MZR_0Z]Zt4qemjOosNkuZE}lHsR$s^[tzsUT14|xSmcMv=/o]Fd
::(6}WcmMUJY4[[6dym[JS-GQ5uy6t@9xpR4Eq=~4Xrd-rCPrFOyYejddM][!X75)#HS3hpQU$a~kuzTJ]o-V7JYeL,9;+?YdAy!qV(n$F(D@_t#TY2)opgwgSV-dIJ
::dZk.yO4N(aRPT}{cuuShSI~L#=!{(^K@ep}1#n$7/yU@V=Tgnw[fpW@RMoSz2Xa]8SDBoS=(8b^Kvh(2lmP5U_z4!#@(thGZxqy372v))#aV3q8!)[$[QcK0O2S!{
::.nlNk-S|c~4e5&wadI(bFEkZrD+UBno~fsp]^pdhg]A(=KhRhia,LE/,=P5-ET$FfYFp1C@bh!XYl1Oz(V6-wMZUAK?9|B+[Dkm.5Y)bB!.,#]q5v7^VC#[KT|~cF
::[}B1!to$Uzz?dvwCuCHYY&Ke,?ic|Z@vud^2Kc+hM.{E#A-EhWzESXsN,J&|.Da@5zRjS|=w~iv3eM=xY{.CvMH3V_Jh7VTnlYQBmQy#QH[Q6XmLWY^oSDa^_glXo
::zQm|y)]|7DSTpP,m;9i4ygeYWrV_D_p{vT1Vg{J}zkP7}yVhW9h5|bX_K#BjudVAdEg~(.;vf}_{vch;?kRCxwU;@3j4_b2Mm/{6iKmszgX01qwn46;fwJ9.fg+ZE
::8!o^lD5jm#a[qtO4q&cC^VqX?rsRP^4#BKAYJv0]I;=|Y7}0Vg5z5dNah[}evLJ}xuzyEkl44(-wM4UXW?uz6wo-6l#sf;=@OS?Ltl}TJQvlAs]O&![SCBCre6B7w
::b8jyMHuLY!vS-zpX8K@e1d#ySOBMTpTK7Im//WKcFxpq^YjE4XAMc3K,RR0#=WRQto&THvuFgxz{N|Sfd293VYSx34z~/;L=kOmthi3!J/XG@cmn^=eas?jDjt/Q.
::xPK$~=gN_S_A/@Yi[yb}+~L7Pz/Zs2|2uQjIrxJOLd;$u$7W_bWabZx@S6NH1OM9,IuSe)K[mj}OAt]YkU})!kV9HVnj=AxFpvO~gi1)CZzaMBb;|j~u0|]jl=(=#
::E{v9xmY|lXmadj;m&2/cCGV1bX;i~PB}=I##}ely&gOeka}iafT2;jzkxU^)LR(@0#jhfpqI89xf+;6^6_cyF3rdSH7GPZ&U!X0tUDRDJU!5;}7sLyG3se@iS.5=@
::x,=tlG9j3OF_YHOHfWt?PHIl(PPx;W+H{it5KfgR(1rTd-G6Suf2y82p5{JHo+SOuo/{yXtJNyAI$PydMOD)2k7^E43F./873BqG1~m[mLdT)Us2oa#)nI~A1VdG!
::Mo]~q{oc=kWf27tND+sER3e.q;U-iW2vP#6#9|3N61YtqPn9Qt6NeMi6L,Q$6UzzA34n=@5|WUZkgAZpkZ6&+NNPyQNNJ=g5.Q0mX,cPcG,0?^5tAB8k5cRjbaXo8
::lNqi_C~8uYQo2$SQ$kmemCQ/_rO=YdlJcc|3A5}H4VhkmE_};sDr6~_Rj5[!Q;$qtu4h?(IIgHw-,+w8$ac|pv3Fs8[xB/doLP_!_eJ&zsxo;,#!PC{;_W!~98)^2
::GNhR?&2?,w#N(xZ6$~qkxs1HV;&~B,?5MuUftz@uV45bH;eTy+,OTVcUz1?IXr$C-;-bAyub$JfQ)V;mt,iP]Wv965]J)r0@/ZUWTYXjst7K3jP#|[9j{[aHmm|/-
::nUH#rf{=_nrbt;(ul#=eD.y4k^$@7|/(38yqFrL-1mwiN#LbB(&L(nm)~t._8Y7XUqLh$TkyJ?oNO#AO7Aw{jj}(|?.YWzt7FOI.OsMcs0fIsf3lJ@nT5056L|SyU
::/B4{k/^d@P/^qVc!qp.x6P2mY]ypXPAAtV
:x86dll:

:amd64dll:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $env:SystemRoot\System32\$tmp -F:* -R; del $tmp -force }

:amd64dll:
::O/bZg000000U_hZ00000EC2ui000000|5a50RR9100000Q2-n{0RRIP07w7/00000004SU=A9q^Q,dxacyvQ=ZBJrqNN/azE[W)M0QpB{lc,s8NLzq^jRX)@fB,n7
::1pp+f0A2v/?N_UGqD31PM-{D7P,cb2xQZ;0w_3}{N)OhD^~&&gCO0#MnISW^u~}m3yi1828v9_5Ymsg.lNVnA7#SL@832H=1NM8,-p^f9vfVdwFH&vOq?W15C@@jR
::+qB#AQpVct/pJ!u+V1$,wh7e1CN;Cwt_YleQgH/)Vw9?aTwtlNt#t+e/y9}S6ieao(26}!G)~[Z.~I3U{{V/p0FV@F-0=l(H~Zb0{XmkiNB{@jXwVpf8(nyMX#mlV
::=dN+8J{#EP;1UfHj]q$;A@V#=v7zkFKM/[=AI#j1e~+ImgzrYS)xXs1]gS)cS!n[hZCjM7QYj!TsVpo(kgBFFpjeSsKt|BhP&2O=OQ@$Iv2pbA3o+6xBr(H21V4#V
::?!W__PGu3)k(c5TFf-sdCkUm8&,&E=tMPY)#^t,){jcT]_Y!3vnS_)|1[M5wApXz7^n+czQ6UMtwzO_#UOf}i+lCDv-7yYcf_)M#+6zd{o)j&3Kt(.[6[o|?kRPEu
::t/K_o{FOe4656Wsn]Ddq-to.4.L[{Ss/9E7=K.^{z|2|Or[CG]XY2;bvdK-h7tvuxOrtru9J7U!69W@l2MIC3j]FrUWNI_u^Im,oD){30LwH7lHAhALn(g2}nr(uI
::nt/Gq?b9]}U#~b3Q}xKxBjKlsD#zdPSSe4JBheTYX#,yk20)fXWu&#~G}Y(mMurUUDWqr.g^]{8;_x}PG!-qV,Q4ha8aFM$LZv=a?=~4X=^p9Q-e&bipKF$TAjAMe
::A[W==5QH9UBs)[{sGzW)YvdmE(}uoB06nTY_7~y8E9Nm0pk+}i&24Ta/y{Aqb54]wrFk~zhivws(oh))Y!ZSOXtSUW;JZOA8=QYTTjfCza|#p(jxtdDW;AP-Qf)$W
::$vX@-SrEO.][(gLKg{Bn;A$=_,W0y$PG7(n1-75Mm[vDW1S?nfF04_teM?y;+,Lh;E#a0wBl?!zXVZM-]T]H4jhf4Q/P|O^QSk&+6,Lk;vFJXIAPO|)Gf9O!wm2Z6
::fZ]q/kl8Klg?yiDK5-{b{f$w-=B?Xs.bnhz/$hM22.8!NhfQjm(;J+ey|1O.n]@i5LcWbr8#26Z;Wd[dcA5rujp[;Pj+5}i?Jf5q)Huam8^acmOy@=nOO_jX5^mFK
::gQhul.0d_zI!FLf=}q_4yJihq0J.hE3tq_eUqMpsW,udBzHkbAU7&#OzNzDLc|mz=T7pH[YbTX00Ck1y85W5t&j)yNMYQ9+PA!BLum$S70,Xg2cwx!_F!to[dV!kz
::uWdA7clDq.-U9R[fIJ_?/x0&8pjiUEj)f7twAof,Ln}#8#iSO7[q)Eilt~iAH|rU(&NYjW+dim3&;wrZU!n)mP#G[b,hX0RZDM22nG![H+)7q3x=aWuU;D.1u/|]l
::AeTwwkLybyD#j/)bKzY@/=~I~m_HdsFFgVUGTu$a8[mEhbV2dBP3F6ARrZ2Qkqf5Mq]P}5DT5^4.QwR_FtB~p_2uo^yUfjc,@/+zm7RhC,93UaT~;ks({5=8L4&!w
::M1a&4(Hh,I!CBFLj{//o9V.I0V2o$TM4}V_1d^&5bcTKryzea7Ya^5,ynF#Jx@ywapx8ZSx=toQ]Z{Dmd~xhEXi]Z4=b.W=S~$l{/-YBAuagj8Qc,hFGyNPifFx~L
::lXA9br(n?C2vxi$wNj!iP(eEN/FB;dMibN3lDQYi7|WZa1otM1SNC@G9&?Axt&8GMerZ^wo_h[TiHQro,SAI!(G+}=BQ9fK(=[NEIxWn7|Cgwgve}Be^yXt39NwQ^
::$M,tG)?DD+EpS?gx-^^^,mo{Zs}UQU9HhN5,D$JyyDP|DnRF0cZAUMQMnI?(!w@}_fnP{?L4_ivT{0^1nuO?pC?XYitKA?QzOO,;QV.lIQ],l3I4qI#DuPIavcMOW
::Eb=L~DO9-RwIJn^l|f#a=D+8{]n8qpI3onrhbh$9vz(3ddKqJHgv9pH/|FtrOE6R{b07LB^3|~ZyQBu&pYFo2R}6bB6qvaWCHy7L&r;PLh4Ek1(}S,4_gGCpGgk0^
::e!B2ASFGfiZB}T&m)AoS|32c=^!Bn_z}kS8aoIsVmlX6ejfaAe@}UmXRqz.Jv6ur]EG=(dF,5s6),~c-v+Hg[nnf3Hv}pTV#d;@Jv}jXbLQ$b/G;,4H[5)mfH6_^g
::2owieG$o{wh8sv,p]!Gp@k]@LHlwr~|AP}L7li@^Lg(..+)x/XKP=J01]~Bd)CMjHVvWpku,e_OQMkr=YjmM?p?)lv($lS[i6~g_lorqrCe8[d..T~@M2IS81[1a9
::m5FwH9-8~Is3hMCNe6H,X+6WJ[ks2VfWnW|Zf-EJMqXK@($2[pFx^ni_/x^o?Z0yg3-9+d)U.=B/omO+.OMg;FplsGW~(9oN,?2=rKHya?Zk45PSgi_WFB#A9C3t4
::;|B30gNvv=i#U-)T4VxKPdfE[U6SozNn&q3AERq)T$KE/6Y!|gGP=kIO;8D@,Sm3nqiX,7NH_K,ZCXSY_[h94C78j;|K,2Xl&WO#z$X=pEkkeU,Snht!oO1oNs;xE
::7CoawzY]X75N];w6C]N+[YCoBwp6^w^Ur]Xzr-MozudZ5]Tb^kIRA$s6#!RmDMN{^e/??a,xyQ5H3OZWhf^3_p-sCH{=4#~ulvK)8GTulk6[_Bdv(8c,YN{7)$EQ.
::e?H4DnA_Wl8xc,a-xTJOGcscG7wafk6OM6Z4g5X=E8;T0T[3Fun8#fm4u@Ym&ch7zY~e2@!T/6+xSHgVCp1r?xAs63c^TD&97K@u-_Ke=1J.au=u|eX6KGEZ48SQo
::DPPZ4J&evZvI?JO@/8Dgim(!i0]TJjU_N{d6&TJqxC|5@Vsa.^)7B|v.,b7Xa?8rDHwtf72Y{+-FcH_R6NaHcDIR3ND=[T09KiYKRbYzCCSkEt=,|dv#0~fS6}Hbp
::B)X}H;szyeME9Y;8IPf2fT6sRf~UL]^f]mjsT1.A_C~kR5Fb-d9b2FlYEx_Lx+wNZp8s|)PDYE}PeJEQSVmH+O@5bY_!.sWr{{-/8CX!}6/O_fl#-!;@oY6V6ADEt
::dm5h{_IZc=R6j.t@71ZzJ$7N~(HSd=;Ep^ch-}vv+(5$9I3/#HbT$6GOd#regkX^187@d#z$$m9B9s_$/V[;!,@R9m!#Osj9XXQ/2^~+[1?ELyE)k[56#vsWx1/vk
::5rj1VzJLZ(!dvJRzf41WrC.KMK9VtRzhz4f0JxCgdA|7#QKsGshEVH$c?XWq+q=hm3Tm)|AP.ohk02$v6J[(++b3GO!yqJ+QqOz[obzM]S8,T+o[9g2m2QQ0EgJaS
::rtsg!]w5q=K~&mV=d!Zq&t?=9VB@X{xk4]EUk,w&-~,y}#e00{dz/R|+{I8Q1Q]6OuNH[.ScCaHj,l3senu?Dds!KH+}Ix0(Q3{@CkIQnk;$7Hm[{~3x#Etn(tm|q
::0__!P3HK/hQbh4g5L}J!Ljo6;r=Tmn;=k!tXYr&QjY$Y|7d~=rb--sSL3.Be)bFsZ$|k7QxGvPB2F0D{3In&adsuy~P(nGLotFtT=L5=|.vpooIoog}2zcld66j$=
::=f_-Os]}BZNG0k@JJ5=RgYSuGIQuO+swg|yQJb#vjH)wAzu-!,;0WsP-pVG-sBTkd1XVbuQu}vhK5uuu6Us7/G,PZ(_l-]z?uRSb#{YD^yEFmD-bY$[G]Wxvur_cG
::JLQ)zK(aG,7f2j9tXbmN3[Y$@U]X[Qup9FcMZJr(m}(2mX5W6hUysk;ai6r4+XnFv2v8P41RQK-JWfVZ3OE8XQaTU6Wb@iKA)U]jdH+RLvn~FMyh~HxH?@ikI]YDf
::cY2.iQfsBEppW5tjoM$}&1nzycb0Z){v_hs0gU,w$.eQ~C?LovGiHlP$j=}Fyq4zLB@OuFW5CA6yZu93fewUo(LX}&!5U?{c+^Lkm)(9L$D5=Ue{J@E6(r&I,e!}8
::;j^m8IGxVCs.|_!66)@rSyj9+CWpR)2BH=.2_.Ci][&(i{.s19{~l7,v1)e0Cwq|Z&P[WBy$S@{XGWqO-~mF#x=,~,$Yay3CTzQuWbMKj5&c8(51$#TV]Reou?1hJ
::LfCz0mgw]?-g(e3C_oU[I)!8HioZT7z^&tqTCwl|M/s_@O$|04ALB/Fu~4S^z7iA@{[0/r;Y6Dum,W6^MEMIieGD_K8P/SCPgk?Wf)+J6]gS-eie?l;@n2om#F=H9
::,FwwZU=NUMa!M9U]xv$[Sm;[i^nVs,,WBlWn/^sxr9w6G-o/}}x($L@#e6BLz^hbGF4aBbJfK3w9GC/~UD^G!@g=t-?4b5C?eh+3(.MzOJYDj8IrnnnH9pW5mOQ@f
::IfZnta)Mq_aec5(.jMH3daW[QJgXI9bOUJp^HyNMr!G(lZa[6XBGt|^N5kS)K|PavB_12^Ryg!qa9P]PI$7aDtjsPrn1=R,GO|e1];jVHko#kMy{X5|D=qhk!n?jP
::t^(Z$_WyBHxkG+rgt+)q2?JB^pbzLUxxWj[L~@{pk|8YeEj_TE=Vu|u!|vjcBZfa=d(7fY-c0brHL$xpm)DfG7_jXNMCd/GO$8H4lNAjZU(&=[(fI?zLm{!l@)jJ8
::o2JN)WI#M~D)b|SYgVs]=afbjM]@)iU6@1ogkajpUeE^M.-Dbad3&yer5l,p-P$FeotHh[^WAv{DZFWyP5}rQ1/wI&BTic$FUiMC02$wXJ2M^QIW)/fnvf=GD7a=v
::!(EWeL.OWj3De;ll0VG|t-7$?#/0D,8@ZS01?2{&cL.@P7v$6kIOdIjuQZ1fUk0B!^b,REw++eOtk{0KpF[f5k/_kMr[(L?TUblnG6_e5qwux@J(w)O-)srI[kqqH
::!HpzX/J-u^cP!i9.2#FiNN4yfw#M&uEezQnDCg?I#MsvTmOpN^.=dGpV|C^A?6[Fz5rD_x/n{Mz8)0az,CTzBc8[B{]IRUWJ0jF&jr#/{tl,(vSN+|T#8Y}9Ww~GR
::p;hI[hzL+@8h1wXCTG8D3O0j($fWVFIdn109I,?D8AB}-]I|RmME&={O8dO/_5);;IBld@7vM)S{Pkzwz+vdKMh[&Q_-2y+tP2}=5|d1SZ-/cLJkl9W(krO|(X3)s
::/U!qtC8ugQp-B#WkGDx{ioGu#I2R3aK[/8=HgAbO=&7-Y/)zofn;5gDb&(SCe80OLIudlF7_^(LHtQ@@9gqEm)gPD9=l02lz)#^oLptdj{lmljB=!5_Mj?C0WQTbT
::v}M55N1|]9haG1eI/kdw99k6+kWGvuaRrzezIJ/|iQ-5[/ifZyk=&|pLUcAtG,9Ha@Z+]IRCv|+lVfG@WDD7UhuTmc3O@+n5B[2y^T7-E9qiubeF9MeJO^X#pdA5d
::0KgFd5_c{Xi~__[gVCo1pcsGw1mlD.k-97GZs#imOMjX~f&YcCp-tOf.;I?3zgF([x7@||+Ss&X&.idEru=$+ac=w,3P|nD[z/ZmlR)IWB2MjT,#_ht-;i;L;P9MF
::)H[?hXN~JK{b4Nezb0Inu|JO]w=KlpTanp)jlT-7@Lw!5C4Ws.8C)|,X6T;/r.=aRi#^^mUG$+xADPb].oO{)Twi5;PyRCT?1OdipZTy~Pl@i|a[(jgd^ObUYxQ9L
::qs!v|5JF9X4Lz0.=RFP$wAztJYc/S}jiKmfg;M=PqXvlfsVN^7{3#i_6aeEpafc?5KCR.UKR/TxPJl]({ou&tRx78AvAhF8qf8i4Fffw^$Dl}4ZS8hEJTB!#3K8A]
::]?3l9zsJQe,9Qh9yVAoRqgrS@c2oe0v[?mI_TX);,m8B3deTSmG9MvMpoPY-bbTZBM|i?M9yB|?P^Tbp57Ru_1-gRfLYP;O(K;P)]X.f2)V{zJ9^SUTRD^/bG6_S]
::t^6h31dh6w3tNJHkWU?^rgV6tefL$+-#nA+m/s$D0)8Q4HgOs~lW7ArGrd+x7KJ5t-E#EII{$coQ4;K[kD^WdFQS+{p1&MKglt}?E#4[#mCXbG-GIW(@ac8vd8yx$
::kB=Hk9J?aIEaZrs,R/=y7f[tAqAy1UUq=r&bWaHxfIMn;ddH,J3ABH?pFAJg@~_K~d5FF6gm$j#&7I^{;,Ngx0)7$85D~p-l,l|E-ot;oi}!!s@L9I,3[^9u/$F9X
::oA[MKz(w[(ivjq@{V;V]qzoe_zD#z$=,y0Q93{jC7(rt!B])S^LJ0WxI[Jl4CNf3Qt4({-Af]=fk[I?@s;[1[d,+&0ge4r@(3poM$C6QNw-3Pkao;paq-hAVqHi,?
::0w+XrelPiGb7zl}aQ/UwRrV)Kki)++9)NZjo~YkYG;m/1y~.BJBj?3q+MX{M0T&0#LKZ/Goig#6{P),+rOo_wQcu/;5{)gg64]j;Fb6-KdsGV!mV7M$f}3/xn8pt/
::XZd?!{z#E4ddr/}/,#g3acMtG4uI7+K,kIn_=2$&9UBy!a#=hlSU.CF{i2zMnFOdAG$SY8J}X?+aH44/0l4nF6/9lVP#s6)^v8U.EL6.J+dg~VZK~2iOl|iAmm=D)
::$$Re40q9n}Y.ppZnM39Hf4.l4qR31@]p_EwYLi}M+BX;29Jg[1{SE0tRmD$meDRIMHn9SJ,(,I@&lM&IDh0EOzq?Sd[fhrI9eLj9JtN~y4Cmu@bz0pSoH&MlSONPS
::wV[E}#1|4rNH}!ZuxhhMT_L0X$/jqQioOeALWRf2?|q;{1JE0pyre-3QFTKHfsMrLKNLZ$Hw!,S;0iOJws,Vli.d&^g](PneTPbk=U-mV/GpfvZSslFkLDV;9LBpl
::x,7|kp=(e=KVJn|U7D9zTl&L,o1/eH)v-gt4HzQY#s_#zFACt7=WTHk8_!7xS()e]^_hpCLm0nrDfaZ~n3R7beMsR)Qx5/tAN~rX;d50e/)Ig&,]sysqowmGu!M?t
::tyu]?lu#eP;mm~Cslc6rwtB#5h[m|XS@37+A/RB0v[9}b;#${+5OHrl?E(^J[&^zZ{R;{0SVPC)Ftcl6wYOgP,XH$^5aZ&-V2/?juB!eI[+Kf6zhG{2gj^yAG;fs/
::+wY?tG&Mn8DY|5G[fWlQSk-OV9ABw?i}IctjPW}+EW]E{eoy}|ik56{7nZc7X[@HK7XP]Lf_wy]K9;@6Ffqe7#EnV[h3LGuDsYH6=xh_6xb;^d0S+671FR6}k}Npf
::75r4oZQ+wo=&[~EhDw?+nYyz7G$.aYdy_5,g~bcSMN$XO^|OQ8;9hE;4pn[^vw9,3QiX!1$&Kija&$X/?3xgnxDk!0+8+b?Z/zHO-d/OIz$Wr7=bC3[@vMZBNQry#
::D3]6?)&KC5e;x9)ZjGWY5GERzBE(&(8@k8/H/1C@#Rg,Qz]H7+(BWd9_sMp-(U^o#p^KEC1CNh~Uv_NKzU#r~RT#[.2pK(A@)+=bBu/UM_aK6fz[I6U@,B!$^XUu1
::X2/$n/j4L8!cxPMP;(]C4{fuV$ZwsLGL@k5I|PCUMfh0mK$Pcj00?d!9f]S34uW0hmK-6N=0s{e@-)rYQ/?C#1kt.2LFa^!w.,7PW$AXoROK6-{}XXh1l(O@^CD_n
::s.Z5SEHkC7nvy2H.~ru(R9S6/a7eAKVn{i/it^=)[ze^{LR,Be2rKfMgqnm!1gQuWrL^BkFmZwi/h2Pj=bJHun@n;6AMX{Ze](qdRM[oq3|Y]O6Mar+#0K.2QY=Ny
::?j^Mq4CHC{AiNeA#7-X!e|Ja)r7rKcep,(7AeB$7BCEDpKmS.55.v|9Ea[a7(p}{F-YP]|4lFH[ttz3k+ive5R}^=}zxr?IcBim]qnqJvX4=YI&?i3gL{)Q[Qm6ba
::{AaD1p~[&AH1t!x?pcc7MTa=17fgONKUxnrjW1T~ibWKKC6_5@qlH;u5ImsJ+D4vrEXl(VCWNGfs&e8i]$-#4QGF|)-RTW,(FN!WL_^pEf[4{qluBqUk8mU,s!u9E
::+TIW!EnRJYGLPbwzh_wVJr,&mxj/9;as{y]Wo~s8T9C=_SMt/J4jV}MpoZX8;n$KJ{A9xKLkVA|3C+;=;^5p,noyA@f_q~wrn3c)=q1jcEzk.UqBlI|g#vY}r#oN}
::y]P&gYPG&.h3qp_77I1CR|pOKo/)8-Cj~s-Oe2GJKiOA[8czX]|Iw7c.#a|d8MjXvYWfm}/d~VO#a6YLS=4mere.mEyruGK=_z{lg?R5ixKZ$Q,WI;dM(70G=wZ6P
::@$5~w_Y_7Rx)LFtk#hwo1l;U.s|e1400j)cF&8ek|6p,BL722C2t,Od!GMIHOA-MU=Vv-KhEYr3h~M);_#d;cY@aIDj.^C7B]TkjngA(U2Mh#Wl[[9=.cz_ok_3;j
::0z7(smRjoAw_;1.b=OndWpNiR9.lqT2c)=2N&v(BmZL.]#Xs6EP1mg=vC4k=G$IfVdmx2#4qEzf.k|vLnS(2)^v]jEmve@VT&/WA@paA?@M~FU(dZfXQ0@-/lST)s
::nkzX-DTXR_i!Y&5Vd1@P$9?eud3#&nR7yj?MSRn.tExyPr9K(cp+[;t#FVSG8xuvBEBfVvUZt|fH}vvhKHI+/N7;{pT1geJgrB7n()GLRW_tI84hE[e]W?|Rn=H~]
::c#pcak}g.Ov+{e8sYF;L(?_8cEIw6Sy,gBhbl&wHPR+~)PCRf|Z?w5JD0pu{e(,C{eVt.ZR2Lf?Ifope5Mq5LlKu0JHP1DV$gqah)[_qCsVkkhm}GKj-5R^nUCjBN
::w$yRlafRrS_[m,7,k8ZESY,=!(d?@h;(IV@L)0FZwiffI6bm}x9L~1a,+C)nDl09N71cZG3K7viI!nR&r;87Yb?$ci&dY3]6qW7O!YPupiE.I.tk!Wkl.&]Q?)&Kr
::d6e$w^{g?A57-o/ELm26Eu9pahuy4{s7+ItkFqZ==W8Qe)=&TcUKl_qNs_Y?n)gQunXbfp!EtT97;G}z6PMF0dOa4pI_kMtk|KGwkeiHEM|HVDZRO+ix}Y6@[?R7t
::q]{#Cx&Kt[Dk4!,=U=hylZlfl?Gy57reHGM=Lg$jS+jC)N$V6!6?-~=QHs6F3pyh1/5!vE@9RHAS2?_4bQNo#&PqFMlsRT?xq9B!=eo]a9F/?&!a@MSuQDwuTh_Z,
::JK_?795LP67#/VNCs8a]@bk_uU.8Q.Gj8f800TMFFDlZ&4IW+Rf~.TwZt~mo+2g{n_r=TTwzNCB-cYL/j;lL(BNwkUTtTdkZ2Sfxd)-0zIZKj7,PdEZD$ArbiXX}?
::V=$P~,LJ+9Q;u+{9p+RDljz^T)xWAD1=yy{_CC;Mwp5&tm^^bgk.0l6#[-JcAyqXN4Q,33y6hEkh&U69sB3}hki#i,PW5p/F7G}w&)}_u449=__C;!MvdZ{{]Njwq
::C+&)^$vv;7tS^FJp1Fs5)H+sL66/V-g&tYAYac]gLEy/M[=A/]KbdE,7OYM2Q@8f6(5O3OS1nBlt;E!l-#NY)pLtru6J1-QI13f2rWn#xhv13U5Hj7IYCgpy(28)R
::l$5S[tCi.4+[G,HF,STdE}xVz)JNXwtbj_mO9kU|94N(y9qII[F-1-(jI)d()Dw|eE.d81hMZB+,wBSe_5DLe(Sf3S{o3hSt|V[qB}ccf$1?uAlajKcJp/ujA!X)J
::QYG27n@E,]Z-ZP+Mvyx.p=w|s$go[MD8E)vQ8&8qvTIFu=by;SOUV)J&gs(XQEaPjl7pg{9{J+=YuPsA(rR,5]OC.jVYRvSxV1Ch;VtynNVFPM(E[NLxb;T-I_,t}
::vR,vOJKSRKtGhHkZOUDIabv|j[L^+;Uo5dNsjc=[lNTG)LO6aLUkufE@qkv&tR#|uc05gXd{4OuR~hypH=j|p]v!vhJwX|JM(/QVE3q1F3!|!ILl8T-MH8yHG5K=0
::KFS9-]H/8C0oukB2$Sa-o0/wvibfeZVHCqSr0azshzHm+&lDGfPFUAuv.YeG+oGDt5^g8IYvY~kt3mk8{bQ=4gRdl4bjrSVtZJ!at,9~Ec9m@O7)pwZtnzhtynr+H
::NlwTqnfB?QhQ0^8bAJ}IF{QRT22jFqW!{=m9R+2{u;JArzOkrsWE,)V8y@DZg.0EkW&]|n-MpLUmigIRM4IXaDfqEc8#q5JbcznfeI-7l(zKx^Tt[XE$!Tw#tChOW
::o1N^AEK30~!1k$$ZOAIM5/e-&)RXDo80{F9D!nIbR[Ms4HpXm(MBmZvartst|54Nwrp3DM-+RbPKa-MNONx,^]N|}?&OiVYtBbfcRcDvtxgza(.P^K~pA$2LunF(B
::q~zlBf/p6moL{hof;&,$-sX6.gMHhDzsHoL&W9+^a?~uH5T!YnG?dM/v3-u1X#S7qOhi9|sOqxaWoHbnt[&rjF{@JIyEPpM/rPUf!hwqEicyQ9i.U$~9YKtljQWVg
::o5ZBR(=k_OQz|A12CW#d811Q(e}t~69x=9YS).|#uBSye#-.]FL5XC;1)(6y#iI;g194/jrN6LvvN17bZB;Q-P?cIiRM,fKBR-c&qBL0?e.14,+K#U_fi5LgQCTT/
::{{[-qWzL11B#RLoQWcT9Rj#+61(DYYWOBkMB=PDIj0L28;Pe5aRh&HUtc,Y[6t!w6C;vhY(f^+EOwXy)P!tv];13OL1WSnE^YdDi[u]m18sKzfypz0+Q/kcJ01)Bv
::f}Op5&EC(u6N{iPXEAcqx_(1&ZZ38/3x#h8=J)qhwNpY=S]rgknZUp_d_&+IR7[LGHmupMs=26HMF8{v@9&|jFO8wo9f~T3Yf^)yT6b~B]0I&K9hX_c.}oNRglMM6
::jG1wPU+ryzLbA0A=_wZ07v6bQvKQe2-^)P(/X{?CZoPXi!VdX,8wwWhWA.mr6OEQ3oOl_O(/hk8s~t6acmd8x{2~8W|7_8DHn-A}OBCaXl[d!3GbPq41{HgXRmHwy
::W.-)1GO/xnVX=xaIx!wt$-5(SqhpO[P^d^2+$HqLc5]!_8^pfy-/VONw.982$SKGnkui}+k!g{2k&f_el5[![WZcN!$mGZp$Rgf?.W6|1Z?G1[TkCD}R+ovomf$8H
::5YIaR[B
:amd64dll:

:x86exe:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $d\$tmp -F:* -R; del $tmp -force }

:x86exe:
::O/bZg00000vK{~c00000EC2ui000000|5a50RR9100000N(o.=0RRIP07L++00000005js0BQgLV{Bz&Zf|pNa4uzdWdQc^MT=k^07P4WfQ;wY20#S?GDBb]000jF
::;nm&MErwD{qEHrEKpp=lE4!StIjVhYGEh-9LZjuPWy,GR62Jlhej|n&_uKW6JX@WuJ4[Z_kYt3O-3H&yT7Xj+o4fURrH}b!/-?A$DH+m5C2L4!4FJpx9Mk~.J=kXZ
::U/CFjr&Rgvoq-h4u$rt9afgky.V1ZJS$m#hvh~{a73J9WY_2l~Yx~|L&G$l}k92Hp8jj}bq(3!5(?aM,uZbN;6eM6P0+iY8xIn]P|NHNMK?Pqe00319Kovt3[14wf
::Gv8x2-Z$F|L5ZeGO8hSfy&OR}5@}vCsQ=|MV#.(700|;!EiMQI=wAwEwmS^a)T42tWFa4HX|+uAA=#L]Hg{?d@eX#1;/l&/?B/Qr@c&!d/p@T?e&DgkS@J|/)Cxv!
::YiYHHOl-=mclY[+rH}t.^hQ_=9cQzpNk7h.3(0R?5@!@Jo=u@rvv(N6tc9_c#IhQ^{@nV1G!ZFl.MPTbqG00FD6It?[hBUDxvbs(5+7}fD^r4,c#+s=]Q)wGfCee8
::xsM;#1g{S&Nb_72^^34Tu0pVq&i90Z_;jS_EOYmnZA$X,49L&_$5OK)wrh5t;k#uCuJbF;P87OqxG66jIDjCB!asve+[Rp=^GRecapW1ufaTT|6tnSgL5Lq,nG17!
::_^H9xIzJik)=_xRS-xF-3g-.;L7IJ^+^b{[v(VP=pbG@9LQ_.E0W}MFiD0p+Lz[T(b0GI&!-^v{a2JC$Gu.6L2tPm(K=}fP6|y)Pb]zI+pb3FDJhB,|ZUN?3ki7H$
::2r^0s0zqI4fENbLbKnzgL)0MN0[4r-{eW}e^/F(e?/d@|NCy!2fmz@$5t{tyKR~UFe7_Lj#$F6c21$R38Bq(ZtAec!MS5PANT=8nI!)Pb$w3O?7v$I7IeBBU!kY-p
::klvN{f,gf7BB/pyCUxy05[+iLC}o&&P.i.[E9rWS;{(za0wz,e?eXya7nNwlDaxU-NjPr]F0(sqn}2DMzi^Jo,Se5.MivntdFL0NCY(!V]A2q6,V0@W24,M+Mjh=8
::F(;^(aEZ},)2FX/lAPFu_vvv99VeC6O3_Qvmpq{igQggsMAzkL8X+yGdmhuD0i9VFc5vX@9Tf!Flv.,b3SDfBbnc9t2TWOtUDE?7lcxWS-IkM-hDLnpL]T&jL+7w9
::u{nv67x?cEdRSkTF^E~mH|6$dgT_hrY0$T4XQNAAJGU6SsX2,Ug?y1^9{AHqv/CLZG;WGi@bYS-;H&3Cv)}@IuTncv@$u;|U6S22o!&Ua=alSCr47d?C&4Z|a!2/(
::U1=8$[tf~bf{EsAS.Vx&=.K7W/L.Tx==RKYYvJ1IY$~RRk.I6A@Mdj@RO/jFhnYYn[50#,^x,8mYIjotRG!)&?mBpJ;.GXOQ{B!(Pwk[(&VeGZCOtyBm[cUflKTF8
::9n4}c(4sBwa|Sh99SG+K-T|2e,G;Jzuv~FFOSdTU@+bRrTAb0KhSNwV#(c6Vh=I4CM0xt5#)q(wNi=@k^M_iVUaG|tJl,wbCPPvD+Ks}?v(3!&8Pei^x)5TNxIq3@
::)HZ.Kf!2,{tQR@CE#6|[qchRl]v!@Ism#ppc0oY5q[J3~ayWfA0Uk;k_!6lWGVo0E/#oOXC-Z2Ca!C&~2d#x98r8ikG33JLeG8(^hKQZ8(5,TI3v3;)-~-FSCK}5b
::no#nRZXg@95UA98Xm5b1LFEYqM0$Zr;;?&(M_$vx{a}n39-eA4fj$eUBGM]V3[$SI,!-7i2]knF=w?j^]wxH0X^k|lOj6?T)IhY[0qYE,h|y.HO9U_V@8ic-1a4(|
::3=0zgh_7j2kcQ4CCLz5{0!RoJ$nnPIhh1obDO/ZuDDoq}HB5xQO$xcTp&iDcSL2UtCn0rQk}~?;^b-8zlqNTYjRsA5W1;#6;WfL8lEC1nqr?hFJfkssE.Je-C1pzF
::id^)GNFx&WJe0~!FXK@e[b/c_08BagZhlgN9e]3ORBXAyI|ztg|8{r]x3H!Fi{#T3a~qi3bcE[FZaO{_q..D24;NJefnrSCnghm](=5RQHK.,hIq5!nB9K5DW]zHw
::rhwb,gvPbM3^(BTVnp$eP2dYzs0^+ZaDH{jRS)NVxAgNZX#UFBykAs|aGy;0jiPW#f}|?jCa7V0e-tY(KbVjp[{{2mF2zOFKOU=USnU8X41MI(-hD64HX5qqioJ;Y
::p}0Zi$48aWXsey}OQ@-6t45zu9DJFoPYOuNX&a,T]xT4~7],Vy@q@lFcwZDyklB^)&9b5hGx(s21oELX[@nVFebC/ZYXpEmJ}Px/LM.2aB[CxQ)fqGBV=UTPq-tLA
::HAWSn&BRYu?X0g@8c3}@P$]pFGRS&n)&Eb6@=d@{2O}u~f)(^7MV36[FX.=F6KZ|0yo=Zv8QMPEy-=xd9wz_FYZ+&mZ~]Yq.-QOZDA~?q4NYSBb!g_hCV~OPV]OOq
::A#7$5D$3Pdll782#-aBb)Z))-eSDp)p?@}m/c7ZBvw;LOs9k5MQHQQQG[L2b{k#hnO^OiR6b.A?No1mqHJC/5BhBzrnRK3GU!hiIZU!12w2pWvrV2IVY!u-{[FN.2
::CG?@Oh-o?+;W4B#;-?tQCtLYdGzD3AFUdDuD(mSc,XZR6bDQ)}h,[J.SI-AFUanfXRkQwoL9~Plr51u!NphusW?i$0c2+d.KY~]lEef1Osp|?N3gA8C.vxw9mvx|H
::ML80+sItx5Ws,x0qjJV)/QFf^g;h=,w=?]mEs$=t;[5CGo!R~8mvoh,,U|&[+qm@W38U|@8a@xJ#4{Cilfa{;,E+zQWn?BN=X@[efk|ZPZ$48]__mOk3@m.=dZSlS
::D$5#k)rj9eRfVM@BZ_~N6b#mY6y1-_k|5Sn#D.neT!zN^,l&#fnU5j/)Vd~tC6nJBxeI&oz;fvUa6V(s#Wko_x8)7mqJ5/+OA),{IPh|XoC6L2vg-VG3JTwli|S#n
::WnDl-Xh+=GEIH3N.{5JHa?nU.hmx,73bvr3_b(?i)K7N]wiSx,R^Q1O?$o!0fR{nTp|=R,;O;N-h.$P)5K^5[4}d4qYKmCvy~FAU,]8)]3+,aLeHZ6IZ-esWs=T7V
::+vl)3PmoOkEp3C1u9t=m=uAx?^q$j#FIv7+n88RPi)4&e3)w;hk9~Pn@;w^((.VUhM@UJ^Mo}Iov=9N|8wZoCvEz;(0yqu~&qE(UGjW9fiYDIM.kuwhl0i[IL08Kc
::D.|@weChO}vuCjUjt^R^4m}V1$87OHXsYOr-TXEab;H!VA=jmkLgGV]HUoPT0fd/;84Q2^ANEl0{hkQ=f9?m_S4Fw.I]OL9.3,DJcy36q)1-Uu1{Xbh@1p[}wD7Eq
::uRH60Jn)r#^;&ZE$u2hpctdBr(OKV^?[4d]oj&5xcWRFAl#6hLH,x6emhNzLtfwf3eg-Qek$#xwd}Udggkb86WfSl.bQU(GA$fAMxi7CVO7bOM4=uT$V4b7cY)t/~
::]~u$uaPVBkYA0j2.ar|c5VPph&YDf#h4LA2snwO1cQNX.@^4kM$g/&|gQQ6fMbbAg9(Yw3q;i}vqgU;Z1tI0q{1)Jjsw2i,n[2l9RH_wT;sds]B5afhs1WHOOM@hX
::L^icEHiKH-vzQdI75=PDIcPuIXYo_9r{,ADI_QIF!1$GXXG?y1{]hA]x@}o?XK;,Neo1WUdswiX2P#$y5;x](k1$0F3T3IzvoL+TlRqp5i($M]vQcG=m!C_unlHSU
::XxFE&TB)y+(Dttd12K9),/lKt!&n$yGm3#SHJZ6H9.#i6mgP2LFM0gvN@THi;z9=dWw;2MS~#X,g.FB/[_##+avx}eCDk8CwYL35u#,gCC|XlPM{V/.2!x8|MWi2A
::_uOOqD!iEms-#2Ccc,u&n3tvI0RP(d;+PCk!aX5a,s2)gPsI2c@fQl,r2NGEZ]d/2{{0hlDzRW(g=Fn1RD#ajPfFgfe|gh#Lm]Z57nccY_FC4?p!jd2tQS)P@7~oR
::$qLD2aBl=1)V?&q57mPKbI;bJ)3eTld/ZE1ef6xj8tJuDbcD_4/Vtl&Na~-rX!v8ol[Ps7cBlPNaRi]{QG25kgfN@BtgVYqkRMYpz-Aj_J)/Xn74Nd$!@8h.3f3iR
::crB1GsH4;$M#KhL1J7e|#{.x,M})g;=6+m{FVFsbj2oXo=ZxOY{c_7nfvp+#B/I^$xo+AO(scNe7oRtDVDasW-2zU&-]OU4.)NRI2vm#8@0p|yUOv[2@;WF-1;d_)
::reT51IjF(D#,5Q-gA2myeo0B+I]67zYU)_=P8IxM,a3fkPxe{36TAeln$qh_3lJeoX~,JiGvkS@=(,SW7UA7-tY=sW-iy}qgBuY)4$Fo,SXCuff&20M=X4xazH(u9
::Bei}DbHzolqGC;$8!2x/pIC,r;!zI9Lw3iLaI9(.I#nd&ILsXn9EqoP1dG_L8|n26Xt{#gXlY~Acutpek,A)X[XW^74(TIm?kn!A&2+jJ^5Ge|uINpNwj]8q8Tab;
::wS=8d9+#3BBHORt0V=g5XQyRFoO89xK1O,p^yJyPuC6M_[NU4Z$Y1!$GX2-JoZB,t.UM(V7E;_}R4Nhp1CNmp4VfzAH1IUW&bf4|6;|^fCz49$,LN&yVWdH2y0=ad
::;)SsZp=tbiMc;SPYwMlT1hA,Fq&1@aDx-U[=FX|}eVi#=73g/cBjxD$a7lbh#WV1uI!hf&K(mOXrX?(z?9OE8{T~53q;2=FNnOg_(ay5BIeVASq[-dcdk1kSEoS$c
::j9qc}b.7KAOIA~aoyyb^.4LNba2S,+ZpZ4M]te2_86N6G+4-D[QQfHJN}4H!YJ[@S/LZaw2LJxBnY6=@O}7C,4yDwvifo6CFg$|w3beFUMEg|NJkR}H.encM@[qpo
::b]+L{BjU(^W_.MNF;HIZo;AF.11sfnM;gp!_WXccf{w4wnB(~!4mA3mmc&=EdeB;o#I,n}&kP1Gr8zgO)Cn4(J{1=W$.2p(y3]ptJvRU{JiN$7VC]hX=oTAx]X/8;
::X=E8RfsVHB5=j18p~ChVW!r=vw#QGiWKDR)run1f{ttDR1OwYy+(&w8Jtb3l+(8M(#ZQR)=(8oKG?o!6AEeA5nU?bNw=bA1&cvwi5,4}rI6QG3=P4ELa;0!djfKHj
::-C~xZ]R~u.g[2u6n,y_R@u?XLX#5&=V[3x4Kb=KG|4y-kwdk}WbZkqn!ib73fCW&kMp9TWS|0oBB=qCD7]~Q1C{TDrq!j8)pJFJxQ[N1v0T7mq8TSf|B|&3ue-C#s
::]qOX.Mg[|hX+Tnt;Ut;qey7Yq&.h~E^dzuP@OgMy,+r8NbB{S@CuHqu3EMV4zA_3EXl~U}UI3eIXZ[5bLrpv&ZZ^+wMX~bQ[Y~tqFNm&#gQgsD_0{]avFMlc,mrW8
::2i)Mv${breta]FkiJ_poz,K5$hDp&-3EWuuPz1cJTPA5.@o.OvB)DpD&Um^}-)k@ZFPT-d|G8^Ei{QH)=BY-GoOr/[D,NnFI3X#x/RThiEtM;q,q@L^Za/e9mn3dT
::dc/}3tr7-om.|Z=e6kDv?+i_ki@Q3Wrp03AXJ-Cvx#aQ/h0P3an5V|hldsoPU3w1JJn9Cs&e5~Sdnk}7I(;V9MzUnNSj^ghvUT#!vL{Qggfl@^kBfQTEY1z~Q+ZH]
::Px=Fvao7KfCFZ_],KpVuVJoFB{P0ujd6|rxN#wHQ4Ka,Rg2mae9n..1?_Qe$R1-iSNKB3HJXfakQ$NdJRi8{l.WVTo^~gI-bg|ue8ffpgwSDwZz.S(ZwPw#_JK7_#
::h2aBvhEyj^p!I[NT&cC?pPe!?yrBFx1)+.,n0|&tyFWDvqE)/;Gm.qjcu7TrfUm^PLoLi.AjoY2-lFEUMuDP7NgViH+StFbGHHds[He-PCd&zwccW&[HZD!|sAk=K
::[?5aeAEo=Mf]!{$x7YTcxx4;2zx6.0MuA!H!CK1PzJAaKcMdtY-{kWhH!?R-jf&#]!QCq61sC)6Fy1{Dtjc;,$e,#uQU8f(FResBv{nnuQ585PcF{RVLwUi42#wL|
::,Z+bUk6d!UM_^)wx9~odBKD&Dp#CS${y=~Htg-_a{lyS@?bTO~48+5xJl|skjA;M9/kaPHPhnQ|_0oZ7XVt;}a.S!1=?cu[Bft8mNJkQWV&fE8Zz]+m2^EBOs.S&g
::Rl9sY&^/b)F{bp2.S,d02[d(hCi05Zc0klLTGX.8HJ]tXv9nFf&M,jl8+UY1@J!,HFODDM7Vr1C{$!8@!{AYU{J^TKu.&^X=qDGug1T5LP5fTN3+plFl6tH)H}2ch
::cCNtf)RlB.lSrxfIFTMp?!_NcyZ0I!9IaiCF{]x!k(2PkeBI,)_fR{.PG[=]!ZW+ms4qG-m)C)[TqPLq;45snC)hMc4uEVGMX9KLQ?xTRr{uKrpMQ_qt[Q0f1;rIU
::j/Fr$V66/z--5{9)CUPO#J0=xr1@,@/T=N7boUQa_f[rH9[?l,+Iyv{]@9Bt0X1pTEj7/9zN],cHLKOT,Kk=wx]^M9L#c=Dd{l#a8^H^p9JAT_tu5{/8r#,XLNb)Z
::^myvpzkW)PZHa,N,R@+XBVwvJZ_1sS88x[t$1jNLG7Pn.]-H=DNY8LkA8Bo08EM4Rm6g(zvbk-S8U@kOD[XJo?P2j{PPb?poAZD/ORG@WaXhy.wEN~i;^w_MQ7dVq
::o/y2GXKDfqWu;59|7paKEuiQZfsK_F]2m4vh$]u8(Q8Gbu}^;80$xAYL0X]F)j6qJ+!0SZ3eUMfVcUl2FHtzS5ssKtfMJHAV_euMq=r$!,-buZx~Lp~?dJDw/D9Br
::{c]U#yDGNxKTUJ@+o8@e&?9xg[Og9ukLoQto6wAc0^raL!|jFo@E)yFG-Q,EdO;o2.ZQC^uHmetY.haA{UkggrdGAV?R9a4w.A_ze6.B)tj?LL@|ZDwP!jHKJ-#hx
::eAyKct(K}~?t7|c19==ZA8H=+Vf52I+fhe-cEdpnsSWhFbrgEgIax|Ctafn689K6zW}sX&Iv4_f^rH$GFxhO~HCzNs2IS[U[DJKXtCam8@1Qza]FY;AgwLSvR)HU8
::&f51J5t7eU|2TtUyp,{ny6DoSy[ewn0+F2.ZrjgdmR{gh|I=H9iC4]5.9O5YssV_b=Ak#MESTY09xn^M.lt_VO{Z7bZ.p?Ed2pO|s-E4Y1v[]&]!zo0VcS9ww4F,o
::m]1/i{.F+b!@aI/bv_tzARO?[_K9MHLPi#6s[#,6TcFTZ8?i@gk_ZQJ-XH$lq9lvG&AXJ.=y;m~fTGEg(ZzAIZj!-K5q7HkKe01&tt_LabaU$hQ5g)d&LUscWIGe9
::PAPT&WKW+3px,ei{J&fynH@e6_d6CP&Oys[o.kA(!Wlki[{9];]dG|Db6I;w5$-?!!JcE^GQ)yq&mlCOl/(E^ddwC_w@DOW.1]|-ayE&6uA1i]t@)4$2zT]Ga;)Dg
::9J3y^Smey_,MC^#SE{xM7!Zhef-aj43h&]aiXvH5ONC/iDROSt=F7,)ySDp6S[Hw43=qEhIM&Roi;^V8503c{RGloxT2PSwNhb4]@Up-G+.GLsLLjSd+-;P6km5Jx
::_kpt}w&Y!KPFq,Ay##yKzaX[agUm|)Mig]xB$U2!l#W5zn&/_,=e5_J3~g?p#f3wYkch8-#]f}E=?,/9iUErT(b-y|7+1SF4.PC]7Y&oFGII9R)_8NckprE|6]#sW
::@!QG8xEZ/TpY@ORLVK|R)TthHv5bnKGood4Hvn^{)cePcy$jsmyUo10K$?3Ln!k5]IC?ZADn{h@ZSwZo0AF$L?jw#;+gMnrQZCh?W=CJ[M.f(&739~u=)p}}j&GVG
::IGNR7i)lv-_g4~2Bze_l5cwp@H0|D$^)GVSHB2Tf?t#Dy!@A[Ao!tCC4o,u)N~i{8a^8D~saKIjGaudgJYI7DXfjh/,BKt}z5DybHJUklG=W[?H/Xs47zp_+{cOC9
::BtDxII)-QxSD6p9cW-$PGKGcEyhf,r/T7-m;#(JrNK+17NdWT)X+{iL)9SvD^g+U(MwWU0q.jX/8./[=)vwU}{/_h?KtVmyoqly8CLgrnE]KML#}Tu~rY^f$5L4[|
::Lu|8=FO9W_9~.W-N!e1X)~tW5oVxK;]&Aak-;aTOwuPOY1P-6(m/d$rc006)@el!~=r6S7UDgUnC7j.VC[r.=oAc^|py}+E#p&d[Mt|{D/hkC-XkT6UXkP-b/-/3|
::cUSC]?gE@K0n=9snUx~.U]dUBuypR/!KVGx+=V5((e36cJHh=Qd+Z.;zWU?_;VxBrCBLRaE}PQm)xuL)smUzKO(00marq$E1D(VI=p?nI9?S$5+kcaMyS-H=SaANj
::J4ez}mo.]!zXDyXx1OgbI^bEb&kL9L!V=n~3GC]#DH}_q+XJF$R=7wJ5~Y)PZfUf-@f,=!q6jO/oe$!v0jRK{B+26{yO2)cRP^=p^b8_zeeZ8nv2&Cs5Sy=$Lspu)
::@!btLrRgXcPk!!WQ,HWXqyJ~Kd3e[gA!wFmR385Rob}oLH6Z}2C.kVrc58C6j]6GqHygyK8{H_C2yor,tyDWZz2@ZLU#M5+zxA=&^}=CC&TD1|8-0bQz;-=Y9IB,s
::Y=~A62)4}lcXIWeMKdo(+hfRr7v7_4!lUWL^s84/H;xSvs!rRqN$1ZZoXLBmUOC1~q91q6kHg#v)i+(J0p|pk7C@0P00bCdV{noI[B#5H83=WOfB}GFiE~IxKyib{
::2p~cr^yB--iZf;GumwN_1LDFU41&i,2b6,#-W.zK8CV4X;@9Z~88kECb8$n/2n.;}7z#o#Ac89kEof_NPyq&S93)R6gupliSq7{pFoXdWVxZ1J2m,i,2Z,x7MfCq7
::-1K-St(!fhPhKv(k}&f}liBvqj)HpZT&V_~#6Tb)1/zoE3#;TaKwYF5B]f2@?,Fl!?[4Lznuq/iA7/abu-fmL5VDBl2wV]~Aag^VLn;MqNb.;EM2NPCy]/2jencJ#
::k9;c]$oG.Ov9~cWu_#hVF,?n4Mtk.rVJ^W]W3j)ji(bKvVu~bINV-SDE~&4HBdM=Mzr/)#OBj/ilE[Ou!u{k?fxfmT?[v1]V@4nh$Yb=ee=Lt7W2j8|V|_,DD8^u9
::jGdgF9.nbcA~p$SleUw,6TVF^o8(v/Jo!HnPi(Lhr1#@aByc42&Kto1=}&gd/$$qDmdr-$MutYVM&G96LpGj^lg.HtneO8~H71?t^-,))Oco}al3kMxWUd-P6gXS;
::FAHU^GPkm[Ws[@xGPoJvnO.tZ-2C1U3~?yw/l7&e^hrA#ml0,4ndoMErCuSH?6gc{s+nvGuiml.4j-]q(]3Vl0C(UJ4P!o^GECmVc0;q&LuY~Z2fqi,2hGFpVEs7$
::5PcYZFnu)C/WHzMj0keVHzIcl_Vtf.!b5~)1n7v$[+GyLy(hg|1]=ZuobMLBXT2aC5$5#XICtVZ{SMr,]/quF@;35EdJk|Ovwwvj&OAv#^mBRA{22eJKA]w8Utuq^
::m+&$5cKOo}^-S?S5=K5uht)MM7J8}_Pl!E454l2(hYUhmM9dO.d4xD)I8r#WJ[OwJN19[cnB$n246.I!kuIi.?0-VSx+?|wD@ui4go.qu(zHj_-p+/lv$2_6!ejd}
::KZeJAjQS3LQlUY_F3Yv#pZF+VCy2[JYUDx&h!YYg|4+^tm_Y;;C?0tSYuMR=ER^ef4s0LX9[qsLnM]2ySQ~-D4wW71H{AIU.C!]Sl]G6DfCmH,6dpJ{fOsMCgW_w9
::hs-O,9~vJxA36_f9w0bG[){s6mIp8oWF9a$LGuT~913|bShd5U2hSW3aY&$gu];QrAhH)GT,S/0=pqD!Rb!y|VvPtkOH[Fu5l3M/1RWtT4UupN08A?Fn1V5JVgkiP
::iwT)&7!x!ma87hifKNash+g1ogk&$fCM..~kjQ3YgoG3lVMt/#5-6)+k@=$UNYq6O@3jMX0Du
:x86exe:

:amd64exe:
Add-Type -Language CSharp -TypeDefinition @"
 using System.IO; public class BAT85{ public static void Decode(string tmp, string s) { MemoryStream ms=new MemoryStream(); n=0;
 byte[] b85=new byte[255]; string a85="0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$&()+,-./;=?@[]^_{|}~";
 int[] p85={52200625,614125,7225,85,1}; for(byte i=0;i<85;i++){b85[(byte)a85[i]]=i;} bool k=false;int p=0; foreach(char c in s){
 switch(c){ case'\0':case'\n':case'\r':case'\b':case'\t':case'\xA0':case' ':case':': k=false;break; default: k=true;break; }
 if(k){ n+= b85[(byte)c] * p85[p++]; if(p == 5){ ms.Write(n4b(), 0, 4); n=0; p=0; } } }         if(p>0){ for(int i=0;i<5-p;i++){
 n += 84 * p85[p+i]; } ms.Write(n4b(), 0, p-1); } File.WriteAllBytes(tmp, ms.ToArray()); ms.SetLength(0); }
 private static byte[] n4b(){ return new byte[4]{(byte)(n>>24),(byte)(n>>16),(byte)(n>>8),(byte)n}; } private static long n=0; }
"@; function X([int]$r=1){ $tmp="$r._"; [BAT85]::Decode($tmp, $f[$r+1]); expand $d\$tmp -F:* -R; del $tmp -force }

:amd64exe:
::O/bZg00000!XN-u00000EC2ui000000|5a50RR9100000N(o.=0RRIP08Rh]00000005js0BQgLV{Bz&Zf|pNa4uzdWdJ2ilOu2.08U$gfQ;wo2jFoN0Jl?HBq9KS
::008cfr?;tDkd,4b?GCKCnmgo^dROl1myYJ8Y[swTP1jY!@&qo78NeWY-&N09^};1hHr_Y9(.i&X?+U{D4J1V3T.t~y[NrC7gAuv$x6wA}pz^=v0APe[s&8KHtqj}V
::TY5IvrZCKGsl^3=tMEw[L91Gg)i,enb^]KESYEUQ6PnPL[Rm1^]OpKairi,f_k7_#d8/Zycui;2LNP+/?8^.jVvHV/3ZIgA/&Vni[4x?0ATj]|0surq15oWyy7zi^
::HB6K?5D5ivNx_HH2x8P2uomP0e@(w3P+Okihc$G8JOM=jN9Tgi2lQOs7XlGjmMdOv2e^psS;i,uh!i&mG#9}&3zwGp/f5_)@F|O,.]LA#aWZMV;b2^cYm9@+Yf-kN
::Wu3[lzAt^+/I.BTzuf=ct|_e/D!OGB$;aNt-,G.)k?f3LU/N1?L&(=eOt?t-&pU/puAE|c)k@C8lM^/[wLofdS]+ZAzPp5~0klq9+~vBWYsSrg,h#f?$X^Tv!5?2?
::fhb,_4&PYTWFJ8ANKu2(oAEFH)flX?_!t$&/Y7rYGdKP&S84@oa!3=&vr^|phoNH4jJEU{I$3PmiL,MfP3_;8Ff7W)-J32J-n^&{)PmjMw4,i/o$K4a^~7~nS{bN$
::M})t9wB41(W$F@zS8$A3$+ML}kOrh+2(RxSbgA78ec|B=mYMh!X9CN5@A7Rn@1i73R#/FkYxogzoNL6730C,U$O;J1;+ifuaHKli/lk_p194KhQ-($VF~]1VFjfmt
::[k9]?6IsaS+ME5yT7kWzYjNEpSi}Af^xYvU1bW$9/[5.;etIM/TIKXRQKYv_lY9awS,LO(=QBUT?R&ZtBGX1bjp7dcH1~Ko-DQIxqjy4/(v.1Ve437#G+0e1H&{.P
::{+fqhnGR0qQuMl}VA}w?|D?!m;DJ(c{,t,tF#B/Q{;&9RKo$.6l3GDt)qTt$a6SvE4)a;2[WI-x=LPV2OlE#Q]U@J/6al{qOV+$r#zw7qQN#]GwC{&QdW+bCyf{FH
::.l2n+d;}@K8ybp7]M/!z(YZp&g;4Q&,p;5AdZ1YHoMsGz=0V3Q]aSrn,4trDV)@3$D;tr9rKh;H#s_xH&Cpxjfoyto9KFG/j!dZ/R0i^mi{7/ZJNzsVRy;Qh#tDv!
::(j=CM;y_ah1tU}p.0v[/!usOKsV@g1JyAdNkOq.IPy9]f0WV&D#YtUJ1Bg$}w4+YGaIJrOu0PFEdZ)$5yHKwrYx@0YaUz/+yIOZCHn1l8C0#J)E0R$DqE!E!xn@s7
::e4IXnLHPJEo~gEXB3$y0uBchan,nkO8$/+w4D9@r6QtW6-Mv2[#EJ&0]oYbaXz{yz&Jd&pa5xNenY}|y!!k!Y#n,wzV4]hj//YvTi_]ZpAuWp8|5=Y3q3N4$Gv#tP
::on@|8au7l]Z8BC;=J[#iM^hP=AB2OwiveWFDHkbU/~j2eS0[+|(}7!10j53pgZgjt#NAE?8y]$FX-IJ8=f4=^yb@Dq)D7BQip~tEZ4RI~o++yp;n@.&;gB0kxtBzg
::m-MhEUP#PQuK0u}SnWNAwtgGA$z5AJ=oJ1+Y4~wPp9=83gn8G]pDwzHEG{VJt6|7=hf}@SoIxQ^ipf_i/qN;_C7g&2DJ=Xar+?XlK0//sHmvQ0Qh^/;M{i@aICoDL
::Km8)&d~l2;.4kFP/0f)yLKBG-vX+pY.Sc_ql9tg2yov514yG6Ap+j61Ckp_eLzM0D=0kJ-,5]~QxU=_vfJPo^_,n._ug_Si|LYeTe4l.x8I9jUt&S8JWk.S!2]9g$
::o}Y(bs@dEU;{=lRUHz,ZEW058ca&!-0q(NpsuPenCS2H1OVvG9Vk.T~wT67Ec_BF)Kh&4t(W#U8EW-2qPqBc4@N(B&RS/+6AM6r44y65S__v2Qa#cTS!@0}zB01mG
::VMTCckW0]}6]S$VDIL+8cCES4iLPaVr{6x_vMtJ539;[W^~gZ[fOFSRuy4id^2WpxFJh!3IY!zCNy13(eRtFD)5x^4(WK=.gw~adMd8JdzasL]09Kin;$)dl)Kp0I
::9?1TXF2thKtNn=weZ/LE;!oI.hL13uH)Fooj^xEI_xqSr_upw~F@o^[8uK#PNWU7UmC63PD_oQS7KpWM_Ov~;K1$VoJUcj_^MN&0x.wMZy-{9I6!?B#UFC=bQ9rng
::[?J3wubU^g-RcQmP=O@&tfLX4g^$a$]l[]JVL?yLA@36FON&uJt5nU0?zfRI57/J^U37G$Gd#s8]rIj;e_m{lB^[p=lGu$!r8t-e??Vp$v+J{@iX,^^cRxVM7}~3a
::yLgs24M/uqVQ];sq^]3Nk{I~Y5NtGkc~XhuUwi1uqx[(_xwBv(OUimXvV;Q?2^.@taC{V?WE5~RpAwNi&zop~1r(ZubmzscFDa1O#^-EoahfJ7l{Z?Wq{jx&KJ&aQ
::uUabHg(mFBqA-BU5e+y[4^jFxpeTs.)PRPIa~5FN432lDO7qY0gQcLYpyvybh3LKkmq_aVko!;.!A=lll}|h42Mu8Rt+_e98(C2[]whw]/p~d,kAr-xvH2_&9+P7y
::;U[1DwZ/l}Gk!k@S#Ii^b#dvWlP[@|BmLeSPpOz|D.{=x.PUm_&.EPRXVimL;XVb1Co?j#Zay9Kj-iTHW&_$U38d#DqDVD^D.iMK-#M1_=5g#g@mws3yVHV_s8bi9
::as[NqZ0T@;VH&[b4rS$Os[6G09_x3~jzNDi^#/JR#y@g&js)oB)!DzrAdl4Nh~+rf78eHeV+s_EFb0,hl[}/})4YQJ6vv)S0|Pcrb@RnYUIsc;WK7rZnPCY-hU#kT
::teoab0{[Rs$(oD|7JW;vkE3M1o$;-),tB&Q)}o~71NyqxvghU.Q^m=LJPBDJ]o5_B?pE,beT8N+Dkv7VwvZ(Ewa;]|I)?cjQN]eh07rpr$TxVr/}Xe;{O4MA?[5[_
::E,tc}6n}}RYQ)ACwp/Im6r_aA!U7~LU=MByEbV,uZ^XoQ8GlN3hLu1K1nVeMAszQHUt09rQf?nlSlkfIG[z4bX|.)rv5fEwCx.uCwK]]@gFz6^{d-@b5dDm=tF6ol
::pjzd+eQiZ9UxTxj&/CcO&LOm^-ti;Ua1rEu.rZPs75,NBY0~^/vE{|hSzJ!bLAF7-AF8nAuNy4g!~MEkAX,oz7eFXJilezFUr8Radx0#-U)snZh6[F)_QK0b8C@eY
::DzkL[2bbnu;;xKYe)60vkL9hLoJR_kht[b@O_rI,Zj6MI2c7JPVdEI+f.mM8Igiyi{5.y;zY6w[y+qvikG(fQC8JReJBG7Wv7giY!uHS8)U?+!HZ(p2wd1[M]{eiQ
::e/^~|+w?fGvMO6R|0s?^Op)$CpOonNV3o]VGp9ZF8B#N)gnxxxiq?upk!trQD[P;Yj9?c2[MdYlJM=-jourz~I+1@{YTT7eUcT}Ysjj_or2QzILejM7l)~]RpC}.V
::&c)0ah!Uu?Vcl1N#U-hoPT6bTE[cy0lQjmU]Cx3,mFWjUiSfn.iCVc^Khqgi)Gl4bhpi9=nnM5H?e{CPJ.x1y=-q4D&r,{Uj)w^@bvjcUi?t#JY}V^m^@rZ&K;{;d
::+KOXvK^b,6^/l!X7c3caYJ;Ilk3sl[,wr{J)Oy})a0qr43@6$CrzB5{VzGOY^noJ(rP4z&PyAKYBl&LfI3J_~KOt5rZpLC7BQPBZR{xjSub;K039mSkd#1B1/q$Nf
::C4X4bcfzgbPh]UHS.sRk+tLouwh#PJr!6hSnxnk@Q.XEG,H8k_eW5j,8S]KuDS~Z=W~wTEf1W?_@nri#X}_AjSF)i^ndAFUSH.{$w-&{oI.=hm|LgaR8mwK=r$PI#
::eaB!r-Ov{YBF-d2/703ks0(H|55h4[eN;w[7goI[?_8e+POv[y@#@][X1[s!A+v6ap.dH0]bHPJlAagn]o{3+bdO|dz(Q=cTtq|iFM4F@+.zg#N3hItRy]NOo4)e1
::@@6jJKZ?HGXHu^KaVpu|R]t(LR9s+.d]v?$-V2NHb(/)sI+7_rAcD8e.}{uB-n!b99/FZUJwN6e@5Uf8)lKi,,hSj_hc}~Npyi$[hfT.PWdB;oTEeCVN[=RV8g#t)
::.|Vv1J4r!.g#F{0)T}-P$k[sp{nF?#8{25G75Tf/IL~|8JXd-L1;O{)^r@!uO7xxcZMMzRNDpj(3v/Och1Y{ldGa4W.01cu^ol^_]9Y4#?}gWiN2}2mpV|QQS)3HZ
::iSFOfA?xb7^w)xp6o&8=#kYe]{5GHVnNo9Im;!xy;~&lHc2hV+w~I-5;hbOqz#N^oQ13E}W.Bf[[fkroHRZFm(6lVNSHfm?bx&~db-UiJzbLYwKYeRt|KQ6/Q|VN;
::;}?}Q^@vz_,K=i1F0HrNO[1}C-[x(1BiuN9aV}XNex@_cRw}gy1@)x;o/B5VgY$n#Qp,tIfs66&L$hhw3@0NORUAaheZ##6$)a]0jUj59n[Q)aN9^r#V??hA{_Cy{
::0;,n4+NG#$2$Q.gU09LRY#2LY-/711p1ml/uv}Kyt|!6Tb2|?pM7I!s,u^o([ulpx{_jA21DqQUd4^i(C8y#N9I3xnnKyOcHI9_m]Fko?m$D&?Q&X]FLjy0zh}+n1
::JyZq!tv$Anm|-(},uAig2a9_C3Aj&mpff!]Eamr.Bis[0vo8i+zn^izw0aSOo|LWIL#Vct|Gtvb(u8Ct5|vse6jNTt0ibHI$pER2Obx,(vz?5uCN-}}t4XK|?;^CF
::uGWXD+K-j#f|Sm~]D[OkZhA$9P,lUnsnYKhIAtat1;)0kPtoM5Tt^0tF6{X~IW-mT0)Sj1lCQ^9U}I!A9&gEsrm-QXa6{9)pf{C6li=.HwYQZ}/lb]O?9[AmwmIC8
::H[f]cIIV)sE~mG8ySZ(!W~F^mZBc;F/C]{d,=zgFU1/0={wzwi0L^Q,+fsrh)ukEyeem~5=kqoH-dr]YQVLD&njD(,ql!y4]IzWSQ|KqGWM4^=8(u^MiPy-Pk)aF[
::|Bi/z8D~96U85nTpj$9G$vv;6$aYEzIYg[G{!p{A8r)+vfw@F,BOF(X@U#L[V3F1n5X(;thlT/p0!5bz&s[~@B[VbM5Ii;f,dICT$i]oCn_|HB0^gx&6ZZwArGl57
::Mi@TgEpwud+)5=Bu}THjeOf2FBR=N#IWq$F37@5#UCW-S^9xF72iF&P6Md^;Mjd+KG0c-3a)3BI{@mJMLJgvn]$1[fVM,sPHWuuNAwqlWO4vp2R)$p[a|4CkR4W.o
::tL]Qo=T0(a,K]VS4ehP(,[js4oiTZmy)8CE7m8[qQza-ab@mu3bXgDTIimE}iF@Vr09[6[Sx$[w!70UfUfM?^#qRLhtr)TCk|0n/{0~qe1xlij^yAxpB@H+XMJ!U^
::rW|lr8-A!~(HNr;6!yJ9.~Kyks2pJI4C#eW=WW@k!7@!Mpp2K7g@^guDYe+Z!Z[T_HgC]4SBcsHF4qXEix|wV&T#qG{Tom@+y[=8ltQ=v|6MR5Qh@Lil5J4)G~|dE
::RjDRd#m$56r,0q^[o@j-tHHJ#Y)+2FZIb/v2haa=JI9[1JApxo4($UMhZ{H]mE!H#?yIf+ZN?wenC{rOZaLqaBGBppvwnpTVK@|.XH]/P51vBdf.|Gq!3D5CtYM@F
::,XlPRGKYZ2)n~aehy;+i-pWBU{o~cA{7ce+So]1|.PO|juIMBCePJI_T,Rk-+h)4FPi(m82Z}nN^AskHbi}~+D!R27Ge.Hp0FIBfkHI0LQ|.o$@Hv_~cK@]m{IoI1
::b3+ZU!}4eU.XVOL63?7cF!.VF!SBP_Ucym+y^mWNv?i}2-5SgZ]6?Kk[,)qKh;AugFnu9#bJj2xbJoKI!{85b9|{k~]+y4|WzIg^JUne6.Qgm_[E8Nu2d?[2IN|uP
::_Ox1G)+_RiSjKEILS8p!#(8CP@PogX(&7T[9zs7r4BmObPp,f_(wLNim$c)eg8#HbHct2sV=.dUygyT4SAx&w{#)9vy7~SrjuEHs!0@1x0B|bO=JqX-HK5u6V.Bjo
::5Yi&98=y$hSkWMe1iH6t[^j[|;8e+5QRcUIP-(J.zWKW^o5Q1uk&Ml7aE0X&/Y1/1xtv0Iqc?&Rgwu^X=s8kWf@tm=G#mY6w6|t7FKS[pBb5tU8Yto;&]b8|UU-?&
::Ip+y{TMi8Ifg+CDC/mc~Ie@G|)^G(SIGUc);rf43RBA~~LnN.5TpN0A5{V1QEubUB^,MvvLq!DR.4aKf+L02cg+9Pb1?hOHz6;2ZB!WYNhs[(;g2s+Iy1iX6OO(aE
::OGj@0Kq&#x451z(#6u{{otZ&niA1EAVhC6o2#nCn8jAB[9t7-nAiVJ{I]/@OAu2fzk^2Qa?Pxo.@m#z|y4V)7gtOnYrzF|oX!S.KbXi!_7PToyP5dC3(hAoORu@2/
::$+-l5mk~3Hjk?_|ORbD5]ojIQln_l}fns#qKY,V(EQx3q;;1+J5S&25Zk7QQxlv1R1Yx~2II2/U{Mr$/n.MHEU=|6.7euQ~goeDANYo|khAHiKEG-Kj5Fx?lOetwE
::CaTcFG5xJ#TnUvWWkG;#1qnbMmM#+59O{ZB&MpmK96TwxXR{?n,&D;/HAf^VsF$haVrlr[rrb;bI??4|y?NISyYNsn7vaw6/OiuktK3YYB{agtejRj4ZfRxHgf[Q]
::rS)mRK/;8ix-f25flXLzw/;KgL.eUoDIui}u;83FY6y@kj1ZLnlX9$TRp@.T7)sG7Ary,7ZXOxZ+P?IzN.#lgv=-Z+J4[t07(+cu[_-Qa.De4)2_rFz0hP,;RFa?|
::r-^W#-N9}Aw_B.SRHaW7DM17mVN7W?uoM~~QTn)k!jYt5mt;0pD(c38fEYnx0RCeco;dTrsSz}@/Y#psQ#s8V2u!9&DAY&5z#7qVyC95HjcKYHiE!U26c|]KsqNNq
::2a[&tO2D^ZLTc!#ER/)K3i?L_IwmTfUka[A7EXvUO;$2B_Bh2;mm?zRzON{~,.g8|LRvy.09=Aa!|.T+!gIx2&]Ol.n]4w7q+PRqB/qJmOKpaPN/3Zy,HQMabd0Ee
::CDN-ZIAYgpGs474+^^DN4m[|kKG+_uA16GfG4$!CUca8!6@7|O7)]F]T$=-8I2D?V-(G9@{k_6?.-UX)b}Qs8+xi7[gX4B.5tOaH&fLZ&]0MD(41|yv?3P543csaA
::Sdo3ZjQxJxYd9[PC;xHt-SPJNw8p(6lNfn_R9KX/nn$8H)60fGw!O@]aMXj##!-mO19,g!gnxPvG&vO_SE0KGS/K}@TUsa[fr+{lZ$L)X,CMC-9@9A(CvujT1+)kG
::VNGzd.;jNqq)r.|z0sr?WI=00M4jea)H]v)5UR06YmMfzC/Xvnmo,thV4aC]jFCgD8kOUHf(wlXTht[WBp4eGnQXsN!Zjl@lSc&ewUaoSOxvinC8B8HwTpg4]MEWR
::=G9E|HI62&6l@m$!bq9+QwxofrKE~,[h{k[FFV{/tNI3@{_s~=BVWIU1{amo2u&Y?I3gJXAzf3+TH(j66=,z=,{C.mGWmy+7A6{]i.Mu(#JnfDi7V1,AVDXk+gsaG
::1|v6,pyYrw,PPp|?fMzd5/_b^VXf5]I3dUhQB$Uis/6$xyZpug,#$KhxttekRT^j#lj3qnx]of-+r6+i[tlx=r5vdFu8}Wpx.NgtSarVuua=O.goYGGu?mphMyjCQ
::H?btM)hy]}GvvKAyCh{-sp7W;TF;vvc3TAXx578T2gFX4@h91x3XW[F^JyKk(C|#!Qwl]9UMS6NRAs5y?tFpPDL_/eV.lErd6K$dLypX9xkflexe@mT??IuSHL4U(
::8Ek7M&-j1tIW#kXZB,+X-WW7iIkXv=E^j!eQi.dZH&J!y)}8Ek@H1b}|Ei7F#lzeXB1o,JYxVJ~njz+)7]1k.cMX?Uad@(;9G1,9R7+XhT2^v1I;}Z?6s{VIiX,6L
::jHoGX=/{zbMloQd=h.A7B{XcMispTPRTS;mk08_NsX?dVX=~/ZG+r+s;uXdJ8D.FQ]tQV&$0v!pI+-g.g|i(@U@xHsN[=cPp#HG8^-|;xHkkwNf-VY6rI[WMz4Ewe
::(TGxuJTWl#e#7kHsPL1h]-=5yb;-5/sv)8&W8=v}e9qI(^J&8DUwVae#GQBma7(6@Oh}OE[iRPo@ZpDco!!qY_sE3PuFeB9&c{Sj-S{$Be$DD!e)YvUg=j+sHV|Yo
::6pC3|wYJkXw^.0EOx|Qn+?zH{c9,i_iw~0AGk;S_c^89q5rmoFs/Z|.nbDw/{LQwp8+w]srZMKin/G)dAP[2V!M#Mjo7Xf/.Z$1QGWLI]v/5#!e=_f#z)8d/]t$##
::AeG.#M)ZpqCF;/6A2rqxdk-Fx+xrf|SZ$jhz_H.Ip4Uv)]j4H1N-G~WEGZSIV)@Y4ROks.P)U?jRuzGTSy~vZ98n5]g3lMh34RhqgeY016^K[.f,6U+b(O4r71sq;
::z/dMvxvxFEDjW]B6;hILy2]JAJiqy92V/kPaQS[yz,Lv(zKY+.pzLh@]y&VGTCh{A{KiE;=wX3Ozpg(2y{SX~+R2n_yb866b?o7FRNPv#3Zc6;Ag{WLqJg-NE1mHk
::3P^])S72+)]?4A4D~9w8JV7_h!!K~RhlvZ|31?((IEgUMyl}zJQVS.Kg3S8WSo0Gw3V0B{PiL_,Olx)qQA&S2Y5[YQ,piiHqGH9BuK?f&N2^nur?a]5Qh&({s67_L
::$tK}hGm5,eF8Cmv0X2)xFR;1,YHPoFp~ImHf@x.GIr)Wia3Gj4NH1/rL,JSOv_eoxHn@rIMz3vyzn8@tSQ&)xc-YJFHoAqB8?a|+C2/U==)JaY1X-lF^D69j.QRL}
::[D}{OyNC-!Xwpl(h/gN56dsbn=Iy}R0{N}4#trwc(eF{#=z?Ur)YD.rAJL4qt!ia|i/dIQh-fgQeY;FEYHuuUJZ6Hh)#qOkcKLDBo!pP)kCZ.52|ItN{laRQbRt;=
::1Hbk}cSfUhfZGJ+PY)Hjnescj]LK_AG+t9CrNL#QdMJp?ZP8Q9+I#|28;FJvH;tf[skBiGuN4mP4qlv2#bDd@h?]nBddFg34!8du8g3B!WYQfO9KYYcF44zy3;t~x
::_nO=yx}IDm^,[wTcv(wHDwJ/=1ez0xS)HdlfcjGYx[GYN[6-tPPHI2hgT~25V;gjr)ws@o&fF&X9HgJ~Or;KI]s=Kd=TzPWOiswPL##?L@X8tNGMPBKr)CFqGGU^A
::N?cZy9256@o@ViS$BUv2|AyNc2A1Hl$pn0?QiEqj8_I]Mo8JqW7Po64uv~VxF#tk7PD7NKj(xLFMK7DolQ^4MlgVsM+a7X}{TkU&5QV;#-9lu,+LzjTvM9$$C5j3;
::3NIJ95TJvy])ReFLANxEIp+OyaH#YR^Q-16At@5VX=W!4qOp#&uqPuLJ@BL$6.]wev4w3LKJ(hlxD-[$b)F3taJu=zlEe^pkCa2^pN|P9p@Q+O.7jB6_0zoct4JiO
::61&&&{gV4ELU!bssq!zsgwDL(jZuhzpb9j1@{jFsB1a=[tU)e5!}bKn_JQeEH0tK=4wKqfwju5=-/3?c&g)g.(knu,HqHOs2E}8!_G5[yr^-aj_M^dP)klCFpj/EW
::&#;{(THlxW&(njH#)4-!M.4.=OwOM460QRePm]ceg_(pqCt8wzIbdpv8X|IWrlNTF4@v_wLL{z4D]v@w^skk#^L.gdCrI@|&sqkcj+$#Vi-}y+WAWUC+axS2G-iSQ
::AbNvU4+Q;!54ZE=|5vQ(ceU^PhRGYscwpTi#RK#P=o{g0TZ~tP+DCGLy?l}i/V.w2QTctCZ43~@Z[{Wyu;zlFhP@Oyp-bkcaQx]!;1$m_m&3,I(;(6_gWtayS!!IE
::oDD!X!_=,m&VmsGFu@}Q9]yU}H$3r?p21xV5zaAGmBGLafo^B14SPKNW3ZVn$).9T!2{|B(?O,T.us@chTvM{6Z.D|g8SGC#NF2~xB1MOmLahByIKFj!?X&Zz[P;^
::Dwa3tYhVCuq4bqm.hm;GU!rig#8wk)BLGbx7oh?)3uYq)jiBX#fd9}(!Lrym/Q_CR,spm)m=F$}m|;uxP$$Bko&/Uxt_!c)EhruCPPo9?T$rHCmP=kh@Sgi&x=$m[
::kD&qi-3|apE.2U8z,L;O)1^hd?^MiYI18aR2@~hP96UY0o[hJ7nCN(9u@a,Qgy/mw7yPhdC5?Q^NK24B#GQF^Rq#s0.F6(ZZKqAWKmrdPd+Io;$jGc-8XhIqRoZ.^
::bGRRDAA5b7t^#hC4,OSzaS4])o!!yhkKc87#dpU$KspGOml#N^/0&QOV|c5oyQ;fyQ)Nb#@op1b.K4roy0W]nI;|F~I?BAx9bNZ$;zG6,y1Y8Qy1kv^U0;DF.Cqk(
::JG$&u,+iEw,j=[2,r~8{W&pvo#cr8hnO(RRj2$?TV7suL(aT{!vwOCy,)urO-3DNu??TX-@ELKh-Iiarh&Ck)synN[jk~pXjyoQAlkVLFyUKT&cX2zxXW[@;@=J67
::^.=RSch_64c=vo4?9hln6([8{2Hsk{2c8c-TzDpURq(PJo#Dgai]B=V3ghVT$MHJ6V!R&n9$p[vG2Rc+53di/2k#+BGF|}BRbo1Ey3T,I0AK
:amd64exe:

:spptask:
[io.file]::WriteAllText("SvcTrigger.xml",$f[2].Trim(),[System.Text.Encoding]::Unicode);

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
[io.file]::WriteAllText("ReadMe.html",$f[2].Trim(),[System.Text.Encoding]::UTF8);

:readme:
<!DOCTYPE html>
<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>KMS_VL_ALL</title>
    <style type="text/css">
        #nav {
            position: absolute;
            top: 0;
            left: 0;
            bottom: 0;
            width: 200px;
            overflow: auto;
        }

        main {
            position: fixed;
            top: 0;
            left: 200px;
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

            <h1><a name="Overview"></a>KMS_VL_ALL - Smart Activation Script</h1>
    <ul>
      <li>Batch script(s) to automate the activation of supported Windows and Office products using local KMS server emulator, or external server.</li>
    </ul>
    <ul>
      <li>Designed to be unattended and smart enough not to override the permanent activation of products (Windows or Office),<br />
      only non-activated products will be KMS-activated (if supported).</li>
    </ul>
    <ul>
      <li>The ultimate feature of this solution when installed, it will provide 24/7 activation, whenever the system itself request it (renewal, reactivation, hardware change, Edition upgrade, new Office...), without needing interaction from user.</li>
    </ul>
    <ul>
      <li>Some security programs will report infected files due KMS emulating (see source code near the end),<br />
      this is false-positive, as long as you download the file from the trusted Home Page.</li>
    </ul>
    <ul>
      <li>Home Page:<br />
	  <a href="https://forums.mydigitallife.net/posts/838808/">https://forums.mydigitallife.net/posts/838808/</a></li>
	  Backup links:<br />
	  <a href="https://pastebin.com/cpdmr6HZ">https://pastebin.com/cpdmr6HZ</a><br />
	  <a href="https://textuploader.com/1dav8">https://textuploader.com/1dav8</a>
    </ul>
    <hr />
            <br />

            <h2><a name="How"></a>How it works?</h2>
    <ul>
      <li>Key Management Service (KMS) is a genuine activation method provided by Microsoft for volume licensing cutomers (organizations, schools or goverments).<br />
      The machines in those environments (called KMS clients) activate via the environment KMS host server (authorized Microsoft's licensing key), not via Microsoft activation servers.</li>
      <p>For more info, see <a href="https://www.microsoft.com/Licensing/servicecenter/Help/FAQDetails.aspx?id=201#215">here</a> and <a href="https://technet.microsoft.com/en-us/library/ee939272(v=ws.10).aspx#kms-overview">here</a>.</p>
    </ul>
    <ul>
      <li>By design, KMS activation period lasts up to <strong>180 Days</strong> (6 Months) at max, with the ability to renew and reinstate the period at any time.<br />
      With the proper auto renewal configuration, it will be a continuous activation (essentially permanent).</li>
    </ul>
    <ul>
      <li>KMS Emulators (server and client) are sophisticated tools based on the reversed engineered KMS protocol.<br />
      It mimic the KMS server/client communications, and provide a clean activation for the supported KMS clients, without altering or hacking any system files integrity.</li>
    </ul>
    <ul>
      <li>Updates for Windows or Office do not affect or block KMS activation, only a new KMS protocol will not work with local emulator.<br />
    </ul>
    <ul>
      <li>The mechanism of <strong>SppExtComObjPatcher</strong> make it act as ready-on-request KMS server, providing instant activation without external schedule task or manual intervention.<br />
      Incluing auto renewal, auto activation of volume Office afterwards, reactivation because of hardware change, date change, windows or office edition change... etc.</li>
    <p>On Windows 7, later installed Office may require initiating the first activation vis OSPP.vbs or the script, or opening Office program.</p>
    </ul>
    <ul>
      <li>That feature make use of "Image File Execution Options" technique to work, programmed as an Application Verifier custom provider for the system file responsible of KMS process.<br />
      Hence, OS itself handle the DLL injection, allowing the hook to intercept the KMS activation request and write the response on the fly.</li>
    <p>On Windows 8.1/10, it also handle the localhost restriction for KMS activation, and redirect any local/private IP address as it were external (different stack).</p>
    </ul>
    <ul>
      <li>The activation script consist of advanced checks and commands of Windows Management Instrumentation Command <strong>WMIC</strong> utility, that query and execute the methods of Windows and Office licensing classes,<br />
      providing a native activation processing, which is almost identical to the official VBScript tools slmgr.vbs and ospp.vbs, but in automated way.</li>
    </ul>
    <ul>
      <li>The script(s) only access 3 parts of the system (if emulator is used):<br />
      copy or link the file <code>"C:\Windows\System32\SppExtComObjHook.dll"</code><br />
      add the hook registry keys to <code>"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options"</code><br />
      add the osppsvc.exe keys to <code>"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"</code></li>
    </ul>
    <hr />
            <br />

            <h3><a name="Supported"></a>Supported Products</h3>
    <p>Volume-capable:</p>
    <ul>
      <li>Windows 8 / 8.1 / 10 (all official editions, except Windows 10 S)</li>
      <li>Windows 7 (Enterprise /N/E, Professional /N/E, Embedded Standard/POSReady/ThinPC)</li>
      <li>Windows Server 2008 R2 / 2012 / 2012 R2 / 2016 / 2019</li>
      <li>Office Volume 2010 / 2013 / 2016 / 2019</li>
    </ul>
    <p>Unsupported Products:</p>
    <ul>
      <li>Office Retail</li>
      <li>Office UWP Apps</li>
      <li>Windows Editions which do not support KMS activation by design:</li>
      Windows Evaluation Editions<br />
      Windows 7 (Starter, HomeBasic, HomePremium, Ultimate)<br />
      Windows 10 (Cloud "S", IoTEnterprise, IoTEnterpriseS, ProfessionalSingleLanguage... etc)<br />
      Windows Server (Server Foundation, Storage Server, Home Server 2011... etc)
    </ul>
    <p>These editions are only KMS-activatable for <em>45</em> days at max:</p>
    <ul>
      <li>Windows 10 Home edition variants</li>
      <li>Windows 8.1 Core edition variants, Pro with Media Center, Pro Student</li>
    </ul>
    <p>These editions are only KMS-activatable for <em>30</em> days at max:</p>
    <ul>
      <li>Windows 8 Core edition variants, Pro with Media Center</li>
    </ul>
    <p>Notes:</p>
    <ul>
      <li>supported Windows products do not need volume conversion, only the GVLK (KMS key) is needed, which the script will install accordingly.</li>
      <li>KMS activation on Windows 7 have a limitation related to SLIC 2.1 and Windows marker. For more info, see <a href="https://support.microsoft.com/en-us/help/942962">here</a> and <a href="https://technet.microsoft.com/en-us/library/ff793426(v=ws.10).aspx#activation-of-windows-oem-computers">here</a>.</li>
    </ul>
    <hr />
            <br />

            <h3><a name="OfficeR2V"></a>Office Retail to Volume</h3>
    <p>Office Retail must be converted to Volume first, before it can be activated with KMS<br />
    this includes Office C2R 365/2019/2016/2013 installed from default image files (e.g. ProPlus2019Retail.img)</p>
    <p>To do so, you need to use this licensing converter script:<br />
	<a href="https://forums.mydigitallife.net/posts/1150042/">C2R-Retail2Volume</a></p>
    <p>You can use other tools that can convert licensing:</p>
    <ul>
      <li><a href="https://forums.mydigitallife.net/posts/1125229/">OfficeRTool</a></li>
      <li>Office Tool Plus</li>
      <li>Office 2013-2019 C2R Install</li>
    </ul>
    <p>Note: only OfficeRTool support converting and activating Office UWP (modern Windows 10 Apps).</p>
    <hr />
            <br />

            <h1><a name="Using"></a>How To Use</h1>
    <ul>
      <li>Remove any other KMS solutions.</li>
    </ul>
    <ul>
      <li>Temporary suspend Antivirus realtime protection, or exclude the downloaded file and extracted folder from scanning to avoid quarantine.</li>
    </ul>
    <ul>
      <li>Extract the downloaded file contents to a simple path without special characters or long spaces.</li>
    </ul>
    <ul>
      <li>Administrator rights are require to run the activation script(s).</li>
    </ul>
    <ul>
      <li>KMS_VL_ALL offer 3 flavors of activation modes.</li>
    </ul>
    <hr />
            <br />

            <h1><a name="Modes"></a>Activation Modes</h1>
            <br />
            <h2><a name="ModesAut"></a>Auto Renewal</h2>
    <p>Recommended mode, where you need to run <strong>Activate.cmd</strong> once, afterwards, the system itself handle and renew activation per schedule.</p>
    <p>To get this mode:</p>
    <ul>
      <li>first, run the script <strong>AutoRenewal-Setup.cmd</strong>, press Y to approve the installation</li>
      <li>then, run <strong>Activate.cmd</strong></li>
    </ul>
    <p>If you use Antivirus software, it is best to exclude this file from scanning protection:<br /><code>C:\Windows\System32\SppExtComObjHook.dll</code></p>
    <p>If Windows Defender is enabled on Windows 8.1 or 10, AutoRenewal-Setup.cmd adds the required exclusion upon installation.</p>
    <p>Additionally, on Windows 8 and later, AutoRenewal-Setup.cmd <em>duplicate</em> inbox system schedule task <code>SvcRestartTaskLogon</code> to <code>SvcTrigger</code><br />
    this is just a precaution step to insure that auto renewal period is evaluated and respected, it's not directly related to activation itself, and you can manually remove it.</p>
    <p>If you later installed Volume Office product(s), it will be auto activated in this mode.</p>
    <p>You can remove the extracted folder contents, it is not needed after installation.</p>
    <p>Run <strong>AutoRenewal-Setup.cmd</strong> again if you want to remove and uninstall the auto renewal solution.</p>
    <hr />
            <br />

            <h2><a name="ModesMan"></a>Manual</h2>
    <p>Easy mode, where you only need to run <strong>Activate.cmd</strong>, without leaving any KMS emulator traces in the system.</p>
    <p>To get this mode:</p>
    <ul>
      <li>make sure that auto renewal solution is not installed, or remove it</li>
      <li>then, just run <strong>Activate.cmd</strong></li>
    </ul>
    <p>You will have to run <strong>Activate.cmd</strong> again before the KMS activation period expire.</p>
    <p>You can run <strong>Activate.cmd</strong> anytime during that period to renew the period to the max interval.</p>
    <p>If <strong>Activate.cmd</strong> is accidentally terminated before it completes, run the script again to clean any leftovers.</p>
    <hr />
            <br />

            <h2><a name="ModesExt"></a>External</h2>
    <p>Standalone mode, where you only need the file <strong>Activate.cmd</strong> alone, previously refered to as "Online KMS".</p>
    <p>You can use <strong>Activate.cmd</strong> to activate against trusted external KMS server, without needing other files or using local KMS emulator functions.</p>
    <p>External server can be a web address, or a network IP address (e.g. for local LAN or VM).</p>
    <p>To get this mode:</p>
    <ul>
      <li>run <strong>Activate.cmd</strong> with command line switch <strong>/e</strong> followed by server address, example: <code>Activate.cmd /e pseudo.kms.server</code></li>
      <li><strong>OR</strong></li>
      <li>edit <strong>Activate.cmd</strong> with Notepad (or text editor)</li>
      <li>change <code>External=0</code> to 1</li>
      <li>change <code>KMS_IP=172.16.0.2</code> to the IP/address of the server</li>
      <li>save the script, and then run</li>
    </ul>
    <p>If you later installed Volume Office product(s), it will be auto activated if the external server is still available</p>
    <p>The used server address will be left registered in the system to allow activated products to auto renew against the external server if it is still available,<br />
	otherwise, you need another manual run against new available server.</p>
    <p>If you want to clean the server registration, run <strong>Activate.cmd</strong> in Manual mode once.<br />
	or else, use this external script: <a href="https://github.com/abbodi1406/WHD/raw/master/scripts/Clear-KMS-Cache-20190329.zip">Clear-KMS-Cache</a></p>
    <hr />
            <br />

            <h1><a name="Options"></a>Additional Options</h1>
            <br />

            <h3><a name="Unattend"></a>Unattended Switches</h3>
    <p>
      <strong>Activate.cmd</strong> command line switches (case-insensitive)
    </p>
    <ul>
      <li>Unattended (auto exit):<br /><code>/u</code></li>
    </ul>
    <ul>
      <li>Silent (implies Unattended):<br /><code>/s</code></li>
    </ul>
    <ul>
      <li>Silent and create simple log:<br /><code>/s /l</code></li>
    </ul>
    <ul>
      <li>Debug mode (implies Unattended):<br /><code>/d</code></li>
    </ul>
    <ul>
      <li>Silent Debug mode:<br /><code>/s /d</code></li>
    </ul>
    <ul>
      <li>External activation mode:<br /><code>/e pseudo.kms.server</code></li>
    </ul>
    <ul>
      <li>Activate Office only:<br /><code>/o</code></li>
    </ul>
    <ul>
      <li>Activate Windows only:<br /><code>/w</code></li>
    </ul>
    <ul>
      <li>Revert Windows 10 KMS38 to normal KMS:<br /><code>/x</code></li>
    </ul>
    <p>
      <strong>AutoRenewal-Setup.cmd</strong> command line switches (case-insensitive)
    </p>
    <ul>
      <li>Unattended (auto install or remove and exit):<br /><code>/u</code></li>
    </ul>
    <ul>
      <li>Silent (implies Unattended):<br /><code>/s</code></li>
    </ul>
    <ul>
      <li>Silent and create simple log:<br /><code>/s /l</code></li>
    </ul>
    <ul>
      <li>Debug mode (implies Unattended):<br /><code>/d</code></li>
    </ul>
    <ul>
      <li>Silent Debug mode:<br /><code>/s /d</code></li>
    </ul>
    <ul>
      <li>Force installation regardless detection (implies Unattended):<br /><code>/i</code></li>
    </ul>
    <ul>
      <li>Force removal regardless detection (implies Unattended):<br /><code>/r</code></li>
    </ul>
    <ul>
      <li>Do not clear KMS cache:<br /><code>/k</code></li>
    </ul>
    <p>
      <strong>Notes:</strong>
    </p>
    <ul>
      <li>You can combine multiple switches together in any order</li>
      <li>Log switch <code>/l</code> only works with silent switch <code>/s</code></li>
      <li>If Activate.cmd switch <code>/e</code> is specified without KMS server address, it will not have any effect</li>
      <li>If Activate.cmd switches <code>/o</code> and <code>/w</code> are specified together, the last one takes precedence</li>
      <li>If AutoRenewal-Setup.cmd switches <code>/i</code> and <code>/r</code> are specified together, the last one takes precedence</li>
    </ul>
    <p>
      <strong>Examples:</strong>
    </p>
    <pre>
<code>
Activate.cmd /s /e /w pseudo.kms.server
Activate.cmd /d /w /o
Activate.cmd /u /x /e pseudo.kms.server
AutoRenewal-Setup.cmd /s /r /k
AutoRenewal-Setup.cmd /i /u 
AutoRenewal-Setup.cmd /s /l
</code>
    </pre>
    <p>===========</p>
            <br />

            <h3><a name="OptAct"></a>Activation Choice</h3>
    <p>
      <strong>Activate.cmd</strong> is set by default to process and try to activate both Windows and Office.</p>
    <p>However, if you want to turn OFF processing Windows or Office, for whatever reason:</p>
    <ul>
      <li>you afraid it may override permanent activation</li>
      <li>you want to speed up the operation (you have Windows or Office already permanently activated)</li>
      <li>you want to activate Windows or Office later on your terms</li>
    </ul>
    <p>To do that:</p>
    <ul>
      <li>run <strong>Activate.cmd</strong> with command line switch <strong>/o</strong> or <strong>/w</strong>: <code>Activate.cmd /w</code></li>
      <li><strong>OR</strong></li>
      <li>edit <strong>Activate.cmd</strong> with Notepad (or text editor)</li>
      <li>change <code>ActWindows=1</code> to zero 0 if tou want to skip Windows</li>
      <li>change <code>ActOffice=1</code> to zero 0 if you want to skip Office</li>
      <li>save the script, and then run</li>
    </ul>
    <p>Notice:<br />
    the turn OFF choice is not very effective if Windows or Office installation is already Volume (GVLK installed),<br />
    because the system itself may try to reach and KMS activate the products, specially on Windows 8 and later.</p>
    <p>===========</p>
            <br />

            <h3><a name="OptW10"></a>Skip Windows 10 KMS 2038</h3>
    <p>
      <strong>Activate.cmd</strong> is set by default to check and skip Windows 10 activation if KMS 2038 is detected</p>
    <p>However, if you want to to revert to normal KMS activation:</p>
    <ul>
      <li>run <strong>Activate.cmd</strong> with command line switch: <code>Activate.cmd /r</code></li>
      <li><strong>OR</strong></li>
      <li>edit <strong>Activate.cmd</strong> with Notepad (or text editor)</li>
      <li>change <code>SkipKMS38=1</code> to zero 0</li>
      <li>save the script, and then run</li>
    </ul>
    <p>Notice:<br />
    if <code>SkipKMS38</code> is ON, Windows will always get checked and processed, even if <code>ActWindows</code> is OFF.</p>
    <p>===========</p>
            <br />

            <h3><a name="OptKMS"></a>Advanced KMS Options</h3>
    <p>You can modify KMS-related options by editing <strong>Activate.cmd</strong> prior running.</p>
    <ul>
      <li>
        <strong>KMS_RenewalInterval</strong>
        <br />
        Set the interval for KMS auto renewal schedule (default is 10080 = weekly)<br />
        this only have much effect on Auto Renwal or External modes<br />
        allowed values in minutes: from 15 to 43200</li>
    </ul>
    <ul>
      <li>
        <strong>KMS_ActivationInterval</strong>
        <br />
        Set the interval for KMS reattempt schedule for failed activation renewal, or unactivated products to attemp activation<br />
        this does not affect the overall KMS period (180 Days), or the renewal schedule<br />
        allowed values in minutes: from 15 to 43200</li>
    </ul>
    <ul>
      <li>
        <strong>KMS_HWID</strong>
        <br />
        Set the Hardware Hash for local KMS emulator server (only affect Windows 8.1/10)<br />
        0x prefix is mandatory</li>
    </ul>
    <ul>
      <li>
        <strong>KMS_Port</strong>
        <br />
        Set TCP port for KMS communications</li>
    </ul>
    <hr />
            <br />

            <h2><a name="Check"></a>Check Activation Status</h2>
    <p>You can use those scripts to check the status of Windows and Office products.</p>
    <p>Both scripts do not require running as administrator, a double-cick to run is enough.</p>
    <p>
      <strong>Check-Activation-Status.cmd</strong>:</p>
    <ul>
      <li>query and execute official licensing VBScripts: slmgr.vbs for Windows, ospp.vbs for Office</li>
      <li>it can show exact date on when will Windows Volume activation will expire</li>
      <li>Office 2010 ospp.vbs show little info</li>
    </ul>
    <p>
      <strong>Check-Activation-Status-Alternative.cmd</strong>:</p>
    <ul>
      <li>query and execute native WMI functions, no vbscripting involved</li>
      <li>it show extra more info (SKU ID, key channel)</li>
      <li>it does not show expiration date for Windows</li>
      <li>it show more detailed info for Office 2010</li>
      <li>it can show status of Office UWP apps</li>
    </ul>
    <hr />
            <br />

            <h2><a name="Setup"></a>Setup Preactivate</h2>
    <p>To preactivate the system during installation, copy <code>$oem$</code> folder to <code>sources</code> folder in the installation media (iso/usb).</p>
    <p>If you already use another <strong>setupcomplete.cmd</strong>, rename this one to <strong>KMS_VL_ALL.cmd</strong> or similar name<br />
    then add a command to run it in your setupcomplete.cmd, example:<br />
	<code>call KMS_VL_ALL.cmd</code></p>
    <p>Use <strong>AutoRenewal-Setup.cmd</strong> if you want to uninstall the project afterwards.</p>
    <p>Notes:</p>
    <ul>
      <li>The included <strong>setupcomplete.cmd</strong> support the <em>Additional Options</em> described previously, except Unattended Switches.</li>
      <li>Use <strong>AutoRenewal-Setup.cmd</strong> if you want to uninstall the project afterwards.</li>
      <li>In Windows 8 and later, running setupcomplete.cmd is disabled if the default installed key for the edition is OEM Channel.</li>
    </ul>
    <hr />
            <br />

            <h2><a name="Debug"></a>Troubleshooting</h2>
    <p>If the activation failed at first run:</p>
    <ul>
      <li>Run <strong>Activate.cmd</strong> one more time.</li>
      <li>Reboot the system and try again.</li>
      <li>Check that Antivirus software is not blocking "C:\Windows\SppExtComObjHook.dll".</li>
      <li>Check System integrity, open command prompt as administrator, and execute these command respectively:<br />
      for Windows 8.1 and 10 only: <code>Dism /online /Cleanup-Image /RestoreHealth</code><br />
      then, for any OS: <code>sfc /scannow</code></li>
    </ul>
    <p>For Windows 7, if have the errors described in <a href="https://support.microsoft.com/en-us/help/4487266">KB4487266</a>, execute the suggested fix.</p>
    <p>If you got Error <strong>0xC004F035</strong> on Windows 7, it means your Machine is not qualified for KMS activation. For more info, see <a href="https://support.microsoft.com/en-us/help/942962">here</a> and <a href="https://technet.microsoft.com/en-us/library/ff793426(v=ws.10).aspx#activation-of-windows-oem-computers">here</a>.</p>
    <p>If you got Error <strong>0x80040154</strong>, it is mostly related to misconfigured Windows 10 KMS38 activation, rearm the system and start over, or revert to Normal KMS.</p>
    <p>If you got Error <strong>0xC004E015</strong>, it is mostly related to misconfigured Office retail to volume conversion, try to reinstall system licenses:<br /><code>cscript //Nologo %SystemRoot%\System32\slmgr.vbs /rilc</code></p>
    <p>If you got one of these Errors on Windows Server, verify that the system is properly converted from Evaluation to Retail/Volume:<br /><strong>0xC004E016</strong> - <strong>0xC004F014</strong> - <strong>0xC004F034</strong></p>
    <p>If the activation still failed after the above tips, you may enable the debug mode to help determine the reason:</p>
    <ul>
      <li>run <strong>Activate.cmd</strong> with command line switch: <code>Activate.cmd /d</code></li>
      <li><strong>OR</strong></li>
      <li>edit <strong>Activate.cmd</strong> with Notepad (or text editor)</li>
      <li>change <code>_Debug=0</code> to 1</li>
      <li>save the script, and then run</li>
      <li>wait until command prompt window is closed and Debug.log is created</li>
      <li>upload or post the log file on the home page (MDL forums) for inspection</li>
    </ul>
    <p>Final tip, you may try to rebuild licensing Tokens.dat as suggested in <a href="https://support.microsoft.com/en-us/help/2736303">KB2736303</a> (this may require you to repair Office afterwards).</p>
    <hr />
            <br />

            <h2><a name="Source"></a>Source Code</h2>
            <br />
            <h3><a name="srcAvrf"></a>SppExtComObjHookAvrf</h3>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1508167/">https://forums.mydigitallife.net/posts/1508167/</a>
      <br />
      <a href="https://0x0.st/zPKX.7z">https://0x0.st/zPKX.7z</a>
    </p>
    <h4 id="visual-studio">Visual Studio:</h4>
    <p>launch shortcut Developer Command Prompt for VS 2017 (or 2019)<br />
    execute:<br />
    <code>MSBuild SppExtComObjHook.sln /p:configuration="Release" /p:platform="Win32"</code><br />
    <code>MSBuild SppExtComObjHook.sln /p:configuration="Release" /p:platform="x64"</code></p>
    <h4 id="mingw-gcc">MinGW GCC:</h4>
    <p>download mingw-w64<br />
    <a href="https://sourceforge.net/projects/mingw-w64/files/i686-8.1.0-release-win32-sjlj-rt_v6-rev0.7z">Windows x86</a><br />
    <a href="https://sourceforge.net/projects/mingw-w64/files/x86_64-8.1.0-release-win32-sjlj-rt_v6-rev0.7z">Windows x64</a><br />
    both can compile 32-bit &amp; 64-bit binaries<br />
    extract and place SppExtComObjHook folder inside mingw32 or mingw64 folder<br />
    run <code>_compile.cmd</code></p>
            <br />

            <h3><a name="srcDebg"></a>SppExtComObjPatcher</h3>
    <h4 id="visual-studio-1">Visual Studio:</h4>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1457558/">https://forums.mydigitallife.net/posts/1457558/</a>
      <br />
      <a href="https://0x0.st/zHAP.7z">https://0x0.st/zHAP.7z</a>
    </p>
    <h4 id="mingw-gcc-1">MinGW GCC:</h4>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1462101/">https://forums.mydigitallife.net/posts/1462101/</a>
      <br />
      <a href="https://0x0.st/zHAK.7z">https://0x0.st/zHAK.7z</a>
    </p>
    <hr />
            <br />

            <h1><a name="Credits"></a>Credits</h1>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1508167/">namazso</a> - SppExtComObjHook, IFEO AVrf custom provider.<br />
      <a href="https://forums.mydigitallife.net/posts/862774">qad</a> - SppExtComObjPatcher, IFEO Debugger.<br />
      <a href="https://forums.mydigitallife.net/posts/1448556/">Mouri_Naruto</a> - SppExtComObjPatcher-DLL<br />
      <a href="https://forums.mydigitallife.net/posts/1462101/">os51</a> - SppExtComObjPatcher ported to MinGW GCC, Retail/MAK checks examples.<br />
      <a href="https://forums.mydigitallife.net/posts/309737/">MasterDisaster</a> - Original script, WMI methods.<br />
      <a href="https://forums.mydigitallife.net/posts/1296482/">qewpal</a> - KMS-VL-ALL script.<br />
      <a href="https://forums.mydigitallife.net/members/1108726/">Windows_Addict</a> - suggestions, ideas and documentation help.<br />
      <a href="https://forums.mydigitallife.net/members/846864/">NormieLyfe</a> - GVLK categorize, Office checks help.<br />
      <a href="https://forums.mydigitallife.net/members/2574/">mxman2k</a>, <a href="https://forums.mydigitallife.net/members/120394/">rpo</a>, <a href="https://forums.mydigitallife.net/members/presto1234.647219/">presto1234</a> - scripting suggestions.<br />
      <a href="https://forums.mydigitallife.net/members/80361/">Nucleus</a>, <a href="https://forums.mydigitallife.net/members/104688/">Enthousiast</a>, <a href="https://forums.mydigitallife.net/members/293479/">s1ave77</a>, <a href="https://forums.mydigitallife.net/members/325887/">l33tisw00t</a>, <a href="https://forums.mydigitallife.net/members/77147/">LostED</a>, <a href="https://forums.mydigitallife.net/members/1023044/">Sajjo</a> and MDL Community for interest, feedback and assistance.</p>
    <p>
      <a href="https://forums.mydigitallife.net/posts/1343297/">abbodi1406</a> - KMS_VL_ALL author</p>

            <h2><a name="acknow"></a>Acknowledgements</h2>
    <p>
      <a href="https://forums.mydigitallife.net/forums/51/">MDL forums</a> - the home of the latest and current emulators.<br />
      <a href="https://forums.mydigitallife.net/posts/838505">mikmik38</a> - first reversed source of KMSv5 and KMSv6.<br />
      <a href="https://forums.mydigitallife.net/threads/41010/">CODYQX4</a> - easy to use KMSEmulator source.<br />
      <a href="https://forums.mydigitallife.net/threads/50234/">Hotbird64</a> - the resourceful vlmcsd tool, and KMSEmulator source development.<br />
      <a href="https://forums.mydigitallife.net/threads/50949/">cynecx</a> - SECO Injector bypass, SppExtComObj KMS functions.<br />
      <a href="https://forums.mydigitallife.net/posts/856978">deagles</a> - SppExtComObjHook Injector.<br />
      <a href="https://forums.mydigitallife.net/posts/839363">deagles</a> - KMSServerService.<br />
      <a href="https://forums.mydigitallife.net/posts/1475544/">ColdZero</a> - CZ VM System.<br />
      <a href="https://forums.mydigitallife.net/posts/1476097/">ColdZero</a> - KMS ePID Generator.<br />
      <a href="https://forums.mydigitallife.net/posts/838023">kelorgo</a>, <a href="http://forums.mydigitallife.net/posts/838114">bedrock</a> - TAP adapter TunMirror bypass.<br />
      <a href="https://forums.mydigitallife.net/posts/1259604/">mishamosherg</a> - WinDivert FakeClient bypass.<br />
      <a href="https://forums.mydigitallife.net/posts/860489">Duser</a> - KMS Emulator fork.<br />
      ZWT, nosferati87, crony12, FreeStyler, Phazor - KMS Emulator development.</p>
        </div>
    </main>

    <nav id="nav">
        <div class="innertube">
            <h1>Index</h1>
            <a href="#Overview">Overview</a><br />
            <a href="#How">How it works?</a><br />
            <a href="#Supported">Supported Products</a><br />
            <a href="#OfficeR2V">Office Retail to Volume</a><br />
            <a href="#Using">How To Use</a><br /><br />
            <a href="#Modes">Activation Modes</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ModesAut">Auto Renewal</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ModesMan">Manual</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#ModesExt">External</a><br /><br />
            <a href="#Options">Additional Options</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#Unattend">Unattended</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#OptAct">Activation Choice</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#OptW10">KMS38 Win10</a><br />
            &nbsp;&nbsp;&nbsp;<a href="#OptKMS">KMS Options</a><br /><br />
            <a href="#Check">Check Activation Status</a><br />
            <a href="#Setup">Setup Preactivate</a><br />
            <a href="#Debug">Troubleshooting</a><br /><br />
            <a href="#Source">Source Code</a><br />
            &nbsp;&nbsp;<a href="#srcAvrf">SppExtComObjHookAvrf</a><br />
            &nbsp;&nbsp;<a href="#srcDebg">SppExtComObjPatcher</a><br /><br />
            <a href="#Credits">Credits</a><br />
        </div>
    </nav>
  </body>
</html>
:readme:

:DoDebug
set _dDbg=No
setlocal
call "%~f0" !_para!
endlocal
set _dDbg=Yes
echo.
echo Done.
echo Press any key to continue...
pause >nul
goto :MainMenu

:E_Admin
echo %_err%
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit.
if %_Debug% EQU 1 goto :eof
if %Unattend% EQU 1 goto :eof
pause >nul
goto :eof

:E_PS
echo %_err%
echo Windows PowerShell is required for this script to work.
echo.
echo Press any key to exit.
if %_Debug% EQU 1 goto :eof
if %Unattend% EQU 1 goto :eof
pause >nul
goto :eof

:E_DLL
echo %_err%
echo "%SystemRoot%\System32\SppExtComObjHook.dll" creation failed.
echo Verify that Antivirus protection is OFF or the file path is excluded.
echo.
echo Turn External option ON to activate via external KMS Server.
goto :TheEnd

:UnsupportedVersion
echo %_err%
echo Unsupported OS version Detected.
echo Project is supported only for Windows 7/8/8.1/10 and their Server equivalent.
:TheEnd
if exist "!_temp!\ReadMe.html" del /f /q "!_temp!\ReadMe.html"
echo.
if %Unattend% EQU 0 echo Press any key to exit.
%_Pause%
goto :eof

----- Begin wsf script --->
<package>
   <job id="ELAV">
       <script language="VBScript">
           Set strArg=WScript.Arguments.Named
           If Not strArg.Exists("File") Then
               Wscript.Echo "Switch /File:<File> is missing."
               WScript.Quit 1
           End If
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