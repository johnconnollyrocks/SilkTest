[ ] use "centrixs.inc"  
[ ] use "Registry.inc"
[ ] use "msw32.inc" 
[ ] // APPSTATES
[+] appstate AdobeReader ()   
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[ ] AdobeReader.sTag = "Adobe Reader*"
		[+] if !(AdobeReader.Exists ())
			[ ] RunStartMenu("acrord32")
			[ ] // STRING sCmdLine = '"C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"'
			[ ] // STRING sWorkingDir = "c:\temp"
			[ ] // AdobeReader.Start(sCmdLine,sWorkingDir,"",20)
			[ ] AdobeReader.SetActive()
			[ ] // Print(sCmdLine)
			[ ] // Print(sWorkingDir)
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[+] if AdobeReader.Exists()
			[ ] AdobeReader.SetActive()
			[ ] AdobeReader.PressKeys("<alt>")
			[ ] AdobeReader.TypeKeys("<f>")
			[ ] AdobeReader.ReleaseKeys("<alt>")
			[ ] AdobeReader.TypeKeys("<x>")
	[ ] 
[+] appstate WinZip() 
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[+] if !(WinZip.Exists ())
			[ ] WinZip.Start("C:\Program Files\WinZip\WINZIP32.exe")
			[ ] WinZip.SetActive()
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[+] if WinZip.Exists ()
			[ ] WinZip.File.Exit.Pick()
[+] appstate SAV()
	[ ] STRING sCmdLine = '"C:\Program Files\SAV\VPC32.exe"'
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[+] if !SymantecAntiVirus.Exists()
			[ ] RunStartMenu(sCmdLine)
			[ ] sleep(2)
			[+] if SymantecAntiVirusLogin.Exists()
				[ ] SymantecAntiVirusLogin.SetActive()
				[ ] SymantecAntiVirusLogin.Username.TypeKeys("admin")
				[ ] SymantecAntiVirusLogin.Password.TypeKeys("symantec")
				[ ] SymantecAntiVirusLogin.OK.Click()
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[ ] SymantecAntiVirus.Exit()
		[ ] 
[+] appstate Outlook() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sCmdLine = sGlobalReadPath+"Outlook.bat"
	[ ] STRING sWorkingDir = sGlobalReadPath
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[+] if !MicrosoftOutlook.Exists()
			[ ] RunStartMenu("outlook")
			[+] if Reminder.Exists()
				[ ] Reminder.DismissAll.Click()
			[ ] // SYS_Execute(sCmdLine,sReturn)
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[ ] // Click Send/Recei ve Button
		[ ] sleep(3)
		[ ] MicrosoftOutlook.TypeKeys("<F9>")
		[ ] sleep(2)
		[ ] // MicrosoftOutlook.MsoDockTop.Standard.Click (1, 397, 11)
		[+] while OutlookSendReceiveProgress.Exists()
			[ ] sleep(1)
		[ ] MicrosoftOutlook.TypeKeys("<esc>")
		[ ] sleep(3)
		[ ] MicrosoftOutlook.Close()
		[ ] sleep(3)
[+] appstate Word() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sWorkingDir = sGlobalReadPath
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[ ] MicrosoftWord.sTag = "*Microsoft Word"
		[+] if !MicrosoftWord.Exists()
			[ ] RunStartMenu("winword")
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[ ] MicrosoftWord.TypeKeys("<esc>")
		[ ] sleep(3)
		[ ] MicrosoftWord.Close()
		[+] if WordSaveChanges.Exists()
			[ ] WordSaveChanges.No.Click()
		[ ] 
[+] appstate Excel() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sWorkingDir = sGlobalReadPath
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[ ] MicrosoftExcel.sTag = "Microsoft Excel*"
		[+] if !MicrosoftExcel.Exists()
			[ ] RunStartMenu("excel")
			[ ] // SYS_Execute(sCmdLine,sReturn)
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[ ] MicrosoftExcel.TypeKeys("<esc>")
		[ ] sleep(3)
		[ ] MicrosoftExcel.Close()
	[ ] 
[+] appstate PowerPoint() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sWorkingDir = sGlobalReadPath
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[ ] MicrosoftPowerPoint.sTag = "Microsoft PowerPoint*"
		[+] if !MicrosoftPowerPoint.Exists()
			[ ] RunStartMenu("powerpnt")
			[ ] // SYS_Execute(sCmdLine,sReturn)
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[ ] MicrosoftPowerPoint.TypeKeys("<esc>")
		[ ] sleep(3)
		[ ] MicrosoftPowerPoint.Close()
	[ ] 
[+] appstate Access() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sWorkingDir = sGlobalReadPath
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[ ] MicrosoftAccess.sTag = "Microsoft Access*"
		[+] if !MicrosoftAccess.Exists()
			[ ] RunStartMenu("msaccess")
			[ ] // SYS_Execute(sCmdLine,sReturn)
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if GetTestCaseErrorCount( )== 0 
			[ ] addTestsPassedCount(1)
		[+] else 
			[ ] addTestsFailedCount(GetTestCaseErrorCount())
		[ ] MicrosoftAccess.Close()
	[ ] 
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] //
[ ] //                  UTILITY  TESTCASES
[ ] //
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] // TEST COUNT FUNCTIONS IN
[ ] // SILKTESTCENTRIXS.INI file
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[+] testcase initTestsPassedCount() appstate none
	[ ] STRING sSection = "TESTCOUNTERS"
	[ ] STRING sName = "TestsPassedCount"
	[ ] INT iTestsPassedCount = 0
	[ ] STRING sTestsPassedCount 
	[ ] HINIFILE hIniFile
	[ ] STRING sFile = "{sGlobalWritePath}SILKTESTCENTRIXS.INI"
	[ ] hIniFile = SYS_IniFileOpen (sFile) // Open the file
	[ ] SYS_IniFileSetValue (hIniFile, sSection, sName, [STRING]iTestsPassedCount)
	[ ] SYS_IniFileClose(hIniFile)
	[ ] // Print("iTestsPassedCount", iTestsPassedCount)
	[ ] 
[+] addTestsPassedCount(INT iTestsToAdd)
	[ ] STRING sSection = "TESTCOUNTERS"
	[ ] STRING sName = "TestsPassedCount"
	[ ] INT iTestsPassedCount = 0
	[ ] STRING sTestsPassedCount 
	[ ] HINIFILE hIniFile
	[ ] STRING sFile = "{sGlobalWritePath}\SILKTESTCENTRIXS.INI"
	[ ] hIniFile = SYS_IniFileOpen (sFile) // Open the file
	[ ] sTestsPassedCount = SYS_IniFileGetValue(hIniFile, sSection, sName)
	[ ] iTestsPassedCount = val(sTestsPassedCount) + iTestsToAdd
	[ ] SYS_IniFileSetValue (hIniFile, sSection, sName, [STRING]iTestsPassedCount)
	[ ] SYS_IniFileClose(hIniFile)
	[ ] // Print("iTestsPassedCount", iTestsPassedCount)
	[ ] 
[+] INT getTestsPassed()
	[ ] STRING sSection = "TESTCOUNTERS"
	[ ] STRING sName = "TestsPassedCount"
	[ ] INT iTestsPassedCount = 0
	[ ] STRING sTestsPassedCount 
	[ ] HINIFILE hIniFile
	[ ] STRING sFile = "{sGlobalWritePath}\SILKTESTCENTRIXS.INI"
	[ ] hIniFile = SYS_IniFileOpen (sFile) // Open the file
	[ ] sTestsPassedCount = SYS_IniFileGetValue(hIniFile, sSection, sName)
	[ ] SYS_IniFileClose(hIniFile)
	[ ] // Print("sTestsPassedCount", sTestsPassedCount)
	[ ] return val(sTestsPassedCount)
[ ] 
[+] testcase initTestsFailedCount() appstate none
	[ ] STRING sSection = "TESTCOUNTERS"
	[ ] STRING sName = "TestsFailedCount"
	[ ] INT iTestsFailedCount = 0
	[ ] STRING sTestsFailedCount 
	[ ] HINIFILE hIniFile
	[ ] STRING sFile = "{sGlobalWritePath}\SILKTESTCENTRIXS.INI"
	[ ] hIniFile = SYS_IniFileOpen (sFile) // Open the file
	[ ] SYS_IniFileSetValue (hIniFile, sSection, sName, [STRING]iTestsFailedCount)
	[ ] SYS_IniFileClose(hIniFile)
	[ ] // Print("iTestsFailedCount", iTestsFailedCount)
	[ ] 
[+] addTestsFailedCount(INT iTestsToAdd )
	[ ] STRING sSection = "TESTCOUNTERS"
	[ ] STRING sName = "TestsFailedCount"
	[ ] INT iTestsFailedCount = 0
	[ ] STRING sTestsFailedCount 
	[ ] HINIFILE hIniFile
	[ ] STRING sFile = "{sGlobalWritePath}\SILKTESTCENTRIXS.INI"
	[ ] hIniFile = SYS_IniFileOpen (sFile) // Open the file
	[ ] sTestsFailedCount = SYS_IniFileGetValue(hIniFile, sSection, sName)
	[ ] iTestsFailedCount = val(sTestsFailedCount) + iTestsToAdd
	[ ] SYS_IniFileSetValue (hIniFile, sSection, sName, [STRING]iTestsFailedCount)
	[ ] SYS_IniFileClose(hIniFile)
	[ ] // Print("iTestsFailedCount", iTestsFailedCount)
	[ ] 
[+] INT getTestsFailed()
	[ ] STRING sSection = "TESTCOUNTERS"
	[ ] STRING sName = "TestsFailedCount"
	[ ] INT iTestsFailedCount = 0
	[ ] STRING sTestsFailedCount 
	[ ] HINIFILE hIniFile
	[ ] STRING sFile = "{sGlobalWritePath}\SILKTESTCENTRIXS.INI"
	[ ] hIniFile = SYS_IniFileOpen (sFile) // Open the file
	[ ] sTestsFailedCount = SYS_IniFileGetValue(hIniFile, sSection, sName)
	[ ] SYS_IniFileClose(hIniFile)
	[ ] // Print("sTestsFailedCount", sTestsFailedCount)
	[ ] return val(sTestsFailedCount)
[ ] 
[+] testcase PrintTestBanner(STRING sBanner) appstate none
	[ ] Print("")
	[ ] Print("**********          "+sBanner+"          **********")
	[ ] Print("")
[+] testcase PrintTestsPassed() appstate none
	[ ] Print("")
	[ ] Print("**** TOTAL TESTS PASSED = "+"{GetTestsPassed()}")
	[ ] // Print("**** GetTestsPassedCount = "+"{GetTestsPassedCount()}")
	[ ] // Print("**** GetTestsPassedCount = "+ "{GetTestsPassedCount()}")
	[ ] Print("")
[+] testcase PrintTestsFailed() appstate none
	[ ] Print("")
	[ ] Print("**** TOTAL TESTS FAILED = "+"{GetTestsFailed()}")
	[ ] // Print("**** GetTestsFailedCount = "+"{GetTestsFailedCount()}")
	[ ] // Print("**** GetTestsPassedCount = "+ "{GetTestsFailedCount()}")
	[ ] Print("")
[+] testcase PrintFileOpenTestsFailed() appstate none
	[+] if lsFailedFileOpenTests != NULL
		[ ] Print("The Following Files failed to open")
		[ ] ListPrint(lsFailedFileOpenTests)
[+] testcase testcaseNotImplemented(STRING sTestCaseID) appstate none
	[ ] Print("Test Case {sTestCaseID} not implemented")
[+] CreateProcess(STRING sUsername, STRING sDomain, STRING sPassword, STRING sApplication)
	[ ] INT iProcessReturn
	[ ] INT iLogonReturn
	[ ] STRING sProcessInfo
	[ ] STRING iPrimaryToken
	[ ] 
	[ ] iLogonReturn = LogonUserW(sUsername, sDomain, sPassword, logon32_logon_network_cleartext, logon32_provider_default, iPrimaryToken)
	[ ] // Print("Primary Token is ", iPrimaryToken) 
	[ ] // Print("Return from LogonUser is ", iLogonReturn) 
	[ ] iProcessReturn = CreateProcessWithLogonW(sUsername, sDomain, sPassword, 0x00000001,sApplication, NULL, 0x00080000, NULL, NULL, NULL,sProcessInfo )
	[ ] // Print("Return from CreateProcessWithLogonW is ", iProcessReturn) 
[+] createProcessLogon(STRING sNewUser, STRING sServerAddress)
	[ ] LIST OF STRING lsReturn
	[ ] STRING sReturn
	[ ] STRING sCommandPath = ""
	[ ] // STRING sPreviousUsername
	[ ] // sPreviousUsername = SYS_GetEnv ("USERNAME")
	[ ] // Print("sPreviousUsername = "+sPreviousUsername)
	[ ] // Print(Substr(gethostname(),8,4))
	[+] if Substr(gethostname(),8,4) == "CS01" 
		[ ] CMD.sTag = "[DialogBox]C:\WINDOWS\system32\CMD.exe*"
		[ ] sCommandPath = "C:\WINDOWS\system32\CMD.exe"
	[+] if Substr(gethostname(),8,4) == "DC01" || Substr(gethostname(),8,4) == "DC02"
		[ ] CMD.sTag = "[DialogBox]C:\WINNT\system32\CMD.exe*"
		[ ] sCommandPath = "C:\WINNT\system32\CMD.exe"
	[ ] CMD.Start(sCommandPath)
	[ ] CMD.SetActive()
	[ ] CMD.TypeKeys("net use H: /delete <enter>")
	[ ] CMD.TypeKeys("Y")
	[ ] CMD.TypeKeys("net use S: /delete <enter>")
	[ ] CMD.TypeKeys("net use Z: /delete <enter>")
	[ ] // CMD.TypeKeys("taskkill /f /im explorer.exe<enter>")
	[ ] CMD.Close()
	[ ] //   /u BLK0520\silka /p C3ntr!X$2k3
	[ ] // SYS_EXECUTE("taskkill /f  /s {sServerAddress} /u BLK0520\silka /p C3ntr!X$2k3 /im explorer.exe", lsReturn)
	[ ] SYS_EXECUTE("taskkill /f  /im explorer.exe", lsReturn)
	[+] for each sReturn in lsReturn
		[ ] // Print(sReturn)
	[+] if Substr(gethostname(),8,4) == "CS01" 
		[ ] CreateProcess(sNewUser, "blk0520.navy.usa.cfe.cmil.mil", "C3ntr!X$2k3", "C:\windows\explorer.exe")
	[+] if Substr(gethostname(),8,4) == "DC01" || Substr(gethostname(),8,4) == "DC02"
		[ ] CreateProcess(sNewUser, "blk0520.navy.usa.cfe.cmil.mil", "C3ntr!X$2k3", "C:\WINNT\explorer.exe")
	[+] if ExplorerPathNotFound.Exists()
		[ ] ExplorerPathNotFound.SetActive()
		[ ] ExplorerPathNotFound.OK.Click()
	[+] if Notepad.Exists()
		[ ] Notepad.SetActive()
		[ ] Notepad.File.Exit.Pick()
	[ ] // SYS_EXECUTE("taskkill /f /im C:\Program Files\Borland\SilkTest\Agent.exe")
	[ ] // CMD.Start("C:\WINDOWS\system32\CMD.exe")
	[ ] RunStartMenu("CMD")
	[ ] CMD.TypeKeys("net use H: \\BLK0520DC01\composeusers<enter>")
	[ ] CMD.TypeKeys("net use S: \\BLK0520DC01\sharedrive<enter>")
	[ ] CMD.TypeKeys("net use Z: \\BLK0520DC01\composeusers\{sNewUser}<enter>")
	[+] if ExplorerPathNotFound.Exists()
		[ ] ExplorerPathNotFound.SetActive()
		[ ] ExplorerPathNotFound.OK.Click()
	[+] if Notepad.Exists()
		[ ] Notepad.SetActive()
		[ ] Notepad.File.Exit.Pick()
	[ ] // CMD.TypeKeys("net use L: \\BLK0520DC01\sysvol<enter>")
	[ ] // CMD.TypeKeys("L:<enter>")
	[ ] // CMD.TypeKeys("cd blk0520.navy.usa.cfe.cmil.mil\scripts<enter>")
	[ ] // CMD.TypeKeys("login.exe<enter>")
	[ ] CMD.Close()
	[ ] // MapDC01ComposeUsersDir("BLK0520")
	[ ] 
[+] BOOLEAN StartProgramAsync (STRING sProgramPath)  
	[+] if (ShellExecuteW (Desktop.GetHandle (), "open", sProgramPath, " ", "", 5) <= 32)		
		[ ] return FALSE	
	[+] else		
		[+] return TRUE
			[ ] 
	[ ] 
[+] testcase testCMD() appstate none
	[ ] CMD.TypeKeys("net use H: \\BLK0520DC01\composeusers\<Enter>")
[+] MapDC01ComposeUsersDir(STRING sDomain) 
	[ ] Taskbar.SetActive ()
	[ ] Taskbar.Start.Click ()
	[ ] Taskbar.TypeKeys("R")
	[ ] Run.SetActive ()
	[ ] Run.Open.SetText("explorer")
	[ ] Run.OK.Click()
	[+] // if ExplorerPathNotFound.Exists()
		[ ] // ExplorerPathNotFound.OK.Click()
	[ ] // MyDocuments.WorkerW1.ReBarWindow321.ToolBar1.Favorites
	[ ] MyDocuments.SetActive ()
	[ ] MyDocuments.WorkerW1.ReBarWindow321.ToolBar1.TypeKeys("<Alt-T>")
	[ ] MyDocuments.WorkerW1.ReBarWindow321.ToolBar1.TypeKeys("N")
	[ ] MapNetworkDrive.SetActive ()
	[ ] MapNetworkDrive.MapNetworkDrive.Drive.TypeKeys("H")
	[ ] MapNetworkDrive.MapNetworkDrive.Folder.TypeKeys("\{sDomain}dc01\composeusers\")
	[ ] // MapNetworkDrive.MapNetworkDrive.WindowsCanHelpYouConnectT.Click (1, 122, 36)
	[ ] MapNetworkDrive.Finish.Click ()
[ ] 
[+] ConnectCS01() 
	[ ] Agent.SetOption(OPT_SET_TARGET_MACHINE, CS01)
	[ ] Print("")
	[ ] Print("**********          "+"STARTING TESTS ON CS01 {CS01}"+"          **********")
	[ ] Print("")
[+] ConnectDC01() 
	[ ] Agent.SetOption(OPT_SET_TARGET_MACHINE, DC01)
	[ ] Print("")
	[ ] Print("**********          "+"STARTING TESTS ON DC01 {DC01}"+"          **********")
	[ ] Print("")
[+] ConnectDC02()
	[ ] Agent.SetOption(OPT_SET_TARGET_MACHINE, DC02)
	[ ] Print("")
	[ ] Print("**********          "+"STARTING TESTS ON DC02 {DC02}"+"          **********")
	[ ] Print("")
[ ] 
[+] DisconnectCS01() 
	[ ] // Disconnect(CS01)
	[ ] Print("")
	[ ] Print("**********          "+"ENDING TESTS ON CS01 {CS01}"+"          **********")
	[ ] Print("")
[+] DisconnectDC01() 
	[ ] // Disconnect(DC01)
	[ ] Print("")
	[ ] Print("**********          "+"ENDING TESTS ON DC01 {DC01}"+"          **********")
	[ ] Print("")
[+] DisconnectDC02() 
	[ ] // Disconnect(DC02)
	[ ] Print("")
	[ ] Print("**********          "+"ENDING TESTS ON DC02 {DC02}"+"          **********")
	[ ] Print("")
[ ] 
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] //
[ ] //                  OFFICIAL SIT TESTCASES
[ ] //
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] 
[ ] // 2.2.1.1 Outlook Mail CC - A -> B
[+] testcase officeOutlook2_2_1_1() appstate Outlook // Mail Message - User A signon w/o attachment
	[ ] createMail( "silkc", "silkb", "Test Case 2.2.1.1","")
	[ ] sendMail()
[+] testcase switchToUser(STRING sNewUser) appstate none
	[ ] createProcessLogon(sNewUser,CS01)
[+] testcase officeOutlook2_2_1_1Verify() appstate Outlook // Mail Message - User B signon w/o attachment
	[+] if findInboxMailItemID("Test Case 2.2.1.1")==0 
		[ ] LogError("Test Case 2.2.1.1 was not found in the subject line")
		[ ] // verifySubject("Test Case 2.2.1.1", 7)
[ ] // 2.2.2.1 Outlook Mail CC - B -> A
[+] testcase officeOutlook2_2_2_1() appstate Outlook // Mail Message -  User B signon w/o attachment
	[ ] createMail( "silkz", "silka", "Test Case 2.2.2.1","")
	[ ] sendMail()
[+] testcase officeOutlook2_2_2_1Verify() appstate Outlook // Mail Message - User A signon w/o attachment
	[+] if findInboxMailItemID("Test Case 2.2.2.1")==0 
		[ ] LogError("Test Case 2.2.2.1 was not found in the subject line")
	[ ] // verifySubject("Test Case 2.2.2.1", 6)
[ ] // 2.2.3.1 Outlook Attachments
[+] testcase officeOutlook2_2_3_1_2() appstate Outlook // Mail Message -  User A signon w/text attachment w/graphics(602K)
	[ ] createMail( "silkb", "", "Test Case 2.2.3.1.2","CENTRIXS-M_AUTOMATED_SIT_BLOCK_0.doc")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3_1_3() appstate Outlook // Mail Message -  User A signon w/RTF text attachment 
	[ ] createMail( "silkb", "", "Test Case 2.2.3.1.3","Security-HOWTO.rtf")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3_1_4() appstate Outlook // Mail Message -  User A signon w/text attachment w/COM object (excel)(423K)
	[ ] createMail( "silkb", "", "Test Case 2.2.3.1.4","CENTRIXS-M_AUTOMATED_SIT_BLOCK_0.doc")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3_1_5() appstate Outlook // Mail Message -  User A signon w/ppt attachment (9 Meg)
	[ ] createMail( "silkb", "", "Test Case 2.2.3.1.5","bighistory.ppt")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3_1_6() appstate Outlook // Mail Message -  User A signon w/ppt attachment (2 Meg)
	[ ] createMail( "silkb", "", "Test Case 2.2.3.1.6","history.ppt")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3_1_7() appstate Outlook // Mail Message -  User A signon w/access attachment (120 K)
	[ ] createMail( "silkb", "", "Test Case 2.2.3.3.1","db1.db")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3_1_8() appstate Outlook // Mail Message -  User A signon w/access attachment (430 K)
	[ ] createMail( "silkb", "", "Test Case 2.2.3.3.2","inventory_demo.db")
	[ ] sendMail()
[ ] 
[+] testcase officeOutlook2_2_1() appstate Outlook // Mail Message - User A signon w/o attachment
	[ ] createMail( "silkb", "", "Test Case 2.2.1","")
	[ ] sendMail()
[+] testcase officeOutlook2_2_1Verify() appstate Outlook // Mail Message - User B signon w/o attachment
	[+] if findInboxMailItemID("Test Case 2.2.1")==0 
		[ ] LogError("Test Case 2.2.1 was not found in the subject line")
[+] testcase officeOutlook2_2_2() appstate Outlook // Mail Message -  User B signon w/o attachment
	[ ] createMail( "silka", "", "Test Case 2.2.2","")
	[ ] sendMail()
[+] testcase officeOutlook2_2_2Verify() appstate Outlook // Mail Message - User A signon w/o attachment
	[+] if findInboxMailItemID("Test Case 2.2.2")==0 
		[ ] LogError("Test Case 2.2.2 was not found in the subject line")
[+] testcase officeOutlook2_2_3() appstate Outlook // Mail Message -  User A signon w/text attachment w/date(1K)
	[ ] createMail( "silkb", "", "Test Case 2.2.3","doc1.doc")
	[ ] sendMail()
[+] testcase officeOutlook2_2_3Verify() appstate Outlook // Mail Message -  User B signon w/attachment
	[+] if findInboxMailItemID("Test Case 2.2.3")==0 
		[ ] LogError("Test Case 2.2.3 was not found in the subject line")
[+] testcase officeOutlook2_2_4() appstate Outlook // Mail Message -  User B signon w/attachment
	[ ] createMail( "silka", "", "Test Case 2.2.4","doc1.doc")
	[ ] sendMail()
[+] testcase officeOutlook2_2_4Verify() appstate Outlook // Mail Message -  User A signon w/attachment
	[+] if findInboxMailItemID("Test Case 2.2.4")==0 
		[ ] LogError("Test Case 2.2.4 was not found in the subject line")
[+] testcase officeOutlook2_2_5() appstate Outlook // Mail Message -  User A signon
	[ ] createMail( "silkb;silkc", "", "Test Case 2.2.5","doc1.doc")
	[ ] sendMail()
[+] testcase officeOutlook2_2_5VerifyB() appstate Outlook // Mail Message -  User B signon
	[+] if findInboxMailItemID("Test Case 2.2.5")==0 
		[ ] LogError("Test Case 2.2.5 was not found in the subject line")
[+] testcase officeOutlook2_2_5VerifyC() appstate Outlook // Mail Message -  User C signon
	[+] if findInboxMailItemID("Test Case 2.2.5")==0 
		[ ] LogError("Test Case 2.2.5 was not found in the subject line")
[+] testcase officeOutlook2_2_6() appstate Outlook // Mail Message -  User C signon - Reply to User A
	[ ] replymail(findInboxMailItemID("Test Case 2.2.5"))
[+] testcase officeOutlook2_2_6Verify() appstate Outlook // Mail Message -  User B signon
	[+] if findInboxMailItemID("RE: Test Case 2.2.5")==0 
		[ ] LogError("RE: Test Case 2.2.5 was not found in the subject line")
[+] testcase officeOutlook2_2_7() appstate Outlook // Mail Message -  User A signon - Reply
	[ ] replyAllmail(findInboxMailItemID("Test Case 2.2.5"))
[-] testcase officeOutlook2_2_7VerifyA() appstate Outlook // Mail Message -  User A signon
	[+] if findInboxMailItemID("RE: Test Case 2.2.5")==0 
		[ ] LogError("RE: Test Case 2.2.5 was not found in the subject line")
[-] testcase officeOutlook2_2_7VerifyC() appstate Outlook // Mail Message -  User C signon
	[+] if findInboxMailItemID("Test Case 2.2.5")==0 
		[ ] LogError("Test Case 2.2.5 was not found in the subject line")
[+] testcase officeOutlook2_2_8() appstate Outlook // Appointment -  User A signon - Invites User B
	[ ] createAppointment("silkb","Test Case 2.2.8","","Body","attach1.txt")
[-] testcase officeOutlook2_2_8Verify() appstate Outlook // Appointment -  User A signon - Invites User B
	[+] if findInboxMeetingItemID("Test Case 2.2.8")==0 
		[ ] LogError("Test Case 2.2.8 was not found in the subject line")
	[ ] // verifySubjectMeeting("Test Case 2.2.8", 1)
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Word Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] // testcase officeWord2_2_9() appstate none // User A signon
	[ ] // STRING sFilename = "doc1.DOC"
	[ ] // STRING sDate
	[ ] // STRING sOpenFileName = sFilename
	[ ] // STRING sSaveAsFileName = sFilename
	[ ] // openword()
	[ ] // openworddoc(sGlobalReadPath+sOpenFileName)
	[ ] // MicrosoftWord.sTag = "*Microsoft Word"
	[ ] // MicrosoftWord.TypeKeys("<Ctrl-A>")
	[ ] // sDate = DateStr ()
	[ ] // MicrosoftWord.TypeKeys(sDate)
	[ ] // // Select Save As using keystrokes
	[ ] // MicrosoftWord.TypeKeys("<alt-f>")
	[ ] // MicrosoftWord.TypeKeys("<a>")
	[ ] // WordSaveAs.Filename.TypeKeys(sGlobalWritePath+sSaveAsFileName)
	[ ] // WordSaveAs.Filename.TypeKeys("<Enter>")
	[+] // if WordReplace.Exists()
		[ ] // WordReplace.TypeKeys("<Enter>")
	[ ] // closeword()
	[ ] // openOutlook()
	[ ] // createMail("silkb","","Test Case 2.2.9",sFilename)
	[ ] // sendMail()
	[ ] // // MicrosoftOutlook.MsoDockTop.Standard.Click (1, 397, 11)
	[ ] // sleep(15)
	[ ] // closeoutlook()
	[ ] // 
	[ ] // // Compare Files - We will verify that the file has CHANGED
	[+] // if SYS_CompareBinary (sGlobalReadPath+sFilename, sGlobalAdminWritePath+sFilename)
		[ ] // LogError("The files "+sGlobalReadPath+sFilename+" and "+sGlobalAdminWritePath+sFilename+" are the same")
	[ ] // 
[+] testcase officeWord2_2_9() appstate Word // User A signon
	[ ] STRING sFilename = "doc1.DOC"
	[ ] STRING sDate
	[ ] STRING sOpenFileName = sFilename
	[ ] STRING sSaveAsFileName = sFilename
	[ ] openworddoc(sGlobalReadPath+sOpenFileName)
	[ ] MicrosoftWord.sTag = "*Microsoft Word"
	[ ] MicrosoftWord.TypeKeys("<Ctrl-A>")
	[ ] sDate = DateStr ()
	[ ] MicrosoftWord.TypeKeys(sDate)
	[ ] // Select Save As using keystrokes
	[ ] MicrosoftWord.TypeKeys("<alt-f>")
	[ ] MicrosoftWord.TypeKeys("<a>")
	[ ] WordSaveAs.Filename.TypeKeys(sGlobalWritePath+sSaveAsFileName)
	[ ] WordSaveAs.Filename.TypeKeys("<Enter>")
	[+] if WordReplace.Exists()
		[ ] WordReplace.TypeKeys("<Enter>")
		[ ] 
[+] testcase officeOutlook2_2_9() appstate Outlook // User A signon
	[ ] STRING sFilename = "doc1.DOC"
	[ ] STRING sOpenFileName = sFilename
	[ ] STRING sSaveAsFileName = sFilename
	[ ] 
	[ ] createMail("silkb","","Test Case 2.2.9",sFilename)
	[ ] sendMail()
	[ ] // Compare Files - We will verify that the file has CHANGED
	[+] if SYS_CompareBinary (sGlobalReadPath+sFilename, sGlobalAdminWritePath+sFilename)
		[ ] LogError("The files "+sGlobalReadPath+sFilename+" and "+sGlobalAdminWritePath+sFilename+" are the same")
	[ ] 
[ ] 
[-] testcase officeWord2_2_9Verify() appstate Outlook // User B Signon
	[+] if findInboxMailItemID("Test Case 2.2.9")==0 
		[ ] LogError("Test Case 2.2.9 was not found in the subject line")
		[ ] 
[+] testcase officeWord2_2_10_1() appstate Word // User B signon - open and save as document to file share
	[ ] STRING sDate
	[ ] STRING sOpenFileName = "doc1.DOC"
	[ ] STRING sSaveAsFileName = "doc1.DOC"
	[ ] // openword()
	[ ] openworddoc(sGlobalReadPath+sOpenFileName)
	[ ] sleep(3)
	[ ] MicrosoftWord.SetActive()
	[ ] // Select Save As using keystrokes
	[ ] MicrosoftWord.TypeKeys("<alt-f>")
	[ ] MicrosoftWord.TypeKeys("A")
	[ ] WordSaveAs.Filename.TypeKeys (sGlobalWritePath+sSaveAsFileName)
	[ ] WordSaveAs.Filename.TypeKeys ("<Enter>")
	[+] if WordReplace.Exists()
		[ ] WordReplace.TypeKeys("<Enter>")
	[ ] // closeword()
	[ ] // Compare Files - We will verify that the file has NOT CHANGED
	[+] // if !SYS_CompareBinary (sGlobalReadPath+sOpenFileName,sGlobalAdminWritePath+sSaveAsFileName)
		[ ] // LogError("The files "+sGlobalReadPath+sOpenFileName+" and "+sGlobalAdminWritePath+sSaveAsFileName+" are not the same ") 
	[ ] 
[+] testcase officeWord2_2_10_2() appstate Word // User A signon - Word Spell Check
	[ ] STRING sDate
	[ ] STRING sOpenFileName = "CENTRIXS-M_AUTOMATED_SIT_BLOCK_0.doc"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"SILK_MODIFIED_CENTRIXS-M_AUTOMATED_SIT_BLOCK_0.doc" 
	[ ] // openword()
	[ ] openworddoc(sGlobalReadPath+sOpenFileName)
	[ ] MicrosoftWord.SetActive()
	[ ] MicrosoftWord.TypeKeys("<Home>")
	[ ] MicrosoftWord.TypeKeys("<F7>")
	[ ] Spelling.SetActive()
	[+] while !SpellingCheckComplete.Exists()
		[+] while SpellingConfirmProofingTools.Exists()
			[ ] SpellingConfirmProofingTools.OK.Click()
		[ ] Spelling.IgnoreAll.Click()
		[ ] sleep(1)
	[ ] SpellingCheckComplete.OK.Click()
	[ ] // closeword()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Power Point Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase officePowerPoint2_2_11_1() appstate PowerPoint  
	[ ] INTEGER i
	[ ] STRING sOpenFileName = sGlobalReadPath+"bighistory.ppt"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"bighistory.ppt" // "sGlobalWritePath"+"bighistory.ppt"
	[ ] // openPowerPoint()
	[ ] openPowerPointdoc(sOpenFileName)
	[ ] MicrosoftPowerPoint.SetActive ()
	[ ] MicrosoftPowerPoint.TypeKeys("<F5>")
	[ ] sleep(10)
	[ ] MicrosoftPowerPoint.SetActive ()
	[ ] MicrosoftPowerPoint.TypeKeys("<Esc>")
	[ ] // sleep(3)
	[+] // for i = 1 to 50
		[ ] // MicrosoftPowerPoint.TypeKeys("<Down>")
		[ ] // sleep(i)
	[ ] // closepowerpoint()
	[ ] 
[+] testcase officePowerPoint2_2_11_2() appstate PowerPoint 
	[ ] INTEGER i
	[ ] STRING sOpenFileName = sGlobalReadPath+"history.ppt"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"history.ppt" // "sGlobalWritePath"+"bighistory.ppt"
	[ ] // openPowerPoint()
	[ ] openPowerPointdoc(sOpenFileName)
	[ ] MicrosoftPowerPoint.SetActive ()
	[ ] MicrosoftPowerPoint.TypeKeys("<F5>")
	[ ] sleep(10)
	[ ] // MicrosoftPowerPoint.TypeKeys("<Esc>")
	[ ] // sleep(3)
	[+] // for i = 1 to 50
		[ ] // MicrosoftPowerPoint.TypeKeys("<Right>")
		[ ] // sleep(i)
	[ ] // closepowerpoint()
	[ ] 
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Excel Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase officeExcel2_2_12_1() appstate Excel  
	[ ] STRING sOpenFileName = sGlobalReadPath+"book1.xls"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"book1.xls"
	[ ] // openExcel()
	[ ] openExceldoc(sOpenFileName)
	[ ] MicrosoftExcel.SetActive ()
	[ ] sleep(1)
	[ ] // MicrosoftExcel.Capturebitmap("{sGlobalReadPath}"+"\Excelslideshowbitmap2_2_11.bmp")
	[ ] // closeExcel()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Access Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase officeAccess2_2_13_1() appstate Access  
	[ ] STRING sOpenFileName = sGlobalReadPath+"inventory_demo.mdb"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"inventory_demo.mdb"
	[ ] // openAccess()
	[ ] openAccessdoc(sOpenFileName)
	[ ] sleep(1)
	[ ] MicrosoftAccess.SetActive ()
	[ ] sleep(1)
	[ ] // MicrosoftAccess.Capturebitmap("{sGlobalReadPath}"+"\Accessslideshowbitmap2_2_11.bmp")
	[ ] // closeAccess()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Adobe Reader Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase AdobeReader2_2_14_1() appstate AdobeReader 
	[ ] STRING sOpenFileName = sGlobalReadPath+"instructions.pdf"
	[ ] STRING sSaveAsFileName = "sGlobalWritePath"+"instructions.pdf"
	[ ] openAdobeReader()
	[ ] openAdobeReaderdoc(sOpenFileName)
	[ ] sleep(1)
	[ ] // MicrosoftAccess.Capturebitmap("{sGlobalReadPath}"+"\Accessslideshowbitmap2_2_11.bmp")
	[ ] closeAdobeReader()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Winzip Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase WinZip2_2_15() appstate WinZip  
	[ ] // Note: Must have licensed copy (e.g. not and expired demo version)
	[ ] // to run this testcase
	[ ] LIST OF STRING lsFileList  = GetFilenamesFromExcel("WinZip")
	[ ] createWinZip(sGlobalWritePath+"SilkTest.zip", true)
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Symantec AntiVirus Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase SAV2_6_1() appstate SAV // User N/A
	[ ] STRING sOverallTime
	[ ] HTIMER TotalTimer
	[ ] TotalTimer = TimerCreate ("MyTimer")
	[ ] TimerStart (TotalTimer)
	[ ] SymantecAntiVirus.SetActive ()
	[ ] SymantecAntiVirus.Scan.FullScan.Pick ()
	[ ] SymantecAntiVirus.AfxMDIFrame701.AfxFrameOrView702.AfxOleControl701.FullScan.Scan.Click ()
	[+] while SAVFullScan.Text1.GetText() != "Completed"
		[ ] sleep(1)
	[ ] TimerStop (TotalTimer)
	[ ] sOverallTime = TimerStr (TotalTimer)
	[ ] Print ("Full Scan executed in {sOverallTime} seconds")
	[ ] TimerDestroy (TotalTimer)
	[ ] SAVFullScan.Close.Click()
	[+] SymantecAntiVirus.File.Exit.Pick()
		[ ] 
[+] testcase PrintCompleted() appstate none
	[ ] Print(SAVFullScan.Text1.GetText())
	[ ] 
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Print Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] testcase printOutlook2_7_1_1() appstate Outlook 
	[ ] MicrosoftOfficeOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
	[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
	[ ] Message.SetActive()
	[ ] Message.TypeKeys("Test Case 2.7.1.1 - Print from Outlook")
	[ ] Message.MsoDockTop.MenuBar.PressKeys("<Alt>")
	[ ] Message.MsoDockTop.MenuBar.TypeKeys("<F>")
	[ ] Message.MsoDockTop.MenuBar.ReleaseKeys("<Alt>")
	[ ] Message.MsoDockTop.MenuBar.TypeKeys("<P>")
	[ ] OutlookPrint.SetActive()
	[ ] OutlookPrint.TypeKeys("<Enter>")
	[ ] OpenMessage.PressKeys("<Alt>")
	[ ] OpenMessage.TypeKeys("<F4>")
	[ ] OpenMessage.ReleaseKeys("<Alt>")
	[ ] OutlookSaveChanges.No.Click()
	[ ] 
[+] testcase printWord2_7_1_2() appstate Word
	[ ] STRING sDate
	[ ] STRING sOpenFileName = sGlobalReadPath+"doc1.DOC"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"doc1.DOC"
	[ ] 
	[ ] MicrosoftWord.SetActive()
	[ ] MicrosoftWord.TypeKeys("Test Case 2.7.1.2 - Print from Word")
	[ ] MicrosoftWord.MsoDockTop.MenuBar.PressKeys("<Alt>")
	[ ] MicrosoftWord.MsoDockTop.MenuBar.TypeKeys("<F>")
	[ ] MicrosoftWord.MsoDockTop.MenuBar.ReleaseKeys("<Alt>")
	[ ] MicrosoftWord.MsoDockTop.MenuBar.TypeKeys("<P>")
	[ ] WordPrint.SetActive()
	[ ] WordPrint.TypeKeys("<Enter>")
	[ ] MicrosoftWord.PressKeys("<Alt>")
	[ ] MicrosoftWord.TypeKeys("<F4>")
	[ ] MicrosoftWord.ReleaseKeys("<Alt>")
	[+] if WordSaveChanges.Exists()
		[ ] WordSaveChanges.No.Click()
		[ ] 
[+] testcase printPowerPoint2_7_1_3() appstate PowerPoint  // User N/A
	[ ] STRING sOpenFileName = sGlobalReadPath+"bighistory.ppt"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"bighistory.ppt" // "sGlobalWritePath"+"bighistory.ppt"
	[ ] 
	[ ] MicrosoftPowerPoint.SetActive ()
	[ ] MicrosoftPowerPoint.TypeKeys("Test Case 2.7.1.3 - Print from Power Point")
	[ ] MicrosoftPowerPoint.MsoDockTop.MenuBar.PressKeys("<Alt>")
	[ ] MicrosoftPowerPoint.MsoDockTop.MenuBar.TypeKeys("<F>")
	[ ] MicrosoftPowerPoint.MsoDockTop.MenuBar.ReleaseKeys("<Alt>")
	[ ] MicrosoftPowerPoint.MsoDockTop.MenuBar.TypeKeys("<P>")
	[ ] PowerPointPrint.SetActive()
	[ ] PowerPointPrint.TypeKeys("<Enter>")
	[ ] MicrosoftPowerPoint.PressKeys("<Alt>")
	[ ] MicrosoftPowerPoint.TypeKeys("<F4>")
	[ ] MicrosoftPowerPoint.ReleaseKeys("<Alt>")
	[+] if PowerPointSaveChanges.Exists()
		[ ] PowerPointSaveChanges.No.Click()
	[ ] 
[+] testcase printExcel2_7_1_4() appstate Excel // User N/A
	[ ] STRING sOpenFileName = sGlobalReadPath+"book1.xls"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"book1.xls"
	[ ] 
	[ ] MicrosoftExcel.SetActive ()
	[ ] MicrosoftExcel.TypeKeys("Test Case 2.7.1.4 - Print From Excel<Enter>")
	[ ] MicrosoftExcel.EXCEL21.WorksheetMenuBar.PressKeys("<Alt>")
	[ ] MicrosoftExcel.EXCEL21.WorksheetMenuBar.TypeKeys("<F>")
	[ ] MicrosoftExcel.EXCEL21.WorksheetMenuBar.ReleaseKeys("<Alt>")
	[ ] MicrosoftExcel.EXCEL21.WorksheetMenuBar.TypeKeys("<P>")
	[ ] ExcelPrint.SetActive()
	[ ] ExcelPrint.TypeKeys("<Enter>")
	[ ] MicrosoftExcel.PressKeys("<Alt>")
	[ ] MicrosoftExcel.TypeKeys("<F4>")
	[ ] MicrosoftExcel.ReleaseKeys("<Alt>")
	[+] if ExcelSaveChanges.Exists()
		[ ] ExcelSaveChanges.No.Click()
[+] testcase printAccess2_7_1_5() appstate Access // User N/A
	[ ] STRING sOpenFileName = sGlobalReadPath+"inventory_demo.mdb"
	[ ] STRING sSaveAsFileName = sGlobalWritePath+"inventory_demo1.mdb"
	[ ] 
	[ ] openAccessdoc(sOpenFileName)
	[ ] MicrosoftAccess.sTag = "Inventory Calculations Demo"
	[ ] MicrosoftAccess.SetActive ()
	[ ] MicrosoftAccess.MsoDockTop.MenuBar.PressKeys("<Alt>")
	[ ] MicrosoftAccess.MsoDockTop.MenuBar.TypeKeys("<F>")
	[ ] MicrosoftAccess.MsoDockTop.MenuBar.ReleaseKeys("<Alt>")
	[ ] MicrosoftAccess.MsoDockTop.MenuBar.TypeKeys("<P>")
	[ ] AccessPrint.SetActive()
	[ ] AccessPrint.TypeKeys("<Enter>")
	[ ] MicrosoftAccess.PressKeys("<Alt>")
	[ ] MicrosoftAccess.TypeKeys("<F4>")
	[ ] MicrosoftAccess.ReleaseKeys("<Alt>")
	[+] // if AccessSaveChanges.Exists()
		[ ] // AccessSaveChanges.No.Click()
	[ ] 
[+] testcase PrintAdobeReader2_7_1_6() appstate AdobeReader // User N/A
	[ ] STRING sOpenFileName = sGlobalReadPath+"instructions.pdf"
	[ ] STRING sSaveAsFileName = "sGlobalWritePath"+"instructions.pdf"
	[ ] openAdobeReader()
	[ ] openAdobeReaderdoc(sOpenFileName)
	[ ] sleep(1)
	[ ] AdobeReader.PressKeys("<Alt>")
	[ ] AdobeReader.TypeKeys("<F>")
	[ ] AdobeReader.ReleaseKeys("<Alt>")
	[ ] AdobeReader.TypeKeys("<P>")
	[ ] AdobeReaderPrint.TypeKeys("<Enter>")
	[ ] closeAdobeReader()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Diagnostic Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // Version Checking
[+] testcase verifyOutlookVersion() appstate none // User N/A
	[ ] Print("Verifying Outlook Version")
	[ ] verify(getOutlookVersion(), getVersionFromExcel("Outlook"))
[+] testcase verifyOutlookBuild() appstate none // User N/A
	[ ] Print("Verifying Outlook Build")
	[ ] verify(getOutlookBuild(), getBuildFromExcel("Outlook"))
[+] testcase verifyWordVersion() appstate none // User N/A
	[ ] Print("Verifying Word Version")
	[ ] STRING sOSVersion = GetOSVersion()
	[+] if sOSVersion == "Microsoft Windows XP Professional 5.1.2600"
		[ ] verify(getWordVersion(), getVersionFromExcel("WordXP"))
	[+] if sOSVersion == "Microsoft(R) Windows(R) Server 2003, Standard Edition 5.2.3790"
		[ ] verify(getWordVersion(), getVersionFromExcel("Word2003Server"))
[+] testcase verifyWordBuild() appstate none // User N/A
	[ ] STRING sOSVersion
	[ ] Print("Verifying Word Build")
	[ ] sOSVersion = GetOSVersion()
	[+] if sOSVersion == "Microsoft Windows XP Professional 5.1.2600"
		[ ] verify(getWordBuild(), getBuildFromExcel("WordXP"))
	[+] if sOSVersion == "Microsoft(R) Windows(R) Server 2003, Standard Edition 5.2.3790"
		[ ] verify(getWordBuild(), getBuildFromExcel("Word2003Server"))
[+] testcase verifyExcelVersion() appstate none // User N/A
	[ ] Print("Verifying Excel Version")
	[ ] verify(getExcelVersion(), getVersionFromExcel("Excel"))
[+] testcase verifyExcelBuild() appstate none // User N/A
	[ ] Print("Verifying Excel Build")
	[ ] verify(getExcelBuild(), getBuildFromExcel("Excel"))
[+] testcase verifyAccessVersion() appstate none // User N/A
	[ ] Print("Verifying Access Version")
	[ ] verify(getExcelVersion(), getVersionFromExcel("Access"))
[+] testcase verifyAccessBuild() appstate none // User N/A
	[ ] Print("Verifying Access Build")
	[ ] verify(getExcelBuild(), getBuildFromExcel("Access"))
[+] testcase verifyPowerPointVersion() appstate none // User N/A
	[ ] Print("Verifying Power Point Version")
	[ ] verify(getPowerPointVersion(), getVersionFromExcel("PowerPoint"))
[+] testcase verifyPowerPointBuild() appstate none // User N/A
	[ ] Print("Verifying PowerPoint Build")
	[ ] verify(getPowerPointBuild(), getBuildFromExcel("PowerPoint"))
[+] testcase verifyAdobeReaderVersion() appstate AdobeReader // User N/A
	[ ] // openAdobeReader()
	[ ] Print("Verifying Adobe Reader Version")
	[ ] verify(getAdobeReaderVersion(), getVersionFromExcel("AdobeReader"))
	[ ] // closeAdobeReader()
[+] testcase verifyWinZipVersion() appstate WinZip // User N/A
	[ ] Print("Verifying Win Zip Version")
	[ ] verify(getWinZipVersion(), getVersionFromExcel("WinZip"))
[+] testcase verifyJavaVersion() appstate none // User N/A
	[ ] Print("Verifying Java Version")
	[ ] verify(getjavaVersion(), getVersionFromExcel("Java"))
[+] testcase verifyOSVersion() // User N/A
	[ ] Print("Verifying OS Version for", Substr(gethostname(),8,4))
	[+] if Substr(gethostname(),8,4) == "CS01" 
		[ ] // Print(gethostname())
		[ ] verify(getOSVersion(), getVersionFromExcel("WindowsXP"))
	[+] if Substr(gethostname(),8,4) == "DC01" || Substr(gethostname(),8,4) == "DC02"
		[ ] // Print(gethostname())
		[ ] verify(getOSVersion(), getVersionFromExcel("Windows2003Server"))
[+] testcase verifyRegistryEntriesDB() appstate none
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sKey
	[ ] STRING sPath
	[ ] STRING sItem
	[ ] STRING sExpectedEntry = ""
	[ ] STRING sActualEntry = ""
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF STRING lsFilename
	[ ] INT iKey = HKEY_LOCAL_MACHINE
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Path, Key, Entry FROM `Registry$`" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, sPath, sKey, sExpectedEntry)
		[+] if ((sKey == NULL) || (sKey == "'00000000"))
			[ ] Print("Verifying Registry Path {sPath} with empty or 00000000 sKey Column")
			[+] do
				[ ] SYS_GetRegistryKeyNames (iKey, sPath)
			[+] except
				[ ] LogError("***** Registry Path not found {sPath} - exception:",ExceptData())
			[ ] continue
		[+] else
			[ ] sKey = StrTran(sKey,"'", "")
		[ ] Print("Verifying Registry Path {sPath}")
		[ ] sPath = sPath+getRegistryPathWithoutKey(sKey, sExpectedEntry)
		[+] do
			[ ] SYS_GetRegistryKeyNames (iKey, sPath)
		[+] except
			[ ] LogError("***** Registry Path not found {sPath} exception:{ExceptData()}")
		[ ] sKey = getRegistryKeyWithoutPath(sKey, sExpectedEntry)
		[ ] Print("")
		[ ] Print("Verifying Registry Path {sPath} Key {sKey} and expected entry {sExpectedEntry}")
		[+] do
			[ ] sActualEntry = SYS_GetRegistryValue(iKey, sPath, sKey)
		[+] except
			[ ] LogError("***** Registry Key not found - Path = {sPath}, Key = {sKey} exception:",ExceptData())
		[+] if sActualEntry != sExpectedEntry
			[ ] Print("Actual Entry {sActualEntry} does not match Expected entry {sExpectedEntry}")
			[ ] 
[+] verifyRegistryEntriesFromFile(STRING sFile)
	[ ] HFILE hFile
	[ ] BOOLEAN hResult
	[ ] STRING sKey
	[ ] STRING sPath
	[ ] STRING sItem
	[ ] STRING sExpectedEntry = ""
	[ ] STRING sActualEntry = ""
	[ ] STRING sLine
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF STRING lsFilename
	[ ] INT iKey = HKEY_LOCAL_MACHINE
	[ ] INT iTab = 9
	[ ] INT iTestsPassed = 0
	[ ] INT iTestsFailed = 0
	[ ] INT iBreakpoint = 1000
	[ ] // INT iLinecount= 1
	[ ] 
	[ ] hFile = FileOpen(sGlobalWritePath+sFile,FM_READ)
	[+] while FileReadLine(hFile, sLine)
		[ ] sPath = GetField(sLine,Chr(9),1)
		[ ] sKey = GetField(sLine,Chr(9),2)
		[ ] sExpectedEntry = GetField(sLine,Chr(9),3)
		[+] // if !(iLinecount % iBreakpoint)
			[ ] // Print("Break at line ",iLinecount)
		[ ] // iLinecount++
		[+] // if StrPos(sExpectedEntry,  "&crlf")
			[ ] // sExpectedEntry = StrTran(sExpectedEntry, "&crlf", Chr(13) + Chr(10))
		[+] // if sKey == "" 
			[+] // if VERBOSE
				[ ] // Print("Verifying Registry Path {sPath} with empty sKey Column")
			[+] // do
				[ ] // SYS_GetRegistryKeyNames (iKey, sPath)
				[ ] // AddTestsPassedCount(1)
			[+] // except
				[ ] // LogError("***** Registry Path not found - {sPath} exception: {ExceptData()}")
				[ ] // AddTestsFailedCount(1)
			[ ] // continue
		[+] // else
			[ ] // sKey = StrTran(sKey,"'", "")
		[+] // if VERBOSE
			[ ] // Print("Verifying Registry Path {sPath}")
		[ ] // sPath = sPath+getRegistryPathWithoutKey(sKey, sExpectedEntry)
		[+] // do
			[ ] // SYS_GetRegistryKeyNames (iKey, sPath)
			[ ] // AddTestsPassedCount(1)
		[+] // except
			[ ] // LogError("***** Registry Path not found {sPath} exception: {ExceptData()}")
			[ ] // AddTestsFailedCount(1)
		[ ] // sKey = getRegistryKeyWithoutPath(sKey, sExpectedEntry)
		[+] if VERBOSE == TRUE
			[ ] Print("Verifying Registry Path {sPath} Key {sKey} and expected entry {sExpectedEntry}")
		[ ] if sPath != NULL && sKey != NULL && sExpectedEntry != NULL
		[+] do
			[ ] sActualEntry = SYS_GetRegistryValue(iKey, sPath, sKey)
			[+] if StrPos(sActualEntry, chr(12))
				[ ] StrTran(sActualEntry, chr(12), "")
		[+] except
			[ ] LogError("***** Registry Key not found {sPath}, {sKey} exception:",ExceptData())
		[+] do
			[ ] verify(sActualEntry,sExpectedEntry)
			[ ] addTestsPassedCount(1)
		[+] except
			[ ] Print("Actual Entry{sPath}{sKey}{sActualEntry} does not match Expected entry {sPath}{sKey}{sExpectedEntry}")
			[ ] addTestsFailedCount(1)
	[ ] FileClose(hFile)
	[ ] 
[+] testcase verifyRegistryEntries() appstate none
	[ ] INT i = 1
	[ ] STRING sBaseFilename = "Registry"
	[ ] Print(sGlobalWritePath+sBaseFilename+"{i}"+".csv")
	[ ] Print(SYS_FileExists(sGlobalWritePath+sBaseFilename+"{i}"+".csv"))
	[+] while SYS_FileExists(sGlobalWritePath+sBaseFilename+"{i}"+".csv")
		[ ] verifyRegistryEntriesFromFile(sBaseFilename+"{i}"+".csv")
		[ ] i++
	[ ] 
[+] testcase verifyRegistryPathsCSV() appstate none
	[ ] HFILE hFile
	[ ] BOOLEAN hResult
	[ ] STRING sKey
	[ ] STRING sPath
	[ ] STRING sItem
	[ ] STRING sExpectedEntry = ""
	[ ] STRING sActualEntry = ""
	[ ] STRING sLine
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF STRING lsFilename
	[ ] INT iKey = HKEY_LOCAL_MACHINE
	[ ] 
	[ ] hFile = FileOpen(sGlobalWritePath+"Registry.csv",FM_READ)
	[+] while FileReadLine(hFile, sLine)
		[ ] sPath = GetField(sLine,",",1)
		[ ] Print("Verifying Registry Path {sPath} with empty sKey Column")
		[+] do
			[ ] SYS_GetRegistryKeyNames (iKey, sPath)
		[+] except
			[ ] LogError("***** Registry Path not found - {sPath} exception:",ExceptData())
	[ ] FileClose(hFile)
[+] testcase verifyIAVAEntries() appstate none
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN bKeyInvalid
	[ ] STRING sKey
	[ ] STRING sExpectedEntry
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF STRING lsFilename
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Key, Entry FROM `IAVA$`" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, sKey, sExpectedEntry)
		[+] bKeyInvalid = FALSE
			[+] if VERBOSE == TRUE 
				[ ] Print("Verifying Registry Key {sKey} and expected entry {sExpectedEntry}")
			[+] do
				[ ] verify(MatchStr("*Unable to open registry key*",getRegistryEntry(sKey)), FALSE)
				[+] if VERBOSE == TRUE 
					[ ] Print("Found Registry key {sKey}")
				[ ] addTestsPassedCount(1)
			[+] except
				[ ] LogError("ERROR: Expected Registry Key {sKey} not found")
				[ ] bKeyInvalid = TRUE
				[ ] addTestsFailedCount(1)
			[+] do
				[ ] verify(MatchStr("*Invalid root in registry key*",getRegistryEntry(sKey)), FALSE)
			[+] except
				[ ] LogError("ERROR: Expected Registry Key {sKey} has an invalid root")
				[ ] bKeyInvalid = TRUE
				[ ] addTestsFailedCount(1)
			[+] if bKeyInvalid == FALSE
				[+] do
					[ ] verify(getRegistryEntry(sKey) == sExpectedEntry, TRUE)
					[ ] addTestsPassedCount(1)
				[+] except
					[ ] LogError("ERROR: Expected and actual registry entries do not match")
					[ ] Print("	Actual Registry Key: {sKey} Entry: {getRegistryEntry(sKey)} ")
					[ ] Print("	Expected Registry Key: {sKey} Entry: {sExpectedEntry} ")
					[ ] addTestsFailedCount(1)
			[ ] 
		[+] // if MatchStr("*Unable to open registry key*",getRegistryEntry(sKey))
			[ ] // LogError("ERROR: Expected Registry Key {sKey} not found in registry")
			[ ] // addTestsFailedCount(1)
		[+] // else if MatchStr("*Invalid root in registry key*",getRegistryEntry(sKey))
			[ ] // LogError("ERROR: Expected Registry Key {sKey} has an invalid root")
			[ ] // addTestsFailedCount(1)
		[+] // else if getRegistryEntry(sKey) != sExpectedEntry
			[ ] // LogError("ERROR: Expected and actual registry entries do not match")
			[ ] // Print("Actual Registry Key: {sKey} Entry: {getRegistryEntry(sKey)} ")
			[ ] // Print("Expected Registry Key: {sKey} Entry: {sExpectedEntry} ")
			[ ] // addTestsFailedCount(1)
		[+] // else
			[ ] // addTestsPassedCount(1)
[ ] 
[+] testcase IAVA13_1_1_IAVA_2005_A_017() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-017\IAVANumber1"
	[ ] STRING sValue = "IAVA 2005-A-0017"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ] 
[+] testcase IAVA13_1_2_IAVA_2005_A_018() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-018\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0018"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_3_IAVA_2005_A_025() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-025\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0025"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_4_IAVA_2005_A_027() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-027\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0027"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_5_IAVA_2005_A_029() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-029\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0029"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_6_IAVA_2005_A_030() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-030\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0030"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_7_IAVA_2005_A_039() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-039\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0039"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_8_IAVA_2005_A_040() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-040\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0040"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_9_IAVA_2005_A_052() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-052\IAVANumber"
	[ ] STRING sValue = "IAVA 2005-A-0042"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_10_IAVA_2006_A_001() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA06-001\IAVANumber"
	[ ] STRING sValue = "IAVA 2006-A-0001"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_11_IAVA_2006_A_002() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA06-002\IAVANumber"
	[ ] STRING sValue = "IAVA 2006-A-0002"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_12_IAVA_2006_A_003() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA06-003\IAVANumber"
	[ ] STRING sValue = "IAVA 2006-A-0003"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[+] testcase IAVA13_1_13_IAVB_2006_A_003() appstate none // User N/A
	[ ] STRING sKey = "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVB05-010IAVANumber"
	[ ] STRING sValue = "IAVB 2005-B-0010"
	[ ] Print("Verifiying IAVA install {sValue}")
	[ ] verify(getRegistryEntry(sKey), sValue )
	[ ]  
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[ ] // Verify Services  
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[+]  testcase verifyServices() appstate none
	[ ] STRING sExcelWorksheetName  = getServerName() + "Services"
	[ ] LIST OF SERVICEREC lsActualServices = getServicesList()
	[ ] LIST OF SERVICEREC lsExpectedServices = getServicesFromExcel(sExcelWorksheetName)
	[ ] 
	[ ] STRING sItem
	[ ] STRING iListCount 
	[ ] INTEGER iActualIndex
	[ ] INTEGER iExpectedIndex
	[ ] INTEGER iFailed = 0
	[ ] INTEGER iPassed = 0
	[ ] BOOLEAN bExpectedServiceFound = FALSE
	[ ] BOOLEAN bActualServiceFound = FALSE
	[ ] 
	[+] for iExpectedIndex = 1 to ListCount(lsExpectedServices) 
		[ ] bExpectedServiceFound = FALSE
		[+] for iActualIndex = 1 to ListCount(lsActualServices)
			[+] if lsExpectedServices[iExpectedIndex].sDisplayName == lsActualServices[iActualIndex].sDisplayName
				[ ] bExpectedServiceFound = TRUE
				[+] if lsActualServices[iActualIndex].sState != lsExpectedServices[iExpectedIndex].sState
					[ ] Print("***** Expected and Actual States do not match for Service:",lsExpectedServices[iExpectedIndex].sDisplayName)
					[ ] Print("               Actual State:",lsActualServices[iActualIndex].sState)
					[ ] Print("               Expected State:",lsExpectedServices[iExpectedIndex].sState)
					[ ] Print("")
					[ ] addTestsFailedCount(1)
				[+] else if lsActualServices[iActualIndex].sStartMode != lsExpectedServices[iExpectedIndex].sStartMode
					[ ] Print("***** Expected and Actual Start Modes do not match for Service:",lsActualServices[iActualIndex].sDisplayName)
					[ ] Print("                Actual Start Mode:",lsActualServices[iActualIndex].sStartMode)
					[ ] Print("                Expected Start Mode:",lsExpectedServices[iExpectedIndex].sStartMode)
					[ ] Print("")
					[ ] addTestsFailedCount(1)
				[+] else 
					[ ] addTestsPassedCount(1)
		[+] if bExpectedServiceFound == FALSE
			[ ] LogError("***** Expected service {lsExpectedServices[iExpectedIndex].sDisplayName} not found in Actual List")
			[ ] addTestsFailedCount(1)
	[+] for iActualIndex = 1 to ListCount(lsActualServices)
		[ ] bActualServiceFound = FALSE
		[+] for iExpectedIndex = 1 to ListCount(lsExpectedServices) 
			[+] if lsExpectedServices[iExpectedIndex].sDisplayName == lsActualServices[iActualIndex].sDisplayName
				[ ] bActualServiceFound = TRUE
				[+] if lsActualServices[iActualIndex].sState != lsExpectedServices[iExpectedIndex].sState
					[ ] LogError("***** Expected and Actual States do not match for Service:",lsActualServices[iActualIndex].sDisplayName)
					[ ] LogError("               Actual State:",lsActualServices[iActualIndex].sState)
					[ ] LogError("               Expected State:",lsExpectedServices[iExpectedIndex].sState)
					[ ] Print("")
					[ ] addTestsFailedCount(1)
				[+] else if lsActualServices[iActualIndex].sStartMode != lsExpectedServices[iExpectedIndex].sStartMode
					[ ] Print("***** Expected and Actual Start Modes do not match for Service:",lsActualServices[iActualIndex].sDisplayName)
					[ ] Print("                Actual Start Mode:",lsActualServices[iActualIndex].sStartMode)
					[ ] Print("                Expected Start Mode:",lsExpectedServices[iExpectedIndex].sStartMode)
					[ ] Print("")
					[ ] addTestsFailedCount(1)
				[+] else 
					[ ] addTestsPassedCount(1)
		[+] if bActualServiceFound == FALSE
			[ ] LogError("***** Actual service {lsActualServices[iActualIndex].sDisplayName} not found in expected list")
			[ ] addTestsFailedCount(1)
	[ ] 
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[ ] // Get Existing list of Services for Excel
[+] testcase tc_createComposeDLLVersionsCSV()
	[ ] STRING sItem
	[ ] LIST OF VERSIONREC lsComposeDLLs = getComposeDLLs()
	[ ] VERSIONREC VERSION
	[ ] HFILE hFile
	[ ] STRING sFile = "C:\SilkTestData\COMPOSEDLLVERSIONS.CSV"
	[ ] hFile = FileOpen (sFile, FM_WRITE) // Open the file
	[+] for each VERSION in lsComposeDLLs
		[ ] FileWriteLine (hFile, VERSION.sFilename +","+ getFileVersion(VERSION.sFilename))
		[ ] // Print(sItem, ",", getFileVersion(sItem))
	[ ] FileClose(hFile)
[ ] 
[+] testcase createServicesForExcelCSV () appstate none
	[ ] LIST OF STRING lsServices = getServicesExcel()
	[ ] SERVICEREC SERVICE
	[ ] LIST OF SERVICEREC SERVICES
	[ ] STRING sItem
	[ ] STRING sExcelHeader = "DisplayName,State,StartMode"
	[ ] STRING sHost = Left(getHostName(),11)
	[ ] sHost = Right(sHost,4)
	[ ] HFILE hFile
	[ ] STRING sFile = "C:\SilkTestData\{sHost}SERVICES.CSV"
	[ ] Print(sHost)
	[ ] hFile = FileOpen (sFile, FM_WRITE) // Open the file
	[ ] FileWriteLine (hFile, sExcelHeader)
	[+] for each sItem in lsServices
		[ ] FileWriteLine (hFile, sItem)
		[ ] // Print(sItem)
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[ ] // Event Log Functions and Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[+]  InitTestStartTime()
	[ ] DATETIME dTestStartDateTime = GetDateTime()
	[ ] // Store Time Using Enviroment Variable
	[ ] // SYS_SetEnv("SILKTESTSTARTTIME",[STRING]dTestStartDateTime)
	[ ] // Store Time using a file
	[ ] HFILE hFile
	[ ] STRING sFile = "C:\SilkTestData\SILKTESTSTARTTIME.TXT"
	[ ] hFile = FileOpen (sFile, FM_WRITE) // Open the file
	[ ] FileWriteValue (hFile, dTestStartDateTime)
	[ ] FileClose(hFile)
	[ ] Print(dTestStartDateTime)
	[ ] 
[+] LIST OF STRING getEventLogInfo(STRING sLogFile, STRING sEventType)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getEventLogInfo
	[ ] // Purpose:                                  This function creates a  
	[ ] //								contact in Outlook
	[ ] //								The itemtype is hardcoded to 1
	[ ] //								in the VB Script
	[ ] // Inputs:                                   	STRING LogFile
	[ ] //								INT Event Type
	[ ] // LOGFILE TYPES:
	[ ] // APPLICATION
	[ ] // SECURITY
	[ ] // SYSTEM
	[ ] // EVENT TYPES:
	[ ] // const INT EVENTLOGERROR = 1
	[ ] // const INT EVENTLOGWARNING = 2
	[ ] // const INT EVENTLOGINFORMATION = 3
	[ ] // const INT EVENTLOGSECURITYAUDITSUCCESS = 4
	[ ] // const INT EVENTLOGSECURITYAUDITFAILURE = 5
	[ ] 
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getEventLogInfo.vbs "+"""{sLogFile}"""+" "+"""{sEventType}"""
	[ ] SYS_Execute(cmdLine,sReturn)
	[ ] return sReturn
	[ ] 
[+] getCurrentEventLogInfo(STRING sLogFile, STRING sEventType, STRING sFormattedDate)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getEventLogInfo
	[ ] // Purpose:                                  This function creates a  
	[ ] //								contact in Outlook
	[ ] //								The itemtype is hardcoded to 1
	[ ] //								in the VB Script
	[ ] // Inputs:                                   	STRING LogFile
	[ ] //								INT Event Type
	[ ] // LOGFILE TYPES:
	[ ] // 								APPLICATION
	[ ] // 								SECURITY
	[ ] // 								SYSTEM
	[ ] // EVENT TYPES:
	[ ] // 								const INT EVENTLOGERROR = 1
	[ ] // 								const INT EVENTLOGWARNING = 2
	[ ] // 								const INT EVENTLOGINFORMATION = 3
	[ ] // 								const INT EVENTLOGSECURITYAUDITSUCCESS = 4
	[ ] // 								const INT EVENTLOGSECURITYAUDITFAILURE = 5
	[ ] // TIMEWRITTEN				datetime
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getCurrentEventLogInfo.vbs "+"""{sLogFile}"""+" "+"""{sEventType}"""+" "+"""{sFormattedDate}"""
	[ ] 
	[ ] Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] Print("sReturn = "+sItem)
	[ ] 
[+] GetEventLog(STRING sLogfile, STRING sEventType) 
	[ ] ListPrint(getEventLogInfo(sLogfile, sEventType))
[+] testcase GetCurrentEventLog(STRING sLogfile, STRING sEventType) appstate none
	[ ] STRING sDateTime
	[ ] DATETIME dTestStartDateTime
	[ ] // File Based Global
	[ ] HFILE hFile
	[ ] STRING sFile = "C:\SilkTestData\SILKTESTSTARTTIME.TXT"
	[ ] hFile = FileOpen (sFile, FM_READ) // Open the file
	[ ] FileReadValue (hFile, dTestStartDateTime)
	[ ] Print(dTestStartDateTime)
	[ ] sDateTime = [STRING]dTestStartDateTime
	[ ] Print("sDateTime", sDateTime)
	[ ] FileClose(hFile)
	[ ] // Enviroment Variable Global
	[ ] // sFormattedDateTime = SYS_GetEnv("SILKTESTSTARTTIME")
	[ ] // Print("sFormattedDateTime =", sFormattedDateTime)
	[+] if dTestStartDateTime == NULL
		[ ] LogError("dTestStartDateTime = NULL, must call InitTestStartTime before calling this function")
	[ ] // sDateTime = FormatDateTime (dTestStartDateTime, "yyyy-mm-dd hh:mm:ss.fff") 
	[ ] sDateTime = StrTran(sDateTime, "-","")
	[ ] sDateTime = StrTran(sDateTime, ":","")
	[ ] sDateTime = StrTran(sDateTime, " ","")
	[ ] sDateTime = sDateTime + "-480"
	[ ] // Print("Formatted Date Time",sFormattedDateTime)
	[ ] getCurrentEventLogInfo(sLogfile, sEventType,sDateTime)
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[ ] // DLL/File Version Functions and Testcases
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[+] STRING getFileVersion(STRING sFile)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getFilenameVersion
	[ ] // Purpose:                                  This function returns
	[ ] //								the file version number
	[ ] //
	[ ] // Inputs:                                   	STRING sFile
	[ ] // Note: Must Include Path in sFile						
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine 
	[ ] STRING sSystemRoot = GetField(sFile,"%",2)
	[ ] // Print("System Root =",sSystemRoot)
	[+] if sSystemRoot == "SYSTEMROOT"
		[+] if Substr(gethostname(),8,4) == "CS01"
			[ ] sFile = StrTran(sFile,"%SYSTEMROOT%","C:\WINDOWS") 
		[+] if Substr(gethostname(),8,4) == "DC01" || Substr(gethostname(),8,4) == "DC02"
			[ ] sFile = StrTran(sFile,"%SYSTEMROOT%","C:\WINNT")
	[ ] cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getVersion.vbs "+"""{sFile}"""
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[+] if MatchStr("*cannot find the file*",sItem)
			[ ] LogWarning("cannot find the file {sFile} - check path and filename")
			[ ] return("cannot find the file {sFile} - check path and filename")
		[+] else
			[ ] return sItem
	[ ] 
[+] testcase PrintGetVersion() appstate none
	[ ] Print(getFileVersion("C:\WINNT\system32\SQLSRV32.DLL"))
[+] testcase createComposeDLLVersionsForExcelCSV()
	[ ] STRING sItem
	[ ] LIST OF VERSIONREC lsComposeDLLs = getComposeDLLs()
	[ ] VERSIONREC VERSION
	[ ] HFILE hFile
	[ ] STRING sFile = "C:\SilkTestData\COMPOSEDLLVERSIONS.CSV"
	[ ] hFile = FileOpen (sFile, FM_WRITE) // Open the file
	[ ] STRING sExcelHeader = "Filename,Version"
	[+] for each VERSION in lsComposeDLLs
		[ ] FileWriteLine (hFile, VERSION.sFilename +","+ getFileVersion(VERSION.sFilename))
		[ ] // Print(sItem, ",", getFileVersion(sItem))
	[ ] FileClose(hFile)
[+] testcase verifyComposeDLLVersions()
	[ ] STRING sItem
	[ ] INT iExpectedIndex
	[ ] INT iActualIndex
	[ ] INT iTestCountFailed = 0
	[ ] INT iTestCountPassed = 0
	[ ] 
	[ ] BOOLEAN bActualDLLFound = TRUE
	[ ] BOOLEAN bExpectedDLLFound = TRUE
	[ ] LIST OF VERSIONREC lsActualComposeDLLs = getComposeDLLs()
	[ ] LIST OF VERSIONREC lsExpectedComposeDLLs = getComposeDLLsFromExcel()
	[ ] // Print("Verifying DLL {Filename} and expected version {sExpectedVersion}")
	[+] for iExpectedIndex = 1 to ListCount(lsExpectedComposeDLLs) 
		[ ] bExpectedDLLFound = FALSE
		[+] for iActualIndex = 1 to ListCount(lsActualComposeDLLs)
			[+] if lsExpectedComposeDLLs[iExpectedIndex].sFilename == lsActualComposeDLLs[iActualIndex].sFilename
				[ ] bExpectedDLLFound = TRUE
				[+] if lsActualComposeDLLs[iActualIndex].sVersion == NULL || lsActualComposeDLLs[iActualIndex].sVersion == ""
					[ ] Print("***** No version number found for Actual DLL {lsActualComposeDLLs[iActualIndex].sFilename}")
					[ ] continue
				[+] else if lsActualComposeDLLs[iActualIndex].sVersion != lsExpectedComposeDLLs[iExpectedIndex].sVersion
					[ ] LogError("***** Actual and Expected Versions do not match for DLL:",lsActualComposeDLLs[iActualIndex].sFilename)
					[ ] LogError("               Actual Version:",lsActualComposeDLLs[iActualIndex].sVersion)
					[ ] LogError("               Expected Version:",lsExpectedComposeDLLs[iExpectedIndex].sVersion)
					[ ] LogError("")
					[ ] addTestsFailedCount(1)
				[+] else 
					[+] if VERBOSE 
						[ ] Print("Verified actual version for DLL "+lsExpectedComposeDLLs[iExpectedIndex].sFilename+" is "+ lsExpectedComposeDLLs[iExpectedIndex].sVersion+" matches")
					[ ] addTestsPassedCount(1)
		[+] if bExpectedDLLFound == FALSE
			[ ] LogError("***** Expected DLL {lsExpectedComposeDLLs[iExpectedIndex].sFilename} not found")
			[ ] addTestsFailedCount(1)
	[+] for iActualIndex = 1 to ListCount(lsActualComposeDLLs)
		[+] bActualDLLFound = FALSE
			[+] if lsExpectedComposeDLLs[iActualIndex].sVersion == NULL || lsExpectedComposeDLLs[iActualIndex].sVersion == ""
				[ ] Print("***** No version number found for Expected DLL {lsExpectedComposeDLLs[iActualIndex].sFilename}")
				[ ] continue
		[+] for iExpectedIndex = 1 to ListCount(lsExpectedComposeDLLs) 
			[+] if lsExpectedComposeDLLs[iExpectedIndex].sFilename == lsActualComposeDLLs[iActualIndex].sFilename
				[ ] bActualDLLFound = TRUE
				[+] if lsActualComposeDLLs[iActualIndex].sVersion != lsExpectedComposeDLLs[iExpectedIndex].sVersion
					[ ] LogError("***** Expected and Actual Versions do not match for DLL:",lsActualComposeDLLs[iActualIndex].sFilename)
					[ ] LogError("               Actual Version:",lsActualComposeDLLs[iActualIndex].sVersion)
					[ ] LogError("               Expected Version:",lsExpectedComposeDLLs[iExpectedIndex].sVersion)
					[ ] Print("")
					[ ] addTestsFailedCount(1)
				[+] else 
					[+] if VERBOSE
						[ ] Print("Verified expected version for DLL "+lsExpectedComposeDLLs[iExpectedIndex].sFilename+" is "+ lsExpectedComposeDLLs[iExpectedIndex].sVersion+" matches")
					[ ] // addTestCountFailed(1)
		[+] if bActualDLLFound == FALSE
			[ ] LogError("***** Actual DLL {lsActualComposeDLLs[iActualIndex].sFilename} not found")
			[ ] addTestsFailedCount(1)
[+] testcase verifyMDACDLLVersions()
	[ ] STRING sItem
	[ ] INT iIndex
	[ ] INT iTestCountFailed = 0
	[ ] INT iTestCountPassed = 0
	[ ] 
	[ ] BOOLEAN bActualDLLFound = TRUE
	[ ] BOOLEAN bExpectedDLLFound = TRUE
	[ ] LIST OF VERSIONREC lsActualMDACDLLs = getMDACDLLs()
	[ ] LIST OF VERSIONREC lsExpectedMDACDLLs = getMDACDLLsFromExcel()
	[ ] // Print("Verifying DLL {Filename} and expected version {sExpectedVersion}")
	[+] for iIndex = 1 to ListCount(lsExpectedMDACDLLs) 
		[+] do
			[ ] verify(lsActualMDACDLLs[iIndex].sVersion, lsExpectedMDACDLLs[iIndex].sVersion)
		[+] except
			[ ] LogError("***** ERROR: Versions of DLL's do not match for file {lsActualMDACDLLs[iIndex].sFilename}")
			[ ] Print("     Actual Version: {lsActualMDACDLLs[iIndex].sVersion}")
			[ ] Print("     Expected Version:  {lsExpectedMDACDLLs[iIndex].sVersion}")
			[ ] 
[+] // testcase verifyMDACWithComponentChecker() appstate none
	[ ] // // STRING sCMD = "{sGlobalReadPath}cc.exe"
	[ ] // STRING sCMD = "c:\SilkTestData\cc.exe"
	[ ] // 
	[ ] // LIST OF STRING lsReturn
	[ ] // STRING sMsgText
	[+] // // SYS_EXECUTE(sCMD, lsReturn)
		[ ] // // ListPrint(lsReturn)
	[ ] // StartProgramAsync(sCMD)
	[ ] // // SYS_EXECUTE(sCMD,lsReturn)
	[ ] // ListPrint(lsReturn)
	[ ] // ComponentChecker.OK.Click()
	[ ] // ComponentCheckerComplete.SetActive()
	[ ] // sMsgText = ComponentCheckerComplete.Message.GetText()
	[+] // if VERBOSE
		[ ] // Print("Results from Component Checker MDAC version  check:",sMsgText)
	[+] // do
		[ ] // verify(MatchStr( "*matched*", sMsgText), TRUE)
	[+] // except
		[ ] // LogError("Component Checker MDAC version check failed: Message = {sMsgText}")
	[ ] // ComponentCheckerComplete.No.Click()
	[ ] // ComponentCheckerVersion20.SetActive()
	[ ] // ComponentCheckerVersion20.TypeKeys("<alt-f>")
	[ ] // ComponentCheckerVersion20.TypeKeys("x")
	[ ] // 
[+] int GetCharCount (STRING sSource, STRING sFind)
	[ ] int i  
	[ ] int iCount = 0  
	[+] for i = 1 to Len (sSource)    
		[+] if (sSource[i] == sFind)   
			[ ]  iCount++  
	[ ] return iCount
	[ ] 
[+] STRING StrReplaceSingle (STRING sSource, STRING sFind) 
	[ ] int i  
	[ ] int j
	[ ] int iCount = 0 
	[ ] STRING sResult = sSource
	[ ] INT iLength = Len (sSource)    
	[ ] INT iNewLength = Len (sSource) - 1
	[+] for i = 1 to iLength    
		[+] if (sSource[i] == sFind) 
			[+] for j = i to iNewLength
				[ ] sResult[ j ] = sSource[ j +1] 
				[ ] sResult[iLength] = ""
	[ ] return sResult
	[ ] 
[+] testcase ReplaceSingle() appstate none
	[ ]  Print(StrReplaceSingle("abc\\ced","\"))
[+] STRING getRegistryKeyWithoutPath(STRING sPath, STRING sExpectedKey)
	[ ] STRING sKeyWithoutPath = ""
	[ ] INT iCount = GetCharCount(sPath, "\")
	[ ] INT iExpectedKeyCount = GetCharCount(sExpectedKey, "\")
	[ ] INT i
	[ ] // Print("iCount = ", iCount)
	[+] // if iCount == 0 
		[ ] // return sPath
		[ ] //  sKeyWithoutPath = GetField(sPath,"\",iCount+1)
	[+] // // if iExpectedKeyCount == 1
		[ ] // // sKeyWithoutPath = GetField(sPath,"\",(iCount+1) - i ) + sKeyWithoutPath
	[+] // for i = 1 to iExpectedKeyCount
		[ ] // sKeyWithoutPath = GetField(sPath,"\",(iCount+1) - i ) + sKeyWithoutPath
	[ ] sKeyWithoutPath = GetField(sPath,"\",(iCount+1))
	[ ] return sKeyWithoutPath
	[ ] 
[+] STRING getRegistryPathWithoutKey(STRING sPath, STRING sExpectedValue)
	[ ] STRING sPathWithoutKey = ""
	[ ] INT iCount = GetCharCount(sPath, "\")
	[ ] INT i
	[ ] // Print("iCount = ", iCount)
	[+] if iCount == 0 
		[ ] return ""
	[ ] INT iPathLength = Len(sPath)
	[ ] INT iKeyLength = Len(GetField(sPath,"\", iCount))
	[ ] sPathWithoutKey = Left(sPath, iPathLength-iKeyLength)
	[ ] return sPathWithoutKey
	[ ] 
[+] testcase getRegistryKeyCount()
	[ ] INT iKey = HKEY_LOCAL_MACHINE
	[ ] STRING sPath = ""
	[ ] LIST OF STRING lsPaths =  Reg_EnumKeysAll (iKey, sPath)
	[ ] Print(ListCount(lsPaths))
[+] createRegistryEntriesForExcel(INT iKey, STRING sStartPath optional, STRING sFilename)
	[ ] // INTEGER iKey = HKEY_LOCAL_MACHINE // defined in msw32.inc
	[ ] HFILE hFile, hOutputFile, hRegistryExclusionFile
	[ ] STRING sPath
	[ ] STRING sPathNextLine, sPathNextLine2
	[ ] STRING  sDateStr = DateStr()
	[ ] STRING sOutputFilename = "{sGlobalWritePath}Registry1.csv"
	[ ] STRING sEnumKey
	[ ] STRING sKey
	[ ] STRING sValue
	[ ] STRING sItem
	[ ] STRING sRegistryExclusionItem
	[ ] STRING sRegistryExclusionFilename = "{sGlobalWritePath}RegistryExclusionList.csv"
	[ ] // INT iKey =GetRegistryConst("HKEY_LOCAL_MACHINE")
	[ ] LIST OF STRING lsEnumKeysAll
	[ ] LIST OF STRING lsValues
	[ ] LIST OF REGISTRYENTRYREC lsRegistryEntryRec
	[ ] REGISTRYENTRYREC REGISTRYENTRY
	[ ] LIST OF ANYTYPE lsEnumValues
	[ ] BOOLEAN bResult
	[ ] BOOLEAN bFound
	[ ] BOOLEAN bContinueReadingEntries = FALSE
	[ ] INT i
	[ ] INT iTab = 9 // Tab
	[ ] INT iCharCount
	[ ] INT iLineCount = 0
	[ ] INT iMaxFileLines = 32000
	[ ] hFile = FileOpen(sFilename,FM_READ)
	[ ] hOutputFile = FileOpen(sOutputFilename,FM_WRITE)
	[ ] // hOutputFile = FileOpen(sOutputFilename,FM_WRITE,NULL,FT_UNICODE)
	[+] if REGISTRYEXCLUSION == FALSE
		[ ] hRegistryExclusionFile = FileOpen(sRegistryExclusionFilename,FM_WRITE)
		[ ] // hRegistryExclusionFile = FileOpen(sRegistryExclusionFilename,FM_WRITE,NULL,FT_UNICODE)
		[ ] FileWriteLine(hRegistryExclusionFile,"Path" + Chr(iTab) + "Entry"+ Chr(iTab) + "Exception")
	[+] if REGISTRYEXCLUSION == TRUE
		[ ] lsRegistryEntryRec = GetRegistryEntryRec("RegistryExclusionList")
	[ ] // Read in lines until an open square bracket is found
	[+] while FileReadLine(hFile,sPath)
		[ ] // Skip if Start Path not found
		[+] if !StrPos(sStartPath, sPath)
			[ ] FileReadLine(hFile,sPath)
		[ ] // When we encounter an open square bracket, read all lines until the next open square bracket into a list
		[+] if sPath[1] == "["
			[ ] sPath = StrTran(sPath, "[", "")
			[ ] sPath = StrTran(sPath, "]", "")
			[ ] // Skip the first entry- it contains nothing
			[+] if sPath == "HKEY_LOCAL_MACHINE" 
				[ ] continue
			[ ] // Since the root key must be passed separately, remove it
			[ ] sPath = StrTran(sPath,"HKEY_LOCAL_MACHINE\","")
			[ ] // Skip keys that have the same parent
			[+] if !StrPos(sStartPath, sPath)
				[ ] FileReadLine(hFile,sPath)
				[ ] continue
			[ ] // Skip Keys that have been excluded from the registry exclusion list
			[+] if REGISTRYEXCLUSION == TRUE
				[ ] bFound = FALSE
				[+] for each REGISTRYENTRY in lsRegistryEntryRec
					[+] if REGISTRYENTRY.sPath != NULL && REGISTRYENTRY.sEntry == NULL
						[+] if StrPos(REGISTRYENTRY.sPath, sPath)
							[+] if VERBOSE
								[ ] Print("Path",REGISTRYENTRY.sPath,"excluded from test based on entry in RegistryExclusionList")
							[ ] FileReadLine(hFile,sPath)
							[ ] bFound = TRUE
				[+] if bFound == TRUE 
					[ ] continue
			[ ] // END REGISTRYEXCLUSION
			[ ] sPathNextLine = "Init"
			[ ] while sPathNextLine != ""
			[+] // while (sPathNextLine != "" || bContinueReadingEntries)
				[ ] FileReadLine(hFile,sPathNextLine) 
				[+] if sPathNextLine[1] == '"'
					[ ] sItem = GetField(sPathNextLine,'"',2)
					[+] if StrPos("\",sItem)
						[ ] // sItem = StrReplaceSingle(sItem, "\")
						[ ] sItem = StrTran(sItem,"\\", "\")
					[+] if sItem == "" || StrPos("=",sItem) || !StrPos(sStartPath,sPath) 
						[ ] continue
					[+] if REGISTRYEXCLUSION == TRUE
						[ ] bFound = FALSE
						[+] for each REGISTRYENTRY in lsRegistryEntryRec
							[+] if REGISTRYENTRY.sPath != NULL && REGISTRYENTRY.sEntry != NULL
								[+] if StrPos(REGISTRYENTRY.sPath, sPath) && (REGISTRYENTRY.sEntry == sItem)
									[+] if REGISTRYENTRY.sException != NULL
										[ ] Print("Path",REGISTRYENTRY.sPath,"and Key",REGISTRYENTRY.sEntry,"excluded from test  based on previous exception:",REGISTRYENTRY.sException,"in registry exclusion list")
									[+] else
										[ ] Print("Path",REGISTRYENTRY.sPath,"and Key",REGISTRYENTRY.sEntry,"excluded from test  - exception not found in registry exclusion list")
									[ ] FileReadLine(hFile,sPathNextLine)
									[ ] bFound = TRUE
						[+] if bFound == TRUE 
							[ ] continue
					[+] do
						[ ] sValue = SYS_GetRegistryValue (iKey, sPath, sItem)
						[ ] // The following code accounts for multiple cr's and lf's in a single registry entry (value)
						[ ] // This code was wrriten for HKEY_LOCAL_MACHINE\SOFTWARE\INTEL\LANDESK\VirusProtect6\CurrentVersion\ClientConfig\Storages\FileSystem\RealTimeScan\MessageText
						[+] if GetCharCount(sValue, Chr(13)) > 1
							[+] for iCharCount = 1 to GetCharCount(sValue, Chr(13)) * 2
								[ ] FileReadLine(hFile,sPathNextLine) 
							[ ] sValue = StrTran(sValue, Chr(13)+Chr(10),"&crlf")
						[ ] // The following lines are specifically for the paths
						[ ] // [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\IAVA05-017]
						[ ] // [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\IAVA05-018]
						[ ] // etc
						[+] if GetCharCount(sValue, Chr(13)) == 1
							[ ] sValue = StrTran(sValue, Chr(13)+Chr(10),"&crlf")
							[ ] // Read Past Blank Line
							[ ] FileReadLine(hFile,sPathNextLine) 
							[ ] // Read Past Comprises Line
							[ ] FileReadLine(hFile,sPathNextLine) 
						[ ] // Finally, after checking the contents and encoding 
						[ ] FileWriteLine(hOutputFile, sPath  + Chr(iTab) + sItem + Chr(iTab) + sValue)
						[ ] iLineCount++
						[+] if !(iLineCount % iMaxFileLines)
							[ ] FileClose(hOutputFile)
							[ ] INT iFileNameIncrement = (iLineCount/iMaxFileLines) +1
							[ ] sOutputFilename = "{sGlobalWritePath}Registry{iFileNameIncrement}.csv"
							[ ] hOutputFile = FileOpen(sOutputFilename,FM_WRITE)
							[ ] 
					[+] except
						[ ] // if MatchStr("*large*", ExceptData())
						[ ] LogError("An error occured while getting or writing the value in the registry key using the path {getRegistryKeyString(iKey)}{sPath}{sItem}: {ExceptData()}")
						[+] if REGISTRYEXCLUSION == FALSE
							[ ] // Don't bother trying to write double bytes characters to the exclusion file 
							[+] if !MatchStr("*Error: Cannot write double byte characters*",ExceptData())
								[+] do
									[ ] FileWriteLine(hRegistryExclusionFile, sPath + Chr(iTab) + sItem + Chr(iTab) + ExceptData())
								[+] except
									[ ] // PrintCallStack (ExceptCalls ())
									[ ] LogError("An error occured while writing the value in the registry exclusion file using the path {getRegistryKeyString(iKey)}{sPath}{sItem}: {ExceptData()}")
						[ ] // FileReadLine(hFile,sPathNextLine) 
						[ ] // Known problem: skips the following write, but also skips every subsequent sItem until the next path
						[ ] // continue
	[ ] FileClose(hFile)
	[ ] FileClose(hOutputFile)
	[+] if REGISTRYEXCLUSION == FALSE
		[ ] FileClose(hRegistryExclusionFile)
[ ] 
[+] INT getRegistryKeyInt(STRING sKey)
	[ ] // LIST OF REGISTRYKEYSREC isKeys
	[ ] // Build List of Keys
	[+] switch sKey
		[+] case  "HKEY_CLASSES_ROOT"
			[ ] return(0x80000000)
		[+] case  "HKEY_CURRENT_USER"
			[ ] return(0x80000001)
		[+] case  "HKEY_LOCAL_MACHINE"
			[ ] return(0x80000002)
		[+] case  "HKEY_USERS"
			[ ] return(0x80000003)
		[+] case  "HKEY_PERFORMANCE_DATA"
			[ ] return(0x80000004)
		[+] case  "HKEY_CURRENT_CONFIG"
			[ ] return(0x80000005)
		[+] case  "HKEY_DYN_DATA"
			[ ] return(0x80000006)
			[ ] 
[+] STRING getRegistryKeyString(INT iKey)
	[ ] // LIST OF REGISTRYKEYSREC isKeys
	[ ] // Build List of Keys
	[+] switch iKey
		[+] case  0x80000000
			[ ] return("HKEY_CLASSES_ROOT")
		[+] case  0x80000001
			[ ] return("HKEY_CURRENT_USER")
		[+] case  0x80000002
			[ ] return("HKEY_LOCAL_MACHINE")
		[+] case  0x80000003
			[ ] return("HKEY_USERS")
		[+] case  0x80000004
			[ ] return("HKEY_PERFORMANCE_DATA")
		[+] case  0x80000005
			[ ] return("HKEY_CURRENT_CONFIG")
		[+] case 0x80000006
			[ ] return("HKEY_DYN_DATA")
			[ ] 
[+] testcase PrintGetRegistryKeyString()
	[ ] Print(getRegistryKeyString(0x80000000))
[+] testcase PrintGetRegistryKeyNames()
	[ ] STRING sRegPath = "SOFTWARE\INTEL\LANDesk\VirusProtect6\CurrentVersion\ClientConfig\Storages\FileSystem\RealTimeScan"
	[ ] ANYTYPE sItem
	[ ] ListPrint(SYS_GetRegistryKeyNames (getRegistryKeyInt("HKEY_LOCAL_MACHINE"), sRegPath))
[ ] 
[+] PrintCallStack(LIST OF CALL lcCall)
	[ ] CALL Call
	[+] for each Call in lcCall
		[ ] Print ("MODULE: {Call.sModule}", "FUNCTION: {Call.sFunction}", "LINE: {Call.iLine}")
	[ ] 
[+] // testcase verifyRegistryList() appstate none
	[ ] // // testcase verifyRegistryKeysList(LIST OF ANYTYPE laActual, LIST OF ANYTYPE laExpected)
	[ ] // // Function incomplete
	[ ] // // INTEGER iKey = HKEY_LOCAL_MACHINE // defined in msw32.inc
	[ ] // HFILE hFile, hOutputFile
	[ ] // STRING sActualFilename = "C:\DC01.reg"
	[ ] // STRING sExpectedFilename = "C:\DC01Expected.reg"
	[ ] // STRING sPath
	[ ] // STRING sPathNextLine
	[ ] // STRING sPreviousPath = "HARDWARE"
	[ ] // STRING  sDateStr = DateStr()
	[ ] // STRING sEnumKey
	[ ] // STRING sKey
	[ ] // STRING sValue
	[ ] // STRING sItem
	[ ] // LIST OF STRING lsEnumKeysAll
	[ ] // LIST OF STRING lsValues
	[ ] // LIST OF ANYTYPE lsEnumValues
	[ ] // LIST OF ANYTYPE laActual
	[ ] // LIST OF ANYTYPE laExpected
	[ ] // STRING sRegistryKey
	[ ] // REGISTRYENTRYREC REGISTRYENTRY
	[ ] // BOOLEAN bResult
	[ ] // INT i
	[ ] // INT iTab = 9 // Tab
	[ ] // 
	[ ] // // READ ACTUAL .REG FILE INTO ACTUAL LIST
	[ ] // hFile = FileOpen(sActualFilename,FM_READ)
	[+] // while FileReadLine(hFile,sPath)
		[+] // if StrPos(sPreviousPath,sPath)
			[ ] // sPreviousPath = sPath
			[ ] // continue
		[ ] // // A '[' at the beginning of line signifies a path
		[+] // if sPath[1] == "["
			[ ] // sPath = StrTran(sPath, "[", "")
			[ ] // sPath = StrTran(sPath, "]", "")
			[ ] // // for each sRegistryKey in sRegistryKeys
			[ ] // // sRegistryKey = sPath
			[+] // if sPath ==  "HKEY_LOCAL_MACHINE"
					[ ] // continue
			[ ] // sPath = StrTran(sPath,"HKEY_LOCAL_MACHINE\","")
			[ ] // sPathNextLine = "Init"
			[+] // while sPathNextLine != ""
				[ ] // FileReadLine(hFile,sPathNextLine) 
				[+] // if sPathNextLine[1] == '"'
					[ ] // sItem = GetField(sPathNextLine,'"',2)
					[+] // do
						[ ] // // sValue = SYS_GetRegistryValue (getRegistryKeyConst(sRegistryKey), sPath, sItem)
						[ ] // sValue = SYS_GetRegistryValue (HKEY_LOCAL_MACHINE, sPath, sItem)
						[ ] // REGISTRYENTRY.sKey = sRegistryKey
						[ ] // REGISTRYENTRY.sPath = sPath
						[ ] // REGISTRYENTRY.sItem = sItem
						[ ] // REGISTRYENTRY.sValue = sValue
						[ ] // ListAppend(laActual, REGISTRYENTRY)
						[ ] // ListPrint( laActual)
						[ ] // // FileWriteLine(hOutputFile, sPath  + Chr(iTab) + sItem + Chr(iTab) + sValue)
					[+] // except
						[ ] // LogError("An error occured while getting the value in the registry key: {REGISTRYENTRY}{ExceptData()}")
	[ ] // ListPrint(laActual)
	[ ] // FileClose(hFile)
	[ ] // // READ EXPECTED .REG FILE INTO EXPECTED LIST
	[ ] // // hFile = FileOpen(sExpectedFilename,FM_READ)
	[+] // // while FileReadLine(hFile,sPath)
		[+] // // if StrPos(sPreviousPath,sPath)
			[ ] // // sPreviousPath = sPath
			[ ] // // continue
		[+] // // if sPath[1] == "["
			[ ] // // sPath = StrTran(sPath, "[", "")
			[ ] // // sPath = StrTran(sPath, "]", "")
			[+] // // for each sRegistryKey in sRegistryKeys
				[+] // // if sPath == sRegistryKey
					[ ] // // continue
			[ ] // // sPath = StrTran(sPath,"{sRegistryKey}\","")
			[ ] // // sPathNextLine = "Init"
			[+] // // while sPathNextLine != ""
				[ ] // // FileReadLine(hFile,sPathNextLine) 
				[+] // // if sPathNextLine[1] == '"'
					[ ] // // sItem = GetField(sPathNextLine,'"',2)
					[+] // // do
						[ ] // // sValue = SYS_GetRegistryValue (getRegistryKeyConst(sRegistryKey),sPath, sItem)
						[ ] // // REGISTRYENTRY.sKey = sRegistryKey
						[ ] // // REGISTRYENTRY.sPath = sPath
						[ ] // // REGISTRYENTRY.sItem = sItem
						[ ] // // REGISTRYENTRY.sValue = sValue
						[ ] // // ListAppend(laExpected, REGISTRYENTRY)
					[+] // // except
						[ ] // // LogError("An error occured while getting the value in the registry key: {REGISTRYENTRY} {ExceptData()}")
	[ ] // // FileClose(hFile)
	[ ] // // // COMPARE LISTS
	[ ] // // compareLists(laActual, laExpected)
	[ ] // 
[+] // compareLists(LIST OF ANYTYPE laActual, LIST OF ANYTYPE laExpected)
	[ ] // // LIST OF SERVICEREC lsActualServices = getServicesList()
	[ ] // // LIST OF SERVICEREC lsExpectedServices = getServicesFromExcel(sExcelWorksheetName)
	[ ] // STRING sItem
	[ ] // STRING sField
	[ ] // STRING iListCount 
	[ ] // STRING sActualEntry
	[ ] // STRING sExpectedEntry
	[ ] // INTEGER iActualIndex
	[ ] // INTEGER iExpectedIndex
	[ ] // INTEGER iFailed = 0
	[ ] // INTEGER iPassed = 0
	[ ] // BOOLEAN bExpectedItemFound = FALSE
	[ ] // BOOLEAN bActualItemFound = FALSE
	[ ] // LIST OF STRING lsFields
	[ ] // // Check that actual and expected types match
	[+] // do
		[ ] // verify(typeof(laActual), typeof(laExpected))
	[+] // except
		[ ] // LogError("Actual and Expected lists are different types")
	[+] // if laActual == laExpected
		[ ] // Print("Lists are identical")
		[ ] // return
	[+] // for iExpectedIndex = 1 to ListCount(laExpected) 
		[+] // for iActualIndex = 1 to ListCount(laActual)
			[+] // lsFields = FieldsOfRecord(REGISTRYENTRYREC)
				[+] // for each sField in lsFields
					[ ] // sExpectedEntry = sExpectedEntry + laExpected[iExpectedIndex].@sField
					[ ] // sActualEntry = sActualEntry + laActual[iActualIndex].@sField
				[+] // if sActualEntry != sExpectedEntry
					[ ] // Print("***** Expected and Actual fields from {typeof(laActual)} do not match:")
					[ ] // Print("               Actual value :",sActualEntry)
					[ ] // Print("               Expected value:",sExpectedEntry)
					[ ] // Print("")
					[ ] // addTestsFailedCount(1)
				[+] // else 
					[ ] // addTestsPassedCount(1)
	[+] // for iActualIndex = 1 to ListCount(laActual)
		[+] // for iExpectedIndex = 1 to ListCount(laExpected) 
			[+] // lsFields = FieldsOfRecord(REGISTRYENTRYREC)
				[+] // for each sField in lsFields
					[ ] // sExpectedEntry = sExpectedEntry + laExpected[iExpectedIndex].@sField
					[ ] // sActualEntry = sActualEntry + laActual[iActualIndex].@sField
				[+] // if sActualEntry != sExpectedEntry
					[ ] // LogError("***** Actual and Expected States do not match for Service:",laActual[iActualIndex].sDisplayName)
					[ ] // LogError("               Actual State:",laActual[iActualIndex].sState)
					[ ] // LogError("               Expected State:",laExpected[iExpectedIndex].sState)
					[ ] // Print("")
					[ ] // addTestsFailedCount(1)
				[+] // else 
					[ ] // addTestsPassedCount(1)
[ ] 
[+] testcase createRegistryEntries() appstate none
	[ ] createRegistryEntriesForExcel(HKEY_LOCAL_MACHINE,"SOFTWARE","c:\DC01.reg")
	[ ] // ListPrint(Reg_EnumValuesAll (HKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 0\Logical Unit Id 0"))
	[ ] 
[+] testcase PrintSingleRegistryKey() appstate none
	[ ] STRING sReg = SYS_GetRegistryValue(getRegistryKeyInt("HKEY_LOCAL_MACHINE"), "HARDWARE\ACPI\FACS\","00000000")
	[ ] Print(sReg)
	[ ] 
[+] testcase PrintEnumValues() appstate none
	[ ] STRING sRegPath = "HARDWARE\ACPI\DSDT\DELL\dt_ex\00001000"
	[ ] ANYTYPE sItem
	[ ] LIST OF ANYTYPE lsEnumValues = Reg_EnumValues(HKEY_LOCAL_MACHINE, sRegPath) 
	[+] for each sItem in lsEnumValues
		[ ] Print(sItem)
	[ ] 
[+] testcase PrintGetRegistryValue() appstate none
	[ ] // ListPrint(getRegistryEntries("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0"))
	[ ] INTEGER iKey = HKEY_LOCAL_MACHINE // defined in msw32.inc
	[ ] // INTEGER iKey = HKEY_CURRENT_CONFIG  // defined in msw32.inc
	[ ] // STRING sPath = "SOFTWARE\Microsoft\EventSystem\"{26c409cc-ae86-11d1-b616-00805fc79216}\EventClasses\"{FAF53CC4-BD73-4E36-83F1-2B23F46E513E}-"{00000000-0000-0000-0000-000000000000}-"{00000000-0000-0000-0000-000000000000}"
	[ ] STRING sPath = "SOFTWARE\Microsoft\EventSystem\"{26c409cc-ae86-11d1-b616-00805fc79216}\EventClasses\"{FAF53CC4-BD73-4E36-83F1-2B23F46E513E}-"{00000000-0000-0000-0000-000000000000}-"{00000000-0000-0000-0000-000000000000}"
	[ ] // STRING sPath = "Software\Fonts"
	[ ] // STRING sItem = "FIXEDFON.FON"
	[ ] STRING sItem = "Active"
	[ ] Print (SYS_GetRegistryValue (iKey, sPath, sItem))
	[ ] 
[+] // testcase PrintGetRegistryKeyNames() appstate none
	[ ] // // ListPrint(getRegistryEntries("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0"))
	[ ] // INTEGER iKey = HKEY_LOCAL_MACHINE // defined in msw32.inc
	[ ] // STRING sPath = "HARDWARE\DESCRIPTION\System\CentralProcessor"
	[ ] // Print (SYS_GetRegistryKeyNames (iKey, sPath))
	[ ] 
[ ] /////////////////////////////////////////////////////////////////////////////
[ ] // Print OS Information Test Cases
[+] testcase PrintOSAppVersions() appstate none
	[ ] STRING sOSInfo = getOSVersion()
	[ ] STRING sItem 
	[ ] Print("OS Version: ")
	[+] // for each sItem in lsOSInfo 
		[ ] Print(sOSInfo)
	[ ] Print("Outlook Version is "+getOutlookVersion())
	[ ] Print("Outlook Version is "+getOutlookBuild())
	[ ] Print("Word Version is "+getWordVersion())
	[ ] Print("Word Build is "+getWordBuild())
	[ ] Print("Excel Version is "+getExcelVersion())
	[ ] Print("Excel Build is "+getExcelBuild())
	[ ] Print("Access Version is "+getAccessVersion())
	[ ] Print("Access Build is "+getAccessBuild())
	[ ] Print("PowerPoint Version is "+getPowerPointVersion())
	[ ] Print("PowerPoint Build is "+getPowerPointBuild())
[+] testcase PrintAllTestInformation () appstate none
	[ ] LIST OF STRING lslPConfig =getIPConfig()
	[ ] STRING sItem 
	[ ] LIST OF STRING lsOSInfo = getOSInfo()
	[ ] Print("Operating System Version is ", getOSVersion())
	[ ] Print("Java Version is ",getJavaVersion("M:\Program Files\java"))
	[+] for each sItem in lslPConfig 
		[ ] Print(sItem)
	[ ] Print("Operating System Information: ")
	[+] for each sItem in lsOSInfo 
		[ ] Print(sItem)
	[ ] // Print("Machine Configuration Information is :")
	[ ] LIST OF STRING lsServices = getServices()
	[ ] Print("Service Information: ")
	[+] for each sItem in lsServices 
		[ ] Print(sItem)
	[ ] Print("Java Version is ",getJavaVersion("M:\Program Files\java"))
	[ ] 
[+] testcase PrintIAVAEntries() appstate none
	[ ] STRING sItem
	[+] LIST OF STRING lsIAVAPaths = {...}
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-017\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-018\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-025\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-027\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-029\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-030\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-039\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-040\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA05-052\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA06-001\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA06-002\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVA06-001\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVB05-010\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\IAVB05-011\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\MS05-019\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\MS05-019_01\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\MS05-019_02\IAVANumber"
		[ ] "HKLM\SOFTWARE\COMPOSE\Software\IAVA\MS05-019_03\IAVANumber"
	[+] for each sItem in lsIAVAPaths
		[ ] Print(getRegistryEntry(sItem))
		[ ] 
[ ] 
[+] testcase openMSOutlook()
	[ ] openOutlook()
	[ ] closeOutlook()
[+] testcase openSAV4ExchangeTC()
	[ ] openSAV4Exchange()
	[ ] // closeSAV4Exchange()
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[ ] // FUNCTIONS
[ ] ///////////////////////////////////////////////////////////////////////////////////////
[+] createWinZip (STRING sArchiveFileName, BOOLEAN bOverwiteMode) 
	[ ] // Function: WinZipArchive
	[ ] // Purpose: creates an archive and adds items to the archive
	[ ] // Description: This function creates an archive and adds the files listed in the sFileList Parameter
	[ ] // Parameters:
	[ ] // sArchiveFileName: 		The file name to create. 
	[ ] //							The default directory will be used unless the fully qualified path is passed. 
	[ ] // sFileList: 				List of the files to be added to the archive.
	[ ] //							The default directory will be used unless the fully qualified path is passed. 
	[ ] STRING sFileItem
	[ ] BOOLEAN bExists
	[ ] INTEGER i
	[ ] LIST OF STRING lsFileList = GetFilenamesFromExcel("WinZip")
	[ ] 
	[ ] WinZip.SetActive ()
	[ ] WinZip.ReBarWindow321.ToolBar1.New.Click ()
	[ ] NewArchive.SetActive ()
	[ ] STRING sCurrentDir
	[ ] // sCurrentDir = SYS_GetDir ()
	[+] if SYS_FileExists (sArchiveFileName)
		[+] for i = 1 to 1000
			[+] if SYS_FileExists (sArchiveFileName) 
				[ ] sArchiveFileName = Stuff(sArchiveFileName,len(sArchiveFileName)-3, 0, "{i}")
			[+] else
				[ ] break
	[ ] NewArchive.FileName.SetText (sArchiveFileName)
	[ ] NewArchive.OK.Click ()
	[+] if MessageBox.Exists()
		[ ] MessageBox.No.Click ()
		[ ] LogError("File: "+sArchiveFileName+" Already Exists and was not created")
		[ ] return
	[ ] sleep(3)
	[+] for each sFileItem in lsFileList
		[ ] sleep(3)
		[+] if !Add.Exists ()
			[ ] WinZip.ReBarWindow321.ToolBar1.Add.Click ()
		[ ] Add.SetActive ()
		[ ] Add.FileName.SetText (sFileItem)
		[ ] Add.Add.Click ()
	[+] if MessageBox.Exists()
		[ ] MessageBox.No.Click ()
		[ ] LogError("Error occured while adding the file: "+sFileItem)
		[ ] return
		[ ] 
[+] createWinZipArchive (STRING sArchiveFileName, LIST OF STRING sFileList, BOOLEAN bOverwiteMode) 
	[ ] // Function: WinZipArchive
	[ ] // Purpose: creates an archive and adds items to the archive
	[ ] // Description: This function creates an archive and adds the files listed in the sFileList Parameter
	[ ] // Parameters:
	[ ] // sArchiveFileName: 		The file name to create. 
	[ ] //							The default directory will be used unless the fully qualified path is passed. 
	[ ] // sFileList: 				List of the files to be added to the archive.
	[ ] //							The default directory will be used unless the fully qualified path is passed. 
	[ ] STRING sFileItem
	[ ] BOOLEAN bExists
	[ ] INTEGER i
	[ ] 
	[ ] WinZip.SetActive ()
	[ ] WinZip.ReBarWindow321.ToolBar1.New.Click ()
	[ ] NewArchive.SetActive ()
	[ ] STRING sCurrentDir
	[ ] // sCurrentDir = SYS_GetDir ()
	[+] if SYS_FileExists (sArchiveFileName)
		[+] for i = 1 to 100
			[+] if SYS_FileExists (sArchiveFileName) 
				[ ] sArchiveFileName = Stuff(sArchiveFileName,len(sArchiveFileName)-3, 0, "{i}")
			[+] else
				[ ] break
	[ ] NewArchive.FileName.SetText (sArchiveFileName)
	[ ] NewArchive.OK.Click ()
	[+] if MessageBox.Exists()
		[ ] MessageBox.No.Click ()
		[ ] LogError("File: "+sArchiveFileName+" Already Exists and was not created")
		[ ] return
	[ ] sleep(3)
	[+] for each sFileItem in sFileList
		[ ] sleep(3)
		[+] if !Add.Exists ()
			[ ] WinZip.ReBarWindow321.ToolBar1.Add.Click ()
		[ ] Add.SetActive ()
		[ ] Add.FileName.SetText (sFileItem)
		[ ] Add.Add.Click ()
	[+] if MessageBox.Exists()
		[ ] MessageBox.No.Click ()
		[ ] LogError("Error occured while adding the file: "+sFileItem)
		[ ] return
		[ ] 
[ ] // Microsoft Office getVersion functions
[+] STRING getAdobeReaderVersion() 
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getAcrobatVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of Acrobat
	[ ] //								from the about box
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] STRING sVersion
	[ ] // AdobeReader.SetActive ()
	[+] if Upgrade.Exists()
		[ ] Upgrade.GroupBox1.No.Click()
	[ ] AdobeReader.Help.About.Pick ()
	[ ] AboutAdobeReader.SetActive ()
	[ ] sVersion = AboutAdobeReader.Version.GetCaption()
	[ ] AboutAdobeReader.Click (1, 387, 175)
	[ ] return(Left(sVersion,13))
	[ ] 
	[ ] // AboutAdobeReader70.Version70012142004.Click ()
	[ ] // AdobeReader.SetActive ()
	[ ] // AdobeReader.Help.MenuItem3.Pick ()
	[ ] // AboutAdobeReader70.SetActive ()
	[ ] // AboutAdobeReader70.Version70012142004.Click ()
	[ ] 
[+] STRING getWinZipVersion() 
	[ ] STRING sVersion
	[ ] WinZip.Help.AboutWinZip.Pick ()
	[ ] AboutWinZip.SetActive ()
	[ ] sVersion = AboutWinZip.WinZipVersion.GetText()
	[ ] AboutWinZip.OK.Click ()
	[ ] return(sVersion)
	[ ] 
[+] STRING getOutlookVersion()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getOutlookVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of 
	[ ] //								outlook
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getOutlookVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(Left(sItem,4))
		[ ] 
[+] STRING getWordVersion()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getWordVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of 
	[ ] //								word
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getWordVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[+] STRING getExcelVersion()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getExcelVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of 
	[ ] //								excel
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getExcelVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[+] STRING getAccessVersion()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getAccessVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of 
	[ ] //								access
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getAccessVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[+] STRING getPowerPointVersion()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getPowerPointVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of 
	[ ] //								powerpoint
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getPowerPointVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[ ] // Microsoft Office getBuild functions
[+] STRING getOutlookBuild()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getOutlookBuild
	[ ] // Purpose:                                  This function returns 
	[ ] //								the Build of 
	[ ] //								outlook
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getOutlookVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] // return(sItem)
		[ ] return(Right(sItem,4))
		[ ] 
[+] STRING getWordBuild()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getWordBuild
	[ ] // Purpose:                                  This function returns 
	[ ] //								the Build of 
	[ ] //								word
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getWordBuild.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(Right(sItem,4))
		[ ] 
[+] STRING getExcelBuild()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getExcelBuild
	[ ] // Purpose:                                  This function returns 
	[ ] //								the Build of 
	[ ] //								excel
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getExcelBuild.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[+] STRING getAccessBuild()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getAccessBuild
	[ ] // Purpose:                                  This function returns 
	[ ] //								the Build of 
	[ ] //								access
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getAccessBuild.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[+] STRING getPowerPointBuild()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getPowerPointBuild
	[ ] // Purpose:                                  This function returns 
	[ ] //								the Build of 
	[ ] //								powerpoint
	[ ] //
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getPowerPointBuild.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[ ] // Microsoft Office functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Outlook Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] STRING GetMailCaption()
	[ ] return(sMailSubject+" - Message")
	[ ] 
[+] // createVBSContact(STRING sFullName, STRING sEmailAddress)
	[ ] // /////////////////////////////////////////////////////////////////////////////
	[ ] // // Author:                                   	John Connolly  
	[ ] // // Function Name:                  	createContact
	[ ] // // Purpose:                                  This function creates a  
	[ ] // //								contact in Outlook
	[ ] // //								The itemtype is hardcoded to 1
	[ ] // //								in the VB Script
	[ ] // // Inputs:                                   	STRING Full Name
	[ ] // //								STRING Email Address
	[ ] // /////////////////////////////////////////////////////////////////////////////
	[ ] // LIST OF STRING sReturn
	[ ] // STRING sItem
	[ ] // STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\createContact.vbs "+"""{sFullName}"""+" "+"""{sEmailAddress}"""
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] // SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] // Print("sReturn = "+sItem)
	[ ] // 
[+] // createVBSTask(STRING sSubject,  STRING sAttachments)
	[ ] // /////////////////////////////////////////////////////////////////////////////
	[ ] // // Author:                                   	John Connolly  
	[ ] // // Function Name:                  	createContact
	[ ] // // Purpose:                                  This function creates a  
	[ ] // //								contact in Outlook
	[ ] // //								The itemtype is hardcoded to 1
	[ ] // //								in the VB Script
	[ ] // // Inputs:                                   	STRING Full Name
	[ ] // //								STRING Email Address
	[ ] // /////////////////////////////////////////////////////////////////////////////
	[ ] // LIST OF STRING sReturn
	[ ] // STRING sItem
	[ ] // STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\createTaskItem.vbs "+"""{sSubject}"""+" "+"""{sAttachments}"""
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] // SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] // Print("sReturn = "+sItem)
		[ ] // 
[ ] // Mail Message Functions
[+] createMail(STRING sSendto, STRING sCc, STRING sSubject, STRING sAttachments optional)
	[ ] STRING sAttachment 
	[ ] LIST OF STRING lsSendto, lsCc, lsSubject
	[ ] INTEGER i = 1
	[ ] ListAppend(lsSendto, sSendto)
	[ ] ListAppend(lsCc, sCc)
	[ ] ListAppend(lsSubject, sSubject)
	[ ] sleep(5)
	[ ] MicrosoftOutlook.SetActive ()
	[ ] // MicrosoftOutlook.MsoDockTop.Standard.TypeKeys("<Ctrl-N>")
	[ ] MicrosoftOutlook.MsoDockTop.Standard.PressKeys("<Alt>")
	[ ] MicrosoftOutlook.MsoDockTop.Standard.TypeKeys("N")
	[ ] MicrosoftOutlook.MsoDockTop.Standard.ReleaseKeys("<Alt>")
	[ ] MicrosoftOutlook.MsoDockTop.Standard.TypeKeys("M")
	[ ] Message.SetActive ()
	[ ] Message.MsoDockTop.Envelope.DialogBox1.Click(1,93,40)
	[+] if Reminder.Exists()
		[ ] Reminder.DismissAll.Click()
	[ ] Clipboard.SetText(lsSendto)
	[ ] // Message.MsoDockTop.MenuBar.TypeKeys("<Ctrl-V>")
	[ ] // Message.MsoDockTop.MenuBar.Click (1, 50, 10)
	[ ] Message.TypeKeys("<Ctrl-V>")
	[ ] // Message.MsoDockTop.MenuBar.Click (1, 50, 100)
	[ ] Message.MsoDockTop.Envelope.DialogBox1.Click(1,93,65)
	[ ] Clipboard.SetText(lsCc)
	[ ] Message.TypeKeys("<Ctrl-V>")
	[ ] Message.MsoDockTop.Envelope.DialogBox1.Click(1,93,90)
	[ ] Clipboard.SetText(lsSubject)
	[ ] Message.TypeKeys("<Ctrl-V>")
	[+] if sAttachments != NULL
		[ ] sAttachment = GetField(sAttachments ,",",1)
		[+] while sAttachment != ""
			[ ] sAttachment = GetField(sAttachments ,",", i)
			[+] if sAttachment == ""
				[ ] break
			[ ] i ++
			[ ] Message.SetActive ()
			[ ] // Message.MsoDockTop.MenuBar.PressMouse (1, 139, 16)
			[ ] Message.MsoDockTop.MenuBar.PressKeys("<Alt>")
			[ ] Message.MsoDockTop.MenuBar.TypeKeys("I")
			[ ] Message.MsoDockTop.MenuBar.ReleaseKeys("<Alt>")
			[ ] Message.MsoDockTop.MenuBar.TypeKeys("L")
			[ ] InsertFile.SetActive ()
			[ ] sleep(1)
			[ ] InsertFile.Filename.TypeKeys (sGlobalReadPath+sAttachment)
			[ ] sleep(1)
			[ ] InsertFile.TypeKeys("<Enter>")
			[ ] // InsertFile.Click (1, 541, 325)
	[ ] MicrosoftOutlook.SetActive ()
[+] replyMail(INTEGER iItemID optional)
	[ ] STRING sAttachment 
	[ ] INTEGER i = 1
	[ ] // MicrosoftOutlook.SetActive()
	[ ] // // Select the most recent message 
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[ ] // if iItemID is 2 or more, select a subsequent mail message
	[+] if iItemID !=NULL
		[+] for i = 2 to iItemID
			[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] // Open the active message
	[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
	[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
	[ ] // Press the "Reply" button
	[ ] Message.SetActive()
	[ ] Message.PressKeys("<Alt>")
	[ ] Message.TypeKeys("<R>")
	[ ] Message.ReleaseKeys("<Alt>")
	[ ] // Press the "Send" button on the reply window
	[ ] ReplyMessage.SetActive()
	[ ] // ReplyMessage.PressKeys("<Alt>")
	[ ] // ReplyMessage.TypeKeys("S")
	[ ] // ReplyMessage.ReleaseKeys("<Alt>")
	[ ] ReplyMessage.MsoDockTop.Envelope.DialogBox1.MSOGenericControlContainer.Click(1,29,7)
	[ ] // Close the original message
	[ ] OpenMessage.PressKeys("<Alt>")
	[ ] OpenMessage.TypeKeys("<F4>")
	[ ] OpenMessage.ReleaseKeys("<Alt>")
	[ ] 
[+] replyAllMail(INTEGER iItemID optional)
	[ ] STRING sAttachment 
	[ ] INTEGER i = 1
	[ ] // MicrosoftOutlook.SetActive()
	[ ] // Select the most recent message 
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[ ] // if iItemID is 2 or more, select a subsequent mail message
	[+] if iItemID !=NULL
		[+] for i = 2 to iItemID
			[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] // Open the active message
	[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
	[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
	[ ] // Press the "Reply to All" button
	[ ] Message.SetActive()
	[ ] Message.PressKeys("<Alt>")
	[ ] Message.TypeKeys("<L>")
	[ ] Message.ReleaseKeys("<Alt>")
	[ ] // Press the "Send" button on the reply window
	[ ] ReplyMessage.SetActive()
	[ ] // ReplyMessage.PressKeys("<Alt>")
	[ ] // ReplyMessage.TypeKeys("S")
	[ ] // ReplyMessage.ReleaseKeys("<Alt>")
	[ ] ReplyMessage.MsoDockTop.Envelope.DialogBox1.MSOGenericControlContainer.Click(1,29,7)
	[ ] // Close the original message
	[ ] OpenMessage.PressKeys("<Alt>")
	[ ] OpenMessage.TypeKeys("<F4>")
	[ ] OpenMessage.ReleaseKeys("<Alt>")
	[ ] 
[+] verifySubject(STRING sExpectedValue, INTEGER iItemID optional)
	[ ] INTEGER i
	[ ] MicrosoftOfficeOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[+] if iItemID !=NULL
		[+] for i = 2 to iItemID
			[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
	[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
	[ ] verify(ReceivedMessage.AfxWndW1.DialogBox1.Subject.GetText(), sExpectedValue)
	[ ] ReceivedMessage.PressKeys("<Alt>")
	[ ] ReceivedMessage.TypeKeys("<F4>")
	[ ] ReceivedMessage.ReleaseKeys("<Alt>")
	[+] if OutlookSaveChanges.Exists()
		[ ] OutlookSaveChanges.PressKeys("<Alt>")
		[ ] OutlookSaveChanges.TypeKeys("N")
		[ ] OutlookSaveChanges.ReleaseKeys("<Alt>")
	[ ] 
[ ] 
[+] testcase testOutlookSendMail()
	[ ] outlookSendMail("Silka")
	[ ] 
[+] outlookSendMail(STRING sUser)
	[ ] LIST OF SENDMAILREC LSENDMAIL
	[ ] INTEGER index = 1
	[ ] // ListPrint(GetSendMailRecsFromExcel())
	[ ] LSENDMAIL = GetSendMailRecsFromExcel(sUser)
	[ ] openoutlook()
	[+] while index <= ListCount(LSENDMAIL)
		[+] if LSENDMAIL[index].sAttachment == NULL
			[ ] createMail(LSENDMAIL[index].sTo,LSENDMAIL[index].sCc,LSENDMAIL[index].sSubject)
		[+] else
			[ ] createMail(LSENDMAIL[index].sTo,LSENDMAIL[index].sCc,LSENDMAIL[index].sSubject,LSENDMAIL[index].sAttachment)
		[ ] sendMail()
		[ ] index++
		[ ] // Print(LSENDMAIL[2].sAttachment)
		[ ] // Print(LSENDMAIL[3].sAttachment)
		[ ] // Print(LSENDMAIL[4].sAttachment)
	[ ] closeoutlook()
[ ] 
[+] testcase tc_outlookVerifyMailSubject() appstate none
	[ ] outlookVerifyMailSubject("Silka")
[+] outlookVerifyMailSubject(STRING sToUser)
	[ ] INTEGER i, index, iItemID
	[ ] LIST OF SENDMAILREC SENDMAIL
	[ ] SENDMAIL = GetSendMailRecsFromExcel(sToUser)
	[ ] // ListPrint(SENDMAIL)
	[ ] openoutlook()
	[+] for index = 1 to ListCount(SENDMAIL)
		[ ] MicrosoftOfficeOutlook.SetActive()
		[ ] // Go to top item in email list
		[ ] // use a function to find the inbox item ID
		[ ] LogWarning("Verifying email to user {sToUser} with subject {SENDMAIL[index].sSubject}")
		[ ] iItemID = findInboxMailItemID(SENDMAIL[index].sSubject)
		[ ] // Print("Subject: {SENDMAIL[index].sSubject} Item: {iItemID}")
		[+] if iItemID == 0
			[ ] Print("The email with Subject = {SENDMAIL[index].sSubject} was not found in the inbox")
[ ] 
[+] testcase tc_outlookVerifyMailCC() appstate none
	[ ] outlookVerifyMailSubjectCC("Silkd")
	[ ] 
[+] outlookVerifyMailSubjectCC(STRING sToUser)
	[ ] INTEGER i, index, iItemID
	[ ] LIST OF SENDMAILREC SENDMAIL
	[ ] SENDMAIL = GetCCMailRecsFromExcel(sToUser)
	[ ] // ListPrint(SENDMAIL)
	[ ] openoutlook()
	[+] for index = 1 to ListCount(SENDMAIL)
		[ ] MicrosoftOfficeOutlook.SetActive()
		[ ] // Go to top item in email list
		[ ] // use a function to find the inbox item ID
		[ ] LogWarning("Verifying email to user {sToUser} with subject {SENDMAIL[index].sSubject}")
		[ ] iItemID = findInboxMailItemID(SENDMAIL[index].sSubject)
		[ ] // Print("Subject: {SENDMAIL[index].sSubject} Item: {iItemID}")
		[+] if iItemID == 0
			[ ] Print("The email with Subject = {SENDMAIL[index].sSubject} was not found in the inbox")
[ ] 
[+] verifySubjectMeeting(STRING sExpectedValue, INTEGER iItemID optional)
	[ ] INTEGER i
	[ ] MicrosoftOfficeOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[+] if iItemID !=NULL
		[+] for i = 2 to iItemID
			[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
	[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
	[ ] Verify(ReceivedMeeting.AfxWndW1.DialogBox1.Subject.GetText(), sExpectedValue)
	[ ] ReceivedMessage.PressKeys("<Alt>")
	[ ] ReceivedMessage.TypeKeys("<F4>")
	[ ] ReceivedMessage.ReleaseKeys("<Alt>")
	[+] if OutlookSaveChanges.Exists()
		[ ] OutlookSaveChanges.PressKeys("<Alt>")
		[ ] OutlookSaveChanges.TypeKeys("N")
		[ ] OutlookSaveChanges.ReleaseKeys("<Alt>")
	[ ] 
[+] verifyAttachment(INTEGER iItemID)
	[ ] INTEGER i
	[ ] STRING sFilename
	[ ] MicrosoftOfficeOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[+] if iItemID !=NULL
		[+] for i = 2 to iItemID
			[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
	[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
	[ ] Message.SetActive ()
	[ ] Message.PressKeys("<Alt>")
	[ ] Message.TypeKeys("<f>")
	[ ] Message.ReleaseKeys("<Alt>")
	[ ] Message.TypeKeys("<n>")
	[ ] SaveAttachment.SetActive()
	[ ] sFilename = SaveAttachment.Filename.GetText()
	[ ] // Print(SaveAttachment.Filename.GetText())
	[ ] SaveAttachment.Save.Click()
	[ ] sleep(1)
	[+] if replaceExistingFile.Exists()
		[ ] replaceExistingFile.Yes.Click()
	[ ] ReceivedMessage.PressKeys("<Alt>")
	[ ] ReceivedMessage.TypeKeys("<F4>")
	[ ] ReceivedMessage.ReleaseKeys("<Alt>")
	[+] if OutlookSaveChanges.Exists()
		[ ] OutlookSaveChanges.PressKeys("<Alt>")
		[ ] OutlookSaveChanges.TypeKeys("N")
		[ ] OutlookSaveChanges.ReleaseKeys("<Alt>")
	[ ] // Compare Files - We will verify that the file has CHANGED
	[+] if !SYS_CompareBinary(sGlobalReadPath+sFilename,sGlobalWritePath+sFilename)
		[ ] Print("The files "+sGlobalReadPath+sFilename+" and "+sGlobalWritePath+sFilename+" are different and should be different")
	[ ] 
[+] extractAttachment()
	[ ] STRING sSaveFileName
	[ ] ////////////////////////////////////////////////////////////////////////////////////////////
	[ ] // This function opens the first item in the inbox and 
	[ ] // extracts the attachment
	[ ] // It assumes imte number in the inbox of the attachment to be saved 
	[ ] // has already been selected
	[ ] //////////////////////////////////////////////////////////////////////////////////////////
	[ ] MicrosoftOutlook.SetActive ()
	[ ] MicrosoftOutlook.PressKeys("<Ctrl>")
	[ ] MicrosoftOutlook.TypeKeys("<o>")
	[ ] MicrosoftOutlook.ReleaseKeys("<Ctrl>")
	[ ] Message.SetActive ()
	[ ] Message.PressKeys("<Alt>")
	[ ] Message.TypeKeys("<f>")
	[ ] Message.ReleaseKeys("<Alt>")
	[ ] Message.TypeKeys("<n>")
	[ ] SaveAttachment.SetActive()
	[ ] SaveAttachment.Save.Click()
	[ ] sleep(1)
	[+] if replaceExistingFile.Exists()
		[ ] replaceExistingFile.Yes.Click()
	[ ] Message.PressKeys("<Alt>")
	[ ] Message.TypeKeys("<f4>")
	[ ] Message.ReleaseKeys("<Alt>")
[+] sendMail()
	[ ] // Click on the send button
	[ ] MessageDock.SetActive()
	[ ] MessageDock.MsoDockTop.Envelope.DialogBox1.Click (1, 20, 10)
	[ ] // Send Receive button Keyboard shortcut is F9
	[ ] MicrosoftOutlook.SetActive()
	[ ] MicrosoftOutlook.MsoDockTop.Standard.TypeKeys("<F9>")
	[ ] sleep(1)
[+] INTEGER findInboxMailItemID(STRING sSubject)
	[ ] ///////////////////////////////////////////////////////////////////////////
	[ ] // This function finds the first occurence of 
	[ ] // the item passed in the subject line
	[ ] // of the inbox of outlook
	[ ] //////////////////////////////////////////////////////////////////////////
	[ ] INTEGER inboxSize = 20
	[ ] INTEGER i
	[ ] sleep(10)
	[ ] MicrosoftOfficeOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[+] for i = 1 to 20
		[ ] // Wait for in the inbox to update before opening the selected message
		[ ] // Open the active message
		[ ] MicrosoftOfficeOutlook.TypeKeys("<Alt-F>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
		[+] if ReceivedMessage.AfxWndW1.DialogBox1.Subject.GetText() == sSubject
			[ ] ReceivedMessage.TypeKeys("<Alt-F4>")
			[ ] return( i )
		[ ] ReceivedMessage.TypeKeys("<Alt-F4>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] return(0)
[+] INTEGER findInboxMeetingItemID(STRING sSubject)
	[ ] ///////////////////////////////////////////////////////////////////////////
	[ ] // This function finds the first occurence of 
	[ ] // the item passed in the subject line
	[ ] // of the inbox of outlook
	[ ] //////////////////////////////////////////////////////////////////////////
	[ ] INTEGER inboxSize = 20
	[ ] INTEGER i
	[ ] MicrosoftOfficeOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[+] for i = 1 to 20
		[ ] // Open the active message
		[ ] MicrosoftOfficeOutlook.PressKeys("<Alt>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<F>")
		[ ] MicrosoftOfficeOutlook.ReleaseKeys("<Alt>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<O>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<S>")
		[ ] // ReceivedMeeting.AfxWndW1.DialogBox1.Subject.GetText()
		[+] if ReceivedMeeting.AfxWndW1.DialogBox1.Subject.GetText() == sSubject
			[ ] ReceivedMeeting.TypeKeys("<Alt-F4>")
			[ ] return( i )
		[ ] ReceivedMeeting.TypeKeys("<Alt-F4>")
		[ ] MicrosoftOfficeOutlook.TypeKeys("<Down>")
	[ ] return(0)
[+] testcase clearInbox() appstate Outlook
	[ ] INT iNumItems = 100
	[ ] MicrosoftOutlook.SetActive()
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Home>")
	[ ] MicrosoftOfficeOutlook.TypeKeys("<Delete {iNumItems}>")
[ ] 
[+] closeMailSaveNo()
	[ ] MessageDock.PressKeys("<alt>")
	[ ] MessageDock.TypeKeys("<F>")
	[ ] MessageDock.ReleaseKeys("<alt>")
	[ ] MessageDock.TypeKeys("<C>")
	[ ] ConfirmSave.No.Click()
	[ ] 
[+] closeMailSaveYes()
	[ ] MessageDock.PressKeys("<alt>")
	[ ] MessageDock.TypeKeys("<F>")
	[ ] MessageDock.ReleaseKeys("<alt>")
	[ ] MessageDock.TypeKeys("<C>")
	[ ] ConfirmSave.Yes.Click()
	[ ] 
	[ ] 
[ ] // Appointment Functions
[+] createAppointment(STRING sAttendees, STRING sAppointmentSubject, STRING sLocation,  STRING sBody, STRING sAttachments)
	[ ] STRING sAttachment
	[ ] INTEGER i = 1
	[ ] LIST OF STRING lsAttendees, lsAppointmentSubject, lsLocation, lsBody, lsStartDate
	[ ] DATETIME sToday = GetDateTime()
	[ ] DATETIME sTomorrow = AddDateTime(sToday,730)
	[ ] STRING sStartDate = FormatDateTime (sTomorrow, "mm/dd/yy")
	[ ] ListAppend(lsAttendees, sAttendees+";")
	[ ] ListAppend(lsAppointmentSubject, sAppointmentSubject)
	[ ] ListAppend(lsLocation, sLocation)
	[ ] ListAppend(lsBody, sBody)
	[ ] ListAppend(lsStartDate, sStartDate)
	[ ] 
	[ ] // Click on New button, Appointment Menu Item
	[ ] MicrosoftOutlook.SetActive()
	[ ] MicrosoftOutlook.MsoDockTop.Standard.PressKeys("<Alt>")
	[ ] MicrosoftOutlook.MsoDockTop.Standard.TypeKeys("N")
	[ ] MicrosoftOutlook.MsoDockTop.Standard.ReleaseKeys("<Alt>")
	[ ] MicrosoftOutlook.MsoDockTop.TypeKeys("Q")
	[ ] Meeting.SetActive ()
	[ ] // Meeting.Maximize ()
	[ ] Meeting.AfxWndW1.Appointment.To.TypeKeys (sAttendees+"<Tab>")
	[ ] Meeting.AfxWndW1.Appointment.Click(1,85,140)
	[+] if Reminder.Exists()
		[ ] Reminder.DismissAll.Click()
	[ ] Clipboard.SetText(lsAttendees)
	[ ] Meeting.SetActive ()
	[ ] Meeting.TypeKeys("<Alt-E>")
	[ ] Meeting.TypeKeys("P")
	[ ] Meeting.TypeKeys("<Tab>")
	[ ] // Meeting.AfxWndW1.Appointment.Subject.TypeKeys (sAppointmentSubject+"<Tab>")
	[ ] Meeting.AfxWndW1.Appointment.Click(1,85,164)
	[ ] Clipboard.SetText(lsAppointmentSubject)
	[ ] Meeting.TypeKeys("<Ctrl-v>")
	[ ] Meeting.AfxWndW1.Appointment.Click(1,85,188)
	[ ] Meeting.AfxWndW1.Appointment.Location.TypeKeys (sLocation+"<Tab 2>")
	[ ] // Meeting.AfxWndW1.Appointment.Click(1,85,188)
	[ ] // Clipboard.SetText(lsLocation)
	[ ] // Meeting.TypeKeys("<Ctrl-v>")
	[ ] Meeting.AfxWndW1.Appointment.Click(1,102,230)
	[ ] Meeting.AfxWndW1.Appointment.StartDate.TypeKeys (sStartDate+"<Tab 9>")
	[ ] // Meeting.AfxWndW1.Appointment.Click(1,102,230)
	[ ] // Clipboard.SetText(lsStartDate)
	[ ] // Meeting.TypeKeys("<Ctrl-v>")
	[ ] Meeting.AfxWndW1.Appointment.Click(1,20,380)
	[ ] Meeting.AfxWndW1.Appointment.ShowTimeAs3.Body.TypeKeys (sBody)
	[ ] // Meeting.AfxWndW1.Appointment.Click(1,20,380)
	[ ] // Clipboard.SetText(lsBody)
	[ ] // Meeting.TypeKeys("<Ctrl-v>")
	[ ] 
	[ ] sAttachment = GetField(sAttachments ,",",1)
	[+] while sAttachment != ""
		[ ] sAttachment = GetField(sAttachments ,",", i)
		[+] if sAttachment == ""
			[ ] break
		[ ] i ++
		[ ] Meeting.MsoDockTop.MenuBar.PressKeys("<Alt>")
		[ ] sleep(1)
		[ ] Meeting.MsoDockTop.MenuBar.TypeKeys("I")
		[ ] sleep(1)
		[ ] Meeting.MsoDockTop.MenuBar.ReleaseKeys("<Alt>")
		[ ] sleep(1)
		[ ] Meeting.MsoDockTop.MenuBar.TypeKeys("F")
		[ ] sleep(1)
		[ ] InsertFileMeeting.SetActive ()
		[ ] InsertFileMeeting.Filename.TypeKeys (sGlobalReadPath+sAttachment)
		[ ] InsertFileMeeting.Click (1, 541, 325)
	[ ] // Choose Send button
	[ ] Meeting.PressKeys("<Alt>")
	[ ] Meeting.TypeKeys("S")
	[ ] Meeting.ReleaseKeys("<Alt>")
[+] openOutlook() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sCmdLine = sGlobalReadPath+"Outlook.bat"
	[ ] STRING sWorkingDir = sGlobalReadPath
	[+] if !MicrosoftOutlook.Exists()
		[ ] // SYS_Execute(sCmdLine,sReturn)
		[ ] RunStartMenu("outlook")
	[ ] // Print(sCmdLine)
	[ ] // Print(sWorkingDir)
	[ ] // MicrosoftOutlook.Maximize()
[+] closeOutlook()
	[ ] MicrosoftOutlook.Close()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Word Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] openWord() 
	[ ] STRING sCmdLine = '"C:\Program Files\Microsoft Office\OFFICE11\WINWORD.EXE"'
	[ ] STRING sWorkingDir = "{sGlobalReadPath}"
	[+] if !(Word.Exists ())
		[ ] Word.Start(sCmdLine,sWorkingDir,"",20)
		[ ] // RunStartMenu("WINWORD")
	[ ] // Print(sCmdLine)
	[ ] // Print(sWorkingDir)
	[ ] MicrosoftWord.sTag = "*Microsoft Word" 
[+] openWordDoc(STRING sFileName)
	[ ] MicrosoftWord.sTag = "*Microsoft Word"
	[+] if SYS_FileExists(sFileName)
		[ ] MicrosoftWord.SetActive()
		[ ] MicrosoftWord.PressKeys("<ctrl>")
		[ ] MicrosoftWord.TypeKeys("<o>")
		[ ] MicrosoftWord.ReleaseKeys("<ctrl>")
		[ ] sleep(1)
		[ ] WordOpen.Filename.TypeKeys (sFileName)
		[ ] WordOpen.Filename.TypeKeys ("<Enter>")
	[+] else
		[ ] Print("The File ",sFileName," does not exist")
[+] openWordDocs()
	[ ] LIST OF STRING lsFilenames = GetFilenamesFromExcel("Word")
	[ ] STRING sFilename, sFilenameTag
	[ ] MicrosoftWord.sTag = "*Microsoft Word"
	[ ] openword()
	[+] for each sFilename in lsFilenames
		[+] if SYS_FileExists(sFilename)
			[ ] MicrosoftWord.SetActive()
			[ ] // MicrosoftWord.TypeKeys("<Ctrl-O>")
			[ ] // MicrosoftWord.MsoDockTop.MenuBar.PressKeys("<ctrl-o>")
			[ ] // MicrosoftWord.MsoDockTop.MenuBar.TypeKeys("<Ctrl-O>")
			[ ] MicrosoftWord.MsoDockTop.MenuBar.TypeKeys("<Alt-F>")
			[ ] MicrosoftWord.MsoDockTop.MenuBar.TypeKeys("<O>")
			[ ] sleep(2)
			[ ] WordOpen.SetActive()
			[ ] WordOpen.Filename.TypeKeys (sFilename)
			[ ] WordOpen.Filename.TypeKeys ("<Enter>")
			[ ] sFilenameTag = GetField(sFilename,"\",3)
			[ ] // sFilenameTag = GetField(sFilenameTag,".",1)
			[ ] MicrosoftWord.sTag = sFilenameTag + " - Microsoft Word"
			[ ] // MicrosoftWord.SetActive()
			[ ] MicrosoftWord.TypeKeys("<ctrl-W>")
			[ ] Print("Word file "+sFilename+" successfully opened")
			[ ] addTestsPassedCount(1)
		[+] else
			[ ] Print("The File ",sFilename," does not exist")
			[ ] ListAppend(lsFailedFileOpenTests, sFilename)
			[ ] iTestsFailedCount ++
	[ ] closeword()
[+] closeWord()
	[ ] MicrosoftWord.SetActive()
	[ ] MicrosoftWord.PressKeys("<alt>")
	[ ] MicrosoftWord.TypeKeys("<f>")
	[ ] MicrosoftWord.ReleaseKeys("<alt>")
	[ ] MicrosoftWord.TypeKeys("<x>")
	[+] if WordSaveChanges.Exists()
		[ ] WordSaveChanges.No.Click()
	[ ] 
[+] testcase tc_openWordDocs() appstate none
	[ ] openWordDocs()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Excel Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] openExcel() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sCmdLine = '"C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE"'
	[ ] STRING sWorkingDir = "{sGlobalReadPath}"+""
	[+] if !(Excel.Exists ())
		[ ] Excel.Start(sCmdLine,sWorkingDir,"",20)
	[ ] // Print(sCmdLine)
	[ ] // Print(sWorkingDir)
	[ ] MicrosoftExcel.sTag = "Microsoft Excel*"
	[ ] // MicrosoftExcel.Maximize()
[+] openExcelDoc(STRING sFileName)
	[+] if SYS_FileExists(sFileName)
		[ ] MicrosoftExcel.sTag = "Microsoft Excel*"
		[ ] MicrosoftExcel.SetActive()
		[ ] MicrosoftExcel.PressKeys("<ctrl>")
		[ ] MicrosoftExcel.TypeKeys("<o>")
		[ ] MicrosoftExcel.ReleaseKeys("<ctrl>")
		[ ] sleep(1)
		[ ] ExcelOpen.Filename.TypeKeys (sFileName)
		[ ] ExcelOpen.Filename.TypeKeys ("<Enter>")
	[+] else
		[ ] Print("The File ",sFileName," does not exist")
[+] openExcelDocs()
	[ ] LIST OF STRING lsFilenames = GetFilenamesFromExcel("Excel")
	[ ] STRING sFilename, sFilenameTag
	[ ] MicrosoftExcel.sTag = "Microsoft Excel*"
	[ ] openExcel()
	[+] for each sFilename in lsFilenames
		[+] if SYS_FileExists(sFilename)
			[ ] MicrosoftExcel.SetActive()
			[ ] MicrosoftExcel.PressKeys("<ctrl>")
			[ ] MicrosoftExcel.TypeKeys("<o>")
			[ ] MicrosoftExcel.ReleaseKeys("<ctrl>")
			[ ] sleep(1)
			[ ] ExcelOpen.Filename.TypeKeys (sFilename)
			[ ] ExcelOpen.Filename.TypeKeys ("<Enter>")
			[ ] sFilenameTag = GetField(sFilename,"\",3)
			[ ] sFilenameTag = GetField(sFilenameTag,".",1)
			[ ] MicrosoftWord.sTag = "Microsoft Excel - " + sFilenameTag
			[ ] // MicrosoftExcel.TypeKeys("<ctrl-w>")
			[ ] Print("Excel file "+sFilename+" sucessfully opened")
			[ ] addTestsPassedCount(1)
		[+] else
			[ ] Print("The File ",sFilename," does not exist")
			[ ] ListAppend(lsFailedFileOpenTests, sFilename)
			[ ] iTestsFailedCount++
	[ ] closeExcel()
[+] closeExcel()
	[ ] MicrosoftExcel.SetActive()
	[ ] MicrosoftExcel.PressKeys("<alt>")
	[ ] MicrosoftExcel.TypeKeys("<f>")
	[ ] MicrosoftExcel.ReleaseKeys("<alt>")
	[ ] MicrosoftExcel.TypeKeys("<x>")
	[ ] 
[+] testcase tc_openExcelDocs() appstate none
	[ ] openExcelDocs()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	PowerPoint Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] openPowerPoint() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sCmdLine = '"C:\Program Files\Microsoft Office\OFFICE11\POWERPNT.EXE"'
	[ ] STRING sWorkingDir = "{sGlobalReadPath}"+""
	[+] if !(PowerPoint.Exists ())
		[ ] PowerPoint.Start(sCmdLine,sWorkingDir,"",20)
	[ ] // Print(sCmdLine)
	[ ] // Print(sWorkingDir)
	[ ] MicrosoftPowerPoint.sTag = "Microsoft PowerPoint*"
	[ ] MicrosoftPowerPoint.Maximize()
[+] openPowerPointDoc(STRING sFileName)
	[+] if SYS_FileExists(sFileName)
		[ ] MicrosoftPowerPoint.sTag = "Microsoft PowerPoint*"
		[ ] MicrosoftPowerPoint.SetActive()
		[ ] MicrosoftPowerPoint.PressKeys("<ctrl>")
		[ ] MicrosoftPowerPoint.TypeKeys("<o>")
		[ ] MicrosoftPowerPoint.ReleaseKeys("<ctrl>")
		[ ] sleep(1)
		[ ] PowerPointOpen.Filename.TypeKeys (sFileName)
		[ ] PowerPointOpen.Filename.TypeKeys ("<Enter>")
	[+] else
		[ ] Print("The File ",sFileName," does not exist")
		[ ] 
[+] openPowerPointDocs()
	[ ] LIST OF STRING lsFilenames = GetFilenamesFromExcel("PowerPoint")
	[ ] STRING sFilename, sFilenameTag
	[ ] MicrosoftPowerPoint.sTag = "Microsoft PowerPoint*"
	[ ] openPowerPoint()
	[+] for each sFilename in lsFilenames
		[+] if SYS_FileExists(sFilename)
			[ ] MicrosoftPowerPoint.SetActive()
			[ ] MicrosoftPowerPoint.PressKeys("<ctrl>")
			[ ] MicrosoftPowerPoint.TypeKeys("<o>")
			[ ] MicrosoftPowerPoint.ReleaseKeys("<ctrl>")
			[ ] sleep(1)
			[ ] PowerPointOpen.Filename.TypeKeys (sFilename)
			[ ] PowerPointOpen.Filename.TypeKeys ("<Enter>")
			[ ] sFilenameTag = GetField(sFilename,"\",3)
			[ ] // sFilenameTag = GetField(sFilenameTag,".",1)
			[ ] MicrosoftPowerPoint.sTag = "Microsoft PowerPoint - [{sFilenameTag}]"
			[ ] Print("Power Point file "+sFilename+" sucessfully opened")
			[ ] addTestsPassedCount(1)
		[+] else
			[ ] Print("The File ",sFilename," does not exist")
			[ ] ListAppend(lsFailedFileOpenTests, sFilename)
			[ ] iTestsFailedCount++
	[ ] closePowerPoint()
[+] closePowerPoint()
	[ ] MicrosoftPowerPoint.SetActive()
	[ ] MicrosoftPowerPoint.TypeKeys("<esc>")
	[ ] sleep(3)
	[ ] MicrosoftPowerPoint.SetActive()
	[ ] MicrosoftPowerPoint.PressKeys("<alt>")
	[ ] MicrosoftPowerPoint.TypeKeys("<f>")
	[ ] MicrosoftPowerPoint.ReleaseKeys("<alt>")
	[ ] MicrosoftPowerPoint.TypeKeys("<x>")
	[ ] 
[+] testcase tc_openPowerPointDocs() appstate none
	[ ] openPowerPointDocs()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Access Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] openAccess() 
	[ ] LIST OF STRING sReturn
	[ ] STRING sCmdLine = '"C:\Program Files\Microsoft Office\OFFICE11\MSACCESS.EXE"'
	[ ] STRING sWorkingDir = "{sGlobalReadPath}"+""
	[ ] MicrosoftAccess.sTag = "Microsoft Access"
	[+] if !(MicrosoftAccess.Exists())
		[ ] Access.Start(sCmdLine,sWorkingDir,"",20)
	[ ] // Print(sCmdLine)
	[ ] // Print(sWorkingDir)
	[ ] // MicrosoftAccess.Maximize()
[+] openAccessDoc(STRING sFileName)
	[+] if SYS_FileExists(sFileName)
		[ ] MicrosoftAccess.SetActive()
		[ ] MicrosoftAccess.PressKeys("<alt>")
		[ ] MicrosoftAccess.TypeKeys("<f>")
		[ ] MicrosoftAccess.ReleaseKeys("<alt>")
		[ ] MicrosoftAccess.TypeKeys("<o>")
		[ ] sleep(1)
		[ ] AccessOpen.Filename.TypeKeys (sFileName)
		[ ] AccessOpen.Filename.TypeKeys ("<Enter>")
		[+] if SecurityWarning1.Exists()
			[ ] SecurityWarning1.TypeKeys("<Enter>")
		[+] if SecurityWarning2.Exists()
			[ ] SecurityWarning2.OK.Click()
		[+] if SecurityWarning3.Exists()
			[ ] SecurityWarning3.SetActive()
			[ ] SecurityWarning3.TypeKeys("<Tab>")
			[ ] SecurityWarning3.TypeKeys("<Enter>")
	[+] else
		[ ] Print("The File ",sFileName," does not exist")
	[ ] 
[+] openAccessDocs()
	[ ] LIST OF STRING lsFilenames = GetFilenamesFromExcel("Access")
	[ ] STRING sFilename
	[ ] openAccess()
	[+] for each sFilename in lsFilenames
		[+] if SYS_FileExists(sFilename)
			[ ] MicrosoftAccess.SetActive()
			[ ] MicrosoftAccess.PressKeys("<alt>")
			[ ] MicrosoftAccess.TypeKeys("<f>")
			[ ] MicrosoftAccess.ReleaseKeys("<alt>")
			[ ] MicrosoftAccess.TypeKeys("<o>")
			[ ] sleep(1)
			[ ] AccessOpen.Filename.TypeKeys (sFilename)
			[ ] AccessOpen.Filename.TypeKeys ("<Enter>")
			[+] if SecurityWarning1.Exists()
				[ ] SecurityWarning1.TypeKeys("<Enter>")
			[+] if SecurityWarning2.Exists()
				[ ] SecurityWarning2.OK.Click()
			[+] if SecurityWarning3.Exists()
				[ ] SecurityWarning3.SetActive()
				[ ] SecurityWarning3.TypeKeys("<Tab>")
				[ ] SecurityWarning3.TypeKeys("<Enter>")
			[ ] MicrosoftAccess.TypeKeys("<ctrl-w>")
			[ ] Print("Access db  file "+sFilename+" sucessfully opened")
			[ ] addTestsPassedCount(1)
		[+] else
			[ ] Print("The File ",sFilename," does not exist")
			[ ] ListAppend(lsFailedFileOpenTests, sFilename)
			[ ] iTestsFailedCount++
	[ ] closeAccess()
[+] closeAccess()
	[ ] MicrosoftAccess.SetActive()
	[ ] MicrosoftAccess.PressKeys("<alt>")
	[ ] MicrosoftAccess.TypeKeys("<f>")
	[ ] MicrosoftAccess.ReleaseKeys("<alt>")
	[ ] MicrosoftAccess.TypeKeys("<x>")
[+] testcase tc_openAccessDocs() appstate none
	[ ] openAccessDocs()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Adobe Reader Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] openAdobeReader()
	[+] AdobeReader.sTag = "Adobe Reader*"
		[+] if !(AdobeReader.Exists ())
			[ ] STRING sCmdLine = '"C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"'
			[ ] STRING sWorkingDir = "c:\temp"
			[ ] AdobeReader.Start(sCmdLine,sWorkingDir,"",20)
	[ ] 
[+] openAdobeReaderDoc(STRING sFileName)
	[+] if SYS_FileExists(sFileName)
		[ ] AdobeReader.SetActive()
		[ ] AdobeReader.PressKeys("<ctrl>")
		[ ] AdobeReader.TypeKeys("<o>")
		[ ] AdobeReader.ReleaseKeys("<ctrl>")
		[ ] sleep(1)
		[ ] AdobeReaderOpen.FileName.TypeKeys (sFileName)
		[ ] AdobeReaderOpen.FileName.TypeKeys ("<Enter>")
	[+] else
		[ ] Print("The File ",sFileName," does not exist")
		[ ] 
[+] openAdobeReaderDocs()
	[ ] LIST OF STRING lsFilenames = GetFilenamesFromExcel("AdobeReader")
	[ ] STRING sFilename, sFilenameTag
	[ ] MicrosoftPowerPoint.sTag = "Microsoft PowerPoint*"
	[ ] openAdobeReader()
	[+] for each sFilename in lsFilenames
		[+] if SYS_FileExists(sFilename)
			[ ] AdobeReader.SetActive()
			[ ] AdobeReader.PressKeys("<ctrl>")
			[ ] AdobeReader.TypeKeys("<o>")
			[ ] AdobeReader.ReleaseKeys("<ctrl>")
			[ ] sleep(1)
			[ ] AdobeReaderOpen.FileName.TypeKeys (sFilename)
			[ ] AdobeReaderOpen.FileName.TypeKeys ("<Enter>")
			[ ] sFilenameTag = GetField(sFilename,"\",3)
			[ ] // sFilenameTag = GetField(sFilenameTag,".",1)
			[ ] AdobeReader.sTag = "Adobe Reader - [{sFilenameTag}]"
			[ ] Print("Adobe Reader  file "+sFilename+" sucessfully opened")
			[ ] iTestsPassedCount++
		[+] else
			[ ] Print("The File ",sFilename," does not exist")
			[ ] iTestsFailedCount++
	[ ] closeAdobeReader()
[+] closeAdobeReader()
	[ ] AdobeReader.SetActive()
	[ ] AdobeReader.PressKeys("<alt>")
	[ ] AdobeReader.TypeKeys("<f>")
	[ ] AdobeReader.ReleaseKeys("<alt>")
	[ ] AdobeReader.TypeKeys("<x>")
[+] testcase tc_openAdobeReaderdocs() appstate none
	[ ] openAdobeReaderDocs()
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] //	Symantec A/V Functions
[ ] ///////////////////////////////////////////////////////////////////////////////////////////////////////
[+] openSAV4Exchange()
	[ ] STRING scmd = "start "+sGlobalReadPath+"SAV4Exchange.bat"
	[ ] SYS_EXECUTE(scmd)
	[ ] 
[+] closeSAV4Exchange()
	[ ] SAV4Exchange.SetActive()
	[ ] SAV4Exchange.PressKeys("<alt>")
	[ ] SAV4Exchange.TypeKeys("<f>")
	[ ] SAV4Exchange.ReleaseKeys("<alt>")
	[ ] SAV4Exchange.TypeKeys("<x>")
[ ] ////////////////////////////////////////////////////////////////
[ ] // Print Diagnostics Test Cases
[ ] ////////////////////////////////////////////////////////////////
[+] testcase PrintIpConfigAll () appstate none
	[ ] LIST OF STRING lslPConfig = getIPConfigAll()
	[ ] STRING sItem 
	[ ] Print("Machine Configuration Information is :")
	[+] for each sItem in lslPConfig 
		[ ] Print(sItem)
		[ ] 
[+] testcase PrintHostName () appstate none
		[ ] Print(getHostName())
[+] testcase PrintServices () appstate none
	[ ] LIST OF STRING lsServices = getServices()
	[ ] STRING sItem 
	[ ] Print("Service Information: ")
	[+] for each sItem in lsServices 
		[ ] Print(sItem)
		[ ] 
[+] testcase PrintOSVersion () appstate none
	[ ] Print("OS Version ",getOSVersion())
[+] testcase PrintOSInfo () appstate none
	[ ] LIST OF STRING lsOSInfo = getOSInfo()
	[ ] STRING sItem 
	[ ] Print("Service Information: ")
	[+] for each sItem in lsOSInfo 
		[ ] Print(sItem)
		[ ] 
[+] testcase PrintjavaVersion () appstate none
	[ ] Print("Java Version ",getjavaVersion())
[ ] 
[+] testcase PrintLocalDiskInfo () appstate none
	[ ] LIST OF STRING lsDiskInfo = getLocalDiskInfo()
	[ ] STRING sItem 
	[ ] Print("Disk Information: ")
	[+] for each sItem in lsDiskInfo 
		[ ] Print(sItem)
[+] testcase PrintVirtualMemory () appstate none
	[ ] LIST OF STRING lsDiskInfo = getVirtualMemory()
	[ ] STRING sItem 
	[ ] Print("Virtual Memory: ")
	[+] for each sItem in lsDiskInfo 
		[ ] Print(sItem)
[ ] ////////////////////////////////////////////////////////////////
[ ] // Database Functions
[ ] ////////////////////////////////////////////////////////////////
[+] LIST OF SENDMAILREC GetSendMailRecsFromExcel(STRING sUser) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sAttachment
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF SENDMAILREC LSENDMAIL
	[ ] String sWorkSheet = "Outlook$"
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[+] switch sUser
		[+] case "All"
			[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, sFrom, To, Cc, Subject FROM '{sWorkSheet}'" )
		[+] default
			[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, sFrom, To, Cc, Subject FROM `Outlook$` WHERE sFrom = '{sUser}'" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, SENDMAIL.sAttachment, SENDMAIL.sFrom, SENDMAIL.sTo, SENDMAIL.sCc, SENDMAIL.sSubject)
		[+] if SENDMAIL.sCc == NULL 
			[ ] SENDMAIL.sCc = ' '
		[ ] ListAppend(LSENDMAIL, SENDMAIL)
	[ ] // we will now return the values in the record
	[ ] return LSENDMAIL
[+] LIST OF SENDMAILREC GetReceiveMailRecsFromExcel(STRING sUser) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sAttachment
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF SENDMAILREC LSENDMAIL
	[ ] String sWorkSheet = "Outlook$"
	[ ] // "+"{sGlobalReadPath}"+"
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[+] switch sUser
		[+] case "All"
			[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, sFrom, To, Cc, Subject FROM '{sWorkSheet}'" )
		[+] default
			[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, sFrom, To, Cc, Subject FROM `Outlook$` WHERE To = '{sUser}'" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, SENDMAIL.sAttachment, SENDMAIL.sFrom, SENDMAIL.sTo, SENDMAIL.sCc, SENDMAIL.sSubject)
		[+] if SENDMAIL.sCc == NULL 
			[ ] SENDMAIL.sCc = ' '
		[ ] ListAppend(LSENDMAIL, SENDMAIL)
	[ ] // we will now return the values in the record
	[ ] return LSENDMAIL
[+] LIST OF SENDMAILREC GetCCMailRecsFromExcel(STRING sUser) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sAttachment
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF SENDMAILREC LSENDMAIL
	[ ] String sWorkSheet = "Outlook$"
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[+] switch sUser
		[+] case "All"
			[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, sFrom, To, Cc, Subject FROM '{sWorkSheet}'" )
		[+] default
			[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, sFrom, To, Cc, Subject FROM `Outlook$` WHERE Cc = '{sUser}'" )
	[ ] 
    // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, SENDMAIL.sAttachment, SENDMAIL.sFrom, SENDMAIL.sTo, SENDMAIL.sCc, SENDMAIL.sSubject)
		[ ] ListAppend(LSENDMAIL, SENDMAIL)
	[ ] // we will now return the values in the record
	[ ] return LSENDMAIL
[ ] 
[+] testcase PrintLSENDMAILRECS()
	[ ] LIST OF SENDMAILREC LSENDMAIL
	[ ] INTEGER index = 1
	[ ] // ListPrint(GetSendMailRecsFromDB())
	[ ] LSENDMAIL = GetSendMailRecsFromExcel("All")
	[+] while index <= ListCount(LSENDMAIL)
		[ ] Print(LSENDMAIL[index].sAttachment)
		[ ] index++
		[ ] // Print(LSENDMAIL[2].sAttachment)
		[ ] // Print(LSENDMAIL[3].sAttachment)
		[ ] // Print(LSENDMAIL[4].sAttachment)
	[ ] 
[+] SENDMAILREC GetSendMailRecFromExcel(STRING rowID) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sAttachment
	[ ] SENDMAILREC SENDMAIL
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\\Centrixs.xls;UID=;PWD="
	[ ] 
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Attachment, To, Cc, Subject FROM [Attachments$] WHERE rowNumber = '{rowID}'")
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[ ] hResult = DB_FetchNext (hsqlresult, SENDMAIL.sTo, SENDMAIL.sCc, SENDMAIL.sSubject, SENDMAIL.sAttachment )
	[ ] // we will now return the values in the record
	[ ] return SENDMAIL
[+] testcase PrintSendMailRec()
	[ ] INTEGER index = 1
	[ ] SENDMAILREC SENDMAIL = GetSendMailRecFromExcel("1")
	[+] while IsSet(SENDMAIL)
		[+] while index++
			[ ] SENDMAIL = GetSendMailRecFromExcel([STRING]index)
	[ ] Print(SENDMAIL.sTo)
	[ ] Print(SENDMAIL.sCc)
	[ ] Print(SENDMAIL.sSubject)
	[ ] Print(SENDMAIL.sAttachment)
	[ ] 
[ ] 
[+] LIST OF STRING GetFilenamesFromExcel(STRING sWorksheet) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sFilename
	[ ] LIST OF STRING lsFilename
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Filename FROM [{sWorksheet}$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, sFilename)
		[ ] ListAppend(lsFilename, sGlobalReadPath+sFilename)
	[ ] // we will now return the values in the record
	[ ] return lsFilename
	[ ] 
[+] testcase PrintlsFilenames() appstate none
	[ ] LIST OF STRING lsFilenames
	[ ] INTEGER index = 1
	[ ] // ListPrint(GetSendMailRecsFromDB())
	[ ] lsFilenames = GetFilenamesFromExcel("Word")
	[+] while index <= ListCount(lsFilenames)
		[ ] Print(lsFilenames[index])
		[ ] index++
		[ ] // Print(LSENDMAIL[2].sAttachment)
		[ ] // Print(LSENDMAIL[3].sAttachment)
		[ ] // Print(LSENDMAIL[4].sAttachment)
	[ ] 
[+] LIST OF STRING GetColumnFromExcel(STRING sWorksheet, STRING sColumn) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sFilename
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF STRING lsFilename
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT {sColumn} FROM [{sWorksheet}$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, sFilename)
		[ ] ListAppend(lsFilename, sFilename)
	[ ] // we will now return the values in the record
	[ ] return lsFilename
	[ ] 
[+] LIST OF REGISTRYENTRYREC GetRegistryEntryRec(STRING sWorksheet) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sFilename
	[ ] SENDMAILREC SENDMAIL
	[ ] LIST OF REGISTRYENTRYREC lsRegistryEntryRec
	[ ] REGISTRYENTRYREC REGISTRYENTRY
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Path, Entry, Exception FROM [{sWorksheet}$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, REGISTRYENTRY.sPath, REGISTRYENTRY.sEntry, REGISTRYENTRY.sException)
		[ ] ListAppend(lsRegistryEntryRec, REGISTRYENTRY)
	[ ] // we will now return the values in the record
	[ ] return lsRegistryEntryRec
	[ ] 
[+] testcase PrintGetRegistryEntryRec() appstate none
	[ ] LIST OF REGISTRYENTRYREC lsRegistryExclusionList = GetRegistryEntryRec("RegistryExclusionList")
	[ ] ListPrint(lsRegistryExclusionList)
[ ] 
[+] LIST OF SERVICEREC getServicesFromExcel(STRING sWorksheet) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sDisplayName
	[ ] STRING sState
	[ ] STRING sStartMode
	[ ] SERVICEREC SERVICE
	[ ] LIST OF SERVICEREC LSERVICE
	[ ] LIST OF STRING lsServices
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT DisplayName, State, StartMode FROM [{sWorksheet}$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, SERVICE.sDisplayName, SERVICE.sState, SERVICE.sStartMode)
		[ ] ListAppend(LSERVICE, SERVICE)
	[ ] // we will now return the values in the record
	[ ] return LSERVICE
[+] testcase PrintServicesFromExcel()
	[ ] ListPrint(GetServicesFromExcel("CS01Services"))
[+] STRING getVersionFromExcel(STRING sApplication) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sVersion
	[ ] STRING sWorksheet = "Versions"
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Version FROM `{sWorksheet}$` WHERE Application = '{sApplication}'")
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[ ] while DB_FetchNext (hsqlresult, sVersion)
	[ ] // we will now return the values in the record
	[+] if sVersion == NULL || sVersion == ""
		[ ] Print("Application name",sApplication,"not found in excel spreadsheet")
	[ ] return sVersion
[+] testcase PrintVersionFromExcel()
	[ ] Print(GetVersionFromExcel("Outlook"))
[+] STRING getBuildFromExcel(STRING sApplication) 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sBuild
	[ ] STRING sWorksheet = "Versions"
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Build FROM `{sWorksheet}$` WHERE Application = '{sApplication}'")
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[ ] while DB_FetchNext (hsqlresult, sBuild)
	[ ] // we will now return the values in the record
	[ ] return sBuild
[+] LIST OF VERSIONREC getComposeDLLs()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getOSVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //                                                    the OS Version
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] VERSIONREC VERSION
	[ ] LIST OF VERSIONREC LVERSION
	[ ] STRING cmdLine1 = "dir *.dll /B"
	[ ] // Print("cmdLine1 = "+cmdLine1)
	[ ] SYS_SetDrive("c")
	[ ] SYS_SetDir("C:\Program Files\Compose")
	[ ] SYS_Execute(cmdLine1,sReturn)
	[+] for each sItem in sReturn
		[ ] VERSION.sFilename = sItem
		[ ] VERSION.sVersion = getFileVersion(sGlobalComposeDLLPath+sItem)
		[ ] ListAppend(LVERSION,VERSION)
		[ ] // Print("VERSION.sFilename =", VERSION.sFilename)
		[ ] // Print("VERSION.sVersion =", VERSION.sVersion)
	[ ] return LVERSION
	[ ] 
[+] LIST OF VERSIONREC getComposeDLLsFromExcel() 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sDisplayName
	[ ] STRING sState
	[ ] STRING sStartMode
	[ ] VERSIONREC VERSION
	[ ] LIST OF VERSIONREC LVERSION
	[ ] LIST OF STRING lsServices
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Filename, Version FROM [COMPOSEDLLVERSIONS$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, VERSION.sFilename, VERSION.sVersion)
		[ ] ListAppend(LVERSION, VERSION)
	[ ] // we will now return the values in the record
	[ ] return LVERSION
[ ] 
[+] testcase PrintComposeDLLs() appstate none
	[ ] ListPrint(getComposeDLLs())
	[ ] 
[+] testcase tc_createMDACEntriesForExcel() appstate none
	[ ] createMDACEntriesForExcel("c:\SilkTestData\MSACresults.xml")
	[ ] 
[+] createMDACEntriesForExcel(STRING sFilename)
	[ ] HFILE hFile, hOutputFile, hRegistryExclusionFile
	[ ] STRING sLine
	[ ] STRING sReleaseName
	[ ] STRING sVersion
	[ ] STRING sOutputFilename = "{sGlobalWritePath}MDACDLLVERSIONS.csv"
	[ ] LIST OF VERSIONREC LVERSIONREC
	[ ] VERSIONREC VERSION
	[ ] INT i
	[ ] hFile = FileOpen(sFilename,FM_READ)
	[ ] // hOutputFile = FileOpen(sOutputFilename,FM_WRITE)
	[+] while FileReadLine(hFile,sLine)
		[+] if MatchStr("*xmlns*",sLine)
			[ ] VERSION.sFilename = GetField(sLine, '"', 2)
		[+] if MatchStr("*release name*", sLine)
			[ ] // get the release name first
			[ ] VERSION.sVersion = GetField(sLine, '"', 4)
			[ ] ListAppend(LVERSIONREC, VERSION)
	[ ] ListPrint(LVERSIONREC)
	[ ] FileClose(hFile)
	[ ] 
[+] LIST OF VERSIONREC getMDACDLLsFromExcel() 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sDisplayName
	[ ] STRING sState
	[ ] STRING sStartMode
	[ ] VERSIONREC VERSION
	[ ] LIST OF VERSIONREC LVERSION
	[ ] LIST OF STRING lsServices
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Filename, Version FROM [MDACDLLVERSIONS$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, VERSION.sFilename, VERSION.sVersion)
		[ ] STRING sSystemRoot = GetField(VERSION.sFilename,"%",2)
		[ ] // Print("System Root =",sSystemRoot)
		[+] if sSystemRoot == "SYSTEMROOT"
			[+] if Substr(gethostname(),8,4) == "CS01"
				[ ] VERSION.sFilename = StrTran(VERSION.sFilename,"%SYSTEMROOT%","C:\WINDOWS") 
			[+] if Substr(gethostname(),8,4) == "DC01" || Substr(gethostname(),8,4) == "DC02"
				[ ] VERSION.sFilename = StrTran(VERSION.sFilename,"%SYSTEMROOT%","C:\WINNT")
		[ ] ListAppend(LVERSION, VERSION)
	[ ] // we will now return the values in the record
	[ ] return LVERSION
	[ ] 
[+] LIST OF VERSIONREC getMDACDLLs() 
	[ ] HDATABASE hdbc
	[ ] HSQL hsqlresult
	[ ] BOOLEAN hResult
	[ ] STRING sDisplayName
	[ ] STRING sState
	[ ] STRING sStartMode
	[ ] VERSIONREC VERSION
	[ ] LIST OF VERSIONREC LVERSION
	[ ] LIST OF STRING lsServices
	[ ] 
	[ ] STRING gsDSNConnect = "DSN=Silk;DBQ=C:\SilkTestData\Centrixs.xls;UID=;PWD="
	[ ] // We will connect to an ODBC driver that communicates with an excel spreadsheet in a worksheet called userpolicy
	[ ] hdbc = DB_Connect (gsDSNConnect)
	[ ] // The columns in the worksheet represent the fields of the database. 
	[ ] hsqlresult = DB_ExecuteSql (hdbc, "SELECT Filename FROM [MDACDLLVERSIONS$]" )
	[ ] // This statement takes the values returned from the query and puts the values into the passed variables
	[+] while DB_FetchNext (hsqlresult, VERSION.sFilename)
		[ ] STRING sSystemRoot = GetField(VERSION.sFilename,"%",2)
		[ ] // Print("System Root =",sSystemRoot)
		[+] if sSystemRoot == "SYSTEMROOT"
			[+] if Substr(gethostname(),8,4) == "CS01"
				[ ] VERSION.sFilename = StrTran(VERSION.sFilename,"%SYSTEMROOT%","C:\WINDOWS") 
			[+] if Substr(gethostname(),8,4) == "DC01" || Substr(gethostname(),8,4) == "DC02"
				[ ] VERSION.sFilename = StrTran(VERSION.sFilename,"%SYSTEMROOT%","C:\WINNT")
		[ ] VERSION.sVersion = getFileVersion(VERSION.sFilename)
		[ ] ListAppend(LVERSION, VERSION)
	[ ] // we will now return the values in the record
	[ ] return LVERSION
[ ] 
[+] testcase PrintMDACDLLs() appstate none
	[ ] ListPrint(getMDACDLLsFromExcel())
[ ] 
