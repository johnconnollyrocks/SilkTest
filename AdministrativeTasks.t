[ ] use "AdministrativeTasks.inc" 
[ ] // use "centrixs.inc"
[ ] use "centrixs.t"
[-] appstate ODBC() 
	[-] if (GetTestCaseState()==TCS_ENTERING)
		[ ] RunStartMenu("odbcad32.exe")
	[-] if (GetTestCaseState()==TCS_EXITING)
		[ ] ODBCDataSourceAdministrator.SetActive ()
		[ ] ODBCDataSourceAdministrator.OK.Click ()
[+] appstate ComputerManagement() 
	[-] if (GetTestCaseState()==TCS_ENTERING)
		[ ] RunStartMenu("%SystemRoot%\system32\compmgmt.msc /s")
		[ ] // SYS_EXECUTE("%SystemRoot%\system32\compmgmt.msc /s")
	[-] if (GetTestCaseState()==TCS_EXITING)
		[ ] 
[+] appstate ActiveDirectoryUsers ()
	[+] if (GetTestCaseState()==TCS_ENTERING)
		[+] if !ActiveDirectoryUsers.Exists()
			[ ] Taskbar.SetActive ()
			[ ] Taskbar.Start.Click ()
			[ ] StartMenu.SetActive ()
			[ ] StartMenu.TypeKeys("R")
			[ ] Run.SetActive ()
			[ ] Run.Open.TypeKeys("dsa.msc")
			[ ] Run.OK.Click ()
	[+] if (GetTestCaseState()==TCS_EXITING)
		[+] if HomeDirectoryExists.Exists()
			[ ] HomeDirectoryExists.OK.Click()
		[ ] ActiveDirectoryUsers.Close ()
[ ] //////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // Testcases
[ ] //////////////////////////////////////////////////////////////////////////////////////////////////////
[-] testcase tc_createDSN(STRING sDSN, BOOLEAN sReplaceExisting) appstate ODBC
	[ ] createDSN(sDSN,sReplaceExisting)
	[ ] 
[-] testcase tc_runDefrag(STRING sDrive) appstate ComputerManagement
	[ ] runDefrag(sDrive)
	[ ] 
[+] testcase tc_createStandardUser(STRING sUser, STRING sPassword) appstate ActiveDirectoryUsers
	[ ] createStandardUser(sUser,sPassword)
[+] testcase tc_createAdminUser(STRING sUser, STRING sPassword) appstate ActiveDirectoryUsers
	[ ] createAdminUser(sUser,sPassword)
[-] testcase tc_deleteUser(STRING sUser) appstate ActiveDirectoryUsers
	[ ] deleteUser(sUser)
[-] testcase tc_ChangeUserPassword(STRING sUser, STRING sPassword) appstate ActiveDirectoryUsers
	[ ] changeUserPassword(sUser,sPassword)
[+] testcase tc_disableAccount(STRING sUser) appstate ActiveDirectoryUsers
	[ ] disableAccount(sUser)
[+] testcase tc_enableAccount(STRING sUser) appstate ActiveDirectoryUsers
	[ ] enableAccount(sUser)
[+] testcase tc_lockoutAccount(STRING sUser) appstate ActiveDirectoryUsers
	[ ] lockoutAccount(sUser, getHostNameSuffix())
[+] testcase tc_unlockAccount(STRING sUser) appstate ActiveDirectoryUsers
	[ ] unlockAccount(sUser)
[+] testcase tc_verifyAccount(STRING sUser) appstate ActiveDirectoryUsers
	[ ] verifyAccount(sUser)
[+] testcase tc_setAccountExpirationDate(BOOLEAN bImmediate) appstate ActiveDirectoryUsers
	[ ] setAccountExpirationDate("user1", bImmediate)
[-] testcase tc_RemoveHomeDirectory(STRING sHome) appstate none
	[ ] deleteHomeDirectory(sHome)
	[ ] 
[+] testcase PrintGetServerName() appstate none
	[ ] Print(getServerName())
[+] testcase PrintGetHostName() appstate none
	[ ] Print(getHostName())
[+] testcase PrintGetHostNameSuffix() appstate none
	[ ] Print(getHostNameSuffix())
[+] testcase PrintGetHostNamePrefix() appstate none
	[ ] Print(getHostNamePrefix())
[ ] //////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // Functions
[ ] //////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // Active Directory fuinctions
[ ] /////////////////////////////////////////////////////////////////////////////////////////////////////
[+] createStandardUser(STRING sUser, STRING sPassword)
	[+] if getServerName() == "DC01"
		[ ] STRING sListItem = Lower(getHostNamePrefix()) + Lower(getServerName()) + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[ ] // ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [BLK0520DC01.blk0520.navy.usa.cfe.cmil.mil]/blk0520.navy.usa.cfe.cmil.mil")
		[ ] // ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [BLK0520DC01.blk0520.navy.usa.cfe.cmil.mil]/blk0520.navy.usa.cfe.cmil.mil/COMPOSE Users and Computers")
		[ ] // ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [BLK0520DC01.blk0520.navy.usa.cfe.cmil.mil]/blk0520.navy.usa.cfe.cmil.mil/COMPOSE Users and Computers/COMPOSE Users")
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("_UserTemplate;User;This is a template used to create User accounts;;;;;;;;;;;;;;;;;;;;;;")
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.TypeKeys("<Alt-A>")
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.TypeKeys("C")
		[ ] 
		[ ] CopyObjectUser1.SetActive ()
		[ ] CopyObjectUser1.DialogBox1.FirstName.SetText (sUser)
		[ ] CopyObjectUser1.DialogBox1.UserLogonName.SetText (sUser)
		[ ] CopyObjectUser1.Next.Click ()
		[ ] sleep(1)
		[+] if UserAlreadyExists.Exists()
			[ ] LogWarning("***** WARNING: User {sUser} could not be added - already an existing user")
			[ ] UserAlreadyExists.OK.Click()
			[ ] CopyObjectUser1.SetActive ()
			[ ] CopyObjectUser1.Cancel.Click ()
			[ ] return
		[ ] CopyObjectUser2.SetActive ()
		[ ] CopyObjectUser2.DialogBox1.Password.SetText (Decrypt ("xhNwAdrsMdaDxCr/"))
		[ ] CopyObjectUser2.DialogBox1.ConfirmPassword.SetText (Decrypt ("xhNwAdrsMdaDxCr/"))
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.UserMustChangePasswordAtN.Uncheck ()
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.UserCannotChangePassword.Check ()
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.PasswordNeverExpires.Check ()
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.AccountIsDisabled.Uncheck ()
		[ ] CopyObjectUser2.Next.Click ()
		[ ] 
		[ ] CopyObjectUser3.SetActive ()
		[ ] CopyObjectUser3.Next.Click ()
		[ ] CopyObjectUser3.Next.Click ()
	[-] else
		[ ] LogWarning("testcase tc_createStandardUser was not run on host {getHostName()} must be run on DC01")
		[ ] 
[+] createAdminUser(STRING sUser, STRING sPassword)
	[+] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[ ] 
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("_AdminTemplate;User;This is a template used to create Admin accounts;;;;;;;;;;;;;;;;;;;;;;")
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.TypeKeys("<Alt-A>")
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.TypeKeys("C")
		[ ] 
		[ ] CopyObjectUser1.SetActive ()
		[ ] CopyObjectUser1.DialogBox1.FirstName.SetText (sUser)
		[ ] CopyObjectUser1.DialogBox1.UserLogonName.SetText (sUser)
		[ ] CopyObjectUser1.Next.Click ()
		[ ] sleep(1)
		[+] if UserAlreadyExists.Exists()
			[ ] LogWarning("***** WARNING: User {sUser} could not be added - already an existing user")
			[ ] UserAlreadyExists.OK.Click()
			[ ] CopyObjectUser1.SetActive ()
			[ ] CopyObjectUser1.Cancel.Click ()
			[ ] return
		[ ] CopyObjectUser2.SetActive ()
		[ ] CopyObjectUser2.DialogBox1.Password.SetText (Decrypt ("xhNwAdrsMdaDxCr/"))
		[ ] CopyObjectUser2.DialogBox1.ConfirmPassword.SetText (Decrypt ("xhNwAdrsMdaDxCr/"))
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.UserMustChangePasswordAtN.Uncheck ()
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.UserCannotChangePassword.Check ()
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.PasswordNeverExpires.Check ()
		[ ] sleep(1)
		[ ] CopyObjectUser2.DialogBox1.AccountIsDisabled.Uncheck ()
		[ ] CopyObjectUser2.Next.Click ()
		[ ] 
		[ ] CopyObjectUser3.SetActive ()
		[ ] CopyObjectUser3.Next.Click ()
		[ ] CopyObjectUser3.Next.Click ()
	[-] else
		[ ] LogWarning("testcase tc_createAdminUser was not run on host {getHostName()} must be run on DC01")
	[ ] 
[-] deleteUser(STRING sUser)
	[-] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return
		[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.TypeKeys("<alt-a>d")
		[ ] ActiveDirectoryConfirm.Yes.Click()
	[-] else
		[ ] LogWarning("testcase tc_deleteUser was not run on host {getHostName()} must be run on DC01")
[-] changeUserPassword(STRING sUser, STRING sNewPassword)
	[-] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return
		[ ] ActiveDirectoryUsers.SetActive()
		[ ] ActiveDirectoryUsers.SizeableRebar1.ReBarWindow321.ToolBar1.Action.Click ()
		[ ] // ActiveDirectoryUsers.TypeKeys("<alt-a>")
		[ ] ActiveDirectoryUsers.TypeKeys("e")
		[ ] ActiveDirectoryUsers.TypeKeys("<enter>")
		[ ] ResetPassword.SetActive()
		[ ] ResetPassword.NewPassword.SetText(sNewPassword)
		[ ] ResetPassword.ConfirmPassword.SetText(sNewPassword)
		[ ] ResetPassword.OK.Click()
		[ ] ResetPasswordConfirm.OK.Click()
	[-] else
		[ ] LogWarning("testcase tc_changeUserPassword was not run on host {getHostName()} must be run on DC01")
	[ ] 
[-] disableAccount(STRING sUser)
	[-] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return
		[ ] ActiveDirectoryUsers.SetActive()
		[ ] ActiveDirectoryUsers.SizeableRebar1.ReBarWindow321.ToolBar1.Action.Click ()
		[ ] // ActiveDirectoryUsers.TypeKeys("<alt-a>")
		[ ] ActiveDirectoryUsers.TypeKeys("s")
		[ ] // ActiveDirectoryUsers.TypeKeys("<enter>")
		[ ] sleep(3)
		[+] if MessageBox.Exists()
			[ ] MessageBox.SetActive ()
			[ ] MessageBox.OK.Click ()
		[+] else
			[ ] LogError("The user {sUser} is already disabled")
	[-] else
		[ ] LogWarning("testcase tc_disableAccount was not run on host {getHostName()} must be run on DC01")
[-] enableAccount(STRING sUser)
	[+] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return
		[ ] ActiveDirectoryUsers.SetActive()
		[ ] ActiveDirectoryUsers.SizeableRebar1.ReBarWindow321.ToolBar1.Action.Click ()
		[ ] // ActiveDirectoryUsers.TypeKeys("<alt-a>")
		[ ] ActiveDirectoryUsers.TypeKeys("e")
		[ ] ActiveDirectoryUsers.TypeKeys("<enter>")
		[ ] sleep(3)
		[+] if MessageBox.Exists()
			[ ] MessageBox.SetActive ()
			[ ] MessageBox.OK.Click ()
		[+] else
			[ ] LogError("The user {sUser} is already enabled")
	[-] else
		[ ] LogWarning("testcase tc_enableAccount was not run on host {getHostName()} must be run on DC01")
	[ ] 
[+] lockoutAccount(STRING sUser, STRING sDomain)
	[ ] STRING sCMD = "runas /user:{sUser}@{sDomain} notepad.exe"
	[ ] LIST OF STRING sReturn
	[ ] INT i
	[+] for i = 1 to 3
		[ ] sCMD = "runas /user:{sUser}@{sDomain} notepad.exe"
		[ ] SYS_EXECUTE(sCMD,sReturn)
		[ ] // ListPrint(sReturn)
		[ ] sCMD = chr(13)+chr(10)
		[ ] SYS_EXECUTE(sCMD,sReturn)
		[ ] // ListPrint(sReturn)
[+] unlockAccount(STRING sUser)
	[-] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return
		[ ] ActiveDirectoryUsers.SetActive()
		[ ] ActiveDirectoryUsers.SizeableRebar1.ReBarWindow321.ToolBar1.Action.Click ()
		[ ] // ActiveDirectoryUsers.TypeKeys("<alt-a>")
		[ ] ActiveDirectoryUsers.TypeKeys("r")
		[ ] // ActiveDirectoryUsers.TypeKeys("<enter>")
		[ ] // UserProperties.PageList1.Select ("Account")
		[ ] UserProperties.DialogBox1.TypeKeys("<shift-tab>")
		[ ] UserProperties.DialogBox1.TypeKeys("<ctrl-tab 2>")
		[+] if UserProperties.DialogBox1.AccountIsLockedOut.IsChecked()
			[ ] UserProperties.DialogBox1.AccountIsLockedOut.Uncheck ()
		[+] else
			[ ] LogError("The account {sUser} is not locked out as expected")
		[ ] UserProperties.DialogBox1.TypeKeys("<enter>")
	[-] else
		[ ] LogWarning("testcase tc_unlockAccount was not run on host {getHostName()} must be run on DC01")
		[ ] 
[+] BOOLEAN verifyAccount(STRING sUser)
	[-] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
			[ ] return(TRUE)
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return(FALSE)
			[ ] 
	[-] else
		[ ] LogWarning("testcase tc_verifyAccount was not run on host {getHostName()} must be run on DC01")
[+] setAccountExpirationDate(STRING sUser,BOOLEAN bImmediate)
	[ ] // if bImmediate = FALSE, the account will expire in one month
	[ ] // if bImmediate = TRUE, the account will expire immediately
	[+] if getServerName() == "DC01"
		[ ] STRING sListItem = getHostNamePrefix() + getServerName() + "." + getHostNameSuffix() + "]/" + getHostNameSuffix()
		[ ] ActiveDirectoryChild.SetActive ()
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Expand ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers")
		[ ] ActiveDirectoryChild.MMCViewWindow1.TreeView1.Select ("/Active Directory Users and Computers [{sListItem}"+"/COMPOSE Users and Computers/COMPOSE Users")
		[+] do
			[ ] ActiveDirectoryChild.MMCViewWindow1.ListView1.Select ("{sUser}")
		[+] except
			[ ] LogError("The user {sUser} was not found in Active Directory")
			[ ] return
		[ ] ActiveDirectoryUsers.SetActive()
		[ ] ActiveDirectoryUsers.SizeableRebar1.ReBarWindow321.ToolBar1.Action.Click ()
		[ ] // ActiveDirectoryUsers.TypeKeys("<alt-a>")
		[ ] ActiveDirectoryUsers.TypeKeys("r")
		[ ] // ActiveDirectoryUsers.TypeKeys("<enter>")
		[ ] // UserProperties.PageList1.Select ("Account")
		[ ] UserProperties.DialogBox1.TypeKeys("<shift-tab>")
		[ ] UserProperties.DialogBox1.TypeKeys("<ctrl-tab 2>")
		[ ] UserProperties.DialogBox1.AccountExpires.Select ("End of:")
		[+] if bImmediate
			[ ] UserProperties.DialogBox1.TypeKeys("<tab><down>")
		[ ] UserProperties.DialogBox1.TypeKeys("<enter>") // OK Button
	[-] else
		[ ] LogWarning("testcase tc_setAccountExpirationDate was not run on host {getHostName()} must be run on DC01")
	[ ] 
[-] deleteHomeDirectory(STRING sHome)
	[ ] STRING sCMD
	[ ] LIST OF STRING lsReturn
	[ ] STRING sReturn
	[-] if getHostName() == "DC01"
		[ ] SYS_SetDrive("e")
		[ ] SYS_SETDIR("\COMPOSEUsers")
		[ ] SYS_EXECUTE("rd "+sHome, lsReturn)
	[-] else
		[ ] LogWarning("testcase tc_deleteHomeDirectory was not run on host {getHostName()} must be run on DC01")
	[ ] // ListPrint(lsReturn)
	[ ] 
[ ] /////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // ODBC Functions
[-] createDSN(STRING sDSN, BOOLEAN bReplaceExisting) 
	[ ] ODBCDataSourceAdministrator.SetActive ()
	[ ] ODBCDataSourceAdministrator.PageList1.Select ("User DSN")
	[ ] ODBCDataSourceAdministrator.DialogBox1.Add.Click ()
	[ ] CreateNewDataSource.SetActive ()
	[ ] CreateNewDataSource.CreateNewDataSource.SelectADriverForWhichYou2.Select ("Microsoft Excel Driver (*.xls);4.00.6305.00;Microsoft Corporation;ODBCJT32.DLL;3/24/2005")
	[ ] CreateNewDataSource.Finish.Click ()
	[ ] ODBCMicrosoftExcelSetup.SetActive ()
	[ ] ODBCMicrosoftExcelSetup.DataSourceName.SetText (sDSN)
	[ ] ODBCMicrosoftExcelSetup.Description.SetText (sDSN)
	[ ] ODBCMicrosoftExcelSetup.OK.Click ()
	[-] if Error.Exists()
		[-] if bReplaceExisting == TRUE
			[ ] Error.Yes.Click()
		[-] else
			[ ] Error.No.Click()
	[ ] 
[ ] /////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // Management Console Functions
[+] // verifyMDACversion()
	[ ] // RunStartMenu("{sGlobalReadPath}\cc.exe")
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer1.ATL01039E601.Click (1, 85, 21)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 32, 26)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 33, 45)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 36, 61)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 38, 73)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 38, 92)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 44, 105)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 51, 127)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 47, 144)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 46, 154)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 46, 170)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 46, 190)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 46, 204)
	[ ] // ComponentCheckerVersion20.WTL_SplitterWindow1.WTL_PaneContainer2.ATL01039E101.Click (1, 45, 215)
	[ ] // 
[-] runDefrag(STRING sDriveLetter)
	[ ] INT i = 0
	[-] if getHostName() == "DC01"
		[ ] ComputerManagementLocal.SetActive()
		[ ] ComputerManagementLocal.MMCViewWindow1.TreeView1.Select ("/Computer Management (Local)/Storage/Disk Defragmenter")
		[ ] ComputerManagementDiskDefragPane.MMCViewWindow1.MMCOCXViewWindow1.AtlAxWinEx1.DiskDefragmenter.ListView1.Select ("({sDriveLetter}:*")
		[ ] ComputerManagementDiskDefragPane.MMCViewWindow1.MMCOCXViewWindow1.AtlAxWinEx1.DiskDefragmenter.Defragment.Click ()
		[+] while !DiskDefragmenter.Exists()
			[ ] sleep(1)
			[ ] i++
		[ ] Print("Defrag completed in",i,"Seconds")
		[ ] DiskDefragmenter.Close.Click()
		[ ] ComputerManagement.Close()
	[+] else
		[ ] LogWarning("testcase tc_deleteHomeDirectory was not run on host {getHostName()} must be run on DC01")
[ ] /////////////////////////////////////////////////////////////////////////////////////////////////////
[ ] // 
