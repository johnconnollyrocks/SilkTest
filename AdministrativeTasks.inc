[+] window DialogBox ActiveDirectoryUsers 
	[ ] tag "Active Directory Users and Computers"
	[+] CustomWin SizeableRebar1
		[ ] tag "[SizeableRebar]$13341"
		[-] CustomWin ReBarWindow321
			[ ] tag "[ReBarWindow32]$4096"
			[+] ToolBar ToolBar2
				[ ] tag "$4098"
				[+] PushButton PushButton1
					[ ] tag "$1"
				[+] PushButton PushButton2
					[ ] tag "$2"
				[+] PushButton PushButton3
					[ ] tag "$4"
				[+] PushButton PushButton4
					[ ] tag "$5"
				[+] PushButton PushButton5
					[ ] tag "$7"
				[+] PushButton PushButton6
					[ ] tag "$9"
				[+] PushButton PushButton7
					[ ] tag "$11"
				[+] PushButton PushButton8
					[ ] tag "$12"
				[+] PushButton PushButton9
					[ ] tag "$14"
				[+] PushButton PushButton10
					[ ] tag "$15"
				[+] PushButton PushButton11
					[ ] tag "$17"
				[+] PushButton PushButton12
					[ ] tag "$19"
				[+] PushButton PushButton13
					[ ] tag "$20"
				[+] PushButton PushButton14
					[ ] tag "$21"
				[+] PushButton PushButton15
					[ ] tag "$22"
				[+] PushButton PushButton16
					[ ] tag "$23"
				[+] PushButton PushButton17
					[ ] tag "$24"
			[+] ToolBar ToolBar1
				[ ] tag "$4097"
				[+] PushButton PushButton1
					[ ] tag "$21508"
				[+] PushButton File
					[+] multitag "File"
						[ ] "$21513"
				[+] PushButton Action
					[+] multitag "Action"
						[ ] "$21509"
				[+] PushButton View
					[+] multitag "View"
						[ ] "$21510"
				[+] PushButton xWindow
					[+] multitag "Window"
						[ ] "$21514"
				[+] PushButton Help
					[+] multitag "Help"
						[ ] "$21515"
			[+] CustomWin AfxWnd42u1
				[ ] tag "[AfxWnd42u]$8193"
[+] window ChildWin ActiveDirectoryChild
	[ ] tag "Active Directory Users and Computers*"
	[ ] parent ActiveDirectoryUsers
	[+] CustomWin MMCViewWindow1
		[+] multitag "[MMCViewWindow]#1"
			[ ] "[MMCViewWindow]$59648"
		[+] ListView ListView1
			[+] multitag "#1"
				[ ] "$12786"
			[+] Header Header1
				[+] multitag "#1"
					[ ] "$0"
				[+] PushButton Name
					[+] multitag "Name"
						[ ] "#1"
				[+] PushButton Type
					[+] multitag "Type"
						[ ] "#2"
				[+] PushButton Description
					[+] multitag "Description"
						[ ] "#3"
				[+] PushButton BusinessPhone
					[+] multitag "Business Phone"
						[ ] "#18"
				[+] PushButton City
					[+] multitag "City"
						[ ] "#19"
				[+] PushButton Company
					[+] multitag "Company"
						[ ] "#17"
				[+] PushButton CountryRegion
					[+] multitag "Country?Region"
						[ ] "#15"
				[+] PushButton Department
					[+] multitag "Department"
						[ ] "#16"
				[+] PushButton DisplayName
					[+] multitag "Display Name"
						[ ] "#20"
				[+] PushButton PreWindows2000LogonName
					[+] multitag "Pre-Windows 2000 Logon Name"
						[ ] "#24"
				[+] PushButton EMailAddress
					[+] multitag "E-Mail Address"
						[ ] "#25"
				[+] PushButton ExchangeAlias
					[+] multitag "Exchange Alias"
						[ ] "#23"
				[+] PushButton ExchangeMailboxStore
					[+] multitag "Exchange Mailbox Store"
						[ ] "#21"
				[+] PushButton FirstName
					[+] multitag "First Name"
						[ ] "#22"
				[+] PushButton InstantMessagingHomeServer
					[+] multitag "Instant Messaging Home Server"
						[ ] "#7"
				[+] PushButton InstantMessagingURL
					[+] multitag "Instant Messaging URL"
						[ ] "#8"
				[+] PushButton LastName
					[+] multitag "Last Name"
						[ ] "#6"
				[+] PushButton Modified
					[+] multitag "Modified"
						[ ] "#4"
				[+] PushButton Office
					[+] multitag "Office"
						[ ] "#5"
				[+] PushButton State
					[+] multitag "State"
						[ ] "#9"
				[+] PushButton TargetAddress
					[+] multitag "Target Address"
						[ ] "#13"
				[+] PushButton JobTitle
					[+] multitag "Job Title"
						[ ] "#14"
				[+] PushButton UserLogonName
					[+] multitag "User Logon Name"
						[ ] "#12"
				[+] PushButton X400EMailAddress
					[+] multitag "X.400 E-Mail Address"
						[ ] "#10"
				[+] PushButton ZipCode
					[+] multitag "Zip Code"
						[ ] "#11"
		[+] TreeView TreeView1
			[+] multitag "#1"
				[ ] "$12785"
		[+] CustomWin OwnerDraw1
			[+] multitag "[OwnerDraw]#1"
				[ ] "[OwnerDraw]$12791"
	[+] StatusBar StatusBar1
		[+] multitag "#1"
			[ ] "$4099"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
		[+] StaticText StaticText1
			[+] multitag "#1"
				[ ] "$4097"
[ ] 
[+] window DialogBox ActiveDirectoryConfirm
	[ ] tag "Active Directory"
	[ ] parent ActiveDirectoryUsers
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "#1"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "#2"
			[ ] "$7"
	[+] ListBox TheSelectedObjectHasOther1
		[+] multitag "The selected object has other associated objects. Select those associated objects that you also want to delete."
			[ ] "#1"
			[ ] "$311"
	[+] StaticText AreYouSureYouWantToDelet
		[+] multitag "Are you sure you want to delete this object?"
			[ ] "#1"
			[ ] "$312"
	[+] StaticText TheSelectedObjectHasOther2
		[+] multitag "The selected object has other associated objects. Select those associated objects that you also want to delete."
			[ ] "#2"
			[ ] "$313"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$65535"
[+] window DialogBox ResetPassword
	[ ] tag "Reset Password"
	[ ] parent ActiveDirectoryUsers
	[+] StaticText NewPasswordText
		[+] multitag "New password:"
			[ ] "#1"
			[ ] "$-1[1]"
	[+] TextField NewPassword
		[+] multitag "New password:"
			[ ] "#1"
			[ ] "$220"
	[+] StaticText ConfirmPasswordText
		[+] multitag "Confirm password:"
			[ ] "#2"
			[ ] "$-1[2]"
	[+] TextField ConfirmPassword
		[+] multitag "Confirm password:"
			[ ] "#2"
			[ ] "$222"
	[+] CheckBox UserMustChangePasswordAtN
		[+] multitag "User must change password at next logon"
			[ ] "#1"
			[ ] "$261"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#2"
			[ ] "$2"
	[+] StaticText TheUserMustLogoffAndThen
		[+] multitag "The user must logoff and then logon again for the change to take effect."
			[ ] "#3"
			[ ] "$-1[3]"
[+] window DialogBox ResetPasswordConfirm
	[ ] tag "Active Directory"
	[ ] parent ActiveDirectoryUsers
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$2"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$20"
	[+] StaticText ThePasswordForUserbHasBee
		[+] multitag "The password for userb has been changed."
			[ ] "#1"
			[ ] "$65535"
[+] window DialogBox DisableUserConfirm
	[ ] tag "Active Directory"
	[ ] parent ActiveDirectoryUsers
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$2"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$20"
	[+] StaticText ThePasswordForUserbHasBee
		[+] multitag "The password for userb has been changed."
			[ ] "#1"
			[ ] "$65535"
[ ] 
[+] window DialogBox CopyObjectUser1
	[ ] tag "Copy Object - User"
	[ ] parent ActiveDirectoryUsers
	[-] DialogBox DialogBox1
		[ ] tag "$0"
		[+] CustomWin Icon1
			[+] multitag "[Icon]#1"
				[ ] "[Icon]$263"
		[+] StaticText CreateInText
			[+] multitag "Create in:"
				[ ] "#1"
				[ ] "$-1[1]"
		[+] TextField CreateIn1
			[+] multitag "Create in:"
				[ ] "#1"
				[ ] "$262"
		[+] CustomWin CreateIn2
			[+] multitag "[BlackFrame]Create in:"
				[ ] "[BlackFrame]#1"
				[ ] "[BlackFrame]$-1"
		[+] StaticText FirstNameText
			[+] multitag "First name:"
				[ ] "#2"
				[ ] "$-1[2]"
		[+] TextField FirstName
			[+] multitag "First name:"
				[ ] "#2"
				[ ] "$248"
		[+] StaticText InitialsText
			[+] multitag "Initials:"
				[ ] "#3"
				[ ] "$-1[3]"
		[+] TextField Initials1
			[+] multitag "Initials:"
				[ ] "#3"
				[ ] "$316"
		[+] StaticText LastNameText
			[+] multitag "Last name:"
				[ ] "#4"
				[ ] "$-1[4]"
		[+] TextField LastName
			[+] multitag "Last name:"
				[ ] "#4"
				[ ] "$249"
		[+] StaticText FullNameText
			[+] multitag "Full name:"
				[ ] "#5"
				[ ] "$-1[5]"
		[+] TextField FullName
			[+] multitag "Full name:"
				[ ] "#5"
				[ ] "$250"
		[+] StaticText UserLogonNameText
			[+] multitag "User logon name:"
				[ ] "#6"
				[ ] "$-1[6]"
		[+] TextField UserLogonName
			[+] multitag "User logon name:"
				[ ] "#6"
				[ ] "$707"
		[+] PopupList Initials2
			[+] multitag "Initials:"
				[ ] "#1"
				[ ] "$257"
		[+] TextField UserLogonNamePreWindows21
			[+] multitag "User logon name (pre-Windows 2000):[1]"
				[ ] "#7"
				[ ] "$710"
		[+] StaticText UserLogonNamePreWindows22
			[+] multitag "User logon name (pre-Windows 2000):"
				[ ] "#7"
				[ ] "$-1[7]"
		[+] TextField UserLogonNamePreWindows23
			[+] multitag "User logon name (pre-Windows 2000):[2]"
				[ ] "#8"
				[ ] "$708"
	[+] PageList PageList1
		[+] multitag "#1"
			[ ] "$12320"
	[+] StaticText StaticText1
		[+] multitag "#1"
			[ ] "$12326"
	[+] PushButton Back
		[+] multitag "< Back"
			[ ] "#1"
			[ ] "$12323"
	[+] PushButton Next
		[+] multitag "Next >"
			[ ] "#2"
			[ ] "$12324"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#3"
			[ ] "$2"
[+] // window DialogBox CopyObjectUser1
	[ ] // tag "Copy Object - User"
	[ ] // parent ActiveDirectoryUsersAndCom
	[-] // DialogBox DialogBox1
		[ ] // tag "$0"
		[+] // StaticText PasswordText
			[+] // multitag "Password:"
				[ ] // "#2"
				[ ] // "$65535[2]"
		[+] // TextField Password
			[+] // multitag "Password:"
				[ ] // "#2"
				[ ] // "$243"
		[+] // StaticText ConfirmPasswordText
			[+] // multitag "Confirm password:"
				[ ] // "#3"
				[ ] // "$65535[3]"
		[+] // TextField ConfirmPassword
			[+] // multitag "Confirm password:"
				[ ] // "#3"
				[ ] // "$244"
		[+] // CheckBox UserMustChangePasswordAtN
			[+] // multitag "User must change password at next logon"
				[ ] // "#1"
				[ ] // "$261"
		[+] // CheckBox UserCannotChangePassword
			[+] // multitag "User cannot change password"
				[ ] // "#2"
				[ ] // "$255"
		[+] // CheckBox PasswordNeverExpires
			[+] // multitag "Password never expires"
				[ ] // "#3"
				[ ] // "$258"
		[+] // CheckBox AccountIsDisabled
			[+] // multitag "Account is disabled"
				[ ] // "#4"
				[ ] // "$254"
		[+] // CustomWin Icon1
			[+] // multitag "[Icon]#1"
				[ ] // "[Icon]$263"
		[+] // StaticText CreateInText
			[+] // multitag "Create in:"
				[ ] // "#1"
				[ ] // "$65535[1]"
		[+] // TextField CreateIn1
			[+] // multitag "Create in:"
				[ ] // "#1"
				[ ] // "$262"
		[+] // CustomWin CreateIn2
			[+] // multitag "[BlackFrame]Create in:"
				[ ] // "[BlackFrame]#1"
				[ ] // "[BlackFrame]$65535"
	[+] // PageList PageList1
		[+] // multitag "#1"
			[ ] // "$12320"
	[+] // StaticText StaticText1
		[+] // multitag "#1"
			[ ] // "$12326"
	[+] // PushButton Back
		[+] // multitag "< Back"
			[ ] // "#1"
			[ ] // "$12323"
	[+] // PushButton Next
		[+] // multitag "Next >"
			[ ] // "#2"
			[ ] // "$12324"
	[+] // PushButton Cancel
		[+] // multitag "Cancel"
			[ ] // "#3"
			[ ] // "$2"
[+] window DialogBox CopyObjectUser2
	[ ] tag "Copy Object - User"
	[ ] parent ActiveDirectoryUsers
	[+] DialogBox DialogBox1
		[+] multitag "#1"
			[ ] "$0"
		[+] TextField Password
			[-] multitag "First name:"
				[ ] "#2"
				[ ] "$248"
		[+] TextField ConfirmPassword
			[-] multitag "Initials:"
				[ ] "#3"
				[ ] "$316"
		[+] TextField LastName
			[-] multitag "Last name:"
				[ ] "#4"
				[ ] "$249"
		[+] TextField FullName
			[-] multitag "Full name:"
				[ ] "#5"
				[ ] "$250"
		[+] TextField UserLogonName
			[-] multitag "User logon name:"
				[ ] "#6"
				[ ] "$707"
		[+] CheckBox UserMustChangePasswordAtN
			[+] multitag "User must change password at next logon"
				[ ] "#1"
				[ ] "$261"
		[+] CheckBox UserCannotChangePassword
			[+] multitag "User cannot change password"
				[ ] "#2"
				[ ] "$255"
		[+] CheckBox PasswordNeverExpires
			[+] multitag "Password never expires"
				[ ] "#3"
				[ ] "$258"
		[+] CheckBox AccountIsDisabled
			[+] multitag "Account is disabled"
				[ ] "#4"
				[ ] "$254"
		[ ] 
	[+] PushButton Next
		[-] multitag "Next >"
			[ ] "#2"
			[ ] "$12324"
[+] window DialogBox CopyObjectUser3
	[ ] tag "Copy Object - User"
	[ ] parent ActiveDirectoryUsers
	[+] DialogBox DialogBox1
		[ ] tag "$0"
		[+] CustomWin Icon1
			[+] multitag "[Icon]#1"
				[ ] "[Icon]$25"
		[+] StaticText CreateInText
			[+] multitag "Create in:"
				[ ] "#1"
				[ ] "$1101"
		[+] CustomWin CreateIn1
			[+] multitag "[BlackFrame]Create in:"
				[ ] "[BlackFrame]#1"
				[ ] "[BlackFrame]$1105"
		[+] TextField CreateIn2
			[+] multitag "Create in:"
				[ ] "#1"
				[ ] "$26"
		[+] CheckBox CreateAnExchangeMailbox
			[+] multitag "Create an Exchange mailbox"
				[ ] "#1"
				[ ] "$1009"
		[+] StaticText AliasText
			[+] multitag "Alias:"
				[ ] "#2"
				[ ] "$2003"
		[+] TextField Alias
			[+] multitag "Alias:"
				[ ] "#2"
				[ ] "$1003"
		[+] StaticText ServerText
			[+] multitag "Server:"
				[ ] "#3"
				[ ] "$1104"
		[+] PopupList Server
			[+] multitag "Server:"
				[ ] "#1"
				[ ] "$1120"
		[+] StaticText MailboxStoreText
			[+] multitag "Mailbox Store:"
				[ ] "#4"
				[ ] "$2007"
		[+] PopupList MailboxStore
			[+] multitag "Mailbox Store:"
				[ ] "#2"
				[ ] "$1007"
	[+] PageList PageList1
		[+] multitag "#1"
			[ ] "$12320"
	[+] StaticText StaticText1
		[+] multitag "#1"
			[ ] "$12326"
	[+] PushButton Back
		[+] multitag "< Back"
			[ ] "#1"
			[ ] "$12323"
	[+] PushButton Next
		[+] multitag "Next >"
			[ ] "#2"
			[ ] "$12324"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#3"
			[ ] "$2"
[+] window DialogBox UserAlreadyExists
	[ ] tag "Active Directory"
	[ ] parent CopyObjectUser1
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$2"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$20"
	[+] StaticText TheUserLogonNameYouHaveC
		[+] multitag "The user logon name you have chosen is already in use in this enterprise.??Choose another logon name, and then try again."
			[ ] "#1"
			[ ] "$65535"
[+] window DialogBox HomeDirectoryExists
	[ ] tag "Active Directory"
	[ ] parent CopyObjectUser1
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$2"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$20"
	[+] StaticText TheBLK0520DC01Composeusers
		[+] multitag "The \\BLK0520DC01\composeusers\standard11 home folder was not created because it already exists. The user account has been cre*"
			[ ] "#1"
			[ ] "$65535"
[ ] 
[+] window DialogBox UserProperties
	[ ] tag "*Properties"
	[+] DialogBox DialogBox1
		[ ] tag "$0"
		[+] StaticText UserLogonNameText
			[+] multitag "User logon name:"
				[ ] "#1"
				[ ] "$-1[1]"
		[+] TextField UserLogonName
			[+] multitag "User logon name:"
				[ ] "#1"
				[ ] "$355"
		[+] PopupList PopupList1
			[+] multitag "#1"
				[ ] "$1526"
		[+] TextField UserLogonNamePreWindows21
			[+] multitag "User logon name (pre-Windows 2000):[1]"
				[ ] "#2"
				[ ] "$357"
		[+] StaticText UserLogonNamePreWindows22
			[+] multitag "User logon name (pre-Windows 2000):"
				[ ] "#2"
				[ ] "$-1[2]"
		[+] TextField UserLogonNamePreWindows23
			[+] multitag "User logon name (pre-Windows 2000):[2]"
				[ ] "#3"
				[ ] "$358"
		[+] PushButton LogonHours
			[+] multitag "Logon Hours"
				[ ] "#1"
				[ ] "$166"
		[+] PushButton LogOnTo
			[+] multitag "Log On To"
				[ ] "#2"
				[ ] "$167"
		[+] CheckBox AccountIsLockedOut
			[+] multitag "Account is locked out"
				[ ] "#1"
				[ ] "$158"
		[+] StaticText AccountOptionsText
			[+] multitag "Account options:"
				[ ] "#3"
				[ ] "$3391"
		[+] CustomWin AccountOptions
			[+] multitag "[CHECKLIST]Account options:"
				[ ] "[CHECKLIST]#1"
				[ ] "[CHECKLIST]$175"
			[+] StaticText UserMustChangePasswordAtN
				[+] multitag "User must change password at next logon"
					[ ] "#1"
					[ ] "$0"
			[+] CheckBox CheckBox1
				[+] multitag "#1"
					[ ] "$1"
			[+] StaticText UserCannotChangePasswordTex
				[+] multitag "User cannot change password"
					[ ] "#2"
					[ ] "$4"
			[+] CheckBox CheckBox2
				[+] multitag "#2"
					[ ] "$5"
			[+] StaticText PasswordNeverExpiresText
				[+] multitag "Password never expires"
					[ ] "#3"
					[ ] "$8"
			[+] CheckBox CheckBox3
				[+] multitag "#3"
					[ ] "$9"
			[+] StaticText StorePasswordUsingReversibl
				[+] multitag "Store password using reversible encryption"
					[ ] "#4"
					[ ] "$12"
			[+] CheckBox CheckBox4
				[+] multitag "#4"
					[ ] "$13"
			[+] StaticText AccountIsDisabledText
				[+] multitag "Account is disabled"
					[ ] "#5"
					[ ] "$16"
			[+] CheckBox CheckBox5
				[+] multitag "#5"
					[ ] "$17"
			[+] StaticText SmartCardIsRequiredForInt
				[+] multitag "Smart card is required for interactive logon"
					[ ] "#6"
					[ ] "$20"
			[+] CheckBox CheckBox6
				[+] multitag "#6"
					[ ] "$21"
			[+] StaticText AccountIsSensitiveAndCanno
				[+] multitag "Account is sensitive and cannot be delegated"
					[ ] "#7"
					[ ] "$24"
			[+] CheckBox CheckBox7
				[+] multitag "#7"
					[ ] "$25"
			[+] StaticText UseDESEncryptionTypesForT
				[+] multitag "Use DES encryption types for this account"
					[ ] "#8"
					[ ] "$28"
			[+] CheckBox CheckBox8
				[+] multitag "#8"
					[ ] "$29"
			[+] StaticText DoNotRequireKerberosPreaut
				[+] multitag "Do not require Kerberos preauthentication"
					[ ] "#9"
					[ ] "$32"
			[+] CheckBox CheckBox9
				[+] multitag "#9"
					[ ] "$33"
		[+] StaticText AccountExpiresText
			[+] multitag "Account expires"
				[ ] "#4"
				[ ] "$-1[3]"
		[+] RadioList AccountExpires
			[+] multitag "Account expires"
				[ ] "#1"
				[ ] "$161"
		[+] CustomWin ThursdayApril172008
			[+] multitag "[SysDateTimePick32]Thursday, April 17, 2008"
				[ ] "[SysDateTimePick32]#1"
				[ ] "[SysDateTimePick32]$163"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#2"
			[ ] "$2"
	[+] PushButton Apply
		[+] multitag "Apply"
			[ ] "#3"
			[ ] "$12321"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "#4"
			[ ] "$9"
	[+] PageList PageList1
		[+] multitag "#1"
			[ ] "$12320"
[ ] 
[+] window TaskbarWin Taskbar
	[ ] tag "#1"
	[+] PushButton Start
		[+] multitag "Start"
			[ ] "#1"
			[ ] "$304"
	[+] CustomWin TrayNotifyWnd1
		[+] multitag "[TrayNotifyWnd]#1"
			[ ] "[TrayNotifyWnd]$303"
		[+] CustomWin N241PM
			[+] multitag "[TrayClockWClass]2:41 PM"
				[ ] "[TrayClockWClass]#1"
				[ ] "[TrayClockWClass]$303"
		[+] CustomWin SysPager1
			[+] multitag "[SysPager]#1"
				[ ] "[SysPager]$0"
			[+] ToolBar ToolBar1
				[+] multitag "#1"
					[ ] "$0"
				[+] PushButton SymantecAntiVirus
					[+] multitag "Symantec AntiVirus"
						[ ] "#1"
						[ ] "$0"
		[+] PushButton PushButton1
			[+] multitag "#1"
				[ ] "$1502"
	[+] CustomWin ReBarWindow321
		[+] multitag "[ReBarWindow32]#1"
			[ ] "[ReBarWindow32]$40965"
		[+] ToolBar ToolBar2
			[+] multitag "#2"
				[ ] "$0[2]"
			[+] CheckBox SilkTest
				[+] multitag "SilkTest"
					[ ] "#1"
					[ ] "$3"
			[+] CheckBox RecordWindowDeclarations
				[+] multitag "Record Window Declarations"
					[ ] "#2"
					[ ] "$1"
			[+] CheckBox CDocumentsAndSettingsAdmi
				[+] multitag "C:\Documents and Settings\Administrator\Start Menu\Programs\Administrative Tool"
					[ ] "#3"
					[ ] "$9"
			[+] CheckBox AgentTCPIPBLK0520DC012965
				[+] multitag "Agent ?TCP?IP BLK0520DC01 2965?"
					[ ] "#4"
					[ ] "$4"
		[+] ToolBar ToolBar1
			[+] multitag "#1"
				[ ] "$0[1]"
			[+] PushButton ShowDesktopScf
				[+] multitag "Show Desktop.scf"
					[ ] "#1"
					[ ] "$0"
			[+] PushButton LaunchInternetExplorerBrows
				[+] multitag "Launch Internet Explorer Browser.lnk"
					[ ] "#2"
					[ ] "$1"
			[+] PushButton SilkTestLnk
				[+] multitag "SilkTest.lnk"
					[ ] "#3"
					[ ] "$2"
[+] window DialogBox StartMenu
	[ ] tag "Start Menu"
	[ ] parent Taskbar
	[+] CustomWin DesktopSFTBarHost1
		[-] multitag "[DesktopSFTBarHost]#1"
			[ ] "$1"
		[-] ListView ListView1
			[-] multitag "#1"
				[ ] "$0"
[ ] 
[-] window DialogBox ODBCDataSourceAdministrator
	[ ] tag "ODBC Data Source Administrator"
	[ ] parent MicrosoftODBCAdministrator
	[-] DialogBox DialogBox1
		[ ] tag "$0"
		[+] StaticText UserDataSources1Text
			[+] multitag "User Data Sources:[1]"
				[ ] "#1"
				[ ] "$65535[1]"
		[+] ListView UserDataSources
			[+] multitag "User Data Sources:"
				[ ] "#1"
				[ ] "$1872"
			[+] Header Header1
				[+] multitag "#1"
					[ ] "$0"
				[+] PushButton Name
					[+] multitag "Name"
						[ ] "#1"
				[+] PushButton Driver
					[+] multitag "Driver"
						[ ] "#2"
		[+] PushButton Add
			[+] multitag "Add"
				[ ] "#1"
				[ ] "$1840"
		[+] PushButton Remove
			[+] multitag "Remove"
				[ ] "#2"
				[ ] "$1870"
		[+] PushButton Configure
			[+] multitag "Configure"
				[ ] "#3"
				[ ] "$1881"
		[+] StaticText AnODBCUserDataSourceStore
			[+] multitag "An ODBC User data source stores information about how to connect to the indicated data provider.   A User data source is only *"
				[ ] "#3"
				[ ] "$65535[3]"
		[+] CustomWin Icon1
			[+] multitag "[Icon]#1"
				[ ] "[Icon]$65535"
		[+] StaticText UserDataSources2Text
			[+] multitag "User Data Sources:[2]"
				[ ] "#2"
				[ ] "$65535[2]"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#2"
			[ ] "$2"
	[+] PushButton Apply
		[+] multitag "Apply"
			[ ] "#3"
			[ ] "$12321"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "#4"
			[ ] "$9"
	[+] PageList PageList1
		[-] multitag "#1"
			[ ] "$12320"
[ ] 
[+] window DialogBox DataSourcesODBCProperties
	[ ] tag "Data Sources (ODBC) Properties"
[+] window DialogBox MicrosoftODBCAdministrator
	[ ] tag "Microsoft ODBC Administrator"
[+] window DialogBox CreateNewDataSource
	[ ] tag "Create New Data Source"
	[ ] parent ODBCDataSourceAdministrator
	[-] DialogBox CreateNewDataSource
		[+] multitag "Create New Data Source"
			[ ] "$0"
		[+] StaticText SelectADriverForWhichYou1
			[+] multitag "Select a driver for which you want to set up a data source."
				[ ] "#1"
				[ ] "$65535"
		[-] ListView SelectADriverForWhichYou2
			[+] multitag "Select a driver for which you want to set up a data source."
				[ ] "#1"
				[ ] "$1855"
			[+] Header Header1
				[+] multitag "#1"
					[ ] "$0"
				[+] PushButton Name
					[+] multitag "Name"
						[ ] "#1"
				[+] PushButton Version
					[+] multitag "Version"
						[ ] "#2"
				[+] PushButton Company
					[+] multitag "Company"
						[ ] "#3"
				[+] PushButton File
					[+] multitag "File"
						[ ] "#4"
				[+] PushButton xDate
					[+] multitag "Date"
						[ ] "#5"
		[+] CustomWin BitMap1
			[+] multitag "[BitMap]#1"
				[ ] "[BitMap]$65535"
	[+] PageList PageList1
		[+] multitag "#1"
			[ ] "$12320"
	[+] StaticText StaticText1
		[+] multitag "#1"
			[ ] "$12326"
	[+] PushButton Back
		[+] multitag "< Back"
			[ ] "#1"
			[ ] "$12323"
	[+] PushButton Finish
		[+] multitag "Finish"
			[ ] "#2"
			[ ] "$12325"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#3"
			[ ] "$2"
[+] window DialogBox ODBCMicrosoftExcelSetup
	[ ] tag "ODBC Microsoft Excel Setup"
	[ ] parent ODBCDataSourceAdministrator
	[+] StaticText DataSourceNameText
		[+] multitag "Data Source Name:"
			[ ] "#1"
			[ ] "$1090"
	[+] TextField DataSourceName
		[+] multitag "Data Source Name:"
			[ ] "#1"
			[ ] "$1015"
	[+] StaticText DescriptionText
		[+] multitag "Description:"
			[ ] "#2"
			[ ] "$1091"
	[+] TextField Description
		[+] multitag "Description:"
			[ ] "#2"
			[ ] "$1026"
	[+] StaticText DatabaseText
		[+] multitag "Database"
			[ ] "#3"
			[ ] "$1079"
	[+] StaticText VersionText
		[+] multitag "Version:"
			[ ] "#4"
			[ ] "$1106"
	[+] PopupList Version
		[+] multitag "Version:"
			[ ] "#1"
			[ ] "$1066"
	[+] StaticText Workbook1Text
		[+] multitag "Workbook:[1]"
			[ ] "#5"
			[ ] "$1076"
	[+] StaticText Workbook2Text
		[+] multitag "Workbook:[2]"
			[ ] "#6"
			[ ] "$115"
	[+] PushButton SelectWorkbook
		[+] multitag "Select Workbook"
			[ ] "#4"
			[ ] "$116"
	[+] CheckBox UseCurrentDirectory
		[+] multitag "Use Current Directory"
			[ ] "#1"
			[ ] "$1056"
	[+] StaticText DriverText
		[+] multitag "Driver"
			[ ] "#7"
			[ ] "$1078"
	[+] StaticText RowsToScanText
		[+] multitag "Rows to Scan:"
			[ ] "#8"
			[ ] "$-1"
	[+] TextField RowsToScan
		[+] multitag "Rows to Scan:"
			[ ] "#3"
			[ ] "$129"
	[+] CheckBox ReadOnly
		[+] multitag "Read Only"
			[ ] "#2"
			[ ] "$1018"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#2"
			[ ] "$2"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "#3"
			[ ] "$1011"
	[+] PushButton Options
		[+] multitag "Options>>"
			[ ] "#5"
			[ ] "$134"
[+] window DialogBox Error
	[ ] tag "Error"
	[ ] parent ODBCDataSourceAdministrator
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "#1"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "#2"
			[ ] "$7"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#3"
			[ ] "$2"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$20"
	[+] StaticText TheDatasourceNamedSilkbA
		[+] multitag "The datasource named 'Silkb' already exists.  Replace it with this definition?"
			[ ] "#1"
			[ ] "$65535"
[ ] ////////////////////////////////////////////////////////////////////////////
[ ] // MICROSOFT MANAGEMENT CONSOLE
[ ] ///////////////////////////////////////////////////////////////////////////
[+] window DialogBox ComputerManagement
	[ ] tag "Computer Management"
	[+] CustomWin SizeableRebar1
		[+] multitag "[SizeableRebar]#1"
			[ ] "[SizeableRebar]$13341"
		[+] CustomWin ReBarWindow321
			[+] multitag "[ReBarWindow32]#1"
				[ ] "[ReBarWindow32]$4096"
			[+] ToolBar ToolBar2
				[+] multitag "#2"
					[ ] "$4098"
				[+] PushButton PushButton1
					[+] multitag "#1"
						[ ] "$1"
				[+] PushButton PushButton2
					[+] multitag "#2"
						[ ] "$2"
				[+] PushButton PushButton3
					[+] multitag "#3"
						[ ] "$4"
				[+] PushButton PushButton4
					[+] multitag "#4"
						[ ] "$5"
				[+] PushButton PushButton5
					[+] multitag "#5"
						[ ] "$17"
			[+] ToolBar ToolBar1
				[+] multitag "#1"
					[ ] "$4097"
				[+] PushButton PushButton1
					[+] multitag "#1"
						[ ] "$21508"
				[+] PushButton File
					[+] multitag "File"
						[ ] "#2"
						[ ] "$21513"
				[+] PushButton Action
					[+] multitag "Action"
						[ ] "#3"
						[ ] "$21509"
				[+] PushButton View
					[+] multitag "View"
						[ ] "#4"
						[ ] "$21510"
				[+] PushButton xWindow
					[+] multitag "Window"
						[ ] "#5"
						[ ] "$21514"
				[+] PushButton Help
					[+] multitag "Help"
						[ ] "#6"
						[ ] "$21515"
			[+] CustomWin AfxWnd42u1
				[+] multitag "[AfxWnd42u]#1"
					[ ] "[AfxWnd42u]$8193"
[+] window ChildWin ComputerManagementLocal
	[ ] tag "Computer Management (Local)"
	[ ] parent ComputerManagement
	[-] CustomWin MMCViewWindow1
		[+] multitag "[MMCViewWindow]#1"
			[ ] "[MMCViewWindow]$59648"
		[+] ListView ListView1
			[+] multitag "#1"
				[ ] "$12786"
			[+] Header Header1
				[+] multitag "#1"
					[ ] "$0"
				[+] PushButton Name
					[+] multitag "Name"
						[ ] "#1"
		[+] TreeView TreeView1
			[+] multitag "#1"
				[ ] "$12785"
	[+] StatusBar StatusBar1
		[+] multitag "#1"
			[ ] "$4099"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
		[+] StaticText StaticText1
			[+] multitag "#1"
				[ ] "$4097"
[ ] 
[ ] // DISK DEFRAGMENTOR
[+] window ChildWin ComputerManagementDiskDefragPane
	[ ] tag "Computer Management (Local)\Storage\Disk Defragmenter"
	[ ] parent ComputerManagement
	[+] CustomWin MMCViewWindow1
		[+] multitag "[MMCViewWindow]#1"
			[ ] "[MMCViewWindow]$59648"
		[+] CustomWin MMCOCXViewWindow1
			[+] multitag "[MMCOCXViewWindow]#1"
				[ ] "[MMCOCXViewWindow]$12789"
			[+] CustomWin AtlAxWinEx1
				[+] multitag "[AtlAxWinEx]#1"
					[ ] "[AtlAxWinEx]$12578784"
				[+] CustomWin DiskDefragmenter
					[+] multitag "[ATL:6D42AC58]Disk Defragmenter"
						[ ] "[ATL:6D42AC58]#1"
						[ ] "[ATL:6D42AC58]$8881864"
					[+] PushButton Analyze
						[+] multitag "Analyze"
							[ ] "#1"
							[ ] "$8100"
					[+] PushButton Defragment
						[+] multitag "Defragment"
							[ ] "#2"
							[ ] "$8101"
					[+] PushButton Pause
						[+] multitag "Pause"
							[ ] "#3"
							[ ] "$8103"
					[+] PushButton Stop
						[+] multitag "Stop"
							[ ] "#4"
							[ ] "$8102"
					[+] PushButton ViewReport
						[+] multitag "View Report"
							[ ] "#5"
							[ ] "$8500"
					[+] ListView ListView1
						[+] multitag "#1"
							[ ] "$0"
						[+] Header Header1
							[+] multitag "#1"
								[ ] "$0"
							[+] PushButton Volume
								[+] multitag "Volume"
									[ ] "#1"
							[+] PushButton SessionStatus
								[+] multitag "Session Status"
									[ ] "#2"
							[+] PushButton FileSystem
								[+] multitag "File System"
									[ ] "#3"
							[+] PushButton Capacity
								[+] multitag "Capacity"
									[ ] "#4"
							[+] PushButton FreeSpace1
								[+] multitag "Free Space"
									[ ] "#5"
							[+] PushButton FreeSpace2
								[+] multitag "% Free Space"
									[ ] "#6"
		[+] TreeView TreeView1
			[+] multitag "#1"
				[ ] "$12785"
	[+] StatusBar StatusBar1
		[+] multitag "#1"
			[ ] "$4099"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
		[+] StaticText StaticText1
			[+] multitag "#1"
				[ ] "$4097"
[ ] 
[+] window DialogBox DiskDefragmenter
	[ ] tag "Disk Defragmenter"
	[ ] parent ComputerManagement
	[+] PushButton ViewReport
		[+] multitag "View Report"
			[ ] "#1"
			[ ] "$1021"
	[+] PushButton Close
		[+] multitag "Close"
			[ ] "#2"
			[ ] "$1020"
	[+] TextField TextField1
		[+] multitag "#1"
			[ ] "$1011"
[+] window DialogBox DefragmentationReport
	[ ] tag "Defragmentation Report"
	[ ] parent ComputerManagement
	[+] CustomWin EtchedHorz1
		[+] multitag "[EtchedHorz]#1"
			[ ] "[EtchedHorz]$1036"
	[+] StaticText VolumeInformationText
		[+] multitag "Volume information:"
			[ ] "#1"
			[ ] "$1004"
	[+] ListView VolumeInformation
		[+] multitag "Volume information:"
			[ ] "#1"
			[ ] "$203"
		[+] Header Header1
			[+] multitag "#1"
				[ ] "$0"
	[+] StaticText FilesThatDidNotDefragment1
		[+] multitag "Files that did not defragment:"
			[ ] "#2"
			[ ] "$1006"
	[+] ListView FilesThatDidNotDefragment2
		[+] multitag "Files that did not defragment:"
			[ ] "#2"
			[ ] "$204"
		[+] Header Header1
			[+] multitag "#1"
				[ ] "$0"
			[+] PushButton Fragments
				[+] multitag "Fragments"
					[ ] "#1"
			[+] PushButton FileSize
				[+] multitag "File Size"
					[ ] "#2"
			[+] PushButton FileName
				[+] multitag "File Name"
					[ ] "#3"
	[+] PushButton Print
		[+] multitag "Print"
			[ ] "#1"
			[ ] "$1014"
	[+] PushButton SaveAs
		[+] multitag "Save As"
			[ ] "#2"
			[ ] "$1013"
	[+] PushButton Close
		[+] multitag "Close"
			[ ] "#3"
			[ ] "$2"
	[+] TextField TextField1
		[+] multitag "#1"
			[ ] "$1011"
[ ] 
[ ] 
[ ] 
[ ] 
