[ ] use "winclass.inc" 
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] // Global Const Paths for Datafiles
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] const sGlobalReadPath = "c:\SilkTestData\"
[ ] const sGlobalVBReadPath = "c:\vb\"
[ ] const sGlobalWritePath = "C:\"
[ ] const sGlobalAdminWritePath = "T:\"
[ ] const sGlobalComposeDLLPath = "C:\Program Files\COMPOSE\"
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] // Server IP Consts
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] const DCO1 = "149.199.32.70"
[ ] const DCO2 = "149.199.32.71"
[ ] const CSO1 = "149.199.32.72"
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] STRING sCurrentMachine = DC01
[ ] BOOLEAN bSAVAsync
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] // Test Case Globals
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] INT iTestsPassedCount = 0
[ ] INT iTestsFailedCount = 0
[ ] LIST OF STRING lsFailedFileOpenTests
[ ] // Set this variable to show all verification info
[ ] BOOLEAN VERBOSE = FALSE
[ ] // If REGISTRYEXCLUSION == TRUE the entries in 
[ ] //  the RegistryListList worksheet in Centrixs.xls will
[ ] // cause the script the verification of those entries
[ ] // If REGISTRYEXCLUSION == FALSE  the entries in
[ ] //  the RegistryListList worksheet in Centrixs.xls will
[ ] // a  tab delimited
[ ] // file called RegistryExclusionList.csv will be written to the sGlobalWritePath 
[ ] BOOLEAN REGISTRYEXCLUSION = TRUE
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] // Global IP Addresses for servers
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] const STRING DC01 = "140.199.32.70"
[ ] const STRING DC02 = "140.199.32.71"
[ ] const STRING CS01 = "140.199.32.72"
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] //CONSTS FOR EVENT LOG INFO
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] const INT EVENTLOGERROR = 1
[ ] const INT EVENTLOGWARNING = 2
[ ] const INT EVENTLOGINFORMATION = 3
[ ] const INT EVENTLOGSECURITYAUDITSUCCESS = 4
[ ] const INT EVENTLOGSECURITYAUDITFAILURE = 5
[ ] /////////////////////////////////////////////////////////////////////////////////////////
[ ] // AGENT OPTIONS
[ ] // Agent.SetOption(OPT_WINDOW_TIMEOUT,35)
[ ] ////////////////////////////////////////////////////////////////////////////////////////
[ ] //  CONSTANTS
[ ] ///////////////////////////////////////////////////////////////////
[ ] // VB Constants for Outlook
[ ] const olAppointmentItem = 1
[ ] const olContactItem = 2
[ ] const olJournalItem = 4 
[ ] const olMailItem = 0
[ ] const olNoteItem = 5
[ ] const olPostItem = 6
[ ] const olTaskItem = 3 
[ ] ///////////////////////////////////////////////////////////////////
[ ] // LogonUser Constants
[ ] const int logon32_logon_network_cleartext = 8
[ ] const int logon32_provider_default = 0
[ ] ///////////////////////////////////////////////////////////////////
[ ] //  GLOBALS
[ ] ///////////////////////////////////////////////////////////////////
[ ] STRING sMailSubject
[ ] STRING sCurrentUser
[+] LIST OF STRING sRegistryKeys = {...}
	[ ] "HKEY_CLASSES_ROOT"
	[ ] "HKEY_CURRENT_USER"
	[ ] "HKEY_LOCAL_MACHINE"
	[ ] "HKEY_USERS"
	[ ] "HKEY_PERFORMANCE_DATA"
	[ ] "HKEY_CURRENT_CONFIG"
	[ ] "HKEY_DYN_DATA"
[ ] ///////////////////////////////////////////////////////////////////
[ ] //  TYPES
[ ] ///////////////////////////////////////////////////////////////////
[+] type SENDMAILREC is record
	[ ] STRING sFrom
	[ ] STRING sTo
	[ ] STRING sCc
	[ ] STRING sSubject
	[ ] STRING sAttachment
[+] type SERVICEREC is record
	[ ] STRING sDisplayName
	[ ] STRING sState
	[ ] STRING sStartMode
[+] type VERSIONREC is record
	[ ] STRING sFilename
	[ ] STRING sVersion
	[ ] 
[+] type REGISTRYKEYSREC is record
	[ ] STRING sKey
	[ ] INT iConst
[+] type REGISTRYENTRYREC is record
	[ ] // STRING sRootKey
	[ ] STRING sPath
	[ ] STRING sEntry
	[ ] STRING sException
[+] type ADMINISTRATIVETEMPLATEREC is record
	[ ] STRING sGPO
	[ ] STRING sKeyName
	[ ] STRING sValue
	[ ] STRING sState
[ ] ///////////////////////////////////////////////////////////////////
[ ] //  DLL
[ ] ///////////////////////////////////////////////////////////////////
[+] dll "shell32"	
	[ ] int ShellExecuteW (int hwnd, String lpOperation, String lpFile, string lpParameters,String lpDirectory, int nShowCmd)
	[ ] // Function Description	
	[ ] // This function is going to execute an application and allow SilkTest to continue running	
	[ ] // before application has closed.	
	[ ] // Input parameter
	[ ] // STRING sProgramPath = specifies absolute path to the program that needs to be executed
	[ ] // return		
	[ ] // BOOLEAN = TRUE if program was exectued, FALSE if execution failed
[+] dll "Advapi32"
	[ ] INT CreateProcessWithLogonW(STRING lpUsername, STRING lpDomain, STRING lpPassword, INT dwLogonFlags, STRING lpApplicationName, STRING lpCommandLine optional, INT dwCreationFlags optional, STRING lpEnvironment optional, STRING lpCurrentDirectory optional, STRING lpStartupInfo optional,  out STRING lpProcessInfo)
	[ ] INT LogonUserW(STRING lpszUsername, STRING lpszDomain, STRING lpszPassword, INT dwLogonType, INT dwLogonProvider, out STRING phToken)
	[ ] 
[ ] ///////////////////////////////////////////////////////////////////
[ ] //  GENERIC/MISC WINDOWS
[ ] ///////////////////////////////////////////////////////////////////
[+] window DialogBox Desktop1
	[ ] tag "Desktop[1]"
	[+] CustomWin WorkerW1
		[+] multitag "[WorkerW]#1"
			[ ] "$40965"
		[+] CustomWin ReBarWindow321
			[+] multitag "[ReBarWindow32]#1"
				[ ] "$40965"
			[+] ToolBar ToolBar1
				[+] multitag "#1"
					[ ] "$0"
[+] window DialogBox MapNetworkDrive
	[ ] tag "Map Network Drive"
	[+] DialogBox MapNetworkDrive
		[+] multitag "Map Network Drive"
			[ ] "$0"
		[+] CustomWin BitMap1
			[ ] tag "[BitMap]$-1"
		[+] StaticText WindowsCanHelpYouConnectT
			[+] multitag "Windows can help you connect to a shared network folder and assign a drive letter to the connection so that you can access th*"
				[ ] "$-1[1]"
		[+] StaticText SpecifyTheDriveLetterForT
			[+] multitag "Specify the drive letter for the connection and the folder that you want to connect to:"
				[ ] "$-1[2]"
		[+] StaticText DriveText
			[+] multitag "Drive:"
				[ ] "$-1[3]"
		[+] PopupList Drive
			[+] multitag "Drive:"
				[ ] "$1036"
		[+] StaticText FolderText
			[+] multitag "Folder:"
				[ ] "$-1[4]"
		[+] ComboBox Folder
			[+] multitag "Folder:"
				[ ] "$1034"
		[+] PushButton Browse
			[+] multitag "Browse"
				[ ] "$1035"
		[+] StaticText ExampleServerShareText
			[+] multitag "Example: \\server\share"
				[ ] "$-1[5]"
		[+] CheckBox ReconnectAtLogon
			[+] multitag "Reconnect at logon"
				[ ] "$1038"
		[+] CustomWin ConnectUsingAADifferentU
			[+] multitag "[Link Window]Connect using a <A>different user name<?A>."
				[ ] "[Link Window]$1044"
		[+] CustomWin ASignUpForOnlineStorage
			[+] multitag "[Link Window]<A>Sign up for online storage or connect to a network server<?A>."
				[ ] "[Link Window]$1045"
	[+] PageList PageList1
		[ ] tag "$12320"
	[+] StaticText StaticText1
		[ ] tag "$12326"
	[+] PushButton Back
		[+] multitag "< Back"
			[ ] "$12323"
	[+] PushButton Finish
		[+] multitag "Finish"
			[ ] "$12325"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
[+] window DialogBox MyDocuments
	[ ] tag "My Documents"
	[+] CustomWin WorkerW1
		[ ] tag "[WorkerW]$40965"
		[+] CustomWin ReBarWindow321
			[ ] tag "[ReBarWindow32]$40965"
			[+] ComboBox ComboBox1
				[ ] tag "$41477"
			[+] ToolBar ToolBar2
				[ ] tag "$40960"
				[+] PushButton Back
					[+] multitag "Back"
						[ ] "$1010"
				[+] PushButton Forward
					[+] multitag "Forward"
						[ ] "$1011"
				[+] PushButton Up
					[+] multitag "Up"
						[ ] "$1012"
				[+] PushButton Search
					[+] multitag "Search"
						[ ] "$1013"
				[+] PushButton Folders
					[+] multitag "Folders"
						[ ] "$1014"
				[+] PushButton MoveTo
					[+] multitag "Move To"
						[ ] "$28703"
				[+] PushButton CopyTo
					[+] multitag "Copy To"
						[ ] "$28702"
				[+] PushButton Delete
					[+] multitag "Delete"
						[ ] "$28689"
			[+] CustomWin WorkerW1
				[ ] tag "[WorkerW]$0"
			[+] ToolBar ToolBar1
				[ ] tag "$0"
				[+] PushButton File
					[+] multitag "File"
						[ ] "$32768"
				[+] PushButton Edit
					[+] multitag "Edit"
						[ ] "$32832"
				[+] PushButton View
					[+] multitag "View"
						[ ] "$32896"
				[+] PushButton Favorites
					[+] multitag "Favorites"
						[ ] "$33136"
				[+] PushButton Tools
					[+] multitag "Tools"
						[ ] "$32960"
				[+] PushButton Help
					[+] multitag "Help"
						[ ] "$33024"
	[+] ListView ListView1
		[ ] tag "$1"
		[+] Header Header1
			[ ] tag "$0"
			[+] PushButton Name
				[ ] tag "Name"
			[+] PushButton Size
				[ ] tag "Size"
			[+] PushButton Type
				[ ] tag "Type"
			[+] PushButton DateModified
				[ ] tag "Date Modified"
	[+] CustomWin BaseBar1
		[ ] tag "[BaseBar]$0"
		[+] CustomWin ReBarWindow321
			[ ] tag "[ReBarWindow32]$40965"
			[+] ToolBar ToolBar1
				[ ] tag "$0"
				[+] PushButton Close
					[+] multitag "Close"
						[ ] "$17"
			[+] TreeView TreeView1
				[ ] tag "$100"
[+] window MessageBoxClass MessageBox
	[ ] tag "~ActiveApp/[DialogBox]$MessageBox"
	[+] PushButton OK
		[ ] tag "OK"
	[+] PushButton Cancel
		[ ] tag "Cancel"
	[+] PushButton Yes
		[ ] tag "Yes"
	[+] PushButton No
		[ ] tag "No"
	[+] StaticText Message
		[ ] motif tag "#2"
		[ ] tag "#1"
[+] window DialogBox WelcomeToWindows 
	[ ] STRING sTag
	[ ] tag sTag
	[+] PushButton OK
		[ ] tag "OK"
	[+] TextField Username
		[ ] tag "#1"
	[+] TextField Password
		[ ] tag "#2"
[+] window DialogBox Desktop
	[+] multitag "Desktop"
		[ ] "$MessageBox[1]"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$20"
	[+] StaticText HSilkaProfDataDesktopRefe
		[+] multitag "H:\silka\ProfData\Desktop refers to a location that is unavailable. It could be on a hard drive on this computer, or on a netw*"
			[ ] "$65535"
[+] window DialogBox ExplorerPathNotFound
	[ ] tag "Desktop*"
	[+] // multitag "Windows Explorer"
		[ ] // "$MessageBox[1]"
		[ ] // "Desktop*"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$20"
	[+] StaticText Message
		[+] multitag "H:\silkb refers to a location that is unavailable. It could be on a hard drive on this computer, or on a network. Check to mak*"
			[ ] "$65535"
[+] window MainWin CMD
	[ ] STRING sTag
	[ ] tag sTag
[+] window MainWin Notepad
	[+] multitag "*Notepad"
		[ ] "Untitled*"
		[ ] "$C:\WINDOWS\system32\notepad.exe"
	[+] Menu File
		[ ] tag "File"
		[+] MenuItem New
			[+] multitag "New"
				[ ] "$1"
		[+] MenuItem Open
			[+] multitag "Open"
				[ ] "$2"
		[+] MenuItem Save
			[+] multitag "Save"
				[ ] "$3"
		[+] MenuItem SaveAs
			[+] multitag "Save As"
				[ ] "$4"
		[+] MenuItem PageSetup
			[+] multitag "Page Setup"
				[ ] "$5"
		[+] MenuItem Print
			[+] multitag "Print"
				[ ] "$6"
		[+] MenuItem Exit
			[+] multitag "Exit"
				[ ] "$7"
	[+] Menu Edit
		[ ] tag "Edit"
		[+] MenuItem Undo
			[+] multitag "Undo"
				[ ] "$16"
		[+] MenuItem Cut
			[+] multitag "Cut"
				[ ] "$768"
		[+] MenuItem Copy
			[+] multitag "Copy"
				[ ] "$769"
		[+] MenuItem Paste
			[+] multitag "Paste"
				[ ] "$770"
		[+] MenuItem Delete
			[+] multitag "Delete"
				[ ] "$771"
		[+] MenuItem Find
			[+] multitag "Find"
				[ ] "$21"
		[+] MenuItem FindNext
			[+] multitag "Find Next"
				[ ] "$22"
		[+] MenuItem Replace
			[+] multitag "Replace"
				[ ] "$23"
		[+] MenuItem GoTo
			[+] multitag "Go To"
				[ ] "$24"
		[+] MenuItem SelectAll
			[+] multitag "Select All"
				[ ] "$25"
		[+] MenuItem TimeDate
			[+] multitag "Time?Date"
				[ ] "$26"
	[+] Menu Format
		[ ] tag "Format"
		[+] MenuItem WordWrap
			[+] multitag "Word Wrap"
				[ ] "$32"
		[+] MenuItem Font
			[+] multitag "Font"
				[ ] "$33"
	[+] Menu View
		[ ] tag "View"
		[+] MenuItem StatusBar
			[+] multitag "Status Bar"
				[ ] "$27"
	[+] Menu Help
		[ ] tag "Help"
		[+] MenuItem HelpTopics
			[+] multitag "Help Topics"
				[ ] "$64"
		[+] MenuItem AboutNotepad
			[+] multitag "About Notepad"
				[ ] "$65"
	[+] TextField TextField1
		[ ] tag "$15"
[ ] ///////////////////////////////////////////////////////////////////
[ ] //  WINDOW DECLARATIONS
[ ] ///////////////////////////////////////////////////////////////////
[ ] // Outlook Declarations
[+] window MainWin Outlook
	[+] multitag "Microsoft Outlook"
		[ ] "*Microsoft Outlook"
		[ ] "$C:\Program Files\Microsoft Office\Office11\OUTLOOK.exe"
[+] window DialogBox MicrosoftOutlook
	[ ] tag "*Microsoft Outlook"
	[+] CustomWin AfxWndW1
		[+] multitag "[AfxWndW]#1"
			[ ] "$103"
		[+] CustomWin WunderBar
			[+] multitag "[WUNDERBAR]WunderBar"
				[ ] "$1"
			[+] CustomWin NUIDocumentWindow
				[+] multitag "[NativeHWNDHost]NUIDocumentWindow"
					[ ] "$0"
				[+] CustomWin NetUIHWND1
					[+] multitag "[NetUIHWND]#1"
						[ ] "$0"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "$0[2]"
[+] window DialogBox MicrosoftOfficeOutlook
	[ ] tag "*Microsoft Outlook"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin GoogleDesktopToolbar
			[+] multitag "[MsoCommandBar]Google Desktop Toolbar"
				[ ] "[MsoCommandBar]$0[3]"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] CustomWin WunderBar
			[+] multitag "[WUNDERBAR]WunderBar"
				[ ] "[WUNDERBAR]$1"
			[+] CustomWin NUIDocumentWindow
				[+] multitag "[NativeHWNDHost]NUIDocumentWindow"
					[ ] "[NativeHWNDHost]$0"
				[+] CustomWin NetUIHWND1
					[ ] tag "[NetUIHWND]$0"
	[+] StatusBar StatusBar1
		[ ] tag "$101"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
		[+] CustomWin Rctrl_sbpwnd322
			[ ] tag "[rctrl_sbpwnd32]$0[2]"
		[+] CustomWin Rctrl_sbpwnd321
			[ ] tag "[rctrl_sbpwnd32]$0[1]"
	[+] CustomWin Rctrl_renwnd321
		[ ] tag "[rctrl_renwnd32]$105"
		[+] CustomWin AfxWndW1
			[ ] tag "[AfxWndW]$101"
			[+] CustomWin TableView
				[+] multitag "[SUPERGRID]Table View"
					[ ] "[SUPERGRID]$4704"
				[+] ScrollBar ScrollBar1
					[ ] tag "$4693"
			[+] CustomWin AfxWndW1
				[ ] tag "[AfxWndW]$1"
		[+] CustomWin Rctrl_renwnd321
			[ ] tag "[rctrl_renwnd32]$102"
			[+] CustomWin AfxWndW1
				[ ] tag "[AfxWndW]$0"
				[+] DialogBox DialogBox1
					[ ] tag "$0"
					[+] CustomWin UndeliverableDeliveryStatus
						[+] multitag "[RichEdit20WPT]Undeliverable:Delivery Status Notification (Failure)"
							[ ] "[RichEdit20WPT]$4099"
					[+] CustomWin SystemAdministrator
						[+] multitag "[RichEdit20WPT]System Administrator"
							[ ] "[RichEdit20WPT]$4096"
					[+] StaticText ToText
						[+] multitag "To:"
							[ ] "$4514"
					[+] CustomWin RichEdit20WPT3
						[ ] tag "[RichEdit20WPT]$4098"
					[+] StaticText SentText
						[+] multitag "Sent:"
							[ ] "$4513"
					[+] CustomWin Mon1232007313AM
						[+] multitag "[RichEdit20WPT]Mon 12?3?2007 3:13 AM"
							[ ] "[RichEdit20WPT]$4097"
					[+] CustomWin Sent
						[+] multitag "[EtchedHorz]Sent:"
							[ ] "[EtchedHorz]$4522"
					[+] CustomWin YourMessageDidNotReachSom
						[+] multitag "[RichEdit20WPT]Your message did not reach some or all of the intended recipients.????      Subject:	Test????The following recipient(s) could *"
							[ ] "[RichEdit20WPT]$4100"
					[+] CustomWin NativeFontCtl1
						[ ] tag "[NativeFontCtl]$4534"
					[+] CustomWin Rctrl_renwnd321
						[ ] tag "[rctrl_renwnd32]$4101"
[ ] // Mail Message Declarations
[+] window DialogBox Message 
	[ ] tag "*Message*"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Envelope
			[+] multitag "[MsoCommandBar]Envelope"
				[ ] "[MsoCommandBar]$0[2]"
			[+] DialogBox DialogBox1
				[ ] tag "$0"
				[+] CustomWin MSOGenericControlContainer
					[+] multitag "[MsoCommandBar]MSO Generic Control Container"
						[ ] "[MsoCommandBar]$4524"
				[+] PushButton ToButton
					[+] multitag "To"
						[ ] "$4353"
				[+] PushButton ccButton
					[+] multitag "Cc"
						[ ] "$4354"
				[+] DefinedWin To
					[ ] tag "(73,25-784,48)"
				[+] DefinedWin Cc
					[ ] tag "(73,51-784,73)"
				[+] DefinedWin Subject
					[ ] tag "(74,75-786,99)"
				[ ] 
				[+] // CustomWin To
					[+] // multitag "[RichEdit20WPT];[1]"
						[ ] // "[RichEdit20WPT]$4097"
				[+] // CustomWin cc
					[+] // multitag "[RichEdit20WPT];[2]"
						[ ] // "[RichEdit20WPT]$4098"
				[+] // CustomWin Subject
					[+] // multitag "[RichEdit20WPT]Test1234"
						[ ] // "[RichEdit20WPT]$4102"
				[+] CustomWin EnvInfoBar1
					[ ] tag "[EnvInfoBar]$4525"
		[+] CustomWin EMail
			[+] multitag "[MsoCommandBar]E-mail"
				[ ] "[MsoCommandBar]$0[3]"
			[+] CustomWin RichEdit20W2
				[ ] tag "[RichEdit20W]$28770320"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$18087172"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$18041460"
	[+] CustomWin _WwF1
		[ ] tag "[_WwF]$0"
[+] window DialogBox MessageDock
	[ ] tag "*Message"
	[+] CustomWin _WwF1
		[+] multitag "[_WwF]#1"
			[ ] "$0"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "$0"
		[+] CustomWin Envelope
			[+] multitag "[MsoCommandBar]Envelope"
				[ ] "$0[2]"
			[+] DialogBox DialogBox1
				[+] multitag "#1"
					[ ] "$0"
				[+] CustomWin User4
					[+] multitag "[RichEdit20WPT]user4"
						[ ] "$4102"
	[+] CustomWin Message
		[+] multitag "[_WwG]Message"
			[ ] "$0"
[+] window DialogBox OpenMessage
	[ ] tag "*- Message*"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] DialogBox DialogBox1
			[ ] tag "$0"
			[+] StaticText FromText
				[+] multitag "From:"
					[ ] "$4512"
			[+] CustomWin Usera
				[+] multitag "[RichEdit20WPT]usera"
					[ ] "[RichEdit20WPT]$4097"
			[+] StaticText SentText
				[+] multitag "Sent:"
					[ ] "$4513"
			[+] CustomWin Mon12320071114AM
				[+] multitag "[RichEdit20WPT]Mon 12?3?2007 11:14 AM"
					[ ] "[RichEdit20WPT]$4098"
			[+] StaticText ToText
				[+] multitag "To:"
					[ ] "$4514"
			[+] CustomWin UserbBlk0520NavyUsaCfeCmi
				[+] multitag "[RichEdit20WPT]userb@blk0520.navy.usa.cfe.cmil.mil"
					[ ] "[RichEdit20WPT]$4099"
			[+] StaticText CcText
				[+] multitag "Cc:"
					[ ] "$4515"
			[+] CustomWin Cc
				[+] multitag "[RichEdit20WPT]Cc:"
					[ ] "[RichEdit20WPT]$4100"
			[+] StaticText SubjectText
				[+] multitag "Subject:"
					[ ] "$4516"
			[+] CustomWin TEST
				[+] multitag "[RichEdit20WPT]TEST"
					[ ] "[RichEdit20WPT]$4101"
			[+] StaticText AttachmentsText
				[+] multitag "Attachments"
					[ ] "$4537"
			[+] StaticText MessageText
				[+] multitag "Message"
					[ ] "$4531"
			[+] CustomWin Subject
				[+] multitag "[AfxWndA]Subject:"
					[ ] "[AfxWndA]$4159"
				[+] StaticText MessageText
					[+] multitag "Message"
						[ ] "$0"
				[+] CustomWin AfxWndW1
					[ ] tag "[AfxWndW]$2748"
					[+] CustomWin InternetExplorer_Server1
						[ ] tag "[Internet Explorer_Server]$0"
			[+] CustomWin NativeFontCtl1
				[ ] tag "[NativeFontCtl]$4534"
			[+] CustomWin Rctrl_renwnd321
				[ ] tag "[rctrl_renwnd32]$4096"
[+] window DialogBox ReplyMessage
	[ ] tag "RE*"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Envelope
			[+] multitag "[MsoCommandBar]Envelope"
				[ ] "[MsoCommandBar]$0[2]"
			[+] DialogBox DialogBox1
				[ ] tag "$0"
				[+] CustomWin MSOGenericControlContainer
					[+] multitag "[MsoCommandBar]MSO Generic Control Container"
						[ ] "[MsoCommandBar]$4524"
				[+] PushButton ToButton
					[+] multitag "To"
						[ ] "$4353"
				[+] PushButton CcButton
					[+] multitag "Cc"
						[ ] "$4354"
				[+] CustomWin To
					[ ] tag "[RichEdit20WPT]$4097"
				[+] CustomWin Cc
					[ ] tag "[RichEdit20WPT]$4098"
				[+] CustomWin RETEST
					[+] multitag "[RichEdit20WPT]RE: TEST"
						[ ] "[RichEdit20WPT]$4102"
				[+] CustomWin EnvInfoBar1
					[ ] tag "[EnvInfoBar]$4525"
		[+] CustomWin EMail
			[+] multitag "[MsoCommandBar]E-mail"
				[ ] "[MsoCommandBar]$0[3]"
			[+] CustomWin RichEdit20W2
				[ ] tag "[RichEdit20W]$19518864"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$19518184"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$19483252"
	[+] CustomWin _WwF1
		[ ] tag "[_WwF]$0"
[+] window DialogBox ReceivedMessage
	[+] multitag "*- Message*"
		[ ] "$MessageBox[1]"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] DialogBox DialogBox1
			[ ] tag "$0"
			[+] StaticText FromText
				[+] multitag "From:"
					[ ] "$4512"
			[+] CustomWin Silka
				[+] multitag "[RichEdit20WPT]silka"
					[ ] "[RichEdit20WPT]$4097"
			[+] StaticText SentText
				[+] multitag "Sent:"
					[ ] "$4513"
			[+] CustomWin Thu12132007243PM
				[+] multitag "[RichEdit20WPT]Thu 12?13?2007 2:43 PM"
					[ ] "[RichEdit20WPT]$4098"
			[+] StaticText ToText
				[+] multitag "To:"
					[ ] "$4514"
			[+] CustomWin Silkb
				[+] multitag "[RichEdit20WPT]silkb"
					[ ] "[RichEdit20WPT]$4099"
			[+] StaticText CcText
				[+] multitag "Cc:"
					[ ] "$4515"
			[+] CustomWin Cc
				[+] multitag "[RichEdit20WPT]Cc:"
					[ ] "[RichEdit20WPT]$4100"
			[+] StaticText SubjectText
				[+] multitag "Subject:"
					[ ] "$4516"
			[+] TextField Subject
				[+] multitag "[RichEdit20WPT]Test Case 2.2.9"
					[ ] "[RichEdit20WPT]$4101"
			[+] StaticText AttachmentsText1
				[+] multitag "Attachments:"
					[ ] "$4536"
			[+] StaticText AttachmentsText2
				[+] multitag "Attachments"
					[ ] "$4537"
			[+] TextField RichEdit
				[ ] tag "[RichEdit20WPT]$4104"
			[+] StaticText MessageText
				[+] multitag "Message"
					[ ] "$4531"
			[+] CustomWin Attachments
				[+] multitag "[AfxWndA]Attachments:"
					[ ] "[AfxWndA]$4159"
				[+] StaticText MessageText
					[+] multitag "Message"
						[ ] "$0"
				[+] CustomWin AfxWndW1
					[ ] tag "[AfxWndW]$2748"
					[+] CustomWin InternetExplorer_Server1
						[ ] tag "[Internet Explorer_Server]$0"
			[+] CustomWin NativeFontCtl1
				[ ] tag "[NativeFontCtl]$4534"
			[+] CustomWin Rctrl_renwnd321
				[ ] tag "[rctrl_renwnd32]$4096"
[+] window DialogBox InsertFile
	[ ] tag "Insert File"
	[ ] parent MessageDock
	[+] CustomWin Filename
		[+] multitag "[RichEdit20W]#1"
			[ ] "$54"
[+] window DialogBox ConfirmSave
	[ ] tag "Microsoft Office*"
	[ ] parent MessageDock
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin DoYouWantToSaveTheChange
		[+] multitag "[MSOUNISTAT]Do you want to save the changes to User10?"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox SaveAllAttachments1
	[ ] tag "Save All Attachments"
	[ ] parent MicrosoftOutlook
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$1"
	[+] PushButton Close
		[+] multitag "Close"
			[ ] "$2"
	[+] StaticText AttachmentsText
		[+] multitag "Attachments:"
			[ ] "$4512"
	[+] CustomWin Attachments
		[+] multitag "[REListBox20W]Attachments:"
			[ ] "[REListBox20W]$4624"
	[+] CustomWin NativeFontCtl1
		[ ] tag "[NativeFontCtl]$4873"
[+] window DialogBox SaveAllAttachments2
	[ ] tag "Save All Attachments"
	[ ] parent SaveAllAttachments1
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin SaveAllAttachments
		[+] multitag "[Snake List]Save All Attachments"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin RichEdit20W1
		[ ] tag "[RichEdit20W]$54"
	[+] DefinedWin OK
		[ ] tag "(496,309-585,331)"
		[ ] 
	[+] DefinedWin FolderName
		[ ] tag "(170,308-473,328)"
[+] window DialogBox SaveAttachment
	[ ] tag "Save Attachment"
	[ ] parent Message
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin SaveAttachment
		[+] multitag "[Snake List]Save Attachment"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] TextField Filename
		[ ] tag "[RichEdit20W]$54"
	[+] DefinedWin Save
		[ ] tag "(499,308-585,330)"
[+] window DialogBox replaceExistingFile
	[ ] tag "Microsoft Office Outlook"
	[ ] parent SaveAttachment
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin TextMessage
		[+] multitag "[MSOUNISTAT]The file attach1.txt already exists. Do you want to replace the existing file?"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox OutlookSaveChanges
	[ ] tag "Microsoft Office Word"
	[ ] parent Message
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin DoYouWantToSaveTheChange
		[+] multitag "[MSOUNISTAT]Do you want to save the changes to ?Test Email 8.2.2.1.2.1 Local Shipboard Test??"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox UnsentMessages
	[ ] tag "Microsoft Office Outlook"
	[ ] parent MicrosoftOutlook
	[+] CustomWin ThereAreUnsentMessagesInY
		[+] multitag "[RichEdit20WPT]There are unsent messages in your Outbox. To send messages, Outlook must remain running and connected to your e-mail server. W*"
			[ ] "[RichEdit20WPT]$20"
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
[+] window DialogBox SelectNames
	[ ] tag "Select Names"
	[ ] parent Message
	[+] StaticText ShowNamesFromThe1Text
		[+] multitag "Show Names from the:[1]"
			[ ] "$106"
	[+] PopupList ShowNamesFromThe
		[+] multitag "Show Names from the:"
			[ ] "$116"
	[+] StaticText TypeNameOrSelectFromList1
		[+] multitag "Type Name or Select from List:"
			[ ] "$131"
	[+] CustomWin TypeNameOrSelectFromList2
		[+] multitag "[RichEdit20W]Type Name or Select from List:"
			[ ] "[RichEdit20W]$101"
	[+] CustomWin OUTEXVLB1
		[ ] tag "[OUTEXVLB]$105"
	[+] StaticText ShowNamesFromThe2Text
		[+] multitag "Show Names from the:[2]"
			[ ] "$139"
	[+] PushButton ToButton
		[+] multitag "To ->"
			[ ] "$109"
	[+] PushButton CcButton
		[+] multitag "Cc ->"
			[ ] "$110"
	[+] PushButton BccButton
		[+] multitag "Bcc ->"
			[ ] "$111"
	[+] StaticText MessageRecipientsText
		[+] multitag "Message Recipients"
			[ ] "$132"
	[+] StaticText ToText
		[+] multitag "To"
			[ ] "$206"
	[+] CustomWin To
		[ ] tag "[RichEdit20W]$114"
		[+] // multitag "[RichEdit20W]Message Recipients[1]"
			[ ] // "[RichEdit20W]$113"
	[+] StaticText CcText
		[+] multitag "Cc"
			[ ] "$207"
	[+] CustomWin Cc
		[ ] tag "[RichEdit20W]$114"
		[+] // multitag "[RichEdit20W]Message Recipients[2]"
			[ ] // "[RichEdit20W]$114"
	[+] StaticText BccText
		[+] multitag "Bcc"
			[ ] "$208"
	[+] CustomWin Bcc
		[+] multitag "[RichEdit20W]Message Recipients[3]"
			[ ] "[RichEdit20W]$115"
	[+] PushButton Advanced
		[+] multitag "Advanced"
			[ ] "$209"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$100"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] CustomWin MessageRecipients12
		[+] multitag "[EtchedHorz]Message Recipients[1]"
			[ ] "[EtchedHorz]$134"
	[+] CustomWin MessageRecipients22
		[+] multitag "[EtchedHorz]Message Recipients[2]"
			[ ] "[EtchedHorz]$135"
	[+] CustomWin NativeFontCtl1
		[ ] tag "[NativeFontCtl]$777"
[+] window DialogBox Reminder
	[ ] tag "1 Reminder"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$8344"
	[+] ListView ListView1
		[ ] tag "$8342"
		[+] Header Header1
			[ ] tag "$0"
			[+] StaticText SubjectText
				[ ] tag "Subject"
			[+] StaticText DueInText
				[ ] tag "Due in"
	[+] PushButton DismissAll
		[+] multitag "Dismiss All"
			[ ] "$8345"
	[+] PushButton OpenItem
		[+] multitag "Open Item"
			[ ] "$8332"
	[+] PushButton Dismiss
		[+] multitag "Dismiss"
			[ ] "$8333"
	[+] StaticText ClickSnoozeToBeRemindedAg
		[+] multitag "Click Snooze to be reminded again in:"
			[ ] "$4514"
	[+] CustomWin N5Minutes
		[+] multitag "[REComboBox20W]5 minutes"
			[ ] "[REComboBox20W]$8339"
	[+] PushButton Snooze
		[+] multitag "Snooze"
			[ ] "$8334"
	[+] CustomWin NativeFontCtl1
		[ ] tag "[NativeFontCtl]$4873"
[+] window DialogBox OutlookSendReceiveProgress
	[ ] tag "Outlook Send?Receive Progress"
	[+] CustomWin N1Of2TasksHaveCompletedSu
		[+] multitag "[RichEdit20WPT]1 of 2 Tasks have completed successfully"
			[ ] "[RichEdit20WPT]$4101"
	[+] CustomWin Msctls_progress321
		[ ] tag "[msctls_progress32]$11"
	[+] PushButton CancelAll
		[+] multitag "Cancel All"
			[ ] "$4355"
	[+] PushButton Details
		[+] multitag "<< Details"
			[ ] "$4353"
	[+] CheckBox DonTShowThisDialogBoxDur
		[+] multitag "Don't show this dialog box during Send?Receive"
			[ ] "$4224"
	[+] CustomWin EtchedHorz1
		[ ] tag "[EtchedHorz]$4512"
	[+] ListView ListView1
		[ ] tag "$4640"
		[+] Header Header1
			[ ] tag "$0"
			[+] PushButton Name
				[ ] tag "Name"
			[+] PushButton Progress
				[ ] tag "Progress"
			[+] PushButton Remaining
				[ ] tag "Remaining"
	[+] PushButton CancelTask
		[+] multitag "Cancel Task"
			[ ] "$4354"
	[+] CustomWin MicrosoftExchangeServerSe
		[+] multitag "[RichEdit20WPT]Microsoft Exchange Server - Sending"
			[ ] "[RichEdit20WPT]$4103"
	[+] CustomWin RichEdit20WPT3
		[ ] tag "[RichEdit20WPT]$4102"
	[+] PageList PageList1
		[ ] tag "$4648"
	[+] ToolBar ToolBar1
		[ ] tag "$4352"
		[+] CheckBox CheckBox1
			[ ] tag "$4352"
[ ] 
[ ] // Appointment Declarations
[+] window DialogBox CalendarMicrosoftOutlook
	[ ] tag "Calendar - Microsoft Outlook"
	[+] CustomWin AfxWndW1
		[+] multitag "[AfxWndW]#1"
			[ ] "$103"
		[+] CustomWin WunderBar
			[+] multitag "[WUNDERBAR]WunderBar"
				[ ] "$1"
			[+] CustomWin NUIDocumentWindow
				[+] multitag "[NativeHWNDHost]NUIDocumentWindow"
					[ ] "$0"
				[+] CustomWin NetUIHWND1
					[+] multitag "[NetUIHWND]#1"
						[ ] "$0"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "$0[2]"
[+] window DialogBox UntitledAppointment
	[ ] tag "Untitled - Appointment"
	[+] CustomWin AfxWndW1
		[+] multitag "[AfxWndW]#1"
			[ ] "$103"
		[+] DialogBox DialogBox1
			[+] multitag "#1"
				[ ] "$0"
[+] // window DialogBox Appointment
	[ ] // tag "* - Appointment"
	[+] // CustomWin AfxWndW1
		[+] // multitag "[AfxWndW]#1"
			[ ] // "$103"
		[+] // DialogBox DialogBox1
			[+] // multitag "#1"
				[ ] // "$0"
			[+] // CustomWin Te
				[+] // multitag "[RichEdit20WPT]Te"
					[ ] // "$4100"
			[+] // CustomWin Bu
				[+] // multitag "[REComboBox20W]Bu"
					[ ] // "$4102"
			[+] // CustomWin StartTimeLabel
				[+] // multitag "[AfxWndW]Start time:[1]"
					[ ] // "$4352"
			[+] // CustomWin StartTimeField
				[+] // multitag "[RichEdit20WPT]*[1]"
					[ ] // "$4098"
			[+] // CustomWin Message
				[+] // multitag "[AfxWndA]Show time as:"
					[ ] // "$4159"
				[+] // CustomWin Body
					[+] // multitag "[RichEdit20W]#1"
						[ ] // "$2748"
[+] window DialogBox Appointment
	[ ] tag "* - Appointment"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] DialogBox DialogBox1
			[ ] tag "$0"
			[+] StaticText ToText
				[+] multitag "To"
					[ ] "$4531"
			[+] StaticText SubjectText
				[+] multitag "Subject:"
					[ ] "$4534"
			[+] CustomWin Subject
				[+] multitag "[RichEdit20WPT]Test Case 2.2.7"
					[ ] "[RichEdit20WPT]$4100"
			[+] StaticText LocationText
				[+] multitag "Location:"
					[ ] "$4530"
			[+] CustomWin Location
				[+] multitag "[REComboBox20W]Building 40"
					[ ] "[REComboBox20W]$4102"
			[+] StaticText LabelText
				[+] multitag "Label:"
					[ ] "$4539"
			[+] PopupList Label1
				[+] multitag "Label:"
					[ ] "$4485"
			[+] CustomWin LocationLabel
				[+] multitag "[EtchedHorz]Location:"
					[ ] "[EtchedHorz]$4515"
			[+] StaticText StartTime2Text
				[+] multitag "Start time:[2]"
					[ ] "$4513"
			[+] StaticText StartDateText
				[+] multitag "Start date:"
					[ ] "$4554"
			[+] CustomWin StartDate
				[+] multitag "[RichEdit20WPT]Wed 12?12?2007[1]"
					[ ] "[RichEdit20WPT]$4098"
			[+] CustomWin StartTime1
				[+] multitag "[AfxWndW]Start time:[1]"
					[ ] "[AfxWndW]$4352"
			[+] StaticText StartTime1Text
				[+] multitag "Start time:[1]"
					[ ] "$4555"
			[+] CustomWin N800AM
				[+] multitag "[RichEdit20WPT]8:00 AM"
					[ ] "[RichEdit20WPT]$4096"
			[+] CustomWin StartTime2
				[+] multitag "[AfxWndW]Start time:[2]"
					[ ] "[AfxWndW]$4354"
			[+] StaticText EndTime2Text
				[+] multitag "End time:[2]"
					[ ] "$4512"
			[+] StaticText EndDateText
				[+] multitag "End date:"
					[ ] "$4556"
			[+] CustomWin Wed121220072
				[+] multitag "[RichEdit20WPT]Wed 12?12?2007[2]"
					[ ] "[RichEdit20WPT]$4099"
			[+] CustomWin EndTime1
				[+] multitag "[AfxWndW]End time:[1]"
					[ ] "[AfxWndW]$4353"
			[+] StaticText EndTime1Text
				[+] multitag "End time:[1]"
					[ ] "$4557"
			[+] CustomWin N830AM
				[+] multitag "[RichEdit20WPT]8:30 AM"
					[ ] "[RichEdit20WPT]$4097"
			[+] CustomWin EndTime2
				[+] multitag "[AfxWndW]End time:[2]"
					[ ] "[AfxWndW]$4355"
			[+] CheckBox AllDayEvent
				[+] multitag "All day event"
					[ ] "$4226"
			[+] CustomWin EndTime
				[+] multitag "[EtchedHorz]End time:"
					[ ] "[EtchedHorz]$4517"
			[+] CheckBox Reminder
				[+] multitag "Reminder:"
					[ ] "$4224"
			[+] StaticText ReminderTimeText
				[+] multitag "Reminder time"
					[ ] "$4559"
			[+] CustomWin N15Minutes
				[+] multitag "[RichEdit20WPT]15 minutes"
					[ ] "[RichEdit20WPT]$4480"
			[+] CustomWin AfxWndW5
				[ ] tag "[AfxWndW]$4481"
			[+] PushButton ReminderSound
				[+] multitag "Reminder Sound"
					[ ] "$4358"
			[+] StaticText ShowTimeAsText
				[+] multitag "Show time as:"
					[ ] "$4525"
			[+] PopupList ShowTimeAs1
				[+] multitag "Show time as:"
					[ ] "$4482"
			[+] StaticText ThisIsAnOnlineMeetingUsin
				[+] multitag "This is an online meeting using:"
					[ ] "$4558"
			[+] StaticText NotesText
				[+] multitag "Notes"
					[ ] "$4535"
			[+] CustomWin ShowTimeAs2
				[+] multitag "[AfxWndA]Show time as:"
					[ ] "[AfxWndA]$4159"
				[+] StaticText NotesText
					[+] multitag "Notes"
						[ ] "$0"
				[+] CustomWin Body
					[ ] tag "[RichEdit20W]$2748"
			[+] PushButton Contacts
				[+] multitag "Contacts"
					[ ] "$4361"
			[+] StaticText ContactsText
				[+] multitag "Contacts"
					[ ] "$4537"
			[+] CustomWin ShowTimeAs3
				[+] multitag "[RichEdit20WPT]Show time as:"
					[ ] "[RichEdit20WPT]$4112"
			[+] PushButton Categories
				[+] multitag "Categories"
					[ ] "$4356"
			[+] StaticText CategoriesText
				[+] multitag "Categories"
					[ ] "$4536"
			[+] CustomWin Label2
				[+] multitag "[RichEdit20WPT]Label:"
					[ ] "[RichEdit20WPT]$4107"
			[+] CheckBox Private
				[+] multitag "Private"
					[ ] "$4225"
			[+] CustomWin NativeFontCtl1
				[ ] tag "[NativeFontCtl]$4873"
			[+] CustomWin Rctrl_renwnd321
				[ ] tag "[rctrl_renwnd32]$4105"
		[+] CustomWin Renctl_tabbar1
			[ ] tag "[renctl_tabbar]$32767"
[+] window DialogBox Scheduling
	[ ] tag "* - Appointment"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] DialogBox DialogBox1
			[ ] tag "$0"
			[+] CustomWin Attendees
				[+] multitag "[SUPERGRID]Attendees"
					[ ] "[SUPERGRID]$4542"
			[+] PushButton AddOthers
				[+] multitag "Add Others"
					[ ] "$4356"
			[+] PushButton Options
				[+] multitag "Options"
					[ ] "$4362"
			[+] PushButton PushButton3
				[+] multitag "<<"
					[ ] "$4358"
			[+] PushButton AutoPickNext
				[+] multitag "AutoPick Next >>"
					[ ] "$4357"
			[+] CustomWin Zoom1
				[+] multitag "[BitMap]Zoom:"
					[ ] "[BitMap]$4544"
			[+] StaticText BusyText
				[+] multitag "Busy"
					[ ] "$4545"
			[+] StaticText TentativeText
				[+] multitag "Tentative"
					[ ] "$4547"
			[+] CustomWin Busy
				[+] multitag "[BitMap]Busy"
					[ ] "[BitMap]$4546"
			[+] StaticText OutOfOfficeText
				[+] multitag "Out of Office"
					[ ] "$4549"
			[+] CustomWin Tentative
				[+] multitag "[BitMap]Tentative"
					[ ] "[BitMap]$4548"
			[+] CustomWin MeetingEndTime11
				[+] multitag "[EtchedHorz]Meeting end time:[1]"
					[ ] "[EtchedHorz]$4515"
			[+] StaticText MeetingStartTime2Text
				[+] multitag "Meeting start time:[2]"
					[ ] "$4513"
			[+] StaticText MeetingStartDateText
				[+] multitag "Meeting start date:"
					[ ] "$4866"
			[+] CustomWin Wed121220071
				[+] multitag "[RichEdit20WPT]Wed 12?12?2007[1]"
					[ ] "[RichEdit20WPT]$4098"
			[+] CustomWin MeetingStartTime1
				[+] multitag "[AfxWndW]Meeting start time:[1]"
					[ ] "[AfxWndW]$4352"
			[+] StaticText MeetingStartTime1Text
				[+] multitag "Meeting start time:[1]"
					[ ] "$4865"
			[+] CustomWin N800AM
				[+] multitag "[RichEdit20WPT]8:00 AM"
					[ ] "[RichEdit20WPT]$4096"
			[+] CustomWin MeetingStartTime2
				[+] multitag "[AfxWndW]Meeting start time:[2]"
					[ ] "[AfxWndW]$4354"
			[+] StaticText MeetingEndTime2Text
				[+] multitag "Meeting end time:[2]"
					[ ] "$4512"
			[+] StaticText MeetingEndDateText
				[+] multitag "Meeting end date:"
					[ ] "$4868"
			[+] CustomWin Wed121220072
				[+] multitag "[RichEdit20WPT]Wed 12?12?2007[2]"
					[ ] "[RichEdit20WPT]$4099"
			[+] CustomWin MeetingEndTime12
				[+] multitag "[AfxWndW]Meeting end time:[1]"
					[ ] "[AfxWndW]$4353"
			[+] StaticText MeetingEndTime1Text
				[+] multitag "Meeting end time:[1]"
					[ ] "$4867"
			[+] CustomWin N830AM
				[+] multitag "[RichEdit20WPT]8:30 AM"
					[ ] "[RichEdit20WPT]$4097"
			[+] CustomWin MeetingEndTime21
				[+] multitag "[AfxWndW]Meeting end time:[2]"
					[ ] "[AfxWndW]$4355"
			[+] CustomWin Zoom2
				[+] multitag "[AfxWndW]Zoom:"
					[ ] "[AfxWndW]$4720"
			[+] StaticText NoInformationText
				[+] multitag "No Information"
					[ ] "$4551"
			[+] CustomWin OutOfOffice
				[+] multitag "[BitMap]Out of Office"
					[ ] "[BitMap]$4550"
			[+] CustomWin NativeFontCtl1
				[ ] tag "[NativeFontCtl]$4873"
			[+] StaticText ZoomText
				[+] multitag "Zoom:"
					[ ] "$4871"
			[+] CustomWin N100DayView
				[+] multitag "[REComboBox20W]100% (Day View)"
					[ ] "[REComboBox20W]$4480"
			[+] StaticText CurrentMeetingTimeText
				[+] multitag "Current Meeting Time"
					[ ] "$4553"
			[+] CustomWin BitMap1
				[ ] tag "[BitMap]$4552"
			[+] CustomWin MeetingEndTime22
				[+] multitag "[EtchedHorz]Meeting end time:[2]"
					[ ] "[EtchedHorz]$4516"
			[+] ScrollBar CurrentMeetingTime
				[+] multitag "Current Meeting Time"
					[ ] "$1235"
		[+] CustomWin Renctl_tabbar1
			[ ] tag "[renctl_tabbar]$32767"
[+] window DialogBox Meeting
	[ ] tag "* - Meeting"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] DialogBox Appointment
			[ ] tag "$0"
			[+] PushButton ToButton
				[+] multitag "To"
					[ ] "$4357"
			[+] StaticText ToText
				[+] multitag "To"
					[ ] "$4531"
			[+] CustomWin To
				[+] multitag "[RichEdit20WPT];"
					[ ] "[RichEdit20WPT]$4106"
			[+] StaticText SubjectText
				[+] multitag "Subject:"
					[ ] "$4534"
			[+] CustomWin Subject
				[+] multitag "[RichEdit20WPT]Subject"
					[ ] "[RichEdit20WPT]$4100"
			[+] StaticText LocationText
				[+] multitag "Location:"
					[ ] "$4530"
			[+] CustomWin Location
				[+] multitag "[REComboBox20W]Location"
					[ ] "[REComboBox20W]$4102"
			[+] StaticText LabelText
				[+] multitag "Label:"
					[ ] "$4539"
			[+] PopupList Label1
				[+] multitag "Label:"
					[ ] "$4485"
			[+] CustomWin Location2
				[+] multitag "[EtchedHorz]Location:"
					[ ] "[EtchedHorz]$4515"
			[+] StaticText StartTime2Text
				[+] multitag "Start time:[2]"
					[ ] "$4513"
			[+] StaticText StartDateText
				[+] multitag "Start date:"
					[ ] "$4554"
			[+] CustomWin StartDate
				[+] multitag "[RichEdit20WPT]Wed 12?12?2007[1]"
					[ ] "[RichEdit20WPT]$4098"
			[+] CustomWin StartTime1
				[+] multitag "[AfxWndW]Start time:[1]"
					[ ] "[AfxWndW]$4352"
			[+] StaticText StartTime1Text
				[+] multitag "Start time:[1]"
					[ ] "$4555"
			[+] CustomWin N1130AM
				[+] multitag "[RichEdit20WPT]11:30 AM"
					[ ] "[RichEdit20WPT]$4096"
			[+] CustomWin StartTime2
				[+] multitag "[AfxWndW]Start time:[2]"
					[ ] "[AfxWndW]$4354"
			[+] StaticText EndTime2Text
				[+] multitag "End time:[2]"
					[ ] "$4512"
			[+] StaticText EndDateText
				[+] multitag "End date:"
					[ ] "$4556"
			[+] CustomWin Wed121220072
				[+] multitag "[RichEdit20WPT]Wed 12?12?2007[2]"
					[ ] "[RichEdit20WPT]$4099"
			[+] CustomWin EndTime1
				[+] multitag "[AfxWndW]End time:[1]"
					[ ] "[AfxWndW]$4353"
			[+] StaticText EndTime1Text
				[+] multitag "End time:[1]"
					[ ] "$4557"
			[+] CustomWin N1200PM
				[+] multitag "[RichEdit20WPT]12:00 PM"
					[ ] "[RichEdit20WPT]$4097"
			[+] CustomWin EndTime2
				[+] multitag "[AfxWndW]End time:[2]"
					[ ] "[AfxWndW]$4355"
			[+] CheckBox AllDayEvent
				[+] multitag "All day event"
					[ ] "$4226"
			[+] CustomWin EndTime
				[+] multitag "[EtchedHorz]End time:"
					[ ] "[EtchedHorz]$4517"
			[+] CheckBox Reminder
				[+] multitag "Reminder:"
					[ ] "$4224"
			[+] StaticText ReminderTimeText
				[+] multitag "Reminder time"
					[ ] "$4559"
			[+] CustomWin N15Minutes
				[+] multitag "[RichEdit20WPT]15 minutes"
					[ ] "[RichEdit20WPT]$4480"
			[+] CustomWin AfxWndW5
				[ ] tag "[AfxWndW]$4481"
			[+] PushButton ReminderSound
				[+] multitag "Reminder Sound"
					[ ] "$4358"
			[+] StaticText ShowTimeAsText
				[+] multitag "Show time as:"
					[ ] "$4525"
			[+] PopupList ShowTimeAs1
				[+] multitag "Show time as:"
					[ ] "$4482"
			[+] CustomWin ShowTimeAs2
				[+] multitag "[EtchedHorz]Show time as:"
					[ ] "[EtchedHorz]$4533"
			[+] PushButton MeetingWorkspace
				[+] multitag "Meeting Workspace"
					[ ] "$4362"
			[+] CheckBox ThisIsAnOnlineMeetingUsin1
				[+] multitag "This is an online meeting using:"
					[ ] "$4245"
			[+] StaticText ThisIsAnOnlineMeetingUsin2
				[+] multitag "This is an online meeting using:"
					[ ] "$4558"
			[+] CustomWin MicrosoftNetMeeting
				[+] multitag "[REComboBox20W]Microsoft NetMeeting"
					[ ] "[REComboBox20W]$4483"
			[+] StaticText NotesText
				[+] multitag "Notes"
					[ ] "$4535"
			[+] CustomWin ShowTimeAs3
				[+] multitag "[AfxWndA]Show time as:"
					[ ] "[AfxWndA]$4159"
				[+] StaticText NotesText
					[+] multitag "Notes"
						[ ] "$0"
				[+] CustomWin Body
					[ ] tag "[RichEdit20W]$2748"
			[+] PushButton Contacts
				[+] multitag "Contacts"
					[ ] "$4361"
			[+] StaticText ContactsText
				[+] multitag "Contacts"
					[ ] "$4537"
			[+] CustomWin ShowTimeAs4
				[+] multitag "[RichEdit20WPT]Show time as:"
					[ ] "[RichEdit20WPT]$4112"
			[+] PushButton Categories
				[+] multitag "Categories"
					[ ] "$4356"
			[+] StaticText CategoriesText
				[+] multitag "Categories"
					[ ] "$4536"
			[+] CustomWin Label2
				[+] multitag "[RichEdit20WPT]Label:"
					[ ] "[RichEdit20WPT]$4107"
			[+] CheckBox Private
				[+] multitag "Private"
					[ ] "$4225"
			[+] CustomWin NativeFontCtl1
				[ ] tag "[NativeFontCtl]$4873"
			[+] CustomWin Rctrl_renwnd321
				[ ] tag "[rctrl_renwnd32]$4105"
				[+] CustomWin Rctrl_renwnd321
					[ ] tag "[rctrl_renwnd32]$1001"
		[+] CustomWin Renctl_tabbar1
			[ ] tag "[renctl_tabbar]$32767"
[+] window DialogBox ReceivedMeeting
	[+] multitag "Test Case 2.2.8 - Meeting"
		[ ] "$MessageBox[1]"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin AfxWndW1
		[ ] tag "[AfxWndW]$103"
		[+] DialogBox DialogBox1
			[ ] tag "$0"
			[+] StaticText FromText
				[+] multitag "From:"
					[ ] "$4512"
			[+] CustomWin Silka
				[+] multitag "[RichEdit20WPT]silka"
					[ ] "[RichEdit20WPT]$4096"
			[+] StaticText SentText
				[+] multitag "Sent:"
					[ ] "$4513"
			[+] CustomWin Thu12132007243PM
				[+] multitag "[RichEdit20WPT]Thu 12?13?2007 2:43 PM"
					[ ] "[RichEdit20WPT]$4097"
			[+] StaticText RequiredText
				[+] multitag "Required:"
					[ ] "$4514"
			[+] CustomWin Silkb
				[+] multitag "[RichEdit20WPT]silkb"
					[ ] "[RichEdit20WPT]$4098"
			[+] StaticText OptionalText
				[+] multitag "Optional:"
					[ ] "$4515"
			[+] CustomWin Optional
				[+] multitag "[RichEdit20WPT]Optional:"
					[ ] "[RichEdit20WPT]$4099"
			[+] StaticText SubjectText
				[+] multitag "Subject:"
					[ ] "$4516"
			[+] TextField Subject
				[+] multitag "[RichEdit20WPT]Test Case 2.2.8"
					[ ] "[RichEdit20WPT]$4100"
			[+] CustomWin Subject1
				[+] multitag "[EtchedHorz]Subject:"
					[ ] "[EtchedHorz]$4519"
			[+] StaticText LocationText
				[+] multitag "Location:"
					[ ] "$4517"
			[+] CustomWin Building40
				[+] multitag "[RichEdit20WPT]Building 40"
					[ ] "[RichEdit20WPT]$4101"
			[+] StaticText WhenText
				[+] multitag "When:"
					[ ] "$4518"
			[+] CustomWin FridayDecember142007300
				[+] multitag "[RichEdit20WPT]Friday, December 14, 2007 3:00 PM-3:30 PM."
					[ ] "[RichEdit20WPT]$4102"
			[+] CustomWin When
				[+] multitag "[AfxWndA]When:"
					[ ] "[AfxWndA]$4159"
				[+] CustomWin RichEdit20W1
					[ ] tag "[RichEdit20W]$2748"
			[+] CustomWin Rctrl_renwnd321
				[ ] tag "[rctrl_renwnd32]$4105"
				[+] CustomWin Rctrl_renwnd321
					[ ] tag "[rctrl_renwnd32]$1001"
[+] window DialogBox InsertFileMeeting
	[ ] tag "Insert File"
	[ ] parent Meeting
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin InsertFile
		[+] multitag "[Snake List]Insert File"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox OutlookPrint
	[ ] tag "Print"
	[ ] parent Message
	[+] CustomWin RichEdit20W2
		[ ] tag "[RichEdit20W]$27"
	[+] CustomWin RichEdit20W1
		[ ] tag "[RichEdit20W]$36"
[ ] 
[ ] // Windows Desktop Declarations
[+] window DialogBox DialogBox3
	[ ] tag "#3"
	[+] CustomWin MenuSite1
		[+] multitag "[MenuSite]#1"
			[ ] "$0"
		[+] ToolBar ToolBar1
			[+] multitag "#1"
				[ ] "$0"
			[+] PushButton Run
				[+] multitag "Run"
					[ ] "$401"
[+] window DialogBox Run
	[ ] tag "Run"
	[+] ComboBox Open
		[+] multitag "Open:"
			[ ] "$12298"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$1"
[ ] 
[+] // window DialogBox SentItemsMicrosoftOutlook
	[ ] // tag "Sent Items - Microsoft Outlook"
	[+] // CustomWin MsoDockTop
		[+] // multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] // "[MsoCommandBarDock]$0"
		[+] // CustomWin Standard
			[+] // multitag "[MsoCommandBar]Standard"
				[ ] // "[MsoCommandBar]$0[2]"
			[+] // CustomWin RichEdit20W1
				[ ] // tag "[RichEdit20W]$23339064"
		[+] // CustomWin MenuBar
			[+] // multitag "[MsoCommandBar]Menu Bar"
				[ ] // "[MsoCommandBar]$0[1]"
			[+] // CustomWin RichEdit20W1
				[ ] // tag "[RichEdit20W]$23336532"
		[+] // CustomWin GoogleDesktopToolbar
			[+] // multitag "[MsoCommandBar]Google Desktop Toolbar"
				[ ] // "[MsoCommandBar]$0[3]"
			[+] // CustomWin RichEdit20W1
				[ ] // tag "[RichEdit20W]$23348712"
	[+] // CustomWin AfxWndW1
		[ ] // tag "[AfxWndW]$103"
		[+] // CustomWin WunderBar
			[+] // multitag "[WUNDERBAR]WunderBar"
				[ ] // "[WUNDERBAR]$1"
			[+] // CustomWin NUIDocumentWindow
				[+] // multitag "[NativeHWNDHost]NUIDocumentWindow"
					[ ] // "[NativeHWNDHost]$0"
				[+] // CustomWin NetUIHWND1
					[ ] // tag "[NetUIHWND]$0"
	[+] // StatusBar StatusBar1
		[ ] // tag "$101"
		[+] // DynamicText DynamicText1
			[ ] // tag "#1"
		[+] // DynamicText DynamicText2
			[ ] // tag "#2"
		[+] // CustomWin Rctrl_sbpwnd321
			[ ] // tag "[rctrl_sbpwnd32]$0"
	[+] // CustomWin Rctrl_renwnd321
		[ ] // tag "[rctrl_renwnd32]$105"
		[+] // CustomWin AfxWndW1
			[ ] // tag "[AfxWndW]$101"
			[+] // CustomWin AfxWndW1
				[ ] // tag "[AfxWndW]$1"
			[+] // CustomWin TableView
				[+] // multitag "[SUPERGRID]Table View"
					[ ] // "[SUPERGRID]$4704"
				[+] // ScrollBar ScrollBar1
					[ ] // tag "$4693"
		[+] // CustomWin Rctrl_renwnd321
			[ ] // tag "[rctrl_renwnd32]$102"
			[+] // CustomWin AfxWndW1
				[ ] // tag "[AfxWndW]$0"
				[+] // DialogBox DialogBox1
					[ ] // tag "$0"
					[+] // StaticText StaticText2
						[ ] // tag "$4521"
					[+] // StaticText MessageText
						[+] // multitag "Message"
							[ ] // "$4531"
[+] // window DialogBox Outlook
	[ ] // tag "Microsoft Office Outlook"
	[ ] // parent SentItemsMicrosoftOutlook
	[+] // PushButton No
		[+] // multitag "No"
			[ ] // "$4775"
	[+] // PushButton Help
		[+] // multitag "Help"
			[ ] // "$4776"
	[+] // CustomWin Icon1
		[ ] // tag "[Icon]$4512"
	[+] // StaticText AProgramIsTryingToAutomat1
		[+] // multitag "A program is trying to automatically send e-mail on your behalf.??Do you want to allow this?????If this is unexpected, it may *"
			[ ] // "$4772"
	[+] // CustomWin AProgramIsTryingToAutomat2
		[+] // multitag "[msctls_progress32]A program is trying to automatically send e-mail on your behalf.??Do you want to allow this?????If this is unexpected, it may *"
			[ ] // "[msctls_progress32]$4769"
	[+] // PushButton Yes
		[+] // multitag "Yes"
			[ ] // "$4774"
[ ] 
[+] window DialogBox MicrosoftWord
	[ ] // tag "*Microsoft Word"
	[ ] STRING sTag
	[ ] tag sTag
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$28309164"
		[+] CustomWin Formatting
			[+] multitag "[MsoCommandBar]Formatting"
				[ ] "[MsoCommandBar]$0[2]"
			[+] CustomWin RichEdit20W3
				[ ] tag "[RichEdit20W]$28289976"
			[+] CustomWin RichEdit20W2
				[ ] tag "[RichEdit20W]$28289296"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$28284236"
	[+] CustomWin _WwC1
		[ ] tag "[_WwC]$0"
	[+] CustomWin _WwF1
		[ ] tag "[_WwF]$0"
	[ ] 
[ ] 
[+] window MainWin Word
	[+] multitag "Microsoft Word"
		[ ] "*Microsoft Word"
		[ ] "$C:\Program Files\Microsoft Office\Office11\WINWORD.exe"
[+] window DialogBox WordOpen
	[ ] tag "Open"
	[ ] parent MicrosoftWord
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin Open
		[+] multitag "[Snake List]Open"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox WordSaveAs
	[ ] tag "Save As"
	[ ] parent MicrosoftWord
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin SaveAs
		[+] multitag "[Snake List]Save As"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox WordReplace
	[ ] tag "Microsoft Office Word"
	[ ] parent WordSaveAs
[+] window DialogBox WordPrint
	[ ] tag "Print"
	[ ] parent MicrosoftWord
	[+] CustomWin RichEdit20W2
		[ ] tag "[RichEdit20W]$27"
	[+] CustomWin RichEdit20W1
		[ ] tag "[RichEdit20W]$36"
[+] window DialogBox WordSaveChanges
	[ ] tag "Microsoft Office Word"
	[ ] parent MicrosoftWord
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin DoYouWantToSaveTheChange
		[+] multitag "[MSOUNISTAT]Do you want to save the changes to Doc1.doc?"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox Spelling
	[ ] tag "Spelling and Grammar: English (U.S.)"
	[+] CustomWin _WwO1
		[ ] tag "[_WwO]$18"
	[+] ScrollBar ScrollBar1
		[ ] tag "$19"
	[+] DefinedWin Resume
		[ ] tag "(344,25-441,44)"
	[+] DefinedWin AddToDictionary
		[ ] tag "(345,83-442,104)"
	[+] DefinedWin IgnoreAll
		[ ] tag "(344,54-438,74)"
[+] window DialogBox SpellingConfirmProofingTools
	[ ] tag "Microsoft Office Word"
	[ ] parent Spelling
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin CannotFindProofingToolsFor
		[+] multitag "[MSOUNISTAT]Cannot find proofing tools for Italian (Italy). If you have the proofing tools, try installing or re-installing them.?Contact *"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox SpellingCheckComplete
	[ ] tag "Microsoft Office Word"
	[ ] parent MicrosoftWord
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin TheSpellingAndGrammarCheck
		[+] multitag "[MSOUNISTAT]The spelling and grammar check is complete.?Text marked with ?Do not check spelling or grammar? was skipped."
			[ ] "[MSOUNISTAT]$4001"
[ ] 
[+] window DialogBox MicrosoftExcel
	[ ] STRING sTag
	[ ] tag sTag
	[ ] // tag "Microsoft Excel*"
	[+] CustomWin EXCEL41
		[ ] tag "[EXCEL4]$0"
	[+] CustomWin EXCEL11
		[ ] tag "[EXCEL;]$0"
		[+] ComboBox ComboBox1
			[ ] tag "$13"
	[+] CustomWin EXCEL12
		[ ] tag "[EXCEL<]$0"
	[+] CustomWin EXCEL22
		[ ] tag "[EXCEL2]$0[2]"
		[+] CustomWin TaskPane
			[+] multitag "[MsoCommandBar]Task Pane"
				[ ] "[MsoCommandBar]$0"
			[+] CustomWin GettingStarted
				[+] multitag "[MsoWorkPane]Getting Started"
					[ ] "[MsoWorkPane]$0"
				[+] CustomWin NUIPane1
					[ ] tag "[NUIPane]$0"
					[+] CustomWin NetUIHWND1
						[ ] tag "[NetUIHWND]$0"
						[+] CustomWin CtrlNotifySink1
							[ ] tag "[CtrlNotifySink]$0"
	[+] CustomWin EXCEL21
		[ ] tag "[EXCEL2]$0[1]"
		[+] CustomWin WorksheetMenuBar
			[+] multitag "[MsoCommandBar]Worksheet Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$15992400"
		[+] CustomWin Formatting
			[+] multitag "[MsoCommandBar]Formatting"
				[ ] "[MsoCommandBar]$0[3]"
			[+] CustomWin RichEdit20W2
				[ ] tag "[RichEdit20W]$16010828"
			[+] CustomWin RichEdit20W1
				[ ] tag "[RichEdit20W]$16010708"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
[+] window MainWin Excel
	[+] multitag "Microsoft Excel"
		[ ] "*Microsoft Excel"
		[ ] "$C:\Program Files\Microsoft\Office11\EXCEL.exe"
[+] window DialogBox ExcelOpen
	[ ] tag "Open"
	[ ] parent MicrosoftExcel
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin Open
		[+] multitag "[Snake List]Open"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox ExcelPrint
	[ ] tag "Print"
	[ ] parent MicrosoftExcel
	[+] CustomWin EDTBX2
		[ ] tag "[EDTBX]$35"
	[+] CustomWin EDTBX3
		[ ] tag "[EDTBX]$37"
	[+] CustomWin N1
		[+] multitag "[EDTBX]1"
			[ ] "[EDTBX]$45"
[+] window DialogBox ExcelSaveChanges
	[ ] tag "Microsoft Excel"
	[ ] parent MicrosoftExcel
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin DoYouWantToSaveTheChange
		[+] multitag "[MSOUNISTAT]Do you want to save the changes you made to 'Book1'?"
			[ ] "[MSOUNISTAT]$4001"
[ ] 
[+] window DialogBox MicrosoftPowerPoint
	[ ] STRING sTag
	[ ] tag sTag
	[ ] // tag "Microsoft PowerPoint*"
	[+] CustomWin MsoDockRight
		[+] multitag "[MsoCommandBarDock]MsoDockRight"
			[ ] "[MsoCommandBarDock]$0[2]"
		[+] CustomWin TaskPane
			[+] multitag "[MsoCommandBar]Task Pane"
				[ ] "[MsoCommandBar]$0"
			[+] CustomWin GettingStarted
				[+] multitag "[MsoWorkPane]Getting Started"
					[ ] "[MsoWorkPane]$0"
				[+] CustomWin NUIPane1
					[ ] tag "[NUIPane]$0"
					[+] CustomWin NetUIHWND1
						[ ] tag "[NetUIHWND]$0"
						[+] CustomWin CtrlNotifySink1
							[ ] tag "[CtrlNotifySink]$0"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0[1]"
		[+] CustomWin Formatting
			[+] multitag "[MsoCommandBar]Formatting"
				[ ] "[MsoCommandBar]$0[3]"
		[+] CustomWin Standard
			[+] multitag "[MsoCommandBar]Standard"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MsoDockBottom
		[+] multitag "[MsoCommandBarDock]MsoDockBottom"
			[ ] "[MsoCommandBarDock]$0[3]"
		[+] CustomWin Drawing
			[+] multitag "[MsoCommandBar]Drawing"
				[ ] "[MsoCommandBar]$0"
	[+] StatusBar StatusBar1
		[ ] tag "$1"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
[+] window MainWin PowerPoint
	[+] multitag "Microsoft Power Point"
		[ ] "*Microsoft Power Point"
		[ ] "$C:\Program Files\Microsoft\Office11\POWERPNT.exe"
[+] window DialogBox PowerPointOpen
	[ ] tag "Open"
	[ ] parent MicrosoftPowerPoint
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin Open
		[+] multitag "[Snake List]Open"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
			[+] Header Header1
				[ ] tag "$0"
				[+] PushButton Name
					[ ] tag "Name"
				[+] PushButton Size
					[ ] tag "Size"
				[+] PushButton Type
					[ ] tag "Type"
				[+] PushButton DateModified
					[ ] tag "Date Modified"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox PowerPointSaveAs
	[ ] tag "Save As"
	[ ] parent MicrosoftPowerPoint
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin SaveAs
		[+] multitag "[Snake List]Save As"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox PowerPointSaveChanges
	[ ] tag "Microsoft Office PowerPoint"
	[ ] parent MicrosoftPowerPoint
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin DoYouWantToSaveTheChange
		[+] multitag "[MSOUNISTAT]Do you want to save the changes you made to Presentation1?"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox PowerPointReplace
	[ ] tag "Microsoft Office PowerPoint"
	[ ] parent PowerPointSaveAs
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin TheFileBighistoryPptAlread
		[+] multitag "[MSOUNISTAT]The file bighistory.ppt already exists. Do you want to replace the existing file?"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox PowerPointPrint
	[ ] tag "Print"
	[ ] parent MicrosoftPowerPoint
	[+] StaticText CopiesText
		[+] multitag "Copies"
			[ ] "$25"
	[+] StaticText NumberOfCopiesText
		[+] multitag "Number of copies:"
			[ ] "$26"
	[+] CustomWin NumberOfCopies1
		[+] multitag "[RichEdit20W]Number of copies:"
			[ ] "[RichEdit20W]$27"
	[+] CustomWin NumberOfCopies2
		[+] multitag "[UserControl]Number of copies:"
			[ ] "[UserControl]$28"
	[+] CheckBox Collate
		[+] multitag "Collate"
			[ ] "$29"
	[+] CustomWin Copies
		[+] multitag "[UserControl]Copies"
			[ ] "[UserControl]$30"
	[+] StaticText PrintWhatText
		[+] multitag "Print what:"
			[ ] "$31"
	[+] CustomWin Slides
		[+] multitag "[REComboBox20W]Slides"
			[ ] "[REComboBox20W]$32"
	[+] StaticText HandoutsText
		[+] multitag "Handouts"
			[ ] "$43"
	[+] StaticText SlidesPerPageText
		[+] multitag "Slides per page:"
			[ ] "$39"
	[+] CustomWin N6
		[+] multitag "[REComboBox20W]6"
			[ ] "[REComboBox20W]$44"
	[+] StaticText OrderText
		[+] multitag "Order:"
			[ ] "$38"
	[+] RadioList SlidesPerPage
		[+] multitag "Slides per page:"
			[ ] "$41"
	[+] CustomWin Handouts
		[+] multitag "[UserControl]Handouts"
			[ ] "[UserControl]$45"
	[+] StaticText ColorGrayscaleText
		[+] multitag "Color?grayscale:"
			[ ] "$47"
	[+] CustomWin Grayscale
		[+] multitag "[REComboBox20W]Grayscale"
			[ ] "[REComboBox20W]$48"
	[+] CheckBox ScaleToFitPaper
		[+] multitag "Scale to fit paper"
			[ ] "$34"
	[+] CheckBox FrameSlides
		[+] multitag "Frame slides"
			[ ] "$37"
	[+] CheckBox PrintHiddenSlides
		[+] multitag "Print hidden slides"
			[ ] "$33"
	[+] CheckBox PrintCommentsAndInkMarkup
		[+] multitag "Print comments and ink markup"
			[ ] "$49"
	[+] PushButton Preview
		[+] multitag "Preview"
			[ ] "$46"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] StaticText PrinterText
		[+] multitag "Printer"
			[ ] "$4"
	[+] StaticText NameText
		[+] multitag "Name:"
			[ ] "$5"
	[+] CustomWin HPLaserJet8000SeriesPCL
		[+] multitag "[REComboBox20W]HP LaserJet 8000 Series PCL"
			[ ] "[REComboBox20W]$6"
	[+] PushButton Properties
		[+] multitag "Properties"
			[ ] "$3"
	[+] PushButton FindPrinter
		[+] multitag "Find Printer"
			[ ] "$50"
	[+] StaticText StatusText
		[+] multitag "Status:"
			[ ] "$7"
	[+] CustomWin Status
		[+] multitag "[RichEdit20W]Status:"
			[ ] "[RichEdit20W]$8"
	[+] StaticText TypeText
		[+] multitag "Type:"
			[ ] "$13"
	[+] CustomWin Type
		[+] multitag "[RichEdit20W]Type:"
			[ ] "[RichEdit20W]$14"
	[+] StaticText WhereText
		[+] multitag "Where:"
			[ ] "$9"
	[+] CustomWin Where
		[+] multitag "[RichEdit20W]Where:"
			[ ] "[RichEdit20W]$10"
	[+] StaticText CommentText
		[+] multitag "Comment:"
			[ ] "$11"
	[+] CustomWin Comment
		[+] multitag "[RichEdit20W]Comment:"
			[ ] "[RichEdit20W]$12"
	[+] CheckBox PrintToFile
		[+] multitag "Print to file"
			[ ] "$15"
	[+] StaticText PrintRangeText
		[+] multitag "Print range"
			[ ] "$16"
	[+] RadioList PrintRange1
		[+] multitag "Print range"
			[ ] "$17"
	[+] StaticText CustomShowText
		[+] multitag "Custom Show:"
			[ ] "$30002"
	[+] CustomWin PrintRange2
		[+] multitag "[REComboBox20W]Print range"
			[ ] "[REComboBox20W]$22"
	[+] StaticText SlidesText
		[+] multitag "Slides:"
			[ ] "$30003"
	[+] CustomWin PrintRange3
		[+] multitag "[RichEdit20W]Print range"
			[ ] "[RichEdit20W]$23"
	[+] StaticText EnterSlideNumbersAndOrSli
		[+] multitag "Enter slide numbers and?or slide ranges. For example, 1,3,5-12"
			[ ] "$24"
[ ] 
[+] window DialogBox MicrosoftAccess
	[ ] STRING sTag
	[ ] tag sTag
	[+] CustomWin MsoDockRight
		[+] multitag "[MsoCommandBarDock]MsoDockRight"
			[ ] "[MsoCommandBarDock]$0[2]"
		[+] CustomWin TaskPane
			[+] multitag "[MsoCommandBar]Task Pane"
				[ ] "[MsoCommandBar]$0"
			[+] CustomWin GettingStarted
				[+] multitag "[MsoWorkPane]Getting Started"
					[ ] "[MsoWorkPane]$0"
				[+] CustomWin NUIPane1
					[ ] tag "[NUIPane]$0"
					[+] CustomWin NetUIHWND1
						[ ] tag "[NetUIHWND]$0"
						[+] CustomWin CtrlNotifySink1
							[ ] tag "[CtrlNotifySink]$0"
	[+] CustomWin MsoDockTop
		[+] multitag "[MsoCommandBarDock]MsoDockTop"
			[ ] "[MsoCommandBarDock]$0[1]"
		[+] CustomWin Database
			[+] multitag "[MsoCommandBar]Database"
				[ ] "[MsoCommandBar]$0[2]"
		[+] CustomWin MenuBar
			[+] multitag "[MsoCommandBar]Menu Bar"
				[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin OStatbar1
		[ ] tag "[OStatbar]$0"
[+] window MainWin Access
	[+] multitag "Microsoft Access*"
		[ ] "*Microsoft Access"
		[ ] "$C:\Program Files\Microsoft\Office11\MSACCESS.exe"
[+] window DialogBox AccessOpen
	[ ] tag "Open"
	[ ] parent MicrosoftAccess
	[+] CustomWin MSOGenericControlContainer1
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[1]"
			[ ] "[MsoCommandBar]$0[1]"
	[+] CustomWin MSOGenericControlContainer2
		[+] multitag "[MsoCommandBar]MSO Generic Control Container[2]"
			[ ] "[MsoCommandBar]$0[2]"
	[+] CustomWin Open
		[+] multitag "[Snake List]Open"
			[ ] "[Snake List]$0"
		[+] ListView ListView1
			[ ] tag "$1"
	[+] CustomWin Filename
		[ ] tag "[RichEdit20W]$54"
[+] window DialogBox SecurityWarning
	[ ] tag "Microsoft Office Access"
	[ ] parent MicrosoftAccess
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "$9"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin MicrosoftOfficeAccessCannot
		[+] multitag "[MSOUNISTAT]Microsoft Office Access cannot open C:\vb\db1.mdb due to security restrictions."
			[ ] "[MSOUNISTAT]$4000"
	[+] CustomWin SecuritySettingsRestrictAcc
		[+] multitag "[MSOUNISTAT]Security settings restrict access to the file because it is not digitally signed."
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox SecurityWarning1
	[ ] tag "Microsoft Office Access"
	[ ] parent MicrosoftAccess
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$2"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "$9"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin SecurityWarningUnsafeExpre
		[+] multitag "[MSOUNISTAT]Security Warning: Unsafe expressions are not blocked."
			[ ] "[MSOUNISTAT]$4000"
	[+] CustomWin DoYouWantToBlockUnsafeEx
		[+] multitag "[MSOUNISTAT]Do you want to block unsafe expressions?"
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox SecurityWarning2
	[ ] tag "Microsoft Office Access"
	[ ] parent MicrosoftAccess
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$-1"
	[+] CustomWin YouMustRestartTheApplicati
		[+] multitag "[MSOUNISTAT]You must restart the application for the new security level to take effect."
			[ ] "[MSOUNISTAT]$4001"
[+] window DialogBox SecurityWarning3
	[ ] tag "Security Warning"
	[ ] parent MicrosoftAccess
[+] window DialogBox AccessPrint
	[ ] tag "Print"
	[ ] parent MicrosoftAccess
	[+] StaticText PrinterText
		[+] multitag "Printer"
			[ ] "$1075"
	[+] StaticText NameText
		[+] multitag "Name:"
			[ ] "$1093"
	[+] PopupList Name
		[+] multitag "Name:"
			[ ] "$1139"
	[+] PushButton Properties
		[+] multitag "Properties"
			[ ] "$1025"
	[+] StaticText StatusText
		[+] multitag "Status:"
			[ ] "$1095"
	[+] StaticText ReadyText
		[+] multitag "Ready"
			[ ] "$1099"
	[+] StaticText TypeText
		[+] multitag "Type:"
			[ ] "$1094"
	[+] StaticText HPLaserJet8000SeriesPCLTex
		[+] multitag "HP LaserJet 8000 Series PCL"
			[ ] "$1098"
	[+] StaticText WhereText
		[+] multitag "Where:"
			[ ] "$1097"
	[+] StaticText N19216811199100Text
		[+] multitag "192.168.1.119:9100"
			[ ] "$1101"
	[+] StaticText Comment1Text
		[+] multitag "Comment:[1]"
			[ ] "$1096"
	[+] StaticText Comment2Text
		[+] multitag "Comment:[2]"
			[ ] "$1100"
	[+] CheckBox PrintToFile
		[+] multitag "Print to File"
			[ ] "$1040"
	[+] StaticText PrintRangeText
		[+] multitag "Print Range"
			[ ] "$1072"
	[+] RadioList PrintRange
		[+] multitag "Print Range"
			[ ] "$1056"
	[+] StaticText FromText
		[+] multitag "From:"
			[ ] "$1089"
	[+] TextField From
		[+] multitag "From:"
			[ ] "$1152"
	[+] StaticText ToText
		[+] multitag "To:"
			[ ] "$1090"
	[+] TextField To
		[+] multitag "To:"
			[ ] "$1153"
	[+] StaticText CopiesText
		[+] multitag "Copies"
			[ ] "$1073"
	[+] StaticText NumberOfCopiesText
		[+] multitag "Number of Copies:"
			[ ] "$1092"
	[+] TextField NumberOfCopies
		[+] multitag "Number of Copies:"
			[ ] "$1154"
	[+] CustomWin Copies1
		[+] multitag "[Icon]Copies"
			[ ] "[Icon]$1086"
	[+] CheckBox Collate
		[+] multitag "Collate"
			[ ] "$1041"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] PushButton Setup
		[+] multitag "Setup"
			[ ] "$1965"
	[+] UpDown Copies2
		[+] multitag "Copies"
			[ ] "$9999"
[+] window ChildWin CreatePurchaseOrder
	[ ] tag "Create Purchase Order"
	[ ] parent MicrosoftAccess
	[+] CustomWin OFormSub2
		[ ] tag "[OFormSub]$0[2]"
	[+] CustomWin OFormSub1
		[ ] tag "[OFormSub]$0[1]"
		[+] CustomWin OFormChild1
			[ ] tag "[OFormChild]$0"
			[+] CustomWin OSUI1
				[ ] tag "[OSUI]$0"
			[+] CustomWin OFormSub1
				[ ] tag "[OFormSub]$0[1]"
			[+] CustomWin OFormSub2
				[ ] tag "[OFormSub]$0[2]"
				[+] CustomWin OFormChild1
					[ ] tag "[OFormChild]$0"
					[+] CustomWin OBlank1
						[ ] tag "[OBlank]$4"
					[+] CustomWin OSUI1
						[ ] tag "[OSUI]$0"
					[+] ScrollBar ScrollBar1
						[ ] tag "$2"
					[+] CustomWin OFormSub1
						[ ] tag "[OFormSub]$0[1]"
					[+] CustomWin OFormSub2
						[ ] tag "[OFormSub]$0[2]"
						[+] CustomWin N100004
							[+] multitag "[OKttbx]100004"
								[ ] "[OKttbx]$0"
					[+] CustomWin OFormSub3
						[ ] tag "[OFormSub]$0[3]"
			[+] CustomWin OFormSub3
				[ ] tag "[OFormSub]$0[3]"
	[+] CustomWin OFormSub3
		[ ] tag "[OFormSub]$0[3]"
[ ] 
[ ] // Below Declarations are for Adobe Reader 8.0
[+] window MainWin AdobeReader
	[ ] STRING sTag
	[ ] tag sTag
	[+] // multitag "Adobe Reader"
		[ ] // "$C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe"
	[+] Menu File
		[ ] tag "File"
		[+] MenuItem MenuItem1
			[ ] tag "$6423"
		[+] MenuItem MenuItem2
			[ ] tag "$6001"
		[+] MenuItem MenuItem3
			[ ] tag "$6419"
		[+] Menu Menu4
			[ ] tag "#4"
			[+] MenuItem MenuItem1
				[ ] tag "$6377"
			[+] MenuItem MenuItem2
				[ ] tag "$6379"
			[+] MenuItem MenuItem3
				[ ] tag "$6380"
		[+] MenuItem MenuItem5
			[ ] tag "$6005"
		[+] MenuItem MenuItem6
			[ ] tag "$6006"
		[+] MenuItem MenuItem7
			[ ] tag "$6007"
		[+] MenuItem MenuItem8
			[ ] tag "$6350"
		[+] MenuItem MenuItem9
			[ ] tag "$6008"
		[+] MenuItem Print
			[ ] tag "$6418"
		[+] MenuItem MenuItem11
			[ ] tag "$6225"
		[+] MenuItem MenuItem12
			[ ] tag "$6226"
		[+] MenuItem Exit
			[ ] tag "$6310"
		[+] MenuItem MenuItem14
			[ ] tag "$6014"
		[+] MenuItem MenuItem15
			[ ] tag "$6018"
		[+] MenuItem MenuItem16
			[ ] tag "$6019"
		[+] MenuItem MenuItem17
			[ ] tag "$6021"
		[+] MenuItem MenuItem18
			[ ] tag "$6024"
		[+] MenuItem MenuItem19
			[ ] tag "$6026"
	[+] Menu Edit
		[ ] tag "Edit"
		[+] MenuItem MenuItem1
			[ ] tag "$6027"
		[+] MenuItem MenuItem2
			[ ] tag "$6028"
		[+] MenuItem MenuItem3
			[ ] tag "$6030"
		[+] MenuItem MenuItem4
			[ ] tag "$6031"
		[+] MenuItem MenuItem5
			[ ] tag "$6032"
		[+] MenuItem MenuItem6
			[ ] tag "$6033"
		[+] MenuItem MenuItem7
			[ ] tag "$6205"
		[+] MenuItem MenuItem8
			[ ] tag "$6035"
		[+] MenuItem MenuItem9
			[ ] tag "$6036"
		[+] Menu Menu10
			[ ] tag "#10"
			[+] MenuItem MenuItem1
				[ ] tag "$6421"
			[+] MenuItem MenuItem2
				[ ] tag "$6422"
		[+] MenuItem MenuItem11
			[ ] tag "$6039"
		[+] MenuItem MenuItem12
			[ ] tag "$6041"
		[+] MenuItem MenuItem13
			[ ] tag "$6043"
		[+] MenuItem MenuItem14
			[ ] tag "$6044"
		[+] Menu Menu15
			[ ] tag "#15"
			[+] MenuItem MenuItem1
				[ ] tag "$6045"
			[+] MenuItem MenuItem2
				[ ] tag "$6046"
			[+] MenuItem MenuItem3
				[ ] tag "$6047"
			[+] MenuItem MenuItem4
				[ ] tag "$6048"
		[+] MenuItem MenuItem16
			[ ] tag "$6051"
		[+] MenuItem MenuItem17
			[ ] tag "$6053"
	[+] Menu View
		[ ] tag "View"
		[+] Menu Menu1
			[ ] tag "#1"
			[+] MenuItem MenuItem1
				[ ] tag "$6521"
			[+] MenuItem MenuItem2
				[ ] tag "$6523"
			[+] MenuItem MenuItem3
				[ ] tag "$6519"
			[+] MenuItem MenuItem4
				[ ] tag "$6524"
			[+] MenuItem MenuItem5
				[ ] tag "$6522"
			[+] MenuItem MenuItem6
				[ ] tag "$6520"
			[+] MenuItem MenuItem7
				[ ] tag "$6525"
			[+] MenuItem MenuItem8
				[ ] tag "$6060"
			[+] MenuItem MenuItem9
				[ ] tag "$6061"
			[+] MenuItem MenuItem10
				[ ] tag "$6062"
		[+] Menu Menu2
			[ ] tag "#2"
			[+] MenuItem MenuItem1
				[ ] tag "$6425"
			[+] MenuItem MenuItem2
				[ ] tag "$6427"
			[+] MenuItem MenuItem3
				[ ] tag "$6430"
			[+] MenuItem MenuItem4
				[ ] tag "$6436"
			[+] MenuItem MenuItem5
				[ ] tag "$6437"
			[+] MenuItem MenuItem6
				[ ] tag "$6065"
		[+] Menu Menu3
			[ ] tag "#3"
			[+] MenuItem MenuItem1
				[ ] tag "$6507"
			[+] MenuItem MenuItem2
				[ ] tag "$6518"
			[+] MenuItem MenuItem3
				[ ] tag "$6517"
			[+] MenuItem MenuItem4
				[ ] tag "$6511"
			[+] MenuItem MenuItem5
				[ ] tag "$6506"
			[+] MenuItem MenuItem6
				[ ] tag "$6515"
			[+] MenuItem MenuItem7
				[ ] tag "$6513"
			[+] MenuItem MenuItem8
				[ ] tag "$6514"
			[+] MenuItem MenuItem9
				[ ] tag "$6508"
			[+] MenuItem MenuItem10
				[ ] tag "$6510"
			[+] MenuItem MenuItem11
				[ ] tag "$6516"
			[+] MenuItem MenuItem12
				[ ] tag "$6512"
			[+] MenuItem MenuItem13
				[ ] tag "$6509"
			[+] MenuItem MenuItem14
				[ ] tag "$6068"
			[+] MenuItem MenuItem15
				[ ] tag "$6070"
			[+] MenuItem MenuItem16
				[ ] tag "$6071"
			[+] MenuItem MenuItem17
				[ ] tag "$6072"
			[+] MenuItem MenuItem18
				[ ] tag "$6073"
			[+] Menu Menu19
				[ ] tag "#19"
				[+] MenuItem MenuItem1
					[ ] tag "$6074"
				[+] MenuItem MenuItem2
					[ ] tag "$6075"
				[+] MenuItem MenuItem3
					[ ] tag "$6076"
		[+] MenuItem MenuItem4
			[ ] tag "$6079"
		[+] MenuItem MenuItem5
			[ ] tag "$6081"
		[+] MenuItem MenuItem6
			[ ] tag "$6083"
		[+] MenuItem MenuItem7
			[ ] tag "$6084"
		[+] MenuItem MenuItem8
			[ ] tag "$6085"
		[+] MenuItem MenuItem9
			[ ] tag "$6086"
		[+] MenuItem MenuItem10
			[ ] tag "$6087"
		[+] MenuItem MenuItem11
			[ ] tag "$6088"
		[+] MenuItem MenuItem12
			[ ] tag "$6091"
		[+] Menu Menu13
			[ ] tag "#13"
			[+] MenuItem MenuItem1
				[ ] tag "$6208"
			[+] MenuItem MenuItem2
				[ ] tag "$6209"
			[+] MenuItem MenuItem3
				[ ] tag "$6210"
			[+] MenuItem MenuItem4
				[ ] tag "$6211"
		[+] Menu Menu14
			[ ] tag "#14"
			[+] MenuItem MenuItem1
				[ ] tag "$6094"
			[+] MenuItem MenuItem2
				[ ] tag "$6095"
			[+] MenuItem MenuItem3
				[ ] tag "$6096"
			[+] MenuItem MenuItem4
				[ ] tag "$6097"
			[+] MenuItem MenuItem5
				[ ] tag "$6098"
			[+] MenuItem MenuItem6
				[ ] tag "$6100"
			[+] MenuItem MenuItem7
				[ ] tag "$6101"
			[+] MenuItem MenuItem8
				[ ] tag "$6103"
			[+] MenuItem MenuItem9
				[ ] tag "$6104"
		[+] Menu Menu15
			[ ] tag "#15"
			[+] MenuItem MenuItem1
				[ ] tag "$6108"
			[+] MenuItem MenuItem2
				[ ] tag "$6109"
			[+] MenuItem MenuItem3
				[ ] tag "$6110"
			[+] MenuItem MenuItem4
				[ ] tag "$6111"
		[+] Menu Menu16
			[ ] tag "#16"
			[+] MenuItem MenuItem1
				[ ] tag "$6115"
			[+] MenuItem MenuItem2
				[ ] tag "$6116"
		[+] MenuItem MenuItem17
			[ ] tag "$6120"
		[+] MenuItem MenuItem18
			[ ] tag "$6228"
		[+] MenuItem MenuItem19
			[ ] tag "$6230"
	[+] Menu Document
		[ ] tag "Document"
		[+] Menu Menu1
			[ ] tag "#1"
			[+] MenuItem MenuItem1
				[ ] tag "$6340"
			[+] MenuItem MenuItem2
				[ ] tag "$6315"
			[+] MenuItem MenuItem3
				[ ] tag "$6321"
			[+] MenuItem MenuItem4
				[ ] tag "$6344"
			[+] MenuItem MenuItem5
				[ ] tag "$6317"
		[+] Menu Menu2
			[ ] tag "#2"
		[+] MenuItem MenuItem3
			[ ] tag "$6404"
		[+] MenuItem MenuItem4
			[ ] tag "$6403"
		[+] MenuItem MenuItem5
			[ ] tag "$6206"
		[+] MenuItem MenuItem6
			[ ] tag "$6417"
		[+] MenuItem MenuItem7
			[ ] tag "$6136"
	[+] Menu Tools
		[ ] tag "Tools"
		[+] Menu Menu1
			[ ] tag "#1"
			[+] MenuItem MenuItem1
				[ ] tag "$6488"
			[+] Menu Menu2
				[ ] tag "#2"
				[+] MenuItem MenuItem1
					[ ] tag "$6217"
				[+] MenuItem MenuItem2
					[ ] tag "$6219"
				[+] MenuItem MenuItem3
					[ ] tag "$6220"
				[+] MenuItem MenuItem4
					[ ] tag "$6221"
				[+] MenuItem MenuItem5
					[ ] tag "$6222"
				[+] MenuItem MenuItem6
					[ ] tag "$6223"
				[+] MenuItem MenuItem7
					[ ] tag "$6224"
				[+] MenuItem MenuItem8
					[ ] tag "$6215"
			[+] Menu Menu3
				[ ] tag "#3"
				[+] MenuItem MenuItem1
					[ ] tag "$6490"
			[+] Menu Menu4
				[ ] tag "#4"
				[+] MenuItem MenuItem1
					[ ] tag "$6232"
				[+] MenuItem MenuItem2
					[ ] tag "$6233"
				[+] MenuItem MenuItem3
					[ ] tag "$6235"
				[+] MenuItem MenuItem4
					[ ] tag "$6237"
				[+] Menu Menu5
					[ ] tag "#5"
					[+] MenuItem MenuItem1
						[ ] tag "$6238"
					[+] MenuItem MenuItem2
						[ ] tag "$6239"
					[+] MenuItem MenuItem3
						[ ] tag "$6240"
					[+] MenuItem MenuItem4
						[ ] tag "$6241"
					[+] MenuItem MenuItem5
						[ ] tag "$6242"
					[+] MenuItem MenuItem6
						[ ] tag "$6243"
				[+] Menu Menu6
					[ ] tag "#6"
				[+] Menu Menu7
					[ ] tag "#7"
				[+] Menu Menu8
					[ ] tag "#8"
					[+] MenuItem MenuItem1
						[ ] tag "$6247"
					[+] MenuItem MenuItem2
						[ ] tag "$6248"
					[+] MenuItem MenuItem3
						[ ] tag "$6249"
				[+] MenuItem MenuItem9
					[ ] tag "$6252"
				[+] MenuItem MenuItem10
					[ ] tag "$6253"
				[+] MenuItem MenuItem11
					[ ] tag "$6255"
				[+] MenuItem MenuItem12
					[ ] tag "$6256"
			[+] MenuItem MenuItem5
				[ ] tag "$6493"
			[+] MenuItem MenuItem6
				[ ] tag "$6494"
			[+] MenuItem MenuItem7
				[ ] tag "$6496"
			[+] MenuItem MenuItem8
				[ ] tag "$6497"
			[+] MenuItem MenuItem9
				[ ] tag "$6498"
			[+] MenuItem MenuItem10
				[ ] tag "$6499"
			[+] MenuItem MenuItem11
				[ ] tag "$6501"
			[+] MenuItem MenuItem12
				[ ] tag "$6502"
			[+] MenuItem MenuItem13
				[ ] tag "$6503"
		[+] Menu Menu2
			[ ] tag "#2"
			[+] MenuItem MenuItem1
				[ ] tag "$6470"
			[+] MenuItem MenuItem2
				[ ] tag "$6471"
			[+] MenuItem MenuItem3
				[ ] tag "$6472"
			[+] MenuItem MenuItem4
				[ ] tag "$6473"
			[+] MenuItem MenuItem5
				[ ] tag "$6474"
			[+] MenuItem MenuItem6
				[ ] tag "$6476"
			[+] MenuItem MenuItem7
				[ ] tag "$6477"
			[+] MenuItem MenuItem8
				[ ] tag "$6478"
			[+] MenuItem MenuItem9
				[ ] tag "$6479"
			[+] MenuItem MenuItem10
				[ ] tag "$6480"
			[+] MenuItem MenuItem11
				[ ] tag "$6481"
			[+] MenuItem MenuItem12
				[ ] tag "$6482"
			[+] MenuItem MenuItem13
				[ ] tag "$6483"
			[+] MenuItem MenuItem14
				[ ] tag "$6484"
			[+] MenuItem MenuItem15
				[ ] tag "$6485"
		[+] Menu Menu3
			[ ] tag "#3"
			[+] MenuItem MenuItem1
				[ ] tag "$6453"
			[+] MenuItem MenuItem2
				[ ] tag "$6454"
			[+] MenuItem MenuItem3
				[ ] tag "$6455"
			[+] MenuItem MenuItem4
				[ ] tag "$6456"
		[+] Menu Menu4
			[ ] tag "#4"
			[+] MenuItem MenuItem1
				[ ] tag "$6466"
			[+] MenuItem MenuItem2
				[ ] tag "$6467"
			[+] MenuItem MenuItem3
				[ ] tag "$6468"
			[+] MenuItem MenuItem4
				[ ] tag "$6464"
		[+] Menu Menu5
			[ ] tag "#5"
			[+] MenuItem MenuItem1
				[ ] tag "$6459"
			[+] MenuItem MenuItem2
				[ ] tag "$6460"
	[+] Menu xWindow
		[ ] tag "Window"
		[+] MenuItem MenuItem1
			[ ] tag "$6138"
		[+] Menu Menu2
			[ ] tag "#2"
			[+] MenuItem MenuItem1
				[ ] tag "$6139"
			[+] MenuItem MenuItem2
				[ ] tag "$6140"
			[+] MenuItem MenuItem3
				[ ] tag "$6141"
		[+] MenuItem MenuItem3
			[ ] tag "$6143"
		[+] MenuItem MenuItem4
			[ ] tag "$6145"
		[+] MenuItem MenuItem5
			[ ] tag "$6147"
	[+] Menu Help
		[ ] tag "Help"
		[+] Menu Menu1
			[ ] tag "#1"
			[+] MenuItem MenuItem1
				[ ] tag "$6446"
			[+] MenuItem MenuItem2
				[ ] tag "$6447"
			[+] MenuItem MenuItem3
				[ ] tag "$6448"
			[+] MenuItem MenuItem4
				[ ] tag "$6449"
			[+] MenuItem MenuItem5
				[ ] tag "$6450"
		[+] MenuItem MenuItem2
			[ ] tag "$6151"
		[+] MenuItem About
			[ ] tag "$6158"
		[+] MenuItem MenuItem4
			[ ] tag "$6159"
		[+] MenuItem MenuItem5
			[ ] tag "$6163"
		[+] MenuItem MenuItem6
			[ ] tag "$6164"
		[+] MenuItem MenuItem7
			[ ] tag "$6165"
		[+] MenuItem MenuItem8
			[ ] tag "$6166"
		[+] MenuItem MenuItem9
			[ ] tag "$6167"
		[+] MenuItem MenuItem10
			[ ] tag "$6168"
		[+] MenuItem MenuItem11
			[ ] tag "$6169"
		[+] MenuItem MenuItem12
			[ ] tag "$6207"
	[+] CustomWin AVEaselBarWindow4
		[ ] tag "[AVEaselBarWindow]$59393"
		[+] CustomWin AVEaselBarWindow1
			[ ] tag "[AVEaselBarWindow]$0"
	[+] CustomWin AVEaselBarWindow1
		[ ] tag "[AVEaselBarWindow]$59392"
		[+] CustomWin AVEaselBarWindow1
			[ ] tag "[AVEaselBarWindow]$0"
			[+] CustomWin AVToolBarView7
				[+] multitag "[AVL_AVView]AVToolBarView[7]"
					[ ] "[AVL_AVView]$0[7]"
			[+] CustomWin AVToolBarView5
				[+] multitag "[AVL_AVView]AVToolBarView[5]"
					[ ] "[AVL_AVView]$0[5]"
			[+] CustomWin AVToolBarView6
				[+] multitag "[AVL_AVView]AVToolBarView[6]"
					[ ] "[AVL_AVView]$0[6]"
			[+] CustomWin AVToolBarView4
				[+] multitag "[AVL_AVView]AVToolBarView[4]"
					[ ] "[AVL_AVView]$0[4]"
				[+] CustomWin AVToolBarFlyoutView
					[+] multitag "[AVL_AVView]AVToolBarFlyoutView"
						[ ] "[AVL_AVView]$0"
			[+] CustomWin AVToolBarView3
				[+] multitag "[AVL_AVView]AVToolBarView[3]"
					[ ] "[AVL_AVView]$0[3]"
				[+] CustomWin AVZoomView
					[+] multitag "[AVL_AVView]AVZoomView"
						[ ] "[AVL_AVView]$0[2]"
					[+] TextField TextField1
						[ ] tag "$0"
				[+] CustomWin AVToolBarFlyoutView
					[+] multitag "[AVL_AVView]AVToolBarFlyoutView"
						[ ] "[AVL_AVView]$0[1]"
			[+] CustomWin AVToolBarView2
				[+] multitag "[AVL_AVView]AVToolBarView[2]"
					[ ] "[AVL_AVView]$0[2]"
			[+] CustomWin AVToolBarView1
				[+] multitag "[AVL_AVView]AVToolBarView[1]"
					[ ] "[AVL_AVView]$0[1]"
			[+] CustomWin AVToolBarView8
				[+] multitag "[AVL_AVView]AVToolBarView[8]"
					[ ] "[AVL_AVView]$0[8]"
	[+] CustomWin AVEaselBarWindow2
		[ ] tag "[AVEaselBarWindow]$59394"
		[+] CustomWin AVEaselBarWindow1
			[ ] tag "[AVEaselBarWindow]$0"
	[+] CustomWin AVEaselBarWindow3
		[ ] tag "[AVEaselBarWindow]$59395"
		[+] CustomWin AVEaselBarWindow1
			[ ] tag "[AVEaselBarWindow]$0"
[+] window DialogBox AdobeReaderOpen
	[ ] tag "Open"
	[ ] parent AdobeReader
	[+] StaticText LookInText
		[+] multitag "Look in:"
			[ ] "$1091"
	[+] PopupList LookIn
		[+] multitag "Look in:"
			[ ] "$1137"
	[+] ToolBar LookIn1
		[+] multitag "Look in:[1]"
			[ ] "$1088"
		[+] PushButton PushButton1
			[ ] tag "$40971"
		[+] PushButton PushButton2
			[ ] tag "$40961"
		[+] PushButton PushButton3
			[ ] tag "$40962"
		[+] PushButton PushButton4
			[ ] tag "$40970"
	[+] ToolBar LookIn2
		[+] multitag "Look in:[2]"
			[ ] "$1184"
		[+] PushButton MyRecentDocuments
			[+] multitag "My Recent Documents"
				[ ] "$41060"
		[+] PushButton Desktop
			[+] multitag "Desktop"
				[ ] "$41061"
		[+] PushButton MyDocuments
			[+] multitag "My Documents"
				[ ] "$41062"
		[+] PushButton MyComputer
			[+] multitag "My Computer"
				[ ] "$41063"
		[+] PushButton MyNetworkPlaces
			[+] multitag "My Network Places"
				[ ] "$41064"
	[+] ListView ListView1
		[ ] tag "$1"
	[+] StaticText FileNameText
		[+] multitag "File name:"
			[ ] "$1090"
	[+] ComboBox FileName
		[+] multitag "File name:"
			[ ] "$1148"
	[+] StaticText FilesOfTypeText
		[+] multitag "Files of type:"
			[ ] "$1089"
	[+] PopupList FilesOfType1
		[+] multitag "Files of type:"
			[ ] "$1136"
	[+] PushButton Open
		[+] multitag "Open"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] ScrollBar FilesOfType2
		[+] multitag "Files of type:"
			[ ] "$-1"
	[+] DialogBox DialogBox1
		[ ] tag "$0"
[+] window DialogBox AboutAdobeReader
	[ ] tag "About Adobe Reader 7.0"
	[ ] parent AdobeReader
	[+] PushButton Version
		[+] multitag "Version 7.0.0"
			[ ] "$254"
	[+] StaticText Copyright19842004AdobeSyst
		[+] multitag "Copyright 1984-2004 Adobe Systems Incorporated and its licensors. All rights reserved."
			[ ] "$403"
	[+] PushButton Credits
		[+] multitag "Credits"
			[ ] "$571"
	[+] PushButton PatentAndLegalNotices
		[+] multitag "Patent and Legal Notices"
			[ ] "$570"
[+] window DialogBox Upgrade
	[ ] tag "Upgrade today to New Adobe Reader*"
	[ ] parent AdobeReader
	[+] CustomWin GroupBox1
		[ ] tag "[GroupBox]$0"
		[+] StaticText NewAdobeReader811Availab
			[+] multitag "New Adobe Reader 8.1.1 available now!????Work faster - hardware acceleration?, better forms performance??Work smarter - a clea*"
				[ ] "$0"
		[+] PushButton Yes
			[+] multitag "Yes"
				[ ] "$0[1]"
		[+] PushButton No
			[+] multitag "No"
				[ ] "$0[2]"
		[+] PushButton MoreInfo
			[+] multitag "More Info"
				[ ] "$0[3]"
[+] window DialogBox AdobeReaderPrint
	[ ] tag "Print"
	[ ] parent AdobeReader
	[+] StaticText PrinterText
		[+] multitag "Printer"
			[ ] "$1075"
	[+] StaticText NameText
		[+] multitag "Name:"
			[ ] "$1093"
	[+] PopupList Printer
		[+] multitag "Printer"
			[ ] "$1139"
	[+] PushButton Properties
		[+] multitag "Properties"
			[ ] "$1025"
	[+] StaticText StatusText
		[+] multitag "Status:"
			[ ] "$1095"
	[+] StaticText ReadyText
		[+] multitag "Ready"
			[ ] "$1099"
	[+] StaticText CommentsAndFormsText
		[+] multitag "Comments and Forms:"
			[ ] "$-1[1]"
	[+] PopupList CommentsAndForms
		[+] multitag "Comments and Forms:"
			[ ] "$558"
	[+] RadioList PrintRange
		[+] multitag "Print Range"
			[ ] "$1056"
	[+] TextField From
		[+] multitag "from:"
			[ ] "$1155"
	[+] StaticText ToText
		[+] multitag "to:"
			[ ] "$1090"
	[+] TextField To
		[+] multitag "to:"
			[ ] "$1156"
	[+] StaticText SubsetText
		[+] multitag "Subset:"
			[ ] "$-1[3]"
	[+] PopupList Subset
		[+] multitag "Subset:"
			[ ] "$533"
	[+] CheckBox ReversePages
		[+] multitag "Reverse pages"
			[ ] "$478"
	[+] StaticText CopiesText
		[+] multitag "Copies:"
			[ ] "$1092"
	[+] TextField Copies
		[+] multitag "Copies:"
			[ ] "$1154"
	[+] CheckBox Collate
		[+] multitag "Collate"
			[ ] "$1041"
	[+] StaticText PageScalingText
		[+] multitag "Page Scaling:"
			[ ] "$-1[5]"
	[+] PopupList PageScaling
		[+] multitag "Page Scaling:"
			[ ] "$557"
	[+] CheckBox AutoRotateAndCenter
		[+] multitag "Auto-Rotate and Center"
			[ ] "$563"
	[+] CheckBox ChoosePaperSourceByPDFPag
		[+] multitag "Choose Paper Source by PDF page size"
			[ ] "$564"
	[+] CheckBox PrintToFile
		[+] multitag "Print to file"
			[ ] "$1040"
	[+] PushButton PrintingTips
		[+] multitag "Printing Tips"
			[ ] "$543"
	[+] PushButton Advanced
		[+] multitag "Advanced"
			[ ] "$3"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] StaticText PrintRangeText
		[+] multitag "Print Range"
			[ ] "$1072"
	[+] StaticText FromText
		[+] multitag "from:"
			[ ] "$1089"
	[+] StaticText PageHandling1Text
		[+] multitag "Page Handling[1]"
			[ ] "$-1[4]"
	[+] StaticText PreviewText
		[+] multitag "Preview"
			[ ] "$-1[2]"
	[+] StaticText TypeText
		[+] multitag "Type:"
			[ ] "$1096"
	[+] StaticText HPLaserJet8000SeriesPCLTex
		[+] multitag "HP LaserJet 8000 Series PCL"
			[ ] "$1098"
	[+] StaticText UnitsInches
		[+] multitag "Units: Inches                      Zoom:  94%"
			[ ] "$562"
	[+] PushButton Preview
		[+] multitag "Preview"
			[ ] "$547"
	[+] StaticText PageHandling2Text
		[+] multitag "Page Handling[2]"
			[ ] "$574"
	[+] StaticText N151Text
		[+] multitag "1?5 (1)"
			[ ] "$593"
	[+] Scale N151
		[+] multitag "1?5 (1)"
			[ ] "$594"
	[+] UpDown PageHandling
		[+] multitag "Page Handling"
			[ ] "$9999"
[ ] 
[+] window MainWin WinZip
	[+] multitag "WinZip (Evaluation Version)"
		[ ] "$C:\Program Files\WinZip\WINZIP32.EXE"
	[+] Menu File
		[+] multitag "File"
			[ ] "#1"
		[+] MenuItem NewArchive
			[+] multitag "New Archive"
				[ ] "#1"
				[ ] "$4752"
		[+] MenuItem OpenArchive
			[+] multitag "Open Archive"
				[ ] "#2"
				[ ] "$4753"
		[+] MenuItem FavoriteZipFolders
			[+] multitag "Favorite Zip Folders"
				[ ] "#3"
				[ ] "$114"
		[+] MenuItem NewZipFileOnCDOrDVD
			[+] multitag "New Zip File on CD or DVD"
				[ ] "#4"
				[ ] "$4772"
		[+] MenuItem CloseArchive
			[+] multitag "Close Archive"
				[ ] "#5"
				[ ] "$43"
		[+] Menu ManageArchive
			[+] multitag "Manage Archive"
				[ ] "#6"
			[+] MenuItem Move
				[+] multitag "Move"
					[ ] "#1"
					[ ] "$85"
			[+] MenuItem Copy
				[+] multitag "Copy"
					[ ] "#2"
					[ ] "$87"
			[+] MenuItem Rename
				[+] multitag "Rename"
					[ ] "#3"
					[ ] "$88"
			[+] MenuItem Delete
				[+] multitag "Delete"
					[ ] "#4"
					[ ] "$89"
		[+] Menu SendTo
			[+] multitag "Send To"
				[ ] "#7"
			[+] MenuItem MailRecipient
				[+] multitag "Mail Recipient"
					[ ] "#1"
					[ ] "$101"
			[+] MenuItem FTPSite
				[+] multitag "FTP Site"
					[ ] "#2"
					[ ] "$25167"
			[+] MenuItem CDOrDVD
				[+] multitag "CD or DVD"
					[ ] "#3"
					[ ] "$25166"
		[+] MenuItem Print
			[+] multitag "Print"
				[ ] "#8"
				[ ] "$71"
		[+] MenuItem Properties
			[+] multitag "Properties"
				[ ] "#9"
				[ ] "$83"
		[+] MenuItem CreateShortcut
			[+] multitag "Create Shortcut"
				[ ] "#10"
				[ ] "$112"
		[+] MenuItem Wizard
			[+] multitag "Wizard"
				[ ] "#11"
				[ ] "$111"
		[+] Menu RecentHistory
			[+] multitag "Recent (History)"
				[ ] "#12"
			[+] MenuItem Empty
				[+] multitag "(Empty)"
					[ ] "#1"
					[ ] "$158"
		[+] MenuItem Exit
			[+] multitag "Exit"
				[ ] "#13"
				[ ] "$47"
	[+] Menu Actions
		[+] multitag "Actions"
			[ ] "#2"
		[+] MenuItem Add
			[+] multitag "Add"
				[ ] "#1"
				[ ] "$4756"
		[+] MenuItem Delete
			[+] multitag "Delete"
				[ ] "#2"
				[ ] "$4757"
		[+] MenuItem Extract
			[+] multitag "Extract"
				[ ] "#3"
				[ ] "$4758"
		[+] MenuItem View
			[+] multitag "View"
				[ ] "#4"
				[ ] "$34"
		[+] MenuItem NewFolder
			[+] multitag "New Folder"
				[ ] "#5"
				[ ] "$196"
		[+] MenuItem SelectAll
			[+] multitag "Select All"
				[ ] "#6"
				[ ] "$31"
		[+] MenuItem InvertSelection
			[+] multitag "Invert Selection"
				[ ] "#7"
				[ ] "$32"
		[+] MenuItem SelectFiles
			[+] multitag "Select Files"
				[ ] "#8"
				[ ] "$40"
		[+] MenuItem MakeExeFile
			[+] multitag "Make .Exe File"
				[ ] "#9"
				[ ] "$65"
		[+] MenuItem UUencode
			[+] multitag "UUencode"
				[ ] "#10"
				[ ] "$113"
		[+] MenuItem Split
			[+] multitag "Split"
				[ ] "#11"
				[ ] "$192"
		[+] MenuItem Encrypt
			[+] multitag "Encrypt"
				[ ] "#12"
				[ ] "$198"
		[+] MenuItem Test
			[+] multitag "Test"
				[ ] "#13"
				[ ] "$86"
		[+] MenuItem CheckOut
			[+] multitag "CheckOut"
				[ ] "#14"
				[ ] "$63"
		[+] MenuItem Install
			[+] multitag "Install"
				[ ] "#15"
				[ ] "$35"
	[+] Menu View
		[+] multitag "View"
			[ ] "#3"
		[+] MenuItem AllFilesWinZipStyle
			[+] multitag "All Files (WinZip Style)"
				[ ] "#1"
				[ ] "$194"
		[+] MenuItem FilesByFolderExplorerStyl
			[+] multitag "Files by Folder (Explorer Style)"
				[ ] "#2"
				[ ] "$195"
		[+] MenuItem Folders
			[+] multitag "Folders"
				[ ] "#3"
				[ ] "$199"
		[+] MenuItem Thumbnails
			[+] multitag "Thumbnails"
				[ ] "#4"
				[ ] "$222"
		[+] MenuItem Tiles
			[+] multitag "Tiles"
				[ ] "#5"
				[ ] "$227"
		[+] MenuItem LargeIcons
			[+] multitag "Large Icons"
				[ ] "#6"
				[ ] "$223"
		[+] MenuItem SmallIcons
			[+] multitag "Small Icons"
				[ ] "#7"
				[ ] "$224"
		[+] MenuItem xList
			[+] multitag "List"
				[ ] "#8"
				[ ] "$225"
		[+] MenuItem Details
			[+] multitag "Details"
				[ ] "#9"
				[ ] "$226"
		[+] Menu Sort
			[+] multitag "Sort"
				[ ] "#10"
			[+] MenuItem ByName
				[+] multitag "by Name"
					[ ] "#1"
					[ ] "$125"
			[+] MenuItem ByType
				[+] multitag "by Type"
					[ ] "#2"
					[ ] "$126"
			[+] MenuItem ByDate
				[+] multitag "by Date"
					[ ] "#3"
					[ ] "$127"
			[+] MenuItem BySize
				[+] multitag "by Size"
					[ ] "#4"
					[ ] "$128"
			[+] MenuItem ByCompressionRatio
				[+] multitag "by Compression Ratio"
					[ ] "#5"
					[ ] "$129"
			[+] MenuItem ByPackedSize
				[+] multitag "by Packed Size"
					[ ] "#6"
					[ ] "$130"
			[+] MenuItem ByCRC
				[+] multitag "by CRC"
					[ ] "#7"
					[ ] "$131"
			[+] MenuItem ByFileAttributes
				[+] multitag "by File Attributes"
					[ ] "#8"
					[ ] "$132"
			[+] MenuItem ByPath
				[+] multitag "by Path"
					[ ] "#9"
					[ ] "$133"
			[+] MenuItem ByOriginalOrder
				[+] multitag "by Original Order"
					[ ] "#10"
					[ ] "$134"
		[+] MenuItem UpOneLevel
			[+] multitag "Up one level"
				[ ] "#11"
				[ ] "$197"
		[+] MenuItem Comment
			[+] multitag "Comment"
				[ ] "#12"
				[ ] "$70"
		[+] MenuItem LastOutput
			[+] multitag "Last Output"
				[ ] "#13"
				[ ] "$29"
	[+] Menu Jobs
		[+] multitag "Jobs"
			[ ] "#4"
		[+] MenuItem Create
			[+] multitag "Create"
				[ ] "#1"
				[ ] "$260"
		[+] MenuItem Edit
			[+] multitag "Edit"
				[ ] "#2"
				[ ] "$261"
		[+] MenuItem Run
			[+] multitag "Run"
				[ ] "#3"
				[ ] "$262"
		[+] Menu RunRecentHistory
			[+] multitag "Run Recent (History)"
				[ ] "#4"
			[+] MenuItem Empty
				[+] multitag "(Empty)"
					[ ] "#1"
					[ ] "$2260"
	[+] Menu Options
		[+] multitag "Options"
			[ ] "#5"
		[+] MenuItem Configuration
			[+] multitag "Configuration"
				[ ] "#1"
				[ ] "$45"
		[+] MenuItem ReuseWinZipWindows
			[+] multitag "Reuse WinZip Windows"
				[ ] "#2"
				[ ] "$76"
		[+] MenuItem SaveSettingsOnExit
			[+] multitag "Save Settings on Exit"
				[ ] "#3"
				[ ] "$84"
		[+] MenuItem SaveSettingsNow
			[+] multitag "Save Settings Now"
				[ ] "#4"
				[ ] "$4766"
		[+] MenuItem SetInstallationDefaults
			[+] multitag "Set Installation Defaults"
				[ ] "#5"
				[ ] "$4767"
	[+] Menu Help
		[+] multitag "Help"
			[ ] "#6"
		[+] MenuItem HelpContentsAndIndex
			[+] multitag "Help Contents and Index"
				[ ] "#1"
				[ ] "$4771"
		[+] MenuItem BriefTutorial
			[+] multitag "Brief Tutorial"
				[ ] "#2"
				[ ] "$108"
		[+] MenuItem FrequentlyAskedQuestions
			[+] multitag "Frequently Asked Questions"
				[ ] "#3"
				[ ] "$116"
		[+] MenuItem HintsAndTips
			[+] multitag "Hints and Tips"
				[ ] "#4"
				[ ] "$77"
		[+] MenuItem TipOfTheDay
			[+] multitag "Tip of the Day"
				[ ] "#5"
				[ ] "$78"
		[+] MenuItem EnterRegistrationCode
			[+] multitag "Enter Registration Code"
				[ ] "#6"
				[ ] "$39"
		[+] MenuItem LicenseAgreement
			[+] multitag "License Agreement"
				[ ] "#7"
				[ ] "$52"
		[+] MenuItem OrderingInformation
			[+] multitag "Ordering Information"
				[ ] "#8"
				[ ] "$109"
		[+] MenuItem CheckForUpdate
			[+] multitag "Check for Update"
				[ ] "#9"
				[ ] "$79"
		[+] MenuItem WinZipHomePage
			[+] multitag "WinZip Home Page"
				[ ] "#10"
				[ ] "$117"
		[+] MenuItem AboutWinZip
			[+] multitag "About WinZip"
				[ ] "#11"
				[ ] "$4754"
	[+] MenuItem BuyNow
		[+] multitag "Buy Now!"
			[ ] "#7"
			[ ] "$4773"
	[+] StatusBar StatusBar1
		[+] multitag "#1"
			[ ] "$9995"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
	[+] CustomWin ReBarWindow321
		[+] multitag "[ReBarWindow32]#1"
			[ ] "[ReBarWindow32]$9996"
		[+] ToolBar ToolBar1
			[+] multitag "#1"
				[ ] "$9997"
			[+] PushButton New
				[+] multitag "New"
					[ ] "#1"
					[ ] "$4752"
			[+] PushButton Open
				[+] multitag "Open"
					[ ] "#2"
					[ ] "$4753"
			[+] PushButton Favorites
				[+] multitag "Favorites"
					[ ] "#3"
					[ ] "$114"
			[+] PushButton Add
				[+] multitag "Add"
					[ ] "#4"
					[ ] "$4756"
			[+] PushButton Extract
				[+] multitag "Extract"
					[ ] "#5"
					[ ] "$4758"
			[+] PushButton Encrypt
				[+] multitag "Encrypt"
					[ ] "#6"
					[ ] "$198"
			[+] PushButton View
				[+] multitag "View"
					[ ] "#7"
					[ ] "$34"
			[+] PushButton CheckOut
				[+] multitag "CheckOut"
					[ ] "#8"
					[ ] "$150"
			[+] PushButton Wizard
				[+] multitag "Wizard"
					[ ] "#9"
					[ ] "$111"
			[+] PushButton ViewStyle
				[+] multitag "View Style"
					[ ] "#10"
					[ ] "$209"
	[+] ListView ListView1
		[+] multitag "#1"
			[ ] "$9999"
		[+] Header Header1
			[+] multitag "#1"
				[ ] "$0"
			[+] PushButton Name
				[+] multitag "Name"
					[ ] "#1"
			[+] PushButton Type
				[+] multitag "Type"
					[ ] "#2"
			[+] PushButton Modified
				[+] multitag "Modified"
					[ ] "#3"
			[+] PushButton Size
				[+] multitag "Size"
					[ ] "#4"
			[+] PushButton Ratio
				[+] multitag "Ratio"
					[ ] "#5"
			[+] PushButton Packed
				[+] multitag "Packed"
					[ ] "#6"
			[+] PushButton CRC
				[+] multitag "CRC"
					[ ] "#9"
			[+] PushButton Attributes
				[+] multitag "Attributes"
					[ ] "#8"
			[+] PushButton Path
				[+] multitag "Path"
					[ ] "#7"
[+] window DialogBox NewArchive
	[ ] tag "New Archive"
	[ ] parent WinZip
	[+] StaticText SaveInText
		[+] multitag "Save in:"
			[ ] "$1091"
	[+] PopupList SaveIn
		[+] multitag "Save in:"
			[ ] "$1137"
	[+] ToolBar SaveIn1
		[+] multitag "Save in:[1]"
			[ ] "$1088"
		[+] PushButton PushButton1
			[ ] tag "$40971"
		[+] PushButton PushButton2
			[ ] tag "$40961"
		[+] PushButton PushButton3
			[ ] tag "$40962"
		[+] PushButton PushButton4
			[ ] tag "$40970"
	[+] ToolBar SaveIn2
		[+] multitag "Save in:[2]"
			[ ] "$1184"
		[+] PushButton MyRecentDocuments
			[+] multitag "My Recent Documents"
				[ ] "$41060"
		[+] PushButton Desktop
			[+] multitag "Desktop"
				[ ] "$41061"
		[+] PushButton MyDocuments
			[+] multitag "My Documents"
				[ ] "$41062"
		[+] PushButton MyComputer
			[+] multitag "My Computer"
				[ ] "$41063"
		[+] PushButton MyNetworkPlaces
			[+] multitag "My Network Places"
				[ ] "$41064"
	[+] ListView ListView1
		[ ] tag "$1"
	[+] StaticText FileNameText
		[+] multitag "File name:"
			[ ] "$1090"
	[+] ComboBox FileName
		[+] multitag "File name:"
			[ ] "$1148"
	[+] StaticText FilesOfTypeText
		[+] multitag "Files of type:"
			[ ] "$1089"
	[+] PopupList FilesOfType
		[+] multitag "Files of type:"
			[ ] "$1136"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "$1038"
	[+] ScrollBar ScrollBar1
		[ ] tag "$-1"
	[+] DialogBox DialogBox1
		[ ] tag "$0"
[+] window DialogBox Add
	[ ] tag "Add"
	[ ] parent WinZip
	[+] ComboBox FileName
		[+] multitag "File name:"
			[ ] "$1148"
	[+] PushButton Add
		[+] multitag "Add"
			[ ] "$1"
[+] window DialogBox WinZipMessageBox
	[ ] tag "WinZip"
	[ ] parent WinZip
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "#1"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "#2"
			[ ] "$7"
	[+] PushButton Help
		[+] multitag "Help"
			[ ] "#3"
			[ ] "$9"
	[+] CustomWin Icon1
		[+] multitag "[Icon]#1"
			[ ] "[Icon]$20"
	[+] StaticText CannotCreateCDocumentsAnd
		[+] multitag "Cannot create C:\Documents and Settings\John Connolly\My Documents\New2.zip.  File already exists.  Open existing file?"
			[ ] "#1"
			[ ] "$65535"
[+] window DialogBox AboutWinZip
	[ ] tag "About WinZip"
	[ ] parent WinZip
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#2"
			[ ] "$1"
	[+] PushButton OrderingInfo
		[+] multitag "Ordering Info"
			[ ] "#3"
			[ ] "$703"
	[+] PushButton License
		[+] multitag "License"
			[ ] "#4"
			[ ] "$702"
	[+] PushButton Acknowledgments
		[+] multitag "Acknowledgments"
			[ ] "#5"
			[ ] "$715"
	[+] PushButton SystemInfo
		[+] multitag "System Info"
			[ ] "#6"
			[ ] "$718"
	[+] PushButton Register
		[+] multitag "Register"
			[ ] "#7"
			[ ] "$719"
	[+] PushButton CheckForUpdate
		[+] multitag "Check for Update"
			[ ] "#8"
			[ ] "$720"
	[+] PushButton WebSite
		[+] multitag "Web Site"
			[ ] "#9"
			[ ] "$717"
	[+] PushButton FAQs
		[+] multitag "FAQs"
			[ ] "#10"
			[ ] "$716"
	[+] PushButton PushButton1
		[+] multitag "#1"
			[ ] "$713"
	[+] StaticText WinZipVersion
		[+] multitag "WinZip*"
			[ ] "#1"
			[ ] "$709"
	[+] StaticText Copyright19912007WinZipI
		[+] multitag "Copyright © 1991-2007 WinZip International LLC"
			[ ] "#2"
			[ ] "$710"
	[+] StaticText AllRightsReserved1Text
		[+] multitag "All Rights Reserved[1]"
			[ ] "#3"
			[ ] "$712"
	[+] StaticText AllRightsReserved2Text
		[+] multitag "All Rights Reserved[2]"
			[ ] "#4"
			[ ] "$-1[1]"
	[+] StaticText ThisIsAnUnlicensedCopyFor
		[+] multitag "This is an unlicensed copy for evaluation use only."
			[ ] "#5"
			[ ] "$705"
	[+] StaticText PleaseRememberThatIfYouUs
		[+] multitag "Please remember that if you use this software you are"
			[ ] "#6"
			[ ] "$706"
	[+] StaticText RequiredToPayTheRegistrati
		[+] multitag "required to pay the registration fee, as described in"
			[ ] "#7"
			[ ] "$707"
	[+] StaticText TheLicenseAgreementText
		[+] multitag "the License Agreement"
			[ ] "#8"
			[ ] "$708"
	[+] StaticText ContactPointsAreText
		[+] multitag "Contact points are:"
			[ ] "#9"
			[ ] "$-1[2]"
	[+] StaticText InternetText
		[+] multitag "Internet:"
			[ ] "#10"
			[ ] "$-1[3]"
	[+] StaticText WwwWinzipComText
		[+] multitag "www.winzip.com"
			[ ] "#11"
			[ ] "$721"
	[+] StaticText EMailText
		[+] multitag "E-mail:"
			[ ] "#12"
			[ ] "$-1[4]"
	[+] StaticText ContactWinZipText
		[+] multitag "Contact WinZip"
			[ ] "#13"
			[ ] "$725"
	[+] StaticText PostalMailPOBox540Man
		[+] multitag "Postal Mail: P.O. Box 540, Mansfield, CT 06268 USA"
			[ ] "#14"
			[ ] "$-1[5]"
[+] window DialogBox SymantecAntiVirusLogin
	[ ] tag "Symantec AntiVirus"
	[+] StaticText EnterASymantecAntiVirus
		[+] multitag "Enter a Symantec AntiVirus Server Group user name and password to run Symantec AntiVirus on this computer."
			[ ] "#1"
			[ ] "$-1[1]"
	[+] StaticText ServerGroupText
		[+] multitag "Server Group:"
			[ ] "#2"
			[ ] "$-1[2]"
	[+] TextField ServerGroup
		[+] multitag "Server Group:"
			[ ] "#1"
			[ ] "$1034"
	[+] StaticText UsernameText
		[+] multitag "Username:"
			[ ] "#3"
			[ ] "$-1[3]"
	[+] TextField Username
		[+] multitag "Username:"
			[ ] "#2"
			[ ] "$2032"
	[+] StaticText PasswordText
		[+] multitag "Password:"
			[ ] "#4"
			[ ] "$-1[4]"
	[+] TextField Password
		[+] multitag "Password:"
			[ ] "#3"
			[ ] "$1035"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#2"
			[ ] "$2"
[ ] 
[-] window MainWin SymantecAntiVirus
	[+] multitag "Symantec AntiVirus"
		[ ] // "$C:\Program Files\Symantec AntiVirus\VPC32.exe"
		[ ] "$C:\Program Files\SAV\VPC32.exe"
	[+] Menu File
		[ ] tag "File"
		[+] MenuItem ConfigureHistories
			[+] multitag "Configure Histories"
				[ ] "$32798"
		[+] MenuItem ScheduleUpdates
			[+] multitag "Schedule Updates"
				[ ] "$32799"
		[+] MenuItem LiveUpdate
			[+] multitag "LiveUpdate"
				[ ] "$32796"
		[+] MenuItem ThreatList
			[+] multitag "Threat List"
				[ ] "$32800"
		[+] MenuItem UnloadService
			[+] multitag "Unload Service"
				[ ] "$1021"
		[+] MenuItem Exit
			[+] multitag "Exit"
				[ ] "$57665"
	[+] Menu Edit
		[ ] tag "Edit"
		[+] MenuItem Cut
			[+] multitag "Cut"
				[ ] "$32794"
		[+] MenuItem Paste
			[+] multitag "Paste"
				[ ] "$32795"
		[+] MenuItem Delete
			[+] multitag "Delete"
				[ ] "$32787"
		[+] MenuItem NewStartupScan
			[+] multitag "New Startup Scan"
				[ ] "$32792"
		[+] MenuItem NewCustomScan
			[+] multitag "New Custom Scan"
				[ ] "$32779"
		[+] MenuItem NewScheduledScan
			[+] multitag "New Scheduled Scan"
				[ ] "$32788"
	[+] Menu View
		[ ] tag "View"
		[+] MenuItem AutoProtectScanStatistics
			[+] multitag "Auto-Protect Scan Statistics"
				[ ] "$3"
		[+] MenuItem ScheduledScans
			[+] multitag "Scheduled Scans"
				[ ] "$4"
		[+] MenuItem Quarantine
			[+] multitag "Quarantine"
				[ ] "$5"
		[+] MenuItem BackupItems
			[+] multitag "Backup Items"
				[ ] "$6"
		[+] MenuItem RepairedItems
			[+] multitag "Repaired Items"
				[ ] "$7"
	[+] Menu Scan
		[ ] tag "Scan"
		[+] MenuItem CustomScan
			[+] multitag "Custom Scan"
				[ ] "$0"
		[+] MenuItem QuickScan
			[+] multitag "Quick Scan"
				[ ] "$1"
		[+] MenuItem FullScan
			[+] multitag "Full Scan"
				[ ] "$2"
	[+] Menu Histories
		[ ] tag "Histories"
		[+] MenuItem ThreatHistory
			[+] multitag "Threat History"
				[ ] "$8"
		[+] MenuItem ScanHistory
			[+] multitag "Scan History"
				[ ] "$9"
		[+] MenuItem EventLog
			[+] multitag "Event Log"
				[ ] "$10"
		[+] MenuItem TamperHistory
			[+] multitag "Tamper History"
				[ ] "$11"
	[+] Menu Help
		[ ] tag "Help"
		[+] MenuItem HelpTopics
			[+] multitag "Help Topics"
				[ ] "$32789"
		[+] MenuItem AboutSymantecAntiVirus
			[+] multitag "About Symantec AntiVirus"
				[ ] "$57664"
	[-] CustomWin AfxMDIFrame701
		[ ] tag "[AfxMDIFrame70]$59648"
		[+] CustomWin AfxFrameOrView701
			[ ] tag "[AfxFrameOrView70]$59648"
			[+] TreeView TreeView1
				[ ] tag "$1000"
		[+] CustomWin AfxFrameOrView702
			[ ] tag "[AfxFrameOrView70]$59649"
			[+] CustomWin AfxOleControl701
				[ ] tag "[AfxOleControl70]$0"
				[+] DialogBox FullScan
					[+] multitag "Full Scan"
						[ ] "$0"
					[+] PushButton Help
						[+] multitag "Help"
							[ ] "$9"
					[+] DialogBox FullScan
						[+] multitag "Full Scan"
							[ ] "$0"
						[+] PushButton Options
							[+] multitag "Options"
								[ ] "$292"
						[+] StaticText FullScanText
							[+] multitag "Full Scan"
								[ ] "$313"
						[+] StaticText ScansThisEntireComputerFor
							[+] multitag "Scans this entire computer for viruses and security risks such as adware and spyware.??Use this scan to look in the boot secto*"
								[ ] "$-1"
					[+] StaticText StaticText1
						[ ] tag "$12326"
					[+] PushButton Scan
						[+] multitag "Scan"
							[ ] "$12325"
[-] window DialogBox SAVFullScan
	[ ] tag "Full Scan*"
	[+] ToolBar ToolBar1
		[ ] tag "$59392"
	[+] PushButton RemoveRisk
		[+] multitag "Remove Risk"
			[ ] "$1029"
	[+] PushButton Reboot
		[+] multitag "Reboot"
			[ ] "$1097"
	[+] PushButton Close
		[+] multitag "Close"
			[ ] "$1003"
	[+] CustomWin CWINDOWSFonts1
		[+] multitag "[AfxOleControl70]C:\WINDOWS\Fonts"
			[ ] "[AfxOleControl70]$0"
		[+] ListView ListView1
			[ ] tag "$5000"
			[+] Header Header1
				[ ] tag "$0"
				[+] PushButton PushButton11
					[ ] tag "#11"
				[+] PushButton PushButton12
					[ ] tag "#12"
				[+] PushButton PushButton9
					[ ] tag "#9"
				[+] PushButton PushButton10
					[ ] tag "#10"
				[+] PushButton PushButton15
					[ ] tag "#15"
				[+] PushButton PushButton16
					[ ] tag "#16"
				[+] PushButton PushButton13
					[ ] tag "#13"
				[+] PushButton PushButton14
					[ ] tag "#14"
				[+] PushButton PushButton3
					[ ] tag "#3"
				[+] PushButton PushButton4
					[ ] tag "#4"
				[+] PushButton PushButton1
					[ ] tag "#1"
				[+] PushButton PushButton2
					[ ] tag "#2"
				[+] PushButton PushButton7
					[ ] tag "#7"
				[+] PushButton PushButton8
					[ ] tag "#8"
				[+] PushButton PushButton5
					[ ] tag "#5"
				[+] PushButton PushButton6
					[ ] tag "#6"
	[+] StaticText Text1
		[-] multitag "arial.ttf"
			[ ] "$1021"
	[+] StaticText Text2
		[+] multitag "C:\WINDOWS\Fonts"
			[ ] "$1011"
	[+] CustomWin Animate1
		[+] multitag "[SysAnimate32]Animate1"
			[ ] "[SysAnimate32]$1023"
	[+] StatusBar CWINDOWSFonts2
		[+] multitag "C:\WINDOWS\Fonts"
			[ ] "$59393"
[+] window DialogBox FullScanStartedOn4232008
	[ ] tag "Full Scan started on 4?23?2008 3:01:44 PM"
	[+] ToolBar ToolBar1
		[+] multitag "#1"
			[ ] "$59392"
	[+] PushButton RemoveRisk
		[+] multitag "Remove Risk"
			[ ] "#1"
			[ ] "$1029"
	[+] PushButton Reboot
		[+] multitag "Reboot"
			[ ] "#2"
			[ ] "$1097"
	[+] PushButton Close
		[+] multitag "Close"
			[ ] "#3"
			[ ] "$1003"
	[+] CustomWin AfxOleControl701
		[+] multitag "[AfxOleControl70]#1"
			[ ] "[AfxOleControl70]$0"
		[+] ListView ListView1
			[+] multitag "#1"
				[ ] "$5000"
			[+] Header Header1
				[+] multitag "#1"
					[ ] "$0"
				[+] PushButton PushButton11
					[ ] tag "#11"
				[+] PushButton PushButton12
					[ ] tag "#12"
				[+] PushButton PushButton9
					[ ] tag "#9"
				[+] PushButton PushButton10
					[ ] tag "#10"
				[+] PushButton PushButton15
					[ ] tag "#15"
				[+] PushButton PushButton16
					[ ] tag "#16"
				[+] PushButton PushButton13
					[ ] tag "#13"
				[+] PushButton PushButton14
					[ ] tag "#14"
				[+] PushButton PushButton3
					[ ] tag "#3"
				[+] PushButton PushButton4
					[ ] tag "#4"
				[+] PushButton PushButton1
					[ ] tag "#1"
				[+] PushButton PushButton2
					[ ] tag "#2"
				[+] PushButton PushButton7
					[ ] tag "#7"
				[+] PushButton PushButton8
					[ ] tag "#8"
				[+] PushButton PushButton5
					[ ] tag "#5"
				[+] PushButton PushButton6
					[ ] tag "#6"
	[+] StaticText Completed1Text
		[+] multitag "Completed[1]"
			[ ] "#1"
			[ ] "$1021"
	[+] StaticText Completed2Text
		[+] multitag "Completed[2]"
			[ ] "#2"
			[ ] "$1011"
	[+] CustomWin Animate1
		[+] multitag "[SysAnimate32]Animate1"
			[ ] "[SysAnimate32]#1"
			[ ] "[SysAnimate32]$1023"
	[+] StatusBar StatusBar1
		[-] multitag "#1"
			[ ] "$59393"
[ ] 
[+] window MainWin SAV4Exchange
	[ ] tag "[DialogBox]Symantec MS 4.6 Console for Microsoft Exchange"
	[+] CustomWin SizeableRebar1
		[ ] tag "[SizeableRebar]$13341"
		[+] CustomWin ReBarWindow321
			[ ] tag "[ReBarWindow32]$4096"
			[+] ToolBar ToolBar2
				[ ] tag "$4098"
				[+] PushButton PushButton1
					[ ] tag "$1"
				[+] PushButton PushButton2
					[ ] tag "$2"
				[+] PushButton PushButton3
					[ ] tag "$5"
				[+] PushButton PushButton4
					[ ] tag "$17"
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
[ ] 
[+] window MainWin ComponentCheckerVersion20
	[+] multitag "Component Checker Version 2.0"
		[ ] "$C:\SilkTestData\cc.exe"
	[+] Menu File
		[+] multitag "File"
			[ ] "#1"
		[+] MenuItem Save
			[+] multitag "Save"
				[ ] "#1"
				[ ] "$1001"
		[+] MenuItem Exit
			[+] multitag "Exit"
				[ ] "#2"
				[ ] "$1006"
	[+] Menu View
		[+] multitag "View"
			[ ] "#2"
		[+] MenuItem SummaryReport
			[+] multitag "Summary Report"
				[ ] "#1"
				[ ] "$1012"
		[+] MenuItem FileDetails
			[+] multitag "File Details"
				[ ] "#2"
				[ ] "$1013"
		[+] MenuItem COMDetails
			[+] multitag "COM Details"
				[ ] "#3"
				[ ] "$1014"
		[+] MenuItem RegistryDetails
			[+] multitag "Registry Details"
				[ ] "#4"
				[ ] "$1015"
		[+] MenuItem Refresh
			[+] multitag "Refresh"
				[ ] "#5"
				[ ] "$1016"
	[+] Menu Help
		[+] multitag "Help"
			[ ] "#3"
		[+] MenuItem HelpTopics
			[+] multitag "Help topics"
				[ ] "#1"
				[ ] "$1030"
		[+] MenuItem About
			[+] multitag "About"
				[ ] "#2"
				[ ] "$1031"
		[+] MenuItem LegendOfSymbols
			[+] multitag "Legend of symbols"
				[ ] "#3"
				[ ] "$1032"
	[+] ToolBar ToolBar1
		[+] multitag "#1"
			[ ] "$59392"
		[+] PushButton Save
			[+] multitag "Save"
				[ ] "#1"
				[ ] "$1001"
		[+] PushButton SummaryReport
			[+] multitag "Summary Report"
				[ ] "#2"
				[ ] "$1012"
		[+] PushButton FileDetails
			[+] multitag "File Details"
				[ ] "#3"
				[ ] "$1013"
		[+] PushButton COMDetails
			[+] multitag "COM Details"
				[ ] "#4"
				[ ] "$1014"
		[+] PushButton RegistryDetails
			[+] multitag "Registry Details"
				[ ] "#5"
				[ ] "$1015"
		[+] PushButton Refresh
			[+] multitag "Refresh"
				[ ] "#6"
				[ ] "$1016"
	[+] StatusBar StatusBar1
		[+] multitag "#1"
			[ ] "$59393"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
	[+] CustomWin WTL_SplitterWindow1
		[+] multitag "[WTL_SplitterWindow]#1"
			[ ] "[WTL_SplitterWindow]$457768"
		[+] CustomWin WTL_PaneContainer1
			[+] multitag "[WTL_PaneContainer]#1"
				[ ] "[WTL_PaneContainer]$457864"
			[+] CustomWin ATL01039E601
				[+] multitag "[ATL:01039E60]#1"
					[ ] "[ATL:01039E60]$458172"
		[+] CustomWin WTL_PaneContainer2
			[+] multitag "[WTL_PaneContainer]#2"
				[ ] "[WTL_PaneContainer]$457996"
			[+] CustomWin ATL01039E101
				[+] multitag "[ATL:01039E10]#1"
					[ ] "[ATL:01039E10]$458128"
				[+] Header Header1
					[+] multitag "#1"
						[ ] "$0"
					[+] StaticText ItemNameText
						[+] multitag "Item Name"
							[ ] "#1"
					[+] StaticText ItemTypeText
						[+] multitag "Item Type"
							[ ] "#2"
[+] window DialogBox ComponentChecker
	[ ] tag "Component Checker - Choose Analysis Type"
	[ ] parent ComponentCheckerVersion20
	[+] StaticText StaticText1
		[+] multitag "#1"
			[ ] "$-1[1]"
	[+] RadioList RadioList1
		[+] multitag "#1"
			[ ] "$1000"
	[+] StaticText StaticText2
		[+] multitag "#2"
			[ ] "$-1[2]"
	[+] ComboBox ComboBox1
		[+] multitag "#1"
			[ ] "$1004"
	[+] StaticText StaticText3
		[+] multitag "#3"
			[ ] "$-1[3]"
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "#1"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "#2"
			[ ] "$2"
[+] window DialogBox ComponentCheckerComplete
	[ ] tag "ComponentChecker 2.0"
	[ ] parent ComponentCheckerVersion20
	[+] PushButton Yes
		[+] multitag "Yes"
			[ ] "#1"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "#2"
			[ ] "$7"
	[+] StaticText Message
		[+] multitag "The following product releases were matched:?		?MDAC 2.8 SP2 ON WINDOWS SERVER 2003 SP1??Would you like to re-run the query fo*"
			[ ] "#1"
			[ ] "$65535"
[ ] /////////////////////////////////////////////////////////////////////////////
[ ] // GLOBAL FUNCTIONS
[ ] /////////////////////////////////////////////////////////////////////////////
[+] setTarget(STRING sIPAddress)
	[ ] Agent.SetOption(OPT_SET_TARGET_MACHINE, sIPAddress)
	[ ] sCurrentMachine = sIPAddress
[+] RunStartMenu(STRING sApplication)
	[ ] Taskbar.SetActive ()
	[ ] Taskbar.Start.Click ()
	[ ] // Taskbar.TypeKeys("<Ctrl><Esc>")
	[ ] Taskbar.TypeKeys("R")
	[ ] sleep(1)
	[ ] Run.SetActive ()
	[ ] Run.Open.SetText(sApplication)
	[ ] Run.OK.Click()
[+] STRING getHostNamePrefix()
	[ ] STRING sCMD = "hostname"
	[ ] LIST OF STRING lsReturn
	[ ] SYS_EXECUTE(sCMD,lsReturn)
	[ ] return Left(lsReturn[1], len(lsReturn[1]) - 4)
	[ ] 
[+] STRING getHostNameSuffix()
	[ ] LIST OF STRING lslPConfig = getIPConfigAll()
	[ ] STRING sItem 
	[ ] STRING sReturn = ""
	[ ] INTEGER i, j = 1
	[ ] INTEGER iStartOffset = 39
	[ ] INTEGER iLineNumber = 5
	[ ] // Print("Machine Configuration Information is :")
	[ ] i = 1
	[-] for each sItem in lslPConfig 
		[-] if i == iLineNumber
			[-] for j = 1 to (len(sItem) - iStartOffset) - 1
				[ ] sReturn[  j  ] = sItem[ j + iStartOffset ]
			[ ] return(sReturn) 
			[ ] // return(Right(sItem[i],10))
		[ ] i++
[+] STRING getHostName()
	[ ] STRING sCMD = "hostname"
	[ ] LIST OF STRING lsReturn
	[ ] SYS_EXECUTE(sCMD,lsReturn)
	[ ] return lsReturn[1]
	[ ] 
[-] STRING getServerName()
	[ ] STRING sCMD = "hostname"
	[ ] LIST OF STRING lsReturn
	[ ] SYS_EXECUTE(sCMD,lsReturn)
	[ ] return Right(lsReturn[1],4)
	[ ] 
[ ] 
[+] LIST OF STRING getIPConfigAll()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getipconfig
	[ ] // Purpose:                                  This function returns 
	[ ] //								the host name and all other
	[ ] //								information returned by
	[ ] //                                                   ipconfig -all at the command line
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "ipconfig -all"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[ ] return(sReturn)
	[ ] 
[+] LIST OF STRING getIPConfig()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getipconfig
	[ ] // Purpose:                                  This function returns 
	[ ] //								the host name and all other
	[ ] //								information returned by
	[ ] //                                                   ipconfig -all at the command line
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "ipconfig -all"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
	[ ] 
[+] LIST OF STRING getServices()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getservices
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of all installed services
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getServices.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
	[ ] 
[+] LIST OF SERVICEREC getServicesList()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getservices
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of all installed services
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] LIST OF SERVICEREC LSERVICE
	[ ] SERVICEREC SERVICE
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"getServicesCSV.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] SERVICE.sDisplayName = GetField(sItem,",",1)
		[ ] SERVICE.sState = GetField(sItem,",",2)
		[ ] SERVICE.sStartMode = GetField(sItem,",",3)
		[ ] ListAppend(LSERVICE,SERVICE)
	[ ] return(LSERVICE)
	[ ] 
[+] LIST OF STRING getServicesExcel()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getservices
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of all installed services
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo M:\vb\getServicesExcel.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
	[ ] 
[+] LIST OF STRING getOSInfo()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getservices
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of OS Info
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getOSInfo.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
[+] LIST OF STRING getLocalDiskInfo()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getLocalDiskSpace
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of OS Info
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getLocalDiskInfo.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
[+] LIST OF STRING getSystemMemory()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getSystemMemory
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of OS Info
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getSystemMemory.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
[+] LIST OF STRING getVirtualMemory()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getVirtualMemory
	[ ] // Purpose:                                  This function returns 
	[ ] //								a list of OS Info
	[ ] //								
	[ ] //
	[ ] // Inputs:                                   	none
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getVirtualMemory.vbs "
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] // for each sItem in sReturn
		[ ] return(sReturn)
[+] STRING getRegistryEntry(STRING sRegPath)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getRegistryEntry
	[ ] // Purpose:                                  This function returns 
	[ ] //								the value of the registry 
	[ ] //								at the path specified
	[ ] //
	[ ] // Inputs:                                   	STRING (RegPath)
	[ ] // Return:                                     STRING single entry
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getRegistryEntry.vbs "+"""{sRegPath}"""
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[+] LIST OF STRING getRegistryEntries(STRING sRegPath)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getRegistryEntry
	[ ] // Purpose:                                  This function returns 
	[ ] //								the value of the registry 
	[ ] //								at the path specified
	[ ] //
	[ ] // Inputs:                                   	STRING (RegPath)
	[ ] // Return:                                     STRING single entry
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getRegistryEntry.vbs "+"""{sRegPath}"""
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[+] if StrPos("Invalid root in registry key",sItem)
			[ ] return(NULL)
		[+] else if StrPos("Unable to open registry key",sItem)
			[ ] return(NULL)
	[ ] return(sReturn)
	[ ] 
[+] STRING getJavaVersion(STRING sPath optional)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getJavaVerison
	[ ] // Purpose:                                  This function returns 
	[ ] //								the version of java running
	[ ] //
	[ ] // Inputs:                                   	STRING (expected Subject Line)
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine 
	[ ] INTEGER i
	[ ] // Change Dir to root
	[ ] cmdLine = "cd \"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[ ] cmdLine = "cd {sPath}"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[ ] 
	[ ]  cmdLine = "java -version"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)	
	[ ] i = 1
	[+] // for each sItem in sReturn
			[ ] // sItem = GetField (sItem, """", 2)
			[ ] // Print(sItem)
			[ ] // return(sItem)
		[ ] // i ++
	[+] sItem = sReturn[1]
			[ ] sItem = GetField (sItem, """", 2)
			[ ] // Print(sItem)
			[ ] return(sItem)
		[ ] i ++
[+] STRING getOSVersion()
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] // Author:                                   	John Connolly  
	[ ] // Function Name:                  	getOSVersion
	[ ] // Purpose:                                  This function returns 
	[ ] //                                                    the OS Version
	[ ] /////////////////////////////////////////////////////////////////////////////
	[ ] LIST OF STRING sReturn
	[ ] STRING sItem
	[ ] STRING cmdLine = "cscript //nologo {sGlobalVBReadPath}"+"\getOSVersion.vbs"
	[ ] // Print("cmdLine = "+cmdLine)
	[ ] SYS_Execute(cmdLine,sReturn)
	[+] for each sItem in sReturn
		[ ] return(sItem)
		[ ] 
[ ] 
