[ ] ////////////////////////////////////////////////////// 
[ ] // WINDOWS DECLARATIONS
[ ] //////////////////////////////////////////////////////
[+] window MainWin DominoWebLog
	[ ] tag "[DialogBox]http:??localhost?log.nsf - Microsoft Internet Explorer"
	[ ] // The working directory of the application when it is invoked
	[ ] const sDir = "H:\silktestmanager\ProfData\Desktop"
	[ ] // The command line used to invoke the application
	[ ] const sCmdLine = "C:\Program Files\Internet Explorer\iexplore.exe "
	[ ] // The first window to appear when the application is invoked
	[ ] // const wStartup = ?
	[ ] // The list of windows the recovery system is to leave open
	[ ] // const lwLeaveOpen = {?}
	[+] CustomWin WorkerW1
		[ ] tag "[WorkerW]$40965"
		[+] CustomWin ReBarWindow321
			[ ] tag "[ReBarWindow32]$40965"
			[+] ToolBar ToolBar3
				[ ] tag "$0[2]"
			[+] ComboBox ComboBox1
				[ ] tag "$41477"
			[+] ToolBar ToolBar2
				[ ] tag "$40960"
				[+] PushButton Back
					[+] multitag "Back"
						[ ] "$1014"
				[+] PushButton Forward
					[+] multitag "Forward"
						[ ] "$1015"
				[+] PushButton Stop
					[+] multitag "Stop"
						[ ] "$1016"
				[+] PushButton Refresh
					[+] multitag "Refresh"
						[ ] "$1017"
				[+] PushButton Home
					[+] multitag "Home"
						[ ] "$1018"
				[+] PushButton Search
					[+] multitag "Search"
						[ ] "$1019"
				[+] PushButton Favorites
					[+] multitag "Favorites"
						[ ] "$1020"
				[+] PushButton History
					[+] multitag "History"
						[ ] "$1021"
				[+] PushButton Mail
					[+] multitag "Mail"
						[ ] "$369"
				[+] PushButton Print
					[+] multitag "Print"
						[ ] "$260"
				[+] PushButton Edit
					[+] multitag "Edit"
						[ ] "$371"
				[+] PushButton Discuss
					[+] multitag "Discuss"
						[ ] "$372"
				[+] PushButton Research
					[+] multitag "Research"
						[ ] "$8193"
				[+] PushButton Messenger
					[+] multitag "Messenger"
						[ ] "$8194"
			[+] CustomWin WorkerW1
				[ ] tag "[WorkerW]$0"
			[+] ToolBar ToolBar1
				[ ] tag "$0[1]"
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
	[+] StatusBar StatusBar1
		[ ] tag "$40961"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
		[+] DynamicText DynamicText2
			[ ] tag "#2"
		[+] DynamicText DynamicText3
			[ ] tag "#3"
		[+] DynamicText DynamicText4
			[ ] tag "#4"
		[+] DynamicText DynamicText5
			[ ] tag "#5"
		[+] DynamicText DynamicText6
			[ ] tag "#6"
		[+] DynamicText DynamicText7
			[ ] tag "#7"
		[+] DynamicText DynamicText8
			[ ] tag "#8"
	[+] CustomWin ShellDocObjectView1
		[ ] tag "[Shell DocObject View]$0"
		[+] CustomWin InternetExplorer_Server1
			[ ] tag "[Internet Explorer_Server]$0"
			[+] CustomWin JavaPlugInControlWindow1
				[ ] tag "[Java Plug-in Control Window]$86825292"
				[+] CustomWin SunPluginViewerFrameIExplo
					[ ] tag "[sun.plugin.viewer.frame.IExplorerEmbeddedFrame]$0"
					[+] CustomWin SunAwtCanvas1
						[ ] tag "[SunAwtCanvas]$0"
						[+] CustomWin SunAwtCanvas1
							[ ] tag "[SunAwtCanvas]$0"
							[+] CustomWin SunAwtCanvas1
								[ ] tag "[SunAwtCanvas]$0"
								[+] CustomWin SunAwtCanvas1
									[ ] tag "[SunAwtCanvas]$0"
			[+] CustomWin JavaPlugInControlWindow2
				[ ] tag "[Java Plug-in Control Window]$86876060"
				[+] CustomWin SunPluginViewerFrameIExplo
					[ ] tag "[sun.plugin.viewer.frame.IExplorerEmbeddedFrame]$0"
					[+] CustomWin SunAwtCanvas1
						[ ] tag "[SunAwtCanvas]$0"
						[+] CustomWin SunAwtCanvas1
							[ ] tag "[SunAwtCanvas]$0"
							[+] CustomWin SunAwtCanvas1
								[ ] tag "[SunAwtCanvas]$0"
								[+] CustomWin SunAwtCanvas1
									[ ] tag "[SunAwtCanvas]$0"
	[ ] 
[+] // window MessageBoxClass MessageBox
	[ ] // tag "~ActiveApp/[DialogBox]$MessageBox"
	[+] // PushButton OK
		[ ] // tag "OK"
	[+] // PushButton Cancel
		[ ] // tag "Cancel"
	[+] // PushButton Yes
		[ ] // tag "Yes"
	[+] // PushButton No
		[ ] // tag "No"
	[+] // StaticText Message
		[ ] // motif tag "#2"
		[ ] // tag "#1"
[+] window MainWin DominoCommand
	[ ] tag "[DialogBox]*ITFShipSD*"
	[-] // textField Command
		[ ] // tag "#1"
[+] window MainWin WorkspaceLotusNotes
	[+] multitag "Workspace*"
		[ ] "*Lotus Notes"
		[ ] "$D:\Lotus\Domino\nlnotes.exe"
	[+] Menu File
		[ ] tag "File"
		[+] MenuItem CleanupEmptyPopups
			[+] multitag "?Cleanup Empty Popups?"
				[ ] "$2299"
		[+] Menu New
			[ ] tag "New"
			[+] MenuItem New
				[+] multitag "New"
					[ ] "$2302"
		[+] MenuItem Close
			[+] multitag "Close"
				[ ] "$61536"
		[+] MenuItem Save
			[+] multitag "Save"
				[ ] "$2309"
		[+] MenuItem SaveAsNewVersion
			[+] multitag "Save As New Version"
				[ ] "$2310"
		[+] MenuItem DocumentProperties
			[+] multitag "Document Properties"
				[ ] "$2311"
		[+] Menu Database
			[ ] tag "Database"
			[+] MenuItem CleanupSeparators
				[+] multitag "?Cleanup Separators?"
					[ ] "$2315"
			[+] MenuItem Properties
				[+] multitag "Properties"
					[ ] "$2316"
			[+] MenuItem AccessControl
				[+] multitag "Access Control"
					[ ] "$2317"
			[+] MenuItem Open
				[+] multitag "Open"
					[ ] "$2319"
			[+] MenuItem New
				[+] multitag "New"
					[ ] "$2321"
			[+] MenuItem NewCopy
				[+] multitag "New Copy"
					[ ] "$2323"
			[+] MenuItem Publish
				[+] multitag "Publish"
					[ ] "$2324"
			[+] MenuItem Delete
				[+] multitag "Delete"
					[ ] "$2325"
			[+] MenuItem Archive
				[+] multitag "Archive"
					[ ] "$2326"
			[+] MenuItem RefreshDesign
				[+] multitag "Refresh Design"
					[ ] "$2328"
			[+] MenuItem ReplaceDesign
				[+] multitag "Replace Design"
					[ ] "$2329"
			[+] MenuItem DesignSynopsis
				[+] multitag "Design Synopsis"
					[ ] "$2330"
		[+] Menu Replication
			[ ] tag "Replication"
			[+] MenuItem Replicate
				[+] multitag "Replicate"
					[ ] "$2334"
			[+] MenuItem Settings
				[+] multitag "Settings"
					[ ] "$2335"
			[+] MenuItem History
				[+] multitag "History"
					[ ] "$2336"
			[+] MenuItem NewReplica
				[+] multitag "New Replica"
					[ ] "$2338"
			[+] MenuItem FindReplica
				[+] multitag "Find Replica"
					[ ] "$2339"
		[+] Menu Mobile
			[ ] tag "Mobile"
			[+] MenuItem ChooseCurrentLocation
				[+] multitag "Choose Current Location"
					[ ] "$2345"
			[+] MenuItem EditCurrentLocation
				[+] multitag "Edit Current Location"
					[ ] "$2347"
			[+] MenuItem EditCurrentTimePhone
				[+] multitag "Edit Current Time?Phone"
					[ ] "$2349"
			[+] MenuItem Locations
				[+] multitag "Locations"
					[ ] "$2351"
			[+] MenuItem ServerPhoneNumbers
				[+] multitag "Server Phone Numbers"
					[ ] "$2352"
			[+] MenuItem CallServer
				[+] multitag "Call Server"
					[ ] "$2354"
			[+] MenuItem HangUp
				[+] multitag "Hang Up"
					[ ] "$2355"
		[+] Menu InstantMessaging
			[ ] tag "Instant Messaging"
			[+] MenuItem Chat
				[+] multitag "Chat"
					[ ] "$2360"
			[+] MenuItem StartInstantMeeting
				[+] multitag "Start Instant Meeting"
					[ ] "$2361"
			[+] MenuItem AddToInstantContactList
				[+] multitag "Add to Instant Contact List"
					[ ] "$2362"
			[+] MenuItem ShowHideInstantContactList
				[+] multitag "Show?Hide Instant Contact List"
					[ ] "$2363"
			[+] MenuItem IAmActive
				[+] multitag "I Am Active"
					[ ] "$2365"
			[+] MenuItem IAmAway
				[+] multitag "I Am Away"
					[ ] "$2366"
			[+] MenuItem DoNotDisturbMe
				[+] multitag "Do Not Disturb Me"
					[ ] "$2367"
			[+] MenuItem LogOffInstantMessaging
				[+] multitag "Log Off Instant Messaging"
					[ ] "$2369"
			[+] MenuItem LogOnInstantMessaging
				[+] multitag "Log On Instant Messaging"
					[ ] "$2370"
			[+] MenuItem CancelLogOnInstantMessagin
				[+] multitag "Cancel Log On Instant Messaging"
					[ ] "$2371"
		[+] MenuItem Attach
			[+] multitag "Attach"
				[ ] "$2375"
		[+] MenuItem Import
			[+] multitag "Import"
				[ ] "$2376"
		[+] MenuItem Export
			[+] multitag "Export"
				[ ] "$2377"
		[+] MenuItem PageSetup
			[+] multitag "Page Setup"
				[ ] "$2380"
		[+] MenuItem PrintPreview
			[+] multitag "Print Preview"
				[ ] "$2384"
		[+] MenuItem Print
			[+] multitag "Print"
				[ ] "$2385"
		[+] Menu Preferences
			[ ] tag "Preferences"
			[+] MenuItem UserPreferences
				[+] multitag "User Preferences"
					[ ] "$2389"
			[+] MenuItem LocationPreferences
				[+] multitag "Location Preferences"
					[ ] "$2390"
			[+] MenuItem ClientReconfigurationWizard
				[+] multitag "Client Reconfiguration Wizard"
					[ ] "$2391"
			[+] MenuItem AdministrationPreferences
				[+] multitag "Administration Preferences"
					[ ] "$2393"
			[+] MenuItem ToolbarPreferences
				[+] multitag "Toolbar Preferences"
					[ ] "$2395"
			[+] MenuItem StatusBarPreferences
				[+] multitag "Status Bar Preferences"
					[ ] "$2396"
		[+] Menu Security
			[ ] tag "Security"
			[+] MenuItem UserSecurity
				[+] multitag "User Security"
					[ ] "$2401"
			[+] MenuItem LockDisplay
				[+] multitag "Lock Display"
					[ ] "$2402"
			[+] MenuItem SwitchID
				[+] multitag "Switch ID"
					[ ] "$2403"
		[+] Menu Tools
			[ ] tag "Tools"
			[+] MenuItem CleanupSeparators
				[+] multitag "?Cleanup Separators?"
					[ ] "$2407"
			[+] MenuItem NotesSmartUpgrade
				[+] multitag "Notes Smart Upgrade"
					[ ] "$2409"
			[+] MenuItem ServerAdministration
				[+] multitag "Server Administration"
					[ ] "$2414"
			[+] MenuItem DebugLotusScript
				[+] multitag "Debug LotusScript"
					[ ] "$2417"
			[+] MenuItem RemoteDebugger
				[+] multitag "Remote Debugger"
					[ ] "$2419"
			[+] MenuItem ShowJavaDebugConsole
				[+] multitag "Show Java Debug Console"
					[ ] "$2421"
			[+] MenuItem StopLocalWebPreviewProcess
				[+] multitag "Stop Local Web Preview process"
					[ ] "$2422"
		[+] MenuItem ExitNotes
			[+] multitag "Exit Notes"
				[ ] "$2428"
	[+] Menu Edit
		[ ] tag "Edit"
		[+] MenuItem Undo
			[+] multitag "Undo"
				[ ] "$2723"
		[+] MenuItem Properties
			[+] multitag "Properties"
				[ ] "$2727"
		[+] MenuItem Cut
			[+] multitag "Cut"
				[ ] "$2730"
		[+] MenuItem Copy
			[+] multitag "Copy"
				[ ] "$2731"
		[+] Menu CopyAsLink
			[ ] tag "Copy As Link"
			[+] MenuItem AnchorLink
				[+] multitag "Anchor Link"
					[ ] "$2734"
			[+] MenuItem DocumentLink
				[+] multitag "Document Link"
					[ ] "$2735"
			[+] MenuItem ViewLink
				[+] multitag "View Link"
					[ ] "$2736"
			[+] MenuItem DatabaseLink
				[+] multitag "Database Link"
					[ ] "$2737"
		[+] MenuItem CopySelectedAsTable
			[+] multitag "Copy Selected as Table"
				[ ] "$2739"
		[+] MenuItem Paste
			[+] multitag "Paste"
				[ ] "$2740"
		[+] MenuItem PasteSpecial
			[+] multitag "Paste Special"
				[ ] "$2742"
		[+] MenuItem Delete
			[+] multitag "Delete"
				[ ] "$2743"
		[+] MenuItem Restore
			[+] multitag "Restore"
				[ ] "$2744"
		[+] MenuItem SelectAll
			[+] multitag "Select All"
				[ ] "$2745"
		[+] MenuItem DeselectAll
			[+] multitag "Deselect All"
				[ ] "$2746"
		[+] MenuItem QuoteSelection
			[+] multitag "Quote Selection"
				[ ] "$2747"
		[+] MenuItem FindReplace
			[+] multitag "Find?Replace"
				[ ] "$2749"
		[+] MenuItem FindNext
			[+] multitag "Find Next"
				[ ] "$2750"
		[+] MenuItem CheckSpelling
			[+] multitag "Check Spelling"
				[ ] "$2751"
		[+] MenuItem ExternalLinks
			[+] multitag "External Links"
				[ ] "$2753"
		[+] Menu UnreadMarks
			[ ] tag "Unread Marks"
			[+] MenuItem MarkSelectedRead
				[+] multitag "Mark Selected Read"
					[ ] "$2756"
			[+] MenuItem MarkAllRead
				[+] multitag "Mark All Read"
					[ ] "$2757"
			[+] MenuItem MarkSelectedUnread
				[+] multitag "Mark Selected Unread"
					[ ] "$2758"
			[+] MenuItem MarkAllUnread
				[+] multitag "Mark All Unread"
					[ ] "$2759"
			[+] MenuItem ScanUnread
				[+] multitag "Scan Unread"
					[ ] "$2761"
			[+] MenuItem ScanPreferred
				[+] multitag "Scan Preferred"
					[ ] "$2762"
			[+] MenuItem ExchangeUnreadMarks
				[+] multitag "Exchange Unread Marks"
					[ ] "$2763"
	[+] Menu View
		[ ] tag "View"
		[+] MenuItem RefreshUnreadCount
			[+] multitag "Refresh Unread Count"
				[ ] "$3697"
		[+] MenuItem ArrangeIcons
			[+] multitag "Arrange Icons"
				[ ] "$3699"
		[+] MenuItem StackReplicaIcons
			[+] multitag "Stack Replica Icons"
				[ ] "$3700"
		[+] MenuItem ShowUnread
			[+] multitag "Show Unread"
				[ ] "$3701"
		[+] MenuItem ShowServerNames
			[+] multitag "Show Server Names"
				[ ] "$3702"
		[+] MenuItem GoTo
			[+] multitag "Go To"
				[ ] "$3704"
		[+] MenuItem Agents
			[+] multitag "Agents"
				[ ] "$3705"
		[+] MenuItem Design
			[+] multitag "Design"
				[ ] "$3706"
		[+] MenuItem NoViewsAvailable
			[+] multitag "No Views Available"
				[ ] "$3707"
	[+] Menu Create
		[ ] tag "Create"
		[+] MenuItem CleanupSeparators
			[+] multitag "?Cleanup Separators?"
				[ ] "$3131"
		[+] MenuItem Mail
			[+] multitag "Mail"
				[ ] "$3132"
		[+] MenuItem NoFormsAvailable
			[+] multitag "No Forms Available"
				[ ] "$3133"
		[+] MenuItem WorkspacePage
			[+] multitag "Workspace Page"
				[ ] "$3135"
		[+] MenuItem Bookmark
			[+] multitag "Bookmark"
				[ ] "$3137"
		[+] MenuItem Subscription
			[+] multitag "Subscription"
				[ ] "$3138"
		[+] MenuItem Agent
			[+] multitag "Agent"
				[ ] "$3140"
		[+] MenuItem Folder
			[+] multitag "Folder"
				[ ] "$3141"
		[+] MenuItem View
			[+] multitag "View"
				[ ] "$3142"
		[+] MenuItem Design
			[+] multitag "Design"
				[ ] "$3144"
	[+] Menu Actions
		[ ] tag "Actions"
		[+] MenuItem CleanupSeparators
			[+] multitag "?Cleanup Separators?"
				[ ] "$3050"
		[+] MenuItem AgentPlaceholder
			[+] multitag "?Agent Placeholder?"
				[ ] "$3051"
		[+] MenuItem AddInsPlaceholder
			[+] multitag "?AddIns Placeholder?"
				[ ] "$3052"
	[+] Menu Help
		[ ] tag "Help"
		[+] MenuItem HelpTopics
			[+] multitag "Help Topics"
				[ ] "$3543"
		[+] MenuItem ContextHelp
			[+] multitag "Context Help"
				[ ] "$3547"
		[+] Menu LotusInternetResources
			[ ] tag "Lotus Internet Resources"
			[+] MenuItem DominoNotesDocLibrary
				[+] multitag "Domino & Notes Doc Library"
					[ ] "$3554"
			[+] MenuItem NotesHomePage
				[+] multitag "Notes Home Page"
					[ ] "$3555"
			[+] MenuItem LotusHomePage
				[+] multitag "Lotus Home Page"
					[ ] "$3556"
			[+] MenuItem LotusCustomerSupport
				[+] multitag "Lotus Customer Support"
					[ ] "$3557"
		[+] MenuItem AboutThisDatabase
			[+] multitag "About This Database"
				[ ] "$3560"
		[+] MenuItem UsingThisDatabase
			[+] multitag "Using This Database"
				[ ] "$3561"
		[+] MenuItem AboutNotes
			[+] multitag "About Notes"
				[ ] "$3563"
	[+] CustomWin NOTES_MDICLIENT1
		[ ] tag "[NOTES_MDICLIENT]$123"
	[+] CustomWin FramesetSubprog1
		[ ] tag "[FramesetSubprog]$0"
		[+] CustomWin IconBarHost2
			[+] multitag "[SmrtIcnL1W04]Icon Bar Host[2]"
				[ ] "[SmrtIcnL1W04]$0[2]"
		[+] CustomWin IconBarHost1
			[+] multitag "[SmrtIcnL1W04]Icon Bar Host[1]"
				[ ] "[SmrtIcnL1W04]$0[1]"
			[+] CustomWin IconBarWindow3
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[3]"
					[ ] "[SmrtIcnL1W05]$0[3]"
			[+] CustomWin IconBarWindow1
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[1]"
					[ ] "[SmrtIcnL1W05]$0[1]"
			[+] CustomWin IconBarWindow2
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[2]"
					[ ] "[SmrtIcnL1W05]$0[2]"
			[+] CustomWin IconBarWindow5
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[5]"
					[ ] "[SmrtIcnL1W05]$0[5]"
			[+] CustomWin IconBarWindow4
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[4]"
					[ ] "[SmrtIcnL1W05]$0[4]"
				[+] ComboBox ComboBox1
					[ ] tag "$20000"
			[+] CustomWin IconBarWindow6
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[6]"
					[ ] "[SmrtIcnL1W05]$0[6]"
		[+] CustomWin StatusWindow1
			[ ] tag "[StatusWindow]$0"
		[+] CustomWin FramesetSubprog1
			[ ] tag "[FramesetSubprog]$0"
			[+] CustomWin IconBarHost2
				[+] multitag "[SmrtIcnL1W04]Icon Bar Host[2]"
					[ ] "[SmrtIcnL1W04]$0[2]"
			[+] CustomWin IconBarHost1
				[+] multitag "[SmrtIcnL1W04]Icon Bar Host[1]"
					[ ] "[SmrtIcnL1W04]$0[1]"
			[+] CustomWin FramesetSubprog1
				[ ] tag "[FramesetSubprog]$0[1]"
				[+] CustomWin NotesSubprog1
					[ ] tag "[NotesSubprog]$0"
					[+] CustomWin IRISOdlistContainer1
						[ ] tag "[IRIS.odlistContainer]$1230"
						[+] CustomWin IRISOdlistList1
							[ ] tag "[IRIS.odlistList]$0"
				[+] CustomWin NotesTabSpacer1
					[ ] tag "[NotesTabSpacer]$0"
			[+] CustomWin FramesetSubprog2
				[ ] tag "[FramesetSubprog]$0[2]"
				[+] CustomWin FramesetSubprog1
					[ ] tag "[FramesetSubprog]$0[1]"
					[+] CustomWin NotesTabMgr1
						[ ] tag "[NotesTabMgr]$1"
				[+] CustomWin FramesetSubprog2
					[ ] tag "[FramesetSubprog]$0[2]"
					[+] CustomWin FramesetSubprog1
						[ ] tag "[FramesetSubprog]$0"
						[+] CustomWin NotesSubprog1
							[ ] tag "[NotesSubprog]$0"
							[+] CustomWin IRISOdlistContainer1
								[ ] tag "[IRIS.odlistContainer]$1230"
								[+] CustomWin IRISOdlistList1
									[ ] tag "[IRIS.odlistList]$0"
						[+] CustomWin Databases
							[+] multitag "[NotesPageControls]Databases"
								[ ] "[NotesPageControls]$0"
							[+] CustomWin IRISBmpbutton3
								[ ] tag "[IRIS.bmpbutton]$3"
							[+] CustomWin IRISBmpbutton2
								[ ] tag "[IRIS.bmpbutton]$5"
							[+] CustomWin IRISBmpbutton1
								[ ] tag "[IRIS.bmpbutton]$1"
							[+] CustomWin Sort
								[+] multitag "[IRIS.bmpbutton]Sort"
									[ ] "[IRIS.bmpbutton]$6"
							[+] CustomWin View
								[+] multitag "[IRIS.bmpbutton]View"
									[ ] "[IRIS.bmpbutton]$7"
							[+] CustomWin IRISBmpbutton6
								[ ] tag "[IRIS.bmpbutton]$8"
							[+] CustomWin IRISBmpbutton7
								[ ] tag "[IRIS.bmpbutton]$9"
					[+] CustomWin Workspace
						[+] multitag "[NtDsWnd]Workspace"
							[ ] "[NtDsWnd]$0"
[+] window MainWin NotesStartup
	[+] multitag "Welcome*"
		[ ] "*Lotus Notes"
		[ ] "$D:\Lotus\Domino\nlnotes.exe"
	[+] Menu File
		[ ] tag "File"
		[+] MenuItem CleanupEmptyPopups
			[+] multitag "?Cleanup Empty Popups?"
				[ ] "$2299"
		[+] Menu New
			[ ] tag "New"
			[+] MenuItem New
				[+] multitag "New"
					[ ] "$2302"
		[+] MenuItem Close
			[+] multitag "Close"
				[ ] "$61536"
		[+] MenuItem Save
			[+] multitag "Save"
				[ ] "$2309"
		[+] MenuItem SaveAsNewVersion
			[+] multitag "Save As New Version"
				[ ] "$2310"
		[+] MenuItem DocumentProperties
			[+] multitag "Document Properties"
				[ ] "$2311"
		[+] Menu Database
			[ ] tag "Database"
			[+] MenuItem CleanupSeparators
				[+] multitag "?Cleanup Separators?"
					[ ] "$2315"
			[+] MenuItem Properties
				[+] multitag "Properties"
					[ ] "$2316"
			[+] MenuItem AccessControl
				[+] multitag "Access Control"
					[ ] "$2317"
			[+] MenuItem Open
				[+] multitag "Open"
					[ ] "$2319"
			[+] MenuItem New
				[+] multitag "New"
					[ ] "$2321"
			[+] MenuItem NewCopy
				[+] multitag "New Copy"
					[ ] "$2323"
			[+] MenuItem Publish
				[+] multitag "Publish"
					[ ] "$2324"
			[+] MenuItem Delete
				[+] multitag "Delete"
					[ ] "$2325"
			[+] MenuItem Archive
				[+] multitag "Archive"
					[ ] "$2326"
			[+] MenuItem RefreshDesign
				[+] multitag "Refresh Design"
					[ ] "$2328"
			[+] MenuItem ReplaceDesign
				[+] multitag "Replace Design"
					[ ] "$2329"
			[+] MenuItem DesignSynopsis
				[+] multitag "Design Synopsis"
					[ ] "$2330"
		[+] Menu Replication
			[ ] tag "Replication"
			[+] MenuItem Replicate
				[+] multitag "Replicate"
					[ ] "$2334"
			[+] MenuItem Settings
				[+] multitag "Settings"
					[ ] "$2335"
			[+] MenuItem History
				[+] multitag "History"
					[ ] "$2336"
			[+] MenuItem NewReplica
				[+] multitag "New Replica"
					[ ] "$2338"
			[+] MenuItem FindReplica
				[+] multitag "Find Replica"
					[ ] "$2339"
		[+] Menu Mobile
			[ ] tag "Mobile"
			[+] MenuItem ChooseCurrentLocation
				[+] multitag "Choose Current Location"
					[ ] "$2345"
			[+] MenuItem EditCurrentLocation
				[+] multitag "Edit Current Location"
					[ ] "$2347"
			[+] MenuItem EditCurrentTimePhone
				[+] multitag "Edit Current Time?Phone"
					[ ] "$2349"
			[+] MenuItem Locations
				[+] multitag "Locations"
					[ ] "$2351"
			[+] MenuItem ServerPhoneNumbers
				[+] multitag "Server Phone Numbers"
					[ ] "$2352"
			[+] MenuItem CallServer
				[+] multitag "Call Server"
					[ ] "$2354"
			[+] MenuItem HangUp
				[+] multitag "Hang Up"
					[ ] "$2355"
		[+] Menu InstantMessaging
			[ ] tag "Instant Messaging"
			[+] MenuItem Chat
				[+] multitag "Chat"
					[ ] "$2360"
			[+] MenuItem StartInstantMeeting
				[+] multitag "Start Instant Meeting"
					[ ] "$2361"
			[+] MenuItem AddToInstantContactList
				[+] multitag "Add to Instant Contact List"
					[ ] "$2362"
			[+] MenuItem ShowHideInstantContactList
				[+] multitag "Show?Hide Instant Contact List"
					[ ] "$2363"
			[+] MenuItem IAmActive
				[+] multitag "I Am Active"
					[ ] "$2365"
			[+] MenuItem IAmAway
				[+] multitag "I Am Away"
					[ ] "$2366"
			[+] MenuItem DoNotDisturbMe
				[+] multitag "Do Not Disturb Me"
					[ ] "$2367"
			[+] MenuItem LogOffInstantMessaging
				[+] multitag "Log Off Instant Messaging"
					[ ] "$2369"
			[+] MenuItem LogOnInstantMessaging
				[+] multitag "Log On Instant Messaging"
					[ ] "$2370"
			[+] MenuItem CancelLogOnInstantMessagin
				[+] multitag "Cancel Log On Instant Messaging"
					[ ] "$2371"
		[+] MenuItem Attach
			[+] multitag "Attach"
				[ ] "$2375"
		[+] MenuItem Import
			[+] multitag "Import"
				[ ] "$2376"
		[+] MenuItem Export
			[+] multitag "Export"
				[ ] "$2377"
		[+] MenuItem PageSetup
			[+] multitag "Page Setup"
				[ ] "$2380"
		[+] MenuItem PrintPreview
			[+] multitag "Print Preview"
				[ ] "$2384"
		[+] MenuItem Print
			[+] multitag "Print"
				[ ] "$2385"
		[+] Menu Preferences
			[ ] tag "Preferences"
			[+] MenuItem UserPreferences
				[+] multitag "User Preferences"
					[ ] "$2389"
			[+] MenuItem LocationPreferences
				[+] multitag "Location Preferences"
					[ ] "$2390"
			[+] MenuItem ClientReconfigurationWizard
				[+] multitag "Client Reconfiguration Wizard"
					[ ] "$2391"
			[+] MenuItem AdministrationPreferences
				[+] multitag "Administration Preferences"
					[ ] "$2393"
			[+] MenuItem ToolbarPreferences
				[+] multitag "Toolbar Preferences"
					[ ] "$2395"
			[+] MenuItem StatusBarPreferences
				[+] multitag "Status Bar Preferences"
					[ ] "$2396"
		[+] Menu Security
			[ ] tag "Security"
			[+] MenuItem UserSecurity
				[+] multitag "User Security"
					[ ] "$2401"
			[+] MenuItem LockDisplay
				[+] multitag "Lock Display"
					[ ] "$2402"
			[+] MenuItem SwitchID
				[+] multitag "Switch ID"
					[ ] "$2403"
		[+] Menu Tools
			[ ] tag "Tools"
			[+] MenuItem CleanupSeparators
				[+] multitag "?Cleanup Separators?"
					[ ] "$2407"
			[+] MenuItem NotesSmartUpgrade
				[+] multitag "Notes Smart Upgrade"
					[ ] "$2409"
			[+] MenuItem ServerAdministration
				[+] multitag "Server Administration"
					[ ] "$2414"
			[+] MenuItem DebugLotusScript
				[+] multitag "Debug LotusScript"
					[ ] "$2417"
			[+] MenuItem RemoteDebugger
				[+] multitag "Remote Debugger"
					[ ] "$2419"
			[+] MenuItem ShowJavaDebugConsole
				[+] multitag "Show Java Debug Console"
					[ ] "$2421"
			[+] MenuItem StopLocalWebPreviewProcess
				[+] multitag "Stop Local Web Preview process"
					[ ] "$2422"
		[+] MenuItem ExitNotes
			[+] multitag "Exit Notes"
				[ ] "$2428"
	[+] Menu Edit
		[ ] tag "Edit"
		[+] MenuItem Undo
			[+] multitag "Undo"
				[ ] "$2723"
		[+] MenuItem Properties
			[+] multitag "Properties"
				[ ] "$2727"
		[+] MenuItem Cut
			[+] multitag "Cut"
				[ ] "$2730"
		[+] MenuItem Copy
			[+] multitag "Copy"
				[ ] "$2731"
		[+] Menu CopyAsLink
			[ ] tag "Copy As Link"
			[+] MenuItem AnchorLink
				[+] multitag "Anchor Link"
					[ ] "$2734"
			[+] MenuItem DocumentLink
				[+] multitag "Document Link"
					[ ] "$2735"
			[+] MenuItem ViewLink
				[+] multitag "View Link"
					[ ] "$2736"
			[+] MenuItem DatabaseLink
				[+] multitag "Database Link"
					[ ] "$2737"
		[+] MenuItem CopySelectedAsTable
			[+] multitag "Copy Selected as Table"
				[ ] "$2739"
		[+] MenuItem Paste
			[+] multitag "Paste"
				[ ] "$2740"
		[+] MenuItem PasteSpecial
			[+] multitag "Paste Special"
				[ ] "$2742"
		[+] MenuItem Delete
			[+] multitag "Delete"
				[ ] "$2743"
		[+] MenuItem Restore
			[+] multitag "Restore"
				[ ] "$2744"
		[+] MenuItem SelectAll
			[+] multitag "Select All"
				[ ] "$2745"
		[+] MenuItem DeselectAll
			[+] multitag "Deselect All"
				[ ] "$2746"
		[+] MenuItem QuoteSelection
			[+] multitag "Quote Selection"
				[ ] "$2747"
		[+] MenuItem FindReplace
			[+] multitag "Find?Replace"
				[ ] "$2749"
		[+] MenuItem FindNext
			[+] multitag "Find Next"
				[ ] "$2750"
		[+] MenuItem CheckSpelling
			[+] multitag "Check Spelling"
				[ ] "$2751"
		[+] MenuItem ExternalLinks
			[+] multitag "External Links"
				[ ] "$2753"
		[+] Menu UnreadMarks
			[ ] tag "Unread Marks"
			[+] MenuItem MarkSelectedRead
				[+] multitag "Mark Selected Read"
					[ ] "$2756"
			[+] MenuItem MarkAllRead
				[+] multitag "Mark All Read"
					[ ] "$2757"
			[+] MenuItem MarkSelectedUnread
				[+] multitag "Mark Selected Unread"
					[ ] "$2758"
			[+] MenuItem MarkAllUnread
				[+] multitag "Mark All Unread"
					[ ] "$2759"
			[+] MenuItem ScanUnread
				[+] multitag "Scan Unread"
					[ ] "$2761"
			[+] MenuItem ScanPreferred
				[+] multitag "Scan Preferred"
					[ ] "$2762"
			[+] MenuItem ExchangeUnreadMarks
				[+] multitag "Exchange Unread Marks"
					[ ] "$2763"
	[+] Menu View
		[ ] tag "View"
		[+] MenuItem Refresh
			[+] multitag "Refresh"
				[ ] "$4920"
		[+] MenuItem Ruler
			[+] multitag "Ruler"
				[ ] "$4922"
		[+] Menu Show
			[ ] tag "Show"
			[+] MenuItem HorizontalScrollBar
				[+] multitag "Horizontal Scroll Bar"
					[ ] "$4927"
			[+] MenuItem PageBreaks
				[+] multitag "Page Breaks"
					[ ] "$4929"
			[+] MenuItem HiddenCharacters
				[+] multitag "Hidden Characters"
					[ ] "$4930"
			[+] MenuItem FieldHelp
				[+] multitag "Field Help"
					[ ] "$4931"
			[+] MenuItem HiddenFromNotes
				[+] multitag "Hidden From Notes"
					[ ] "$4933"
			[+] MenuItem PageSource
				[+] multitag "Page Source"
					[ ] "$4934"
			[+] MenuItem PassThruHTML
				[+] multitag "Pass-Thru HTML"
					[ ] "$4935"
			[+] MenuItem JavaAppletsRunning
				[+] multitag "Java Applets Running"
					[ ] "$4936"
		[+] MenuItem ParentPreview
			[+] multitag "Parent Preview"
				[ ] "$4939"
		[+] MenuItem DocumentLinkPreview
			[+] multitag "Document Link Preview"
				[ ] "$4940"
		[+] MenuItem ImageResources
			[+] multitag "Image Resources"
				[ ] "$4942"
		[+] MenuItem ExpandAllSections
			[+] multitag "Expand All Sections"
				[ ] "$4944"
		[+] MenuItem CollapseAllSections
			[+] multitag "Collapse All Sections"
				[ ] "$4945"
		[+] MenuItem SwitchForm
			[+] multitag "Switch Form"
				[ ] "$4946"
		[+] MenuItem GoUpToParentView
			[+] multitag "Go up to Parent View"
				[ ] "$4947"
		[+] MenuItem Encoding
			[+] multitag "Encoding"
				[ ] "$4948"
		[+] MenuItem TextSize
			[+] multitag "Text Size"
				[ ] "$4950"
	[+] Menu Create
		[ ] tag "Create"
		[+] MenuItem CleanupSeparators
			[+] multitag "?Cleanup Separators?"
				[ ] "$3224"
		[+] MenuItem Mail
			[+] multitag "Mail"
				[ ] "$3225"
		[+] MenuItem Bookmark
			[+] multitag "Bookmark"
				[ ] "$3226"
		[+] MenuItem Document
			[+] multitag "Document"
				[ ] "$3227"
		[+] MenuItem PageBreak
			[+] multitag "Page Break"
				[ ] "$3229"
		[+] MenuItem HorizontalRule
			[+] multitag "Horizontal Rule"
				[ ] "$3230"
		[+] MenuItem Section
			[+] multitag "Section"
				[ ] "$3231"
		[+] MenuItem ComputedText
			[+] multitag "Computed Text"
				[ ] "$3232"
		[+] Menu Hotspot
			[ ] tag "Hotspot"
			[+] MenuItem LinkHotspot
				[+] multitag "Link Hotspot"
					[ ] "$3235"
			[+] MenuItem TextPopUp
				[+] multitag "Text Pop-up"
					[ ] "$3236"
			[+] MenuItem Button
				[+] multitag "Button"
					[ ] "$3238"
			[+] MenuItem FormulaPopUp
				[+] multitag "Formula Pop-up"
					[ ] "$3239"
			[+] MenuItem ActionHotspot
				[+] multitag "Action Hotspot"
					[ ] "$3240"
		[+] MenuItem Table
			[+] multitag "Table"
				[ ] "$3243"
		[+] MenuItem Picture
			[+] multitag "Picture"
				[ ] "$3244"
		[+] MenuItem ImageResource
			[+] multitag "Image Resource"
				[ ] "$3245"
		[+] MenuItem InsertResource
			[+] multitag "Insert Resource"
				[ ] "$3246"
		[+] MenuItem Object
			[+] multitag "Object"
				[ ] "$3248"
		[+] MenuItem JavaApplet
			[+] multitag "Java Applet"
				[ ] "$3251"
		[+] MenuItem EmbeddedElement
			[+] multitag "Embedded Element"
				[ ] "$3253"
		[+] MenuItem Design
			[+] multitag "Design"
				[ ] "$3255"
	[+] Menu Actions
		[ ] tag "Actions"
		[+] MenuItem CleanupSeparators
			[+] multitag "?Cleanup Separators?"
				[ ] "$3014"
		[+] MenuItem AgentPlaceholder
			[+] multitag "?Agent Placeholder?"
				[ ] "$3015"
		[+] MenuItem CertificatePlaceholder
			[+] multitag "?Certificate Placeholder?"
				[ ] "$3017"
		[+] MenuItem ActionsPlaceholder
			[+] multitag "?Actions Placeholder?"
				[ ] "$3018"
		[+] MenuItem Categorize
			[+] multitag "Categorize"
				[ ] "$3019"
		[+] MenuItem EditDocument
			[+] multitag "Edit Document"
				[ ] "$3020"
		[+] MenuItem Forward
			[+] multitag "Forward"
				[ ] "$3021"
		[+] MenuItem SendDocument
			[+] multitag "Send Document"
				[ ] "$3022"
		[+] MenuItem MoveToFolder
			[+] multitag "Move To Folder"
				[ ] "$3023"
		[+] MenuItem RemoveFromFolder
			[+] multitag "Remove From Folder"
				[ ] "$3024"
		[+] MenuItem RetrieveEntireDocument
			[+] multitag "Retrieve Entire Document"
				[ ] "$3026"
		[+] MenuItem OpenArchivedDocument
			[+] multitag "Open Archived Document"
				[ ] "$3027"
		[+] MenuItem AddInsPlaceholder
			[+] multitag "?AddIns Placeholder?"
				[ ] "$3028"
		[+] MenuItem EditHTMLAttributes
			[+] multitag "Edit HTML Attributes"
				[ ] "$3031"
		[+] MenuItem PreviewInWebBrowser
			[+] multitag "Preview In Web Browser"
				[ ] "$3033"
	[+] Menu Text
		[ ] tag "Text"
		[+] MenuItem TextProperties
			[+] multitag "Text Properties"
				[ ] "$4672"
		[+] Menu PermanentPen
			[ ] tag "Permanent Pen"
			[+] MenuItem UsePermanentPen
				[+] multitag "Use Permanent Pen"
					[ ] "$4676"
			[+] MenuItem SetPermanentPenStyle
				[+] multitag "Set Permanent Pen Style"
					[ ] "$4677"
		[+] Menu Highlighter
			[ ] tag "Highlighter"
			[+] MenuItem UseYellowHighlighter
				[+] multitag "Use Yellow Highlighter"
					[ ] "$4681"
			[+] MenuItem UsePinkHighlighter
				[+] multitag "Use Pink Highlighter"
					[ ] "$4682"
			[+] MenuItem UseBlueHighlighter
				[+] multitag "Use Blue Highlighter"
					[ ] "$4683"
		[+] MenuItem PassThruHTML
			[+] multitag "Pass-Thru HTML"
				[ ] "$4685"
		[+] MenuItem NormalText
			[+] multitag "Normal Text"
				[ ] "$4687"
		[+] MenuItem Italic
			[+] multitag "Italic"
				[ ] "$4688"
		[+] MenuItem Bold
			[+] multitag "Bold"
				[ ] "$4689"
		[+] MenuItem Underline
			[+] multitag "Underline"
				[ ] "$4690"
		[+] MenuItem Strikethrough
			[+] multitag "Strikethrough"
				[ ] "$4691"
		[+] Menu Effects
			[ ] tag "Effects"
			[+] MenuItem Shadow
				[+] multitag "Shadow"
					[ ] "$4694"
			[+] MenuItem Emboss
				[+] multitag "Emboss"
					[ ] "$4695"
			[+] MenuItem Extrude
				[+] multitag "Extrude"
					[ ] "$4696"
			[+] MenuItem Superscript
				[+] multitag "Superscript"
					[ ] "$4697"
			[+] MenuItem Subscript
				[+] multitag "Subscript"
					[ ] "$4698"
		[+] Menu Color
			[ ] tag "Color"
			[+] MenuItem Black
				[+] multitag "Black"
					[ ] "$4702"
			[+] MenuItem Red
				[+] multitag "Red"
					[ ] "$4703"
			[+] MenuItem Blue
				[+] multitag "Blue"
					[ ] "$4704"
			[+] MenuItem DarkRed
				[+] multitag "Dark Red"
					[ ] "$4705"
			[+] MenuItem DarkGreen
				[+] multitag "Dark Green"
					[ ] "$4706"
			[+] MenuItem DarkBlue
				[+] multitag "Dark Blue"
					[ ] "$4707"
			[+] MenuItem DarkMagenta
				[+] multitag "Dark Magenta"
					[ ] "$4708"
			[+] MenuItem DarkYellow
				[+] multitag "Dark Yellow"
					[ ] "$4709"
			[+] MenuItem DarkCyan
				[+] multitag "Dark Cyan"
					[ ] "$4710"
			[+] MenuItem Other
				[+] multitag "Other"
					[ ] "$4712"
		[+] MenuItem EnlargeSize
			[+] multitag "Enlarge Size"
				[ ] "$4715"
		[+] MenuItem ReduceSize
			[+] multitag "Reduce Size"
				[ ] "$4716"
		[+] Menu AlignParagraph
			[ ] tag "Align Paragraph"
			[+] MenuItem Left
				[+] multitag "Left"
					[ ] "$4724"
			[+] MenuItem Right
				[+] multitag "Right"
					[ ] "$4725"
			[+] MenuItem Center
				[+] multitag "Center"
					[ ] "$4726"
			[+] MenuItem Full
				[+] multitag "Full"
					[ ] "$4727"
			[+] MenuItem NoWrap
				[+] multitag "No Wrap"
					[ ] "$4728"
		[+] Menu xList
			[ ] tag "List"
			[+] MenuItem Bullet
				[+] multitag "Bullet"
					[ ] "$4732"
			[+] MenuItem xNumber
				[+] multitag "Number"
					[ ] "$4733"
			[+] MenuItem Checkmark
				[+] multitag "Checkmark"
					[ ] "$4734"
			[+] MenuItem Circle
				[+] multitag "Circle"
					[ ] "$4735"
			[+] MenuItem Square
				[+] multitag "Square"
					[ ] "$4736"
			[+] MenuItem UppercaseAlphabetic
				[+] multitag "Uppercase Alphabetic"
					[ ] "$4737"
			[+] MenuItem LowercaseAlphabetic
				[+] multitag "Lowercase Alphabetic"
					[ ] "$4738"
			[+] MenuItem UppercaseRoman
				[+] multitag "Uppercase Roman"
					[ ] "$4739"
			[+] MenuItem LowercaseRoman
				[+] multitag "Lowercase Roman"
					[ ] "$4740"
			[+] MenuItem MoveItemUp
				[+] multitag "Move Item Up"
					[ ] "$4742"
			[+] MenuItem MoveItemDown
				[+] multitag "Move Item Down"
					[ ] "$4743"
		[+] Menu Spacing
			[ ] tag "Spacing"
			[+] MenuItem Single
				[+] multitag "Single"
					[ ] "$4747"
			[+] MenuItem OneAndAHalf
				[+] multitag "One and a Half"
					[ ] "$4748"
			[+] MenuItem xDouble
				[+] multitag "Double"
					[ ] "$4749"
		[+] MenuItem Indent
			[+] multitag "Indent"
				[ ] "$4751"
		[+] MenuItem Outdent
			[+] multitag "Outdent"
				[ ] "$4752"
		[+] MenuItem ApplyStyle
			[+] multitag "Apply Style"
				[ ] "$4753"
	[+] Menu Help
		[ ] tag "Help"
		[+] MenuItem HelpTopics
			[+] multitag "Help Topics"
				[ ] "$3543"
		[+] MenuItem ContextHelp
			[+] multitag "Context Help"
				[ ] "$3547"
		[+] Menu LotusInternetResources
			[ ] tag "Lotus Internet Resources"
			[+] MenuItem DominoNotesDocLibrary
				[+] multitag "Domino & Notes Doc Library"
					[ ] "$3554"
			[+] MenuItem NotesHomePage
				[+] multitag "Notes Home Page"
					[ ] "$3555"
			[+] MenuItem LotusHomePage
				[+] multitag "Lotus Home Page"
					[ ] "$3556"
			[+] MenuItem LotusCustomerSupport
				[+] multitag "Lotus Customer Support"
					[ ] "$3557"
		[+] MenuItem AboutThisDatabase
			[+] multitag "About This Database"
				[ ] "$3560"
		[+] MenuItem UsingThisDatabase
			[+] multitag "Using This Database"
				[ ] "$3561"
		[+] MenuItem AboutNotes
			[+] multitag "About Notes"
				[ ] "$3563"
	[+] CustomWin NOTES_MDICLIENT1
		[ ] tag "[NOTES_MDICLIENT]$123"
	[+] CustomWin FramesetSubprog1
		[ ] tag "[FramesetSubprog]$0"
		[+] CustomWin IconBarHost2
			[+] multitag "[SmrtIcnL1W04]Icon Bar Host[2]"
				[ ] "[SmrtIcnL1W04]$0[2]"
		[+] CustomWin IconBarHost1
			[+] multitag "[SmrtIcnL1W04]Icon Bar Host[1]"
				[ ] "[SmrtIcnL1W04]$0[1]"
			[+] CustomWin IconBarWindow3
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[3]"
					[ ] "[SmrtIcnL1W05]$0[3]"
				[+] PopupList PopupList1
					[ ] tag "$302"
				[+] PopupList PopupList2
					[ ] tag "$303"
			[+] CustomWin IconBarWindow1
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[1]"
					[ ] "[SmrtIcnL1W05]$0[1]"
			[+] CustomWin IconBarWindow2
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[2]"
					[ ] "[SmrtIcnL1W05]$0[2]"
			[+] CustomWin IconBarWindow5
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[5]"
					[ ] "[SmrtIcnL1W05]$0[5]"
			[+] CustomWin IconBarWindow4
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[4]"
					[ ] "[SmrtIcnL1W05]$0[4]"
				[+] ComboBox ComboBox1
					[ ] tag "$20000"
			[+] CustomWin IconBarWindow6
				[+] multitag "[SmrtIcnL1W05]Icon Bar Window[6]"
					[ ] "[SmrtIcnL1W05]$0[6]"
		[+] CustomWin StatusWindow1
			[ ] tag "[StatusWindow]$0"
		[+] CustomWin FramesetSubprog1
			[ ] tag "[FramesetSubprog]$0"
			[+] CustomWin IconBarHost2
				[+] multitag "[SmrtIcnL1W04]Icon Bar Host[2]"
					[ ] "[SmrtIcnL1W04]$0[2]"
			[+] CustomWin IconBarHost1
				[+] multitag "[SmrtIcnL1W04]Icon Bar Host[1]"
					[ ] "[SmrtIcnL1W04]$0[1]"
			[+] CustomWin FramesetSubprog1
				[ ] tag "[FramesetSubprog]$0[1]"
				[+] CustomWin NotesSubprog1
					[ ] tag "[NotesSubprog]$0"
					[+] CustomWin IRISOdlistContainer1
						[ ] tag "[IRIS.odlistContainer]$1230"
						[+] CustomWin IRISOdlistList1
							[ ] tag "[IRIS.odlistList]$0"
				[+] CustomWin NotesTabSpacer1
					[ ] tag "[NotesTabSpacer]$0"
			[+] CustomWin FramesetSubprog2
				[ ] tag "[FramesetSubprog]$0[2]"
				[+] CustomWin FramesetSubprog1
					[ ] tag "[FramesetSubprog]$0[1]"
					[+] CustomWin NotesTabMgr1
						[ ] tag "[NotesTabMgr]$1"
				[+] CustomWin FramesetSubprog2
					[ ] tag "[FramesetSubprog]$0[2]"
					[+] CustomWin Welcome
						[+] multitag "[NotesSubprog]Welcome"
							[ ] "[NotesSubprog]$0"
						[+] CustomWin FramesetSubprog1
							[ ] tag "[FramesetSubprog]$1003"
							[+] CustomWin FramesetSubprog1
								[ ] tag "[FramesetSubprog]$0"
								[+] CustomWin NotesSubprog1
									[ ] tag "[NotesSubprog]$1002"
									[+] CustomWin FramesetSubprog1
										[ ] tag "[FramesetSubprog]$1004"
										[+] CustomWin FramesetSubprog1
											[ ] tag "[FramesetSubprog]$0"
											[+] CustomWin FramesetSubprog1
												[ ] tag "[FramesetSubprog]$0"
												[+] CustomWin NotesSubprog2
													[ ] tag "[NotesSubprog]$1002[2]"
													[+] CustomWin FramesetSubprog1
														[ ] tag "[FramesetSubprog]$1004"
														[+] CustomWin FramesetSubprog1
															[ ] tag "[FramesetSubprog]$0"
															[+] CustomWin Untitled
																[+] multitag "[NotesSubprog](Untitled)"
																	[ ] "[NotesSubprog]$1002"
																[+] CustomWin NotesRichText1
																	[ ] tag "[NotesRichText]$0"
												[+] CustomWin Untitled
													[+] multitag "[NotesSubprog](Untitled)"
														[ ] "[NotesSubprog]$1002[1]"
													[+] CustomWin NotesRichText1
														[ ] tag "[NotesRichText]$0"
					[+] CustomWin FramesetSubprog1
						[ ] tag "[FramesetSubprog]$0"
						[+] CustomWin NotesPageControls1
							[ ] tag "[NotesPageControls]$0"
							[+] CustomWin IRISBmpbutton3
								[ ] tag "[IRIS.bmpbutton]$3"
							[+] CustomWin IRISBmpbutton2
								[ ] tag "[IRIS.bmpbutton]$5"
							[+] CustomWin IRISBmpbutton1
								[ ] tag "[IRIS.bmpbutton]$1"
							[+] CustomWin Sort
								[+] multitag "[IRIS.bmpbutton]Sort"
									[ ] "[IRIS.bmpbutton]$6"
							[+] CustomWin View
								[+] multitag "[IRIS.bmpbutton]View"
									[ ] "[IRIS.bmpbutton]$7"
							[+] CustomWin IRISBmpbutton6
								[ ] tag "[IRIS.bmpbutton]$8"
							[+] CustomWin IRISBmpbutton7
								[ ] tag "[IRIS.bmpbutton]$9"
[+] window DialogBox SaveChanges
	[ ] tag "Lotus Notes"
	[ ] parent WorkspaceLotusNotes
	[+] PushButton Yes
		[-] multitag "Yes"
			[ ] "$6"
	[+] PushButton No
		[+] multitag "No"
			[ ] "$7"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$20"
	[+] StaticText DoYouWantToDelete1Docume
		[+] multitag "Do you want to delete 1 document from the database ITFShipSD's Log?"
			[ ] "$65535"
[+] window DialogBox ProgramNotFound
	[ ] tag "Lotus Notes"
	[ ] parent WorkspaceLotusNotes
	[+] PushButton OK
		[+] multitag "OK"
			[ ] "$2"
	[+] CustomWin Icon1
		[ ] tag "[Icon]$20"
	[+] StaticText UnableToLocateProgramText
		[+] multitag "Unable to locate program"
			[ ] "$65535"
[+] window DialogBox OpenDatabase
	[ ] tag "Open Database"
	[ ] parent WorkspaceLotusNotes
	[-] StatusBar StatusBar1
		[ ] tag "$258"
		[+] DynamicText DynamicText1
			[ ] tag "#1"
	[+] StaticText ServerText
		[+] multitag "Server"
			[ ] "$65535[1]"
	[+] ComboBox Server
		[+] multitag "Server"
			[ ] "$50"
	[+] StaticText DatabaseText
		[+] multitag "Database"
			[ ] "$65535[2]"
	[+] CustomWin Database
		[+] multitag "[IRIS.odlistContainer]Database"
			[ ] "[IRIS.odlistContainer]$57"
		[+] CustomWin IRISOdlistList1
			[ ] tag "[IRIS.odlistList]$0"
	[+] StaticText FilenameText
		[+] multitag "Filename"
			[ ] "$46"
	[+] TextField Filename
		[+] multitag "Filename"
			[ ] "$59"
	[+] PushButton Open
		[+] multitag "Open"
			[ ] "$1"
	[+] PushButton Cancel
		[+] multitag "Cancel"
			[ ] "$2"
	[+] PushButton Bookmark
		[+] multitag "Bookmark"
			[ ] "$60"
	[+] PushButton About
		[+] multitag "About"
			[ ] "$49"
	[+] PushButton Browse
		[+] multitag "Browse"
			[ ] "$86"
[ ] //////////////////////////////////////////////////////
[ ] // GLOBALS 
[ ] //////////////////////////////////////////////////////
[ ] 
[ ] 
[ ] 
[ ] 
[ ] 
[ ] 
[ ] 
