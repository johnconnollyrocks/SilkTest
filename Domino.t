[ ] use "Domino.inc" 
[ ] use "centrixs.inc"
[+] testcase ClearDominoEventLog() appstate none
	[ ] WorkspaceLotusNotes.SetActive ()
	[ ] sleep(5)
	[-] if ProgramNotFound.Exists()
		[ ] ProgramNotFound.SetActive()
		[ ] ProgramNotFound.OK.Click()
	[ ] WorkspaceLotusNotes.TypeKeys("<ctrl-o>")
	[ ] OpenDatabase.SetActive()
	[ ] OpenDatabase.Filename.SetText("log.nsf")
	[ ] OpenDatabase.Open.Click()
	[ ] WorkspaceLotusNotes.TypeKeys("<ctrl-a>")
	[ ] WorkspaceLotusNotes.TypeKeys("<delete>")
	[ ] WorkspaceLotusNotes.TypeKeys("<esc>")
	[ ] SaveChanges.Yes.Click()
	[ ] 
[+] testcase RunReplication() appstate none
	[ ] STRING sDate = DateStr()
	[ ] STRING sTime = TimeStr()
	[ ] DominoCommand.SetActive()
	[ ] DominoCommand.TypeKeys("replicate 140.199.32.78 briefs.nsf<enter>")
	[ ] Print(sDate)
	[ ] Print(sTime)
	[ ] 
[+] INTEGER TimeToInt(STRING sDate)
	[ ] STRING sHour = GetField(sDate,":",1)
	[ ] STRING sMinute = GetField(sDate,":",2)
	[ ] STRING sSecond = Left(GetField(sDate,":",3),2)
	[ ] STRING sAMPM = Right(sDate,2)
	[ ] // Print("sAMPM=",sAMPM)
	[-] if sAMPM == "PM"
		[-] if val(sHour) != 12
			[ ] return ((val(sHour)+12)*3600)+(val(sMinute)*60)+(val(sSecond))
		[-] else 
			[ ] return ((val(sHour))*3600)+(val(sMinute)*60)+(val(sSecond))
	[-] if sAMPM == "AM"
		[-] if val(sHour) == 12
			[ ] // Hour will be zero
			[ ] return ((val(sHour)-12))+(val(sMinute)*60)+(val(sSecond))
		[-] else 
			[ ] return ((val(sHour))*3600)+(val(sMinute)*60)+(val(sSecond))
[+] testcase verifyDominoReplication() appstate none
		[ ] verifyReplication("140.199.32.78","testdiscuss.nsf")
[+] verifyReplication(STRING sServerIP, STRING sDBFile) 
	[ ] // LIST OF STRING lsClipboard
	[ ] LIST OF STRING lsDominoLog
	[ ] STRING sDominoLog
	[ ] INTEGER iMinutes
	[ ] STRING sStartDate = DateStr()
	[ ] STRING sStartTime = TimeStr()
	[ ] STRING sDate
	[ ] STRING sTime
	[ ] STRING sEndTime
	[ ] INTEGER iStartTime
	[ ] INTEGER iTime
	[ ] DATETIME DayAndTime
	[ ] DayAndTime = GetDateTime ()
	[ ] // format current date and time
	[+] if getServerName() == "CS01"
		[ ] sStartTime = FormatDateTime (DayAndTime, "hh:nn:ss AM/PM") 
		[ ] sStartDate = FormatDateTime (DayAndTime, "mm/dd/yyyy")
		[ ] INTEGER iStartTimeTolerance = 60
		[ ] INTEGER iReplicateTime= 10
		[ ] 
		[ ] BOOLEAN sReplicateStartedFound = FALSE
		[ ] BOOLEAN sReplicateShutdownFound = FALSE
		[ ] DominoCommand.SetActive()
		[ ] DominoCommand.TypeKeys("replicate {sServerIP} {sDBFile}<enter>")
		[ ] sleep(iReplicateTime)
		[ ] // Print(sDate)
		[ ] // Print(sTime)
		[ ] DominoCommand.SetActive ()
		[ ] // DominoCommand.Command.GetText()
		[ ] DominoCommand.Click (1, 8, -17)
		[ ] DominoCommand.TypeKeys("ES<enter>")
		[ ] // DominoCommand.ReleaseMouse (1, 8, -17)
		[ ] // DominoCommand.PressMouse (1, 55, 61)
		[ ] lsDominoLog = Clipboard.GetText()
		[ ] // ListPrint(lsDominoLog)
		[-] for each sDominoLog in lsDominoLog
			[-] if MatchStr("*Database Replicator Started*", sDominoLog)
				[ ] // Print(sDominoLog)
				[ ] sDate = Left(sDominoLog,10)
				[ ] // Print("sDate =",sDate)
				[ ] // Print("sStartDate =",sStartDate)
				[ ] sStartTime = Left(sStartTime,11)
				[ ] iStartTime = TimeToInt(sStartTime)
				[-] if sDate == sStartDate
					[ ] sTime = Right(Left(sDominoLog,22),11)
					[ ] iTime = TimeToInt(sTime)
					[-] if iStartTime - iTime < iStartTimeTolerance
						[ ] sReplicateStartedFound = TRUE
						[-] if VERBOSE == TRUE
							[ ] Print("Database Replicator Started at {sStartTime}")
			[-] if sReplicateStartedFound == TRUE
				[-] if MatchStr("*Unable to replicate with server*", sDominoLog)
					[-] do
						[ ] verify(MatchStr("*Unable to replicate with server*", sDominoLog),FALSE)
					[-] except
						[ ] LogError("Unable to Replicate with server, log entry = {sDominoLog}")
				[-] else if MatchStr("*Unable to replicate {sDBFile}*", sDominoLog)
					[-] do
						[ ] verify(MatchStr("*Unable to replicate {sDBFile}*", sDominoLog),FALSE)
					[-] except
						[ ] LogError("Unable to Replicate {sDBFile}, log entry = {sDominoLog}")
				[-] else if MatchStr("*Database Replicator Shutdown*", sDominoLog)
					[ ] // Print(sDominoLog)
					[ ] // Print(sDominoLog)
					[ ] // sDate = Left(sDominoLog,10)
					[ ] sEndTime = Right(Left(sDominoLog,22),11)
					[ ] sReplicateShutdownFound = TRUE
					[-] if VERBOSE == TRUE
						[ ] Print("Database Replicator Shutdown at {sEndTime}")
		[-] if sReplicateStartedFound == FALSE
			[ ] LogError("The String Database Replicator Started was not  found within a {iStartTimeTolerance/60} minute range")
		[-] if sReplicateShutdownFound == FALSE
			[ ] LogError("The String Database Replicator Shutdown was not found")
	[+] else
		[ ] LogWarning("testcase verifyDominoReplication not run on host {getServerName()} - hostname must be CS01")
		[ ] 
