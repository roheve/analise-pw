option explicit
'************************************************************************************
'	Changelog:
'	RH	08-05-2011	v0.01	Initial version
'	RH	12-05-2011	v0.02	Added file-browse dialog
'	RH	14-05-2011	v0.03	Removed file-browse dialog, added more command-formatting
'	RH	15-05-2011	v0.04	Added CRC check from forum
'	RH	03-07-2011	v0.05	Added more commands in de command decoder
'	RH	28-07-2011	v0.06	Added commands, detected with firware upgrade
'	RH	08-04-2012	v0.07	Corrected day of week translation
'	RH	08-07-2013	v0.07	Added new discovered Commands
'	RH	09-07-2013	v0.08	Added new discovered Commands
'	RH	02-01-2014	v0.09	Added new discovered Commands 0057 and 0058
'************************************************************************************
'* 	Based in information found everywhere, but mainly here:
'*
'* See http://www.maartendamen.com/category/plugwise-unleashed/
'* See http://www.domoticaforum.eu/viewtopic.php?f=39
'* 
'************************************************************************************

Const SCRIPTVERSION	= "0.09"

Const ForReading	= 1
Const ForWriting	= 2
Const ForAppending	= 8

Dim bDebug : bDebug = False

' A script to filter PlugWise packets from capture file
'
' Do the portmon capture with the following options:
' ShowTime=On; ShowHex=Off; ClockTime=On
' Make shure to set the MaxOutputBytes high enough (I choose 230)  
'
'The result will be line-pairs like this:
'72  14:22:31  PlugwiseConfig  IRP_MJ_WRITE  VCP0  Length 14: ....000AB43C..
'72  14:22:31  SUCCESS  
'77  14:22:31  PlugwiseConfig  IRP_MJ_READ  VCP0  Length 83
'77  14:22:31  SUCCESS  Length 83: ....0000017600C1E919......00110176000D6F0000B835CB0101060D6F0000B1B64B1606FF6B65...
'
' The filter looks *_READ and *_WRITE pairs and extracts the message
'
CONST PW_0000_REP = "0000"	'* general response code

CONST PW_CMD_0006 = "0006"	'* Advertise new Module (rcv)
CONST PW_CMD_0007 = "0007"	'* Advertise new Module (snd)

CONST PW_CMD_0008 = "0008"	'* Reset request
CONST PW_CMD_0009 = "0009"	'* 

CONST PW_000A_CMD = "000A"	'* PlugWise System Initialisation Command
CONST PW_CMD_0011 = "0011"	'* Reporting network status

CONST PW_CMD_0012 = "0012"	'* Power Information Request
CONST PW_CMD_0013 = "0013"	'* Power Information Reply

CONST PW_CMD_0014 = "0014"	'* Obsolete
CONST PW_CMD_0015 = "0015"	'* Obsolete

CONST PW_CMD_0016 = "0016"	'* Time set command

CONST PW_CMD_0017 = "0017"	'* Module Switch Command

CONST PW_CMD_0018 = "0018"	'* Module Presence Request
CONST PW_CMD_0019 = "0019"	'* Module Presence Reply

CONST PW_CMD_0023 = "0023"	'* Device Info Request
CONST PW_CMD_0024 = "0024"	'* Device Info Reply

CONST PW_CMD_003E = "0026"	'* Calibration request
CONST PW_CMD_003F = "0027"	'* Calibration reply 

CONST PW_CMD_0040 = "0040"	'* Some form of inital status request after startup of client software 
' Reply to "0040" is a "0000" message (discovered using source 2.23)

CONST PW_CMD_0048 = "0048"	'* Archive retrieve request
CONST PW_CMD_0049 = "0049"	'* Archive retrieve reply 

CONST PW_CMD_0057 = "0057"	'* PWSetMeasurementIntervalRequest (mac, ConsumptionInterval, ProductionInterval)
CONST PW_CMD_0058 = "0058"	'* PWClearGroupMacTableRequest 

'* Global decoding variabelen
Dim strSequence		' Sequence number in replies
Dim strValidate		' Response Code (001C is ACK, 001E = NACK)
Dim strMACsrc		' Usualy a circle+ (NC)  
Dim strMACdst		' Circle being queried for info

Dim strSchemaOffset	' Used to adress schema packets in a bigblock

Dim strGainA, strGainB, strOffTot, strOffNoice 'used by Module Calibration

Dim strPwrPulse8, strPwrPulse1, strPwrPulseCum 'as for 0012/0013
Dim strPackedTime	' Timestamp in PlugWise format
Dim strFlag1, strFlag2, strFlag3 'Unknown
Dim strNetworkID	' Long Network Code
Dim strShortID		' Short Network Code
Dim strNodeNcID		' NodeNcID, used for command 0018/0019

Dim strBcInterval, strBcEnable 'Command 4A
Dim strCmd07Unk		'Fill-byte for Cmd007 (purpose unknown)	
Dim strUnknown1
Dim strUnknown2
Dim strConsumptionInterval	'* 1, 3, 5, 15, 60 min
Dim strProductionInterval	'* 1, 3, 5, 15, 60 min

Dim strCmd13Unk1
Dim strCmd13Unk2
Dim strCmd13Unk3

Dim strCmd16Unk1	'LogBuffer addres (reset function  (FFFFFFFF is do nothing))
Dim strCmd16Unk2
Dim strHexTime	'* this is the UTC time in hex as in HH mm ss
Dim strHexDoW	'* Day of Week 


Dim strCmd27Unk
Dim strHWversion
Dim strFWversion
Dim strSWversion

Dim strCirclePower	'If the circle is powered On (01) or Off (00)
Dim strCmd24Freq	'Detected mains frequency ("85" is 50Hz)
Dim strModType 		'Module Type (00=Stick, 01=Cicle+, 02=Circle)
Dim strAsciiTime	'Time as ASCII for Circle+ clock
Dim strAsciiDoW		'Day of Week 
Dim strAsciiDate	'Data as ASCII for Circle+ clock

Dim strCmd3fUnk1
Dim strCmd3fUnk2
Dim strCmd3fUnk3

Dim strLogBuffer	'for reading archived powerusage buffers
Dim strPackedTime1	'history buffer data, PackedTime, Value
Dim strPackedData1
Dim strPackedTime2
Dim strPackedData2
Dim strPackedTime3
Dim strPackedData3
Dim strPackedTime4
Dim strPackedData4

Dim objFSO
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

' Create a regular expression object which will be used for various text-processing
Dim oRegExp
Set oRegExp		= New RegExp
oRegExp.Global		= True
oRegExp.IgnoreCase	= True


Dim strLine, aLine, strSender, strReceive, strReply, strMessage, intLine, strDecoded
Dim intWritePos, intReadPos, intLengthPos, intStartPos, intEndPos 
Dim intPair,tLogTime, bRead
Dim strPath
	
'* Location where the portmon logfiles are stored
strPath = "U:\plugwise\capture"

WScript.Echo vbCrLf & "Starting, Using folder " & strPath

Call Main

WScript.Echo vbCrLf & "Finished, read " & intLine & " captured lines"





'****************************************************************************
'* Main program, get the files to proces, and proces them
Sub Main
Dim strFileName,intPosition,strTemp	
Dim objFolder, colFiles, objFile

	'* Process alle files in the folder
	If objFso.FolderExists(strPath) Then
		Set objFolder= objFso.GetFolder(strPath)
		Set colFiles = objFolder.Files
		For each objFile in colFiles
			strFileName = objFile.Name
			ProcessFile strPath,strFileName
		Next
	Else
		Wscript.Echo "ERROR - Capture folder does not exist"
	End If

End Sub

'****************************************************************************
'* Check and prepare the filenames
Sub ProcessFile(ByVal strPath, ByVal strInFile)
Dim strOutFile, strExt
' filter a *.log file and create a *.txt file, only if a *.txt file not already exists 
	strOutFile = ""		
	strExt = Mid(strInFile,instrrev(strInFile,"."))
	If objFso.FileExists(objFso.Buildpath(strPath,strInFile)) Then
		If LCase(strExt) = LCase(".log") Then
			strOutFile = objFso.GetBasename(strInFile) & ".txt"
			If objFso.FileExists(objFso.Buildpath(strPath,strOutFile)) Then
				Wscript.Echo "Skipping - Outfile " & strOutFile & " already exists"		
			Else
				Wscript.Echo "Filtering - Infile " & strInFile & " to outfile " & strOutFile		
				Call FilterCapture(strPath, strInFile, strOutFile)
			End If
		End If
	End If
End Sub


Sub FilterCapture(ByVal strFolder, ByVal strInFile, ByVal strOutFile)
Dim objTSin, objTSout
'open input and output file

Set objTSin  = objFSO.OpenTextFile(objFso.Buildpath(strFolder,strInFile),ForReading)
Set objTSOut = objFSO.OpenTextFile(objFso.Buildpath(strFolder,strOutFile),ForWriting,True)

'Skip first line
If Not objTSin.AtEndOfStream Then strLine = objTSin.ReadLine 
intLine=1

strSender=""
strReceive=""
bRead=False

Do Until objTSin.AtEndOfStream 
'read and filter data
	strLine = objTSin.ReadLine 
	intLine = intLine+1
	intWritePos = Instr(strLine,"IRP_MJ_WRITE")
	intReadPos = Instr(strLine,"IRP_MJ_READ")
	if intWritePos > 0 Then
	'* WRITE
		strLine= LTrim(Mid(strLine,intWritePos+12))
		intLengthPos=Instr(strLine,"Length")
		If intLengthPos>0 Then
			intStartPos = Instr(intLengthPos,strLine,":")
			strLine= LTrim(Mid(strLine,intStartPos+1))
			strSender=Trim(strLine)
		End If	
	ElseIf intReadPos > 0 Then
	'* READ, try if there is data too (Length followed by a number and the maybe a colon
		strLine= LTrim(Mid(strLine,intReadPos+11))
		intLengthPos=Instr(strLine,"Length")
		If bDebug Then Wscript.echo "DebugReadLine: !"& strLine & "!"
		If intLengthPos>0 Then
			'* Ok, now check if there is a colon too (save-file mode)
			strLine= LTrim(Mid(strLine,intLengthPos+6))
			If bDebug Then Wscript.echo "DebugReadLine: "& strLine
			intStartPos = Instr(strLine,":")
			If intStartPos > 0 Then
				'* Finish reading READ data	(savefile mode)
				strLine= LTrim(Mid(strLine,intStartPos+1))
				If right(strLine,1) = chr(9) Then strLine= Left(strLine,Len(strLine)-1) 
				If bDebug Then Wscript.echo "DebugReadLine: "& strLine
				strReceive = strReceive & strLine
				bRead = FALSE
				If bDebug Then Wscript.echo "DebugDetectRead: "& strReceive
			Else
				bRead = TRUE
				'* skip rest of the line, parse next line in logfile mode
			End If
		End If
	ElseIf bRead Then
	'* Finish reading READ data	(logfile mode)
		bRead = FALSE
		intLengthPos=Instr(strLine,"Length")
		If intLengthPos>0 Then
			'* looks like a valid READ reply
			strLine= LTrim(Mid(strLine,intLengthPos+6))
			If bDebug Then Wscript.echo "DebugReadLine: "& strLine
			intStartPos = Instr(strLine,":")
			If intStartPos > 0 Then
				strLine= LTrim(Mid(strLine,intStartPos+1))
				If bDebug Then Wscript.echo "DebugReadLine: "& strLine
				strReceive = strReceive & TRIM(strLine)
			End If	
		End If
	Else
	'* skip unknown (port command) data
	End If	

	'* Write Results
	If Len(strSender)>0 Then
		intStartPos = instr(1,strSender,"....")  '* actually the start-sequence 0x05, 0x05, 0x03, 0x03  
		intEndPos   = instr(5,strSender,"..")    '* actualy the end-sequence 0x0D, 0x0A (or Cr+Lf)
		If (intEndPos > 0) and (intStartPos > 0) Then
			strSender = Mid(strSender,intStartPos+4,(intEndPos-1)-4) '* strip start and end sequence
		End If
		strmessage = "SEND" & vbTAB & DecodeCommand(strSender)
		Wscript.Echo strMessage
		objTSout.Writeline  strMessage
		strSender=""
	End if	
	Do While Len(strReceive)>0
		intStartPos = instr(strReceive,"....")  '* actually the start-sequence 0x05, 0x05, 0x03, 0x03  
		If intStartPos >0 Then
			'* StartOfMessage found, trow leading garbage away

			strReceive = Mid(strReceive,intStartPos)
			'*	 compensate for stray garbage characters before start-sequence
			Do While Len(strReceive) > 4
				If Mid(strReceive,5,1) = "." Then
				'*	 Wscript.Echo "COMPENS" & vbTab & strReceive 
					strReceive = Mid(strReceive, 2)
				Else
					Exit Do
				End If
			Loop
		Else
		'* no start-sequence, wait for more data
			strReply = ""  '* receive not finished
			Exit Do		
		End If
		
		intEndPos = instr(5,strReceive,"..") '* actualy the end-sequence 0x0D, 0x0A (or Cr+Lf)
		If intEndPos > 0 Then
		'* EndOfMessage found, extract it
			strReply   = Mid(strReceive,1+4,(intEndPos-1)-4)
			strReceive = Mid(strReceive,intEndPos+2)
		Else
		'* no end-sequence, wait for more data
		'* 	Wscript.Echo "WAITEND" & vbTab & strReceive & vbTab & "End sequence not found, WAIT for DATA" 
			strReply = ""  '* receive not finished
			Exit Do		
		End If
		
		If Len(strReply) > 0 Then
			strmessage = "RECV" & vbTAB & DecodeCommand(strReply)
			Wscript.Echo strMessage
			objTSout.Writeline  strMessage
			strReply=""
		End if	
	Loop '* decode received message
Loop '* read from inputfile

'close input file and notify
objTSin.Close
objTSout.Close
End Sub


Function DecodeCommand(ByVal strPWmessage)
Dim StrCommand, strCRCmsg, strCRCcalc, strOption, strResult
Dim lngBufferAddress

	strPWmessage = TRIM(UCase(strPWmessage))
	If Len (strPWmessage) < 8 Then
		DecodeCommand = "Invalid (too short): " & strPWmessage
	Else
		strCRCmsg	= Mid(strPWmessage,Len(strPWmessage)+1-4,4) '* CRC code
		strPWmessage	= Left(strPWmessage,Len(strPWmessage)-4)  '* strip CRC from message

		strCRCcalc	= Right("000" & Hex(GetCRC16(strPWmessage)),4)
		If NOT (strCRCmsg = strCRCcalc) Then
			DecodeCommand = "CRC Error [calc="& strCRCcalc & "; recv=" & strCRCmsg & "]:" & strPWmessage
			Exit Function
'			Wscript.Echo "CRC - Error! MSG " &strCRCmsg & " - Calc "& strCRCcalc  
'		Else	
'			Wscript.Echo "CRC - Succes MSG " &strCRCmsg & " - Calc "& strCRCcalc  
		End If
		strCommand	= Mid(strPWmessage,1,4) '* command code
		strOption	= Mid(strPWmessage,5) 	'* could be  empty

		'* probably build a command decoder here...	

		Select Case strCommand
		Case PW_0000_REP	'* reply-xxxx (general)
			strSequence = Mid(strOption,1,4)
			strValidate = Mid(strOption,5,4)
			Select Case strValidate
			Case "001C"
				'* SUCCESS, Ack received
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate
			Case "001E"
				'* ERROR, NAC received
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate
			Case "00D7"
				'* Reply-0016, MACdst 
				strMACsrc = Mid(strOption,9,16)
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & strMACsrc 
			Case "00D8"
				'* Reply-0017 reply (Switch module On ) 
				strMACsrc = Mid(strOption,9,16)
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & strMACsrc 
			Case "00D9"
				'* Reply-0008 01 reply (Mac-Circle+) 
				strMACsrc = Mid(strOption,9,16)
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & strMACsrc 
			Case "00DD"
				'* Reply-0008 00 reply (Mac-Circle+) 
				strMACsrc = Mid(strOption,9,16)
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & strMACsrc 
			Case "00DE"
				'* Reply-0017 reply (Switch module Off ) 
				strMACsrc = Mid(strOption,9,16)
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & strMACsrc 
			Case "00E5"
				'* Reply-0040, MACdst, number 
			'* Write default reply for 'Unimplemented' 0000-cases 
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & Mid(strOption, 9)
			Case "00E8"
				'* Reply-000F, MACdst, number 
			'* Write default reply for 'Unimplemented' 0000-cases 
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & Mid(strOption, 9)
			Case "00F1"
				'* Reply-004A reply 
				strMACsrc = Mid(strOption,9,16)
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & strMACsrc 
			Case "00F4"
				'* Reply-004E, MACdst 
			'* Write default reply for 'Unimplemented' 0000-cases 
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & Mid(strOption, 9)
			Case "00F8"
				'* Reply-0057, MACdst 
			'* Write default reply for 'Unimplemented' 0000-cases 
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & Mid(strOption, 9)
			Case Else
			'* Write default reply for 'Unimplemented' 0000-cases 
				DecodeCommand = strCommand & " " & strSequence & " " & strValidate & " " & Mid(strOption, 9)
				'* Generic 0-responce 
			End Select
		Case "0001"			'* Start creat new network (erase old PAN)
							'* also seems to erase the old network
			DecodeCommand = strCommand & " " & strOption & vbTab & "*** -=-Network Init-=- ***"
		Case "0002"	'* Pair reply (suggested netword ID))
			'RECV	0002 0019 0F FFFFFFFFFFFFFFFF 520D6F0000B1B64B FFFFFFFFFFFFFFFF 520D6F0000B1B64B 5D52 01
			strSequence = Mid(strOption,1,4)	: strResult = strSequence
			strFlag1	= Mid(strOption,5,2)	: strResult = strResult & " " & strFlag1
			strMACsrc	= Mid(strOption,7,16)	: strResult = strResult & " " & strMACsrc 
			strFlag2	= Mid(strOption,23,16)	: strResult = strResult & " " & strFlag2 
			strMACdst	= Mid(strOption,39,16)	: strResult = strResult & " " & strMACdst 
			strNetworkID = Mid(strOption,55,16)	: strResult = strResult & " " & strNetworkID
			strShortID	= Mid(strOption,71,4)	: strResult = strResult & " " & strShortID
			strFlag3	= Mid(strOption,75,2)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult & vbTab & "*** -=-Network PAN sugestion-=- ***"
		Case "0003"	'* Pair unpair reply
			strSequence = Mid(strOption,1,4)	: strResult = strSequence
			strValidate	= Mid(strOption,5,4)	: strResult = strResult & " " & strValidate 
			DecodeCommand = strCommand & " " & strResult & vbTab & "*** -=-Pair/Unpair-=- ***"
		Case "0004"	'* Accept suggested PAN , reply to Circle+
			strValidate	= Mid(strOption,5,4)	: strResult = strResult & " " & strValidate 
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strMACdst	= Mid(strOption,21,16)	: strResult = strResult & " " & strMACdst 
			DecodeCommand = strCommand & " " & strResult & vbTab & "*** -=-Pair-=- ***"
		Case "0005"	'* Acknowledge pair reply
			strSequence	= Mid(strOption,1,4)	: strResult = strSequence
			strValidate	= Mid(strOption,5,4)	: strResult = strResult & " " & strValidate 
			DecodeCommand = strCommand & " " & strResult & vbTab & "*** -=-Pair-=- ***"
		Case "0006"	'* Unconfigured Node advertisement (receive, with dummy seq number)
			strSequence	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			DecodeCommand = strCommand & " " & strResult
		Case "0007"	'* Unconfigured Node - Send (Return2Node)
			strCmd07Unk 	= Mid(strOption,1,2)	: strResult = strCmd07Unk
			strMACsrc 		= Mid(strOption,3,16)	: strResult = strResult & " " & strMACsrc 
			DecodeCommand = strCommand & " " & strResult
		Case "0008"			'* Sort of reset/initialise
			DecodeCommand = strCommand & " " & strOption & vbTab & "*** -Restart!- ***"
		Case "0009"	'* request to removeg a module
			'* expect a 0000 responce with functioncode 00F2
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			strFlag3	= Mid(strOption,17,4)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult

		Case PW_000A_CMD	'* "000A -  Reset/initialize command
			DecodeCommand = strCommand & " " & strOption & vbTab & "*** Initialise ***"
		Case "000B"	'* Command (after firmware update)
			'* expect a 0003 response with functioncode 00D0
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			strFlag3	= Mid(strOption,17,4)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult
		Case "000C"	'* Query for stored firmware version (used with firmware update)
			'* expect a 0010 responce
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand = strCommand & " " & strResult & " " & mid(strOption,21)
		Case "000D"	'* Query , Request a Group ID from circle+
			'* expect a 000E responce
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			strFlag3	= Mid(strOption,17,4)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult & " " & mid(strOption,21)
		Case "000E"	'* Reply, GroupID for new group (8 hex-digits only)
			strSequence	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strMACsrc 
			strUnknown1	= Mid(strOption,21,8)	: strResult = strResult & " " & strUnknown1 'is 9group ID?
			DecodeCommand = strCommand & " " & strResult & " " & mid(strOption,29)
		Case "000F"	'*	Command, clear association or reset module
					'	send before network reconfiguration
			strSequence	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strMACsrc 
			strUnknown1	= Mid(strOption,21,4)	: strResult = strResult & " " & strUnknown1 'some sort of command flags
			DecodeCommand = strCommand & " " & strResult & " " & mid(strOption,29)
		Case "0010"	'* Reply with stored firmware version
			strSequence	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strMACsrc 
			strFWversion = Mid(strOption,21,8)	: strResult = strResult & " " & strFWversion 'stored firmware version
			DecodeCommand = strCommand & " " & strResult & " " & mid(strOption,29)
		Case "0011"	'* Reply-000A (1e)
			strSequence	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strFlag1	= Mid(strOption,21,2)	: strResult = strResult & " " & strFlag1
			strFlag2	= Mid(strOption,23,2)	: strResult = strResult & " " & strFlag2
			strNetworkID	= Mid(strOption,25,16)	: strResult = strResult & " " & strNetworkID
			strShortID	= Mid(strOption,41,4)	: strResult = strResult & " " & strShortID
			strFlag3	= Mid(strOption,45,2)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult & " - " & Mid(strOption,47)
		Case "0012"	'* RSP-Actual Powerusage - Request
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand = strCommand & " " & strResult
		Case "0013"	'* RSP-Actual Powerusage - Reply
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strPwrPulse8	= Mid(strOption,21,4)	: strResult = strResult & " " & strPwrPulse8 
			strPwrPulse1	= Mid(strOption,25,4)	: strResult = strResult & " " & strPwrPulse1 
			strPwrPulseCum	= Mid(strOption,29,8)	: strResult = strResult & " " & strPwrPulseCum
			strCmd13Unk1	= Mid(strOption,37,4)	: strResult = strResult & " " & strCmd13Unk1 
			strCmd13Unk2	= Mid(strOption,41,4)	: strResult = strResult & " " & strCmd13Unk2			
			strCmd13Unk3	= Mid(strOption,45,4)	: strResult = strResult & " " & strCmd13Unk3 
			DecodeCommand = strCommand & " " & strResult  & " {" & Mid(strOption,49) & "}" 
			'  & vbTab & "{" & HexToFloat(strPwrPulseCum) & "}"
		Case "0016"	'* Set or Sync Clock (Packed DateTime)
			' also used for switching a circle on or off (different format)
			strMACsrc	= Mid(strOption,1,16)			: strResult = strMACsrc
 			strPackedTime	= Mid(strOption,17,8)	: strResult = strResult & " " & strPackedTime 
			strCmd16Unk1	= Mid(strOption,25,8)	: strResult = strResult & " " & strCmd16Unk1
			strCmd16Unk2	= Mid(strOption,29,4)	: strResult = strResult & " " & strCmd16Unk2
			strHexTime	= Mid(strOption,33,6)	: strResult = strResult & " " & strHexTime
			strHexDoW		= Mid(strOption,39,2)	: strResult = strResult & " " & strHexDoW
			DecodeCommand = strCommand & " " & strResult & vbTab & "[" & DecodePWpackeddate(strPackedTime) &_
				"] - [" & DecodeHexTime(strHexDoW, strHexTime) & "]"
		Case "0017"	'* Switch Module on/off - Request
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			strFlag3	= Mid(strOption,17,2)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult
		Case "0018"	'* CMD-RoleCall - Request
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			strNodeNcID	= Mid(strOption,17,2)	: strResult = strResult & " " & strNodeNcID
			DecodeCommand = strCommand & " " & strResult
		Case "0019"	'* RSP-RoleCall - Reply
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strMACdst	= Mid(strOption,21,16)	: strResult = strResult & " " & strMACdst 
			strNodeNcID	= Mid(strOption,37,2)	: strResult = strResult & " " & strNodeNcID
			DecodeCommand = strCommand & " " & strResult
		Case "001C"	'* remove Module from network - Request
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			strMACdst		= Mid(strOption,17,16)	: strResult = strResult & " " & strMACdst 
			DecodeCommand = strCommand & " " & strResult
		Case "001D"	'* Remove Module from network - Responce
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strMACdst	= Mid(strOption,21,16)	: strResult = strResult & " " & strMACdst 
			strFlag3	= Mid(strOption,37,2)	: strResult = strResult & " " & strFlag3
			DecodeCommand = strCommand & " " & strResult
		Case "0023"	'* Device Information - Request
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand = strCommand & " " & strResult
		Case "0024"	'* Device Information - Reply
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strPackedTime   = Mid(strOption,21,8)	: strResult = strResult & " " & strPackedTime 
			strLogBuffer	= Mid(strOption,29,8)	: strResult = strResult & " " & strLogBuffer 
			strCirclePower	= Mid(strOption,37,2)	: strResult = strResult & " " & strCirclePower 
			strCmd24Freq	= Mid(strOption,39,2)	: strResult = strResult & " " & strCmd24Freq 
			strHWversion	= Mid(strOption,41,12)	: strResult = strResult & " " & strHWversion 
			strFWversion	= Mid(strOption,53,8)	: strResult = strResult & " " & strFWversion  'running firmware version 
			strModType	= Mid(strOption,61,2)	: strResult = strResult & " " & strModType
			If Clng("&H" & strPackedTime) > 0 Then
				DecodeCommand = strCommand & " " & strResult & vbTab & "[" & DecodePWpackeddate(strPackedTime) & "], Buffer:" & DecodeBufferAddress(strLogBuffer)
			Else
				DecodeCommand = strCommand & " " & strResult
			End If
		Case "0026"	'* Power Calibration - Request
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand = strCommand & " " & strResult
		Case "0027"	'* Power Calibration - Reply
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strGainA	= Mid(strOption,21,8)	: strResult = strResult & " " & strGainA 
			strGainB	= Mid(strOption,29,8)	: strResult = strResult & " " & strGainB 
			strOffTot	= Mid(strOption,37,8)	: strResult = strResult & " " & strOffTot 
			strOffNoice	= Mid(strOption,45,8)	: strResult = strResult & " " & strOffNoice 
			strCmd27Unk	= Mid(strOption,53,2)	: strResult = strResult & " " & strCmd27Unk
			DecodeCommand = strCommand & " " & strResult
		Case "0028"	'* Set or Sync Date, "other date-format"
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			strAsciiTime	= Mid(strOption,17,6)	: strResult = strResult & " " & strAsciiTime 
			strAsciiDoW		= Mid(strOption,23,2)	: strResult = strResult & " " & strAsciiDoW 
			strAsciiDate	= Mid(strOption,25,6)	: strResult = strResult & " " & strAsciiDate	' e.g. 050511 is 5 mei 2011 
			DecodeCommand	= strCommand & " " & strResult & vbTab & "[" & DecodeAsciiDateTime(strAsciiDate, strAsciiDoW, strAsciiTime) & "]"
		Case "0029"	'* CMD0029 - Request
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc
			DecodeCommand	= strCommand & " " & strResult
		Case "003A"	'* CMD0029 - Reply  
			' returns the data-time (clockdata set by CMD0028 (6 chars-time, 2-chars unknown, 6-chars date)
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 		= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strAsciiTime		= Mid(strOption,21,6)	: strResult = strResult & " " & strAsciiTime 
			strAsciiDoW		= Mid(strOption,27,2)	: strResult = strResult & " " & strAsciiDoW 
			strAsciiDate		= Mid(strOption,29,6)	: strResult = strResult & " " & strAsciiDate
			DecodeCommand = strCommand & " " & strResult & vbTab & "[" & DecodeAsciiDateTime(strAsciiDate, strAsciiDoW, strAsciiTime) & "]"
		Case "003B"	'* CMD003B - Schema Definition Data
			strSchemaOffset	= Mid(strOption,1,4)	: strResult = strResult & " " & strSchemaOffset 		
			strPackedData1  = Mid(strOption,5,8)	: strResult = strResult & " " & strPackedData1 
			strPackedData2  = Mid(strOption,13,8)	: strResult = strResult & " " & strPackedData2 
			strPackedData3  = Mid(strOption,21,8)	: strResult = strResult & " " & strPackedData3 
			strPackedData4  = Mid(strOption,29,8)	: strResult = strResult & " " & strPackedData4 
			DecodeCommand = strCommand & " " & strResult
		Case "003C"	'* CMD003C - Schema Validation Request
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc
			strSchemaOffset	= Mid(strOption,17,2)	: strResult = strResult & " " & strSchemaOffset 		
			DecodeCommand = strCommand & " " & strResult
		Case "003D"	'* CMD003D - Schema Validation Reply
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strSchemaOffset	= Mid(strOption,21,2)	: strResult = strResult & " " & strSchemaOffset 		
			DecodeCommand = strCommand & " " & strResult
		Case "003E"	'* CMD003E - Request
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand = strCommand & " " & strResult
		Case "003F"	'* CMD003F - Reply to 003E
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 		= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strHexTime	= Mid(strOption,21,6)	: strResult = strResult & " " & strHexTime 
			strHexDoW	= Mid(strOption,27,2)	: strResult = strResult & " " & strHexDoW 
			strCmd3fUnk1	= Mid(strOption,29,6)	: strResult = strResult & " " & strCmd3fUnk1 
			DecodeCommand = strCommand & " " & strResult & vbTab & "[" & DecodeHexTime(strHexDoW, strHexTime) & "]" 
		Case PW_CMD_0040	'* Command, configure schema settings (expect 0000 reply with 00E7 flag)
			strMACsrc 	= Mid(strOption,1,16)	: strResult = strMACsrc 
			strFlag1	= Mid(strOption,17,2)	: strResult = strResult & " " & strFlag1
			strFlag2	= Mid(strOption,19,2)	: strResult = strResult & " " & strFlag2
			DecodeCommand = strCommand & " " & strResult & " {" & Mid(strOption,23) & "}" 
		Case "0047"	'* Set/Create Group 
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			strBcInterval	= Mid(strOption,17,2)	: strResult = strResult & " " & strBcInterval 
			strBcEnable	= Mid(strOption,19,2)	: strResult = strResult & " " & strBcEnable
			DecodeCommand = strCommand & " " & strMACsrc & " " & strBcInterval & " " & strBcEnable & " " & mid(strOption,21)
		Case "0048"	'* Read LogBuffer, Request
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			strLogBuffer	= Mid(strOption,17,8)	: strResult = strResult & " " & strLogBuffer & vbTab & "Buffer:" & DecodeBufferAddress(strLogBuffer) 
			DecodeCommand = strCommand & " " & strResult
		Case "0049"	'* Read LogBuffer, Reply
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 	= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strPackedTime1  = Mid(strOption,21,8)	: strResult = strResult & " " & strPackedTime1
			strPackedData1  = Mid(strOption,29,8)	: strResult = strResult & " " & strPackedData1 
			strPackedTime2  = Mid(strOption,37,8)	: strResult = strResult & " " & strPackedTime2
			strPackedData2  = Mid(strOption,45,8)	: strResult = strResult & " " & strPackedData2 
			strPackedTime3  = Mid(strOption,53,8)	: strResult = strResult & " " & strPackedTime3
			strPackedData3  = Mid(strOption,61,8)	: strResult = strResult & " " & strPackedData3 
			strPackedTime4  = Mid(strOption,69,8)	: strResult = strResult & " " & strPackedTime4
			strPackedData4  = Mid(strOption,77,8)	: strResult = strResult & " " & strPackedData4 
			strLogBuffer	= Mid(strOption,85,8)	: strResult = strResult & " " & strLogBuffer 
			DecodeCommand = strCommand & " " & strResult & vbTab & "[" & DecodePWpackeddate(strPackedTime1) & "], Buffer:" & DecodeBufferAddress(strLogBuffer)
		Case "004A"	'* Set Broadcast interval
			'* See http://www.domoticaforum.eu/viewtopic.php?f=39&t=4319&start=45#p47271
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			strBcInterval	= Mid(strOption,17,2)	: strResult = strResult & " " & strBcInterval 
			strBcEnable		= Mid(strOption,19,2)	: strResult = strResult & " " & strBcEnable
			DecodeCommand = strCommand & " " & strResult & "  {" & Mid(strOption,21) & "}"
		Case "0056"	'* Set/Create Group (reply from 0047) 
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 		= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strBcInterval	= Mid(strOption,21,2)	: strResult = strResult & " " & strBcInterval 
			strBcEnable		= Mid(strOption,23,2)	: strResult = strResult & " " & strBcEnable
			DecodeCommand = strCommand & " " & strResult & "  {" & Mid(strOption,25) & "}"
		Case PW_CMD_0057	'* CMD0057 - Set Measurement Interval
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			strConsumptionInterval	= Mid(strOption,17,4)	: strResult = strResult & " " & strConsumptionInterval 
			strProductionInterval	= Mid(strOption,21,4)	: strResult = strResult & " " & strProductionInterval
			DecodeCommand = strCommand & " " & strResult & " {" & Mid(strOption,25) & "}" 
		Case PW_CMD_0058	'* CMD0058 - ClearGroupMacTable
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand	= strCommand & " " & strResult & " {" & Mid(strOption,17) & "}"
		Case "005F"	'* CMD005F - Request (just after initialize/add module), is it a reset buffer?
			strMACsrc 		= Mid(strOption,1,16)	: strResult = strMACsrc 
			DecodeCommand	= strCommand & " " & strResult & "  {" & Mid(strOption,17) & "}"
		Case "0060"	'* CMD005F - Reply 
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence
			strMACsrc 		= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			strMACdst 		= Mid(strOption,21,16)	: strResult = strResult & " " & strMACdst '(unknown value, maybe -2 (64-bit) in hex 
			DecodeCommand = strCommand & " " & strResult & " {" & Mid(strOption,23) & "}"
		Case "0061"	'* Advertisement from unconfigured modules
			strSequence 	= Mid(strOption,1,4)	: strResult = strSequence 'dummy seq-number
			strMACsrc 		= Mid(strOption,5,16)	: strResult = strResult & " " & strMACsrc 
			DecodeCommand = strCommand & " " & strResult & "{" & Mid(strOption,21) & "}"
		Case Else
		'*	Unknown Command
			DecodeCommand = strCommand  & " " & strOption
		End Select
		
	End If
End Function


Function DecodePWpackeddate(ByVal strPWdate)
Dim intYear, intMonth, lngMinutesMonth
Dim intDay, intHour, intMinute
	If strPWdate <> "FFFFFFFF" Then
		intYear  = CInt("&H" & Mid(strPWdate,1,2)) + 2000
		intMonth = CInt("&H" & Mid(strPWdate,3,2))
		lngMinutesMonth = CLng("&H" & Mid(strPWdate,5,4))
		
		intDay    = CInt(lngMinutesMonth  \  (60*24) ) +1 
		intMinute = CInt(lngMinutesMonth Mod (60*24) )

		intHour   = CInt(intMinute  \  60)
		intMinute = CInt(intMinute Mod 60) 

		DecodePWpackeddate = (intYear) & "-" & Right(("00" & intMonth),2) &_
		"-" & Right(("00" & intDay),2) & " " & Right(("00" & intHour),2) &_
		":" & Right(("00" & intMinute),2) & " UTC"
	Else
		DecodePWpackeddate = "0000-00-00 00:00 UTC"  '* no date
	End If
End Function


Function DayOfWeek(ByVal intDoW)
	Select Case (intDoW)
	Case 0
		DayOfWeek="UnInitialized(0)"
	Case 1
		DayOfWeek="Monday"
	Case 2
		DayOfWeek="Tuesday"
	Case 3
		DayOfWeek="Wednesday"
	Case 4
		DayOfWeek="Thursday"
	Case 5
		DayOfWeek="Friday"
	Case 6
		DayOfWeek="Saturday"
	Case 7
		DayOfWeek="Sunday"
	Case Else
		DayOfWeek="Unknown(" & intDoW & ")"
	End Select
End Function

Function DecodeHexTime(ByVal strHexDoW, Byval strHexTime)
Dim intDoW
Dim intHour, intMinute, intSecond

	intDoW    = Cint("&H" & strHexDoW)
	intHour   = Cint("&H" & MID(strHexTime,1,2))
	intMinute = Cint("&H" & MID(strHexTime,3,2))
	intSecond = Cint("&H" & MID(strHexTime,5,2))
		
	DecodeHexTime = Right(("00" & intHour),2) & ":" & Right(("00" &_
	intMinute),2) & ":" & Right(("00" & intSecond),2) &_
	" UTC (" & DayOfWeek(intDoW) & ")"
End Function


Function DecodeAsciiDateTime(ByVal strAsciiDate, ByVal strAsciiDoW, Byval strAsciiTime)
Dim intYear, intMonth, intDay
Dim intDoW
Dim intHour, intMinute, intSecond

	intYear   = Cint(MID(strAsciiDate,5,2)) + 2000
	intMonth  = Cint(MID(strAsciiDate,3,2))
	intDay    = Cint(MID(strAsciiDate,1,2))
	intDoW    = Cint(strAsciiDoW)
	intHour   = Cint(MID(strAsciiTime,5,2))
	intMinute = Cint(MID(strAsciiTime,3,2))
	intSecond = Cint(MID(strAsciiTime,1,2))
		
	DecodeAsciiDateTime = (intYear) & "-" & Right(("00" & intMonth),2) & "-" &_
	Right(("00" & intDay),2) & " " & Right(("00" & intHour),2) & ":" &_
	Right(("00" & intMinute),2) & ":" & Right(("00" & intSecond),2) &_
	" UTC (" & DayOfWeek(intDoW) & ")"
End Function


Function DecodeBufferAddress(ByVal strLogBuffer)
'* there might be a problem, as a VBscript long seems only 32 bit, not 64 bit
Dim lngBufferAddress
Dim intBufferNumber
Const BufferOffset = "&H044000"

	on error resume next
	lngBufferAddress = Clng("&H" & strLogBuffer)
	If Err.Number <>0 Then
		Wscript.Echo "Error: Code:" & Err.Number & " converting LogAddress (" & strLogBuffer & ") to integer" 
		Wscript.Echo "Error: Desc:" & Err.Description 
		lngBufferAddress = 0
		Wscript.Sleep(10000)
		Err.Clear
	End If
	on error goto 0
 	intBufferNumber = CLng((lngBufferAddress - BufferOffset) \ 8)
'	intBufferNumber = CLng((lngBufferAddress - 278528) \ 8)
	If intBufferNumber < 0 Then
	'	Usualy this is with buffer = 00000000 (no buffer)
	'	Wscript.Echo "Error: Buffernumber out of range (" & intBufferNumber & ")" 
		intBufferNumber =  0
	End If
	DecodeBufferAddress = Cint(intBufferNumber)
End Function 


'******************************************************************
'* See http://www.domoticaforum.eu/viewtopic.php?f=39&t=5813
'* (richard naninck), convert HEX 2 Float for PlugWise calculations
'******************************************************************
'Label    HexValue    FloatValue
'GainA:   3F802571 -  1.00114262104
'GainB:   B5E5DEDB -  -1.712668904474E-6
'OffTot:  3D014486 -  0.031559489667
'OffRuis: 00000000 -  5.877471754111E-39
'******************************************************************
Function HexToFloat(ByRef strHexData)
Dim intBytes(4)
Dim intExp
Dim intDec
Dim fltVal
Dim i

   For i = 0 To 3
	  intBytes(i) = CLng("&H" & Mid(strHexData, (i * 2) + 1, 2))

   Next
   
   'S Exp      Mantissa
   'x xxxxxxxx xxxxxxxxxxxxxxxxxxxxxxx
   'Low 7 bits of intBytes(0) shifted 1 bit left + 8th bit of intBytes(1) shifted 7 bits right
   intExp = ((intBytes(0) And 127) * 2^1)  + ((intBytes(1) And 128) / 2^7) - 127
   intDec = Round(((((intBytes(1) And 127) * 2^16) + (intBytes(2) * 2^8) + intBytes(3)) / 2^23), 12) + 1
   fltVal = intDec * 2^intExp
   
   If (intBytes(0) And 128) = 128 Then
	  fltVal = fltVal * -1
   End If
   
   HexToFloat = fltVal
End Function


'******************************************************************
'* See http://domoticaforum.eu/viewtopic.php?f=39&t=5803#p45086
'* (bwired), CRC calculation for PlugWise Protocol
'******************************************************************
'* used to calculate the CRC for a plugwise message
'* See also http://domoticaforum.eu/viewtopic.php?f=39&t=4319
'******************************************************************
Function GetCRC16(ByVal strInMessage)
    Dim i , j 
    Dim curVal
    Dim CRC
    Dim Poly

    ' Poly: X16+X12+X5+1
    Poly = 4129
    ' Init CRC value
    CRC = 0
       
    ' For each value
    For i = 1 To Len(strInMessage)
        ' Get Value from the char
        curVal = Asc(Mid(strInMessage, i, 1))
        ' XOR it
        CRC = CRC Xor ((curVal And 255) * 256)
        ' Run trough each bit
        For j = 0 To 7
            If (CRC And 32768) = 32768 Then
                CRC = (CRC * 2) And 65535       ' Shift left
                CRC = (CRC Xor Poly)            ' Sum Poly
            Else
                CRC = (CRC * 2) And 65535       ' Shift left
            End If
        Next
    Next
    GetCRC16 = CRC
End Function
