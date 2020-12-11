'
'	* ApolloSynch: Windows EUD (Windows Me onwards) to Google Drive file synch automation	*
'	**********************************************************************************************
'	* Description: 	standalone script to copy data from Windows terminals to 		*
'	*			Google drives. This can send to any machine and run 			*
'	*			Required internet access and termainal Admin rights.			*
'	*			(Rclone binary copy takes time)					*
'	* Author:		VDS                               					*
'	* Version:		4.0                                     				*
'	* Date:		19/08/2019                                     			*                                  			 
'	**********************************************************************************************

CONST ForReading = 1
CONST ForWriting = 2

'To check same machine has multiple users
DIM user_count
'Array to store Valid user accounts
DIM user_array
user_array = ARRAY()

'Total File size varables
DIM TotalGoodSmallFileSize
DIM TotalGoodLargeFileSize
DIM TotalUnproductiveSmallFileSize
DIM TotalUnproductiveLargeFileSize

'Total C Drive free space
DIM CDriveFreeSpace
	
FileTypeArray = ARRAY(".doc",".dot",".wbk",".docx",".docm",".dotx",".dotm",".docb",".xls",".xlt",".xlm",".xlsx",".xlsm",".xltx",".xltm",".xlsb",".xla",".xlam",".xll",".xlw",".ppt",".pot",".pps",".pptx",".pptm",".potx",".potm",".ppam",".ppsx",".ppsm",".sldx",".sldm",".adn",".accdb",".accdr",".accdt",".accda",".mdw",".accde",".mam",".maq",".mar",".mat",".maf",".laccdb",".ade",".adp",".mdb",".cdb",".mda",".mdn",".mdt",".mdf",".mde",".ldb",".pub",".xps",".odt",".fodt",".ods",".fods",".odp",".fodp",".odg",".fodg",".odf",".rtf",".txt",".tex",".ai",".indd",".psd",".pdf",".pmd",".fm",".abf",".afm",".fb",".pfm",".afm",".act",".ase",".psb",".pdd",".prc",".as",".jsfl",".cel",".ses",".sol",".ppj",".prproj",".flv",".jpeg",".jpg",".png",".gif",".bmp",".tiff",".eps",".raw",".iso",".tar",".tar.gz",".zip",".zipx",".rar")
' 	'MS Office
' 	'---------
' 	'Word
' 	".doc",' – Legacy Word document; Microsoft Office refers to them as "Microsoft Word 97 – 2003 Document"
' 	".dot",' – Legacy Word templates; officially designated "Microsoft Word 97 – 2003 Template"
' 	".wbk",' – Legacy Word document backup; referred as "Microsoft Word Backup Document"
' 	".docx",' – Word document
' 	".docm",' – Word macro-enabled document; same as docx, but may contain macros and scripts
' 	".dotx",' – Word template
' 	".dotm",' – Word macro-enabled template; same as dotx, but may contain macros and scripts
' 	".docb",' – Word binary document introduced in Microsoft Office 2007
' 	'Excel
' 	".xls",' – Legacy Excel worksheets; officially designated "Microsoft Excel 97-2003 Worksheet"
' 	".xlt",' – Legacy Excel templates; officially designated "Microsoft Excel 97-2003 Template"
' 	".xlm",' – Legacy Excel macro
' 	'OOXML
' 	".xlsx",' – Excel workbook
' 	".xlsm",' – Excel macro-enabled workbook; same as xlsx but may contain macros and scripts
' 	".xltx",' – Excel template
' 	".xltm",' – Excel macro-enabled template; same as xltx but may contain macros and scripts
' 	".xlsb",' – Excel binary worksheet (BIFF12)
' 	".xla",' – Excel add-in or macro
' 	".xlam",' – Excel add-in
' 	".xll",' – Excel XLL add-in; a form of DLL-based add-in[1]
' 	".xlw",' – Excel work space; previously known as "workbook"
' 	'PowerPoint
' 	".ppt",' – Legacy PowerPoint presentation
' 	".pot",' – Legacy PowerPoint template
' 	".pps",' – Legacy PowerPoint slideshow
' 	".pptx",' – PowerPoint presentation
' 	".pptm",' – PowerPoint macro-enabled presentation
' 	".potx",' – PowerPoint template
' 	".potm",' – PowerPoint macro-enabled template
' 	".ppam",' – PowerPoint add-in
' 	".ppsx",' – PowerPoint slideshow
' 	".ppsm",' – PowerPoint macro-enabled slideshow
' 	".sldx",' – PowerPoint slide
' 	".sldm",' – PowerPoint macro-enabled slide
' 	'Access
' 	".adn",' –	Access Blank Project Template
' 	".accdb",' –	Access Database (2007 and later)
' 	".accdr",' –	Access Database Runtime (2007 and later)
' 	".accdt",' –	Access Database Template (2007 and later)
' 	".accda",' –	Access Add-In (2007 and later)
' 	".mdw",' –	Access Workgroup, database for user-level security.
' 	".accde",' –	Protected Access Database, with compiled VBA and macros (2007 and later)
' 	".mam",' –	Windows Shortcut: Access Macro
' 	".maq",' –	Windows Shortcut: Access Query
' 	".mar",' –	Windows Shortcut: Access Report
' 	".mat",' –	Windows Shortcut: Access Table
' 	".maf",' –	Windows Shortcut: Access Form
' 	".laccdb",' –	Access lock files (associated with .accdb)
' 	".ade",' –	Protected Access Data Project (not supported in 2013)
' 	".adp",' –	Access Data Project (not supported in 2013)
' 	".mdb",' –	Access Database (2003 and earlier)
' 	".cdb",' –	Access Database (Pocket Access for Windows CE)
' 	".mda",' –	Access Database, used for addins (Access 2, 95, 97),
' 	".mdn",' –	Access Blank Database Template (2003 and earlier)
' 	".mdt",' –	Access Add-in Data (2003 and earlier)
' 	".mdf",' –	Access (SQL Server) detached database (2000)
' 	".mde",' –	Protected Access Database, with compiled VBA and macros (2003 and earlier)
' 	".ldb",' –	Access lock files (associated with .mdb)
' 	'Publisher
' 	".pub",' – a Microsoft Publisher publication
' 	'XPS Document
' 	".xps",' – a XML-based document format used for printing (on Windows Vista and later) and preserving documents.
' 	'Open Docs
' 	'---------
' 	".odt",' 	–	 word processing (text) documents
' 	".fodt",' 	–	 word processing (text) documents
' 	".ods",' 	–	spreadsheets	
' 	".fods",' 	–	spreadsheets
' 	".odp",'		–	presentations	
' 	".fodp",' 	–	presentations
' 	".odg",' 	–	graphics
' 	".fodg",' 	–	graphics
' 	".odf",' 	–	formulae, mathematical equations
' 	'Other Text
' 	-----------
' 	".rtf",' 	–	Rich text format
' 	".txt",'	–	Text file
' 	".tex",'	– LaTeX Source Document.
' 	'ADOBE
' 	'-----
' 	".ai",' – Adobe Illustrator
' 	".indd",' – Adobe InDesign
' 	".psd",' – Adobe Photoshop
' 	".pdf",' – Adobe Acrobat or Adobe Reader
' 	".pmd",' – Adobe PageMaker
' 	".fm",' – Adobe FrameMaker
' 	".abf",' – Adobe Binary Screen Font
' 	".afm",' – Adobe Font Metrics
' 	".fb",' – Printer Font Binary – Adobe
' 	".pfm",' – Printer Font Metrics – Adobe
' 	".afm",' – Adobe Font Metrics
' 	".act",' – Adobe Color Table. Contains a raw color palette and consists of 256 24-bit RGB colour values.
' 	".ase",' – Adobe Swatch Exchange. Used by Adobe Photoshop, Illustrator, and InDesign.
' 	".psb",' – Adobe Photoshop Big image file (for large files)
' 	".pdd",' – Adobe Photoshop Drawing
' 	".prc",' – Adobe PRC (embedded in PDF files)
' 	".as",' – Adobe Flash ActionScript File
' 	".jsfl",' – Adobe JavaScript language
' 	".cel",' – Adobe Audition loop file (Cool Edit Loop)
' 	".ses",' – Adobe Audition multitrack session file
' 	".sol",' – Adobe Flash shared object ("Flash cookie")
' 	".ppj",' - Adobe Premiere Pro video editing file
' 	".prproj",' – Adobe Premiere Pro video editing file
' 	".flv",' – Flash Video File.
' 	'Images
' 	'-----
' 	".jpeg",' - Joint Photographic Experts Group
' 	".jpg",' - Joint Photographic Experts Group
' 	".png",' - Portable Network Graphics
' 	".gif",' - Graphics Interchange Format
' 	".bmp",' - Bit Map format
' 	".tiff",' - Tagged Image File
' 	".eps",' - Encapsulated Postscript
' 	".raw",' - Raw Image Formats
' 	'Compression/Archiving
' 	'---------------------
' 	".iso",' 	- Image file
' 	".tar",' 	- Linux Image file
' 	".tar.gz",'	- Linux Image rar
' 	".zip",' 	- ZIP file
' 	".zipx",' 	- ZIP file
' 	".rar",' 	- WunRAR file

ON ERROR resume NEXT

'Command promt initiation
SET WshShell = WScript.CREATEOBJECT( "WScript.Shell" )
SET objFSO = CREATEOBJECT("Scripting.FileSystemObject")

' Get running agent
strName = WshShell.ExpandEnvironmentStrings( "%USERNAME%" )

' Get profile Path
profilePath = WshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )

' Get Computer Name
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

'Initiate Domain Name
strDomain = "Workgroup"
FQDN = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Domain")

IF FQDN <> "" THEN
	strDomain = FQDN
END IF

IF WScript.Arguments.Count > 0 AND Wscript.Arguments(0) <> "" THEN
	strDomain = Wscript.Arguments(0) & "-" & strDomain
END IF

' Env Variable management
SET objEnvUser = WshShell.Environment("User")
SET objEnvProcess = WshShell.Environment("Process")
SET objEnvSystem = WshShell.Environment("System")

'!!!!!!!!!!!!!!!!!!!WARNING!!!!!!!!!!!!!!!!!This must run before defining Log files otherwise it may retrun errors!!!!!!!!!!!!!!!!!!!
' Check RClone available. Install IF not
CALL CheckRClone(objFSO)

IF WScript.Arguments.Count <= 0 THEN
	Wscript.echo "No Country Code passed in. Terminationg the script."
	Wscript.Quit
END IF

' Log file whene C:/rclone available. !!!!!This must be defined before any logging recorded below!!!!!
SET myLog = objFSO.OpenTextFile("C:\rclone\ApolloSynch.log", ForWriting, True)
SET fileLog = objFSO.OpenTextFile("C:\rclone\ApolloSynchFiles.log", ForWriting, True)
SET statusLog = objFSO.OpenTextFile("C:\rclone\statusLog.csv", ForWriting, True)

myLog.WriteLine "Cloning Starts..."

'Check Force Synch param - When passed the script will ignore the previous synch statuses
forceSynch = 0
IF WScript.Arguments.Count > 0 THEN
	forceSynch = Wscript.Arguments(1)
	IF forceSynch = 1 THEN
		myLog.WriteLine "Force Synch enabled. Re-do everything!!"
	END IF
END IF

' System user check
SET SystemUser = ""
SET oRE = New RegExp
oRE.Pattern = "\$$"
IF (oRE.Test(strName)) THEN
	SystemUser = strName
	myLog.WriteLine "Running as System: " & strName
	myLog.WriteLine "Profile Path: " & profilePath
ELSE
	oRE.Pattern = "\%$"
	IF (oRE.Test(strName)) THEN
		Set objSysInfo = CreateObject( "WinNTSystemInfo" )
		strName = objSysInfo.ComputerName
		profilePath = "C:\Windows\System32\config\systemprofile"
		myLog.WriteLine "Running as System: " & strName
		myLog.WriteLine "Profile Path: " & profilePath
	ELSE
		myLog.WriteLine "Running as user: " & strName
		myLog.WriteLine "Profile Path: " & profilePath
	END IF
END IF

'Main user Dir check to find avaiable machine useres
SET objFolder = objFSO.GetFolder("C:\Users")
SET colSubfolders = objFolder.Subfolders
FOR EACH objSubfolder in colSubfolders
	' Standard user accounts will be ignored
	IF objSubfolder.Name <> SystemUser AND LCase(objSubfolder.Name) <> "all users" AND LCase(objSubfolder.Name) <> "dell" AND LCase(objSubfolder.Name) <> "hp" AND LCase(objSubfolder.Name) <> "acer" AND objSubfolder.Name <> "G4SEngineer" AND LCase(objSubfolder.Name) <> "public" AND LCase(objSubfolder.Name) <> "temp" AND (Not CBool(InStr(LCase(objSubfolder.Name), "ncentral"))) AND (Not CBool(InStr(LCase(objSubfolder.Name), "default"))) AND (Not CBool(InStr(LCase(objSubfolder.Name), "patchupdate"))) AND (Not CBool(InStr(LCase(objSubfolder.Name), "admin"))) THEN
 		CALL ArrayAdd(user_array, objSubfolder.Name)
 		myLog.WriteLine "Profile Found: " & objSubfolder.Name
 		user_count = user_count + 1
	END IF
NEXT

'Getting other fixed Partitions
dArray = GetDriveList(objFSO)

'0 users found. Log error!!!!
IF user_count = 0 THEN
	myLog.WriteLine "Warning!!! No users found. No data will be transferred"
	RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	Wscript.Quit
END IF

' Check Config available. Create IF not
CALL CheckConfig(objFSO)
' RClone Initial Setup
CALL SetupRClone(objFSO)

'1 user found. Cloning all
IF user_count = 1 THEN
	myLog.WriteLine "Just one user found. Getting ready to export data."
	RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	CALL Clone(objFSO, user_array(0))
'Multiple users found. Special cloning will be executed
ELSE
	myLog.WriteLine "Multiple users found. Looping..."
	RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	FOR EACH username in user_array
		CALL Clone(objFSO, username)
	NEXT
END IF

CALL sharedClone(objFSO, "C")

'Processing Partitions
FOR EACH dLetter in dArray
	CALL sharedClone(objFSO, dLetter)
NEXT

myLog.WriteLine "Cloning Ends..."

RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	
SET user_count = Nothing
SET objFSO = Nothing
myLog.Close

'	*********************
'	*     CLONE MAIN   	*
'	*********************

SUB Clone(objFSO, username)
	On error resume next
	'Create File Array Object
	DIM fileArray
	fileArray = ARRAY() '2 diamentioanal: Sub array order: Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
	'Create File Size Array Object
	DIM fileSizeArray
	fileSizeArray = ARRAY(0,0,0,0) '1 diamentioanal: array order: Total Good and Small (Less than 10MB)  File Size, Total Good and Large File Size, Total Unproductive Samll File Size, Total Unproductive Large File Size

	UserPath = "C:\Users\" & username
	
	'Get email from Chrome profile
	DIM email
	
	email = ""
	CALL FindEmail(email, objFSO, UserPath)

	myLog.WriteLine "Cloning data: " & username
	'Clone User Data Only
	CALL CheckUserData(objFSO, username, fileArray, fileSizeArray)

	fileLog.WriteLine "--------------------------------------------------------------------------------"
	fileLog.WriteLine "Username: " & username
	fileLog.WriteLine "Total Good and Small File Size: " & fileSizeArray(1)
	fileLog.WriteLine "Total Good and Large File Size: " & fileSizeArray(0)
	fileLog.WriteLine "Total Unproductive Small File Size: " & fileSizeArray(3)
	fileLog.WriteLine "Total Unproductive Large File Size: " & fileSizeArray(2)
	TotalFileSize = fileSizeArray(0) + fileSizeArray(1) + fileSizeArray(3)
	fileLog.WriteLine "TotalFileSize to Synch: " & TotalFileSize
	fileLog.WriteLine "--------------------------------------------------------------------------------"
	myLog.WriteLine "--------------------------------------------------------------------------------"
	myLog.WriteLine "Total Good and Small File Size: " & fileSizeArray(1)
	myLog.WriteLine "Total Good and Large File Size: " & fileSizeArray(0)
	myLog.WriteLine "Total Unproductive Small File Size: " & fileSizeArray(3)
	myLog.WriteLine "Total Unproductive Large File Size: " & fileSizeArray(2)
	myLog.WriteLine "TotalFileSize to Synch: " & TotalFileSize
	myLog.WriteLine "--------------------------------------------------------------------------------"
	
	'Check Rprevious process record is avaiable
	TotalSavedFileSize = objEnvSystem( "RCLONE_PROCESSED_" & username )
	IF TotalSavedFileSize <> "" THEN
		TotalSavedFileSize = CInt(TotalSavedFileSize)
	ELSE
		TotalSavedFileSize = 0
	END IF
	
	RoundedFileSize = Round(TotalFileSize)
	TotalSavedFileSize = Round(TotalSavedFileSize)
	
	myLog.WriteLine "Checking: " & RoundedFileSize & " (TotalFileSize to Synch) = " & TotalSavedFileSize & " (TotalFileSize Synched last attempt)"
	
	IF TotalSavedFileSize <> RoundedFileSize AND forceSynch <> 1 THEN
		myLog.WriteLine "File size change found. Synching again"
	END IF
	
	'Get available free space in C Drive
	SET objWMIService = GetObject("winmgmts:")
	SET objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='C:'")
	CDriveFreeSpace = objLogicalDisk.FreeSpace/1024
	
	' Checked the drive is unchanged from the last process
	IF forceSynch = 1 OR TotalSavedFileSize <> RoundedFileSize THEN
		'Checking we have enough space to create backup folder
		IF TotalFileSize*2 < CDriveFreeSpace THEN
			myLog.WriteLine "Enough space available in C Drive (" & CDriveFreeSpace & " KB). Creating Backup folders to Synch"
			RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			CALL CreateBackupFolder(objFSO, username, fileArray)
			RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			IF validateEmail(email) THEN
				Set OutPutFile = objFSO.OpenTextFile(unescape("C:\rclone\" & username & "\email.txt") ,8 , True)
				WScript.Sleep 1000
				OutPutFile.WriteLine(email)
				WScript.Sleep 1000
				myLog.WriteLine "Email Found: " & username & " - " & email
			END IF
			RunRClone objFSO, "C:\rclone\" & username, strDomain & "/" & strComputerName & "/" & username, "Backup", "RCLONE_PROCESSED_" & username, RoundedFileSize, ""
			objFSO.DeleteFolder "C:\rclone\" & username
			WScript.Sleep 10000
		ELSE
			'Clone User Data Only
			myLog.WriteLine "Not enough space available in C Drive (" & CDriveFreeSpace & " KB). Synching actual data"
			RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			CALL CloneUserData(objFSO, username, email, "RCLONE_PROCESSED_" & username, RoundedFileSize)
		END IF
	ELSE
		myLog.WriteLine "Profile " & username & " already Processed"
	END IF
	
	RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	WScript.Sleep 1000
END SUB 

'	*************************
'	*  SHARED CLONE MAIN   	*
'	*************************

SUB sharedClone(objFSO, SharedDLetter)
	On error resume next
	'Create File Array Object
	DIM sharedFileArray
	sharedFileArray = ARRAY() '2 diamentioanal: Sub array order: Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
	'Create File Size Array Object
	DIM sharedFileSizeArray
	sharedFileSizeArray = ARRAY(0,0,0,0) '1 diamentioanal: array order: Total Good and Small (Less than 10MB)  File Size, Total Good and Large File Size, Total Unproductive Samll File Size, Total Unproductive Large File Size
	
	myLog.WriteLine "Cloning Shared data: " & strComputerName & " " & SharedDLetter & " Drive"

	' Clone Shared Data
	CALL CheckSharedData(objFSO, strComputerName, sharedFileArray, sharedFileSizeArray, SharedDLetter)
	
	'---- Shared Data push --------
	
	fileLog.WriteLine "--------------------------------------------------------------------------------"
	fileLog.WriteLine "Shared data for " & strComputerName & " " & SharedDLetter & " Drive"
	fileLog.WriteLine "Total Good and Small File Size: " & sharedFileSizeArray(1)
	fileLog.WriteLine "Total Good and Large File Size: " & sharedFileSizeArray(0)
	fileLog.WriteLine "Total Unproductive Small File Size: " & sharedFileSizeArray(3)
	fileLog.WriteLine "Total Unproductive Large File Size: " & sharedFileSizeArray(2)
	TotalSahredFileSize = sharedFileSizeArray(0) + sharedFileSizeArray(1) + sharedFileSizeArray(3)
	fileLog.WriteLine "TotalFileSize to Synch: " & TotalSahredFileSize
	fileLog.WriteLine "--------------------------------------------------------------------------------"
	myLog.WriteLine "--------------------------------------------------------------------------------"
	myLog.WriteLine "Shared data for " & strComputerName & " " & SharedDLetter & " Drive"
	myLog.WriteLine "Total Good and Small File Size: " & sharedFileSizeArray(1)
	myLog.WriteLine "Total Good and Large File Size: " & sharedFileSizeArray(0)
	myLog.WriteLine "Total Unproductive Small File Size: " & sharedFileSizeArray(3)
	myLog.WriteLine "Total Unproductive Large File Size: " & sharedFileSizeArray(2)
	myLog.WriteLine "TotalFileSize to Synch: " & TotalSahredFileSize
	myLog.WriteLine "--------------------------------------------------------------------------------"
	
	'Check Rprevious process record is avaiable
	sharedTotalSavedFileSize = objEnvSystem( "RCLONE_PROCESSED_" & SharedDLetter )
	IF sharedTotalSavedFileSize <> "" THEN
		sharedTotalSavedFileSize = CInt(sharedTotalSavedFileSize)
	ELSE
		sharedTotalSavedFileSize = 0
	END IF
	
	sharedRoundedFileSize = Round(TotalSahredFileSize)
	sharedTotalSavedFileSize = Round(sharedTotalSavedFileSize)
	
	myLog.WriteLine "Checking: " & sharedRoundedFileSize & " (TotalFileSize to Synch) = " & sharedTotalSavedFileSize & " (TotalFileSize Synched last attempt)"
	
	'Get available free space in the Drive
	SET objWMIService = GetObject("winmgmts:")
	SET objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='" & SharedDLetter &":'")
	DriveFreeSpace = objLogicalDisk.FreeSpace/1024
	
	IF sharedTotalSavedFileSize <> sharedRoundedFileSize AND forceSynch <> 1 THEN
		myLog.WriteLine "File size change found. Synching again"
	END IF
	
	' Checked the drive is unchanged from the last process
	IF forceSynch = 1 OR sharedTotalSavedFileSize <> sharedRoundedFileSize THEN
		'Checking we have enough space to create backup folders
		IF TotalSahredFileSize*2 < DriveFreeSpace THEN
			myLog.WriteLine "Enough space available in " & SharedDLetter & " Drive(" & DriveFreeSpace & " KB). Creating Backup folders to Synch"
			RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			CALL CreateSharedBackupFolder(objFSO, strComputerName & "_" & SharedDLetter, sharedFileArray, SharedDLetter)
			RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, SharedDLetter & ":\rclone\" & strComputerName & "_" & SharedDLetter, strDomain & "/" & strComputerName, SharedDLetter, "RCLONE_PROCESSED_" & SharedDLetter, sharedRoundedFileSize, ""
			objFSO.DeleteFolder SharedDLetter & ":\rclone\" & strComputerName & "_" & SharedDLetter
			WScript.Sleep 10000
		ELSE
			' Clone Shared Data
			myLog.WriteLine "Not enough space available in " & SharedDLetter & " Drive(" & DriveFreeSpace & " KB). Synching actual data"
			RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
			CALL CloneSharedData(objFSO, strComputerName, "RCLONE_PROCESSED_" & SharedDLetter, RoundedFileSize, SharedDLetter)
		END IF
	ELSE
		myLog.WriteLine SharedDLetter & " Drive already Processed"
	END IF
	
	RunRClone objFSO, "C:\rclone\ApolloSynch.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	RunRClone objFSO, "C:\rclone\ApolloSynchFiles.log", "Logs/" & strDomain & "/" & strName, "", "", "", 1
	WScript.Sleep 1000
END SUB 

'	*************************************************
'	*     CHECK RCLONE AVAILABLE. INSTALL IF NOT    *
'	*************************************************

SUB CheckRClone(fso)
	'Check RClone path
	IF (fso.FileExists("C:\rclone\rclone.exe")) THEN
	   myLog.WriteLine "RClone exists. Ready to Clone."
	ELSE
		Wscript.echo "RClone is not available. Creating Dir to push."
		'Create RClone Dir
		CALL SureDirectoryExists(fso, "C:\rclone", 1)
		WScript.Sleep 1000
		CALL SureDirectoryExists(fso, unescape(profilePath & "\.config"), 1)
		WScript.Sleep 1000
		CALL SureDirectoryExists(fso, unescape(profilePath & "\.config\rclone"), 1)
		WScript.Sleep 1000

		'Create Password File
		IF NOT (fso.FileExists(unescape(profilePath & "\.config\rclone\pwd.txt"))) THEN
			WScript.Sleep 1000
			Set OutPutFile = fso.OpenTextFile(unescape(profilePath & "\.config\rclone\pwd.txt") ,8 , True)
			WScript.Sleep 1000
			OutPutFile.WriteLine("QazWsxEdc1!")
			WScript.Sleep 1000
		END IF

		'Downloading RClone
		'CALL HTTPDownload(fso, "https://downloads.rclone.org/v1.48.0/rclone-v1.48.0-windows-amd64.zip","C:\rclone")
		'Unzip file
		'CALL Unzip("C:\rclone\rclone-v1.48.0-windows-amd64.zip", "C:\rclone\")
		'Move to standard folder
		'fso.MoveFile "C:\rclone\rclone-v1.48.0-windows-amd64\rclone.exe", "C:\rclone\"
		'Deleting downloaded stuffs
		'fso.DeleteFolder "C:\rclone\rclone-v1.48.0-windows-amd64\"
		'fso.DeleteFile "C:\rclone\rclone-v1.48.0-windows-amd64.zip"
		SET objFSO = Nothing
		Wscript.Quit
	END IF
END SUB


'	*************************************************
'	*     CHECK CONFIG AVAILABLE. CREATE IF NOT     *
'	*************************************************

SUB CheckConfig(fso)
	'Check RClone path
	destFile = unescape(profilePath & "\.config\rclone\rclone.conf")
	
	IF (fso.FileExists(destFile)) THEN
	   myLog.WriteLine "Config exists. Ready to Clone."
	ELSE
		myLog.WriteLine "Config is not available. Creating..."
		'Create Config Dir
		IF (fso.FileExists("C:\rclone\rclone.conf")) THEN
			'Create Config Dir
			'WshShell.SendKeys "md " & unescape("C:\Users\" & strName & "\.config"), 0
			'WshShell.SendKeys "{ENTER}"
			'WshShell.SendKeys "md " & unescape("C:\Users\" & strName & "\.config\rclone"), 0
			'WshShell.SendKeys "{ENTER}"
			'WshShell.SendKeys "copy C:\rclone\rclone.conf " & destFile, 0
			'WshShell.SendKeys "{ENTER}"
			' System uder Rclone path is different
			
			CALL SureDirectoryExists(fso, unescape(profilePath & "\.config"), 1)
			CALL SureDirectoryExists(fso, unescape(profilePath & "\.config\rclone"), 1)
			fso.CopyFile "C:\rclone\rclone.conf", destFile
			'Create Password File
		ElSE
			myLog.WriteLine "No config file to copy"
		END IF
	END IF
	
	IF NOT (fso.FileExists(unescape(profilePath & "\.config\rclone\pwd.txt"))) THEN
		'myLog.WriteLine "PWD File: " & unescape(profilePath & "\.config\rclone\pwd.txt")
		WScript.Sleep 1000
		Set OutPutFile = fso.OpenTextFile(unescape(profilePath & "\.config\rclone\pwd.txt") ,8 , True)
		WScript.Sleep 1000
		OutPutFile.WriteLine("QazWsxEdc1!")
		WScript.Sleep 1000
	END IF
END SUB

'	*******************************
'	*     CHECK USER DATA ONLY    *
'	*******************************

SUB CheckUserData(fso, username, fileArray, fileSizeArray)
	myLog.WriteLine "Checking user data: " & username
	UserPath = "C:\Users\" & username & "\"
	SET FileFolder = fso.GetFolder(UserPath)
	CALL ProcessFolders(fileArray, fileSizeArray, UserPath, FileFolder)
END SUB 

'	*******************************
'	*     CLONE USER DATA ONLY    *
'	*******************************

SUB CloneUserData(fso, username, email, envVar, envVal)
	myLog.WriteLine "Cloning user data: " & username
	UserPath = "C:\Users\" & username & "\Documents"
	RunRClone fso, UserPath, strDomain & "/" & strComputerName & "/" & username, "Documents", "", "", ""
	UserPath = "C:\Users\" & username & "\Desktop"
	RunRClone fso, UserPath, strDomain & "/" & strComputerName & "/" & username, "Desktop", "", "", ""
	UserPath = "C:\Users\" & username & "\Downloads"
	RunRClone fso, UserPath, strDomain & "/" & strComputerName & "/" & username, "Downloads", "", "", ""
	IF validateEmail(email) THEN
		Set OutPutFile = fso.OpenTextFile(unescape("C:\Users\" & username & "\email.txt") ,8 , True)
		WScript.Sleep 1000
		OutPutFile.WriteLine(email)
		WScript.Sleep 1000
		RunRClone fso, "C:\Users\" & username & "\email.txt", strDomain & "/" & strComputerName & "/" & username, "", "", "", ""
		myLog.WriteLine "Email Found: " & username & " - " & email
	END IF
	
	'Setting Sytem env VARS to check the process is executed already and unchanged
	IF envVar <> "" AND envVal <> "" THEN
		objEnvUser(envVar) = envVal
		objEnvProcess(envVar) = envVal
		objEnvSystem(envVar) = envVal
		WScript.Sleep 10000
	END IF
END SUB 

'	*********************************
'	*     CHECK SHARED DATA			*
'	*********************************

SUB CheckSharedData(fso, strComputerName, sharedFileArray, sharedFileSizeArray, SharedDLetter)
	myLog.WriteLine "Cloning shared data for: " & strComputerName & " " & SharedDLetter & " Drive"
	SET FileFolder = fso.GetFolder(SharedDLetter & ":\")
	CALL ProcessSharedFolders(sharedFileArray, sharedFileSizeArray, SharedDLetter, FileFolder)
END SUB

'	*********************************
'	*     CLONE SHARED DATA			*
'	*********************************

SUB CloneSharedData(fso, strComputerName, envVar, envVal, SharedDLetter)
	myLog.WriteLine "Cloning shared data for: " & strComputerName
	SET objMainFolder = fso.GetFolder(SharedDLetter & ":\")
	SET colSharedfolders = objMainFolder.Subfolders
	FOR EACH objSharedfolder in colSharedfolders
		DO
		    IF objSharedfolder.Attributes AND 2 THEN EXIT DO
			'Standard Windows folders will be ignored
			IF objSharedfolder.Name <> "Windows" AND objSharedfolder.Name <> "Program Files" AND objSharedfolder.Name <> "Program Files (x86)" AND objSharedfolder.Name <> "Users" AND objSharedfolder.Name <> "Intel" AND objSharedfolder.Name <> "SWSetup" THEN
		 		SharedPath = SharedDLetter & ":\" & objSharedfolder.Name & "\"
		 		SET FileFolder = fso.GetFolder(SharedPath)
				CALL RunRClone(fso, SharedPath, strDomain & "/" & strComputerName, SharedDLetter & "/" & objSharedfolder.Name, "", "", "")
			END IF
		LOOP WHILE FALSE
	NEXT
	
	'Setting Sytem env VARS to check the process is executed already and unchanged
	IF envVar <> "" AND envVal <> "" THEN
		objEnvUser(envVar) = envVal
		objEnvProcess(envVar) = envVal
		objEnvSystem(envVar) = envVal
		WScript.Sleep 10000
	END IF
END SUB

'	*********************************
'	*     Create Backup Folder		*
'	*********************************

SUB CreateBackupFolder(fso, username, fileArray)
	On error resume next
	
	CALL SureDirectoryExists(fso, unescape("C:\rclone\" & username), 0)

	SavedPath = ""
	FOR EACH FileItem IN fileArray 'FileItem Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
		IF FileItem(0) = 1 OR FileItem(1) = 1 THEN
			IF SavedPath <> FileItem(7) THEN
				CALL SureDirectoryExists(fso, unescape("C:\rclone\" & username & "\" & FileItem(7)), 0)
			END IF
			SavedPath = FileItem(7)
			fso.CopyFile FileItem(3), unescape("C:\rclone\" & username & "\" & FileItem(4))
		END IF
    NEXT
    
    backupFolderSize = fso.GetFolder(unescape("C:\rclone\" & username)).Size
    myLog.WriteLine "Backup folder Size: " & backupFolderSize
END SUB

'	*********************************
'	*     Create Backup Folder		*
'	*********************************

SUB CreateSharedBackupFolder(fso, strComputerName, sharedFileArray, SharedDLetter)
	On error resume next
	
	CALL SureDirectoryExists(fso, unescape(SharedDLetter & ":\rclone\"), 0)
	CALL SureDirectoryExists(fso, unescape(SharedDLetter & ":\rclone\" & strComputerName), 0)

	SavedPath = ""
	FOR EACH FileItem IN sharedFileArray 'FileItem Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
		IF FileItem(0) = 1 OR FileItem(1) = 1 THEN
			IF SavedPath <> FileItem(7) THEN
				CALL SureDirectoryExists(fso, unescape(SharedDLetter & ":\rclone\" & strComputerName & "\" & FileItem(7)), 0)
			END IF
			SavedPath = FileItem(7)
			fso.CopyFile FileItem(3), unescape(SharedDLetter & ":\rclone\" & strComputerName & "\" & FileItem(4))
		END IF
    NEXT
    
    backupFolderSize = fso.GetFolder(unescape(SharedDLetter & ":\rclone\" & strComputerName)).Size
    myLog.WriteLine "Backup folder Size: " & backupFolderSize
END SUB

'	*********************************
'	*     RCLONE INITIAL SETUP		*
'	*********************************

SUB SetupRClone(objFSO)
	' Set Env variable for the config path (!!Not working. Still directing to User profile path C:\Users\<username>\.config\rclone\rclone.config!!) 
	objEnvUser("RCLONE_CONFIG") = "C:/rclone/rclone.conf"
	objEnvProcess("RCLONE_CONFIG") = "C:/rclone/rclone.conf"
	objEnvSystem("RCLONE_CONFIG") = "C:/rclone/rclone.conf"
	' Running RClone config minimized mode
	myLog.WriteLine "Running C:\rclone\rclone.exe --config=C:/rclone/rclone.conf config file "
	
	WScript.Sleep 1000
	WshShell.Run "cmd /c C:\rclone\rclone.exe --config=C:/rclone/rclone.conf config file ", 0, True '(!!--config=C:/rclone/rclone.conf is Not working. Still directing to User profile path C:\Users\<username>\.config\rclone\rclone.config!!)
	WScript.Sleep 1000
END SUB

'	*************************
'	*     RClone Synch		*
'	*************************

SUB RunRClone(objFSO, file_path, folder_path, DestPath, envVar, envVal, NoLog)
	' Copying files
	objEnvUser("RCLONE_CONFIG") = "C:/rclone/rclone.conf"
	objEnvProcess("RCLONE_CONFIG") = "C:/rclone/rclone.conf"
	objEnvSystem("RCLONE_CONFIG") = "C:/rclone/rclone.conf"
	
	gDrivePath = "gdrive:"
	
	IF folder_path <> "" THEN
		gDrivePath = gDrivePath & folder_path 
	END IF
	
	IF DestPath <> "" THEN
		gDrivePath = gDrivePath & "/" & DestPath
	END IF
	
	IF NoLog = "" THEN
		myLog.WriteLine "Running C:\rclone\rclone.exe copy " & file_path & " " & gDrivePath
	END IF
	WScript.Sleep 1000
	IF NoLog = "" THEN
		myLog.WriteLine "Synching Started"
	END IF
	'Running RClone
	WshShell.Run "cmd /c C:\rclone\rclone.exe copy " & file_path &  " " & gDrivePath & " < " & profilePath & "\.config\rclone\pwd.txt", 1, True
	IF NoLog = "" THEN
		myLog.WriteLine "Synching Ends"
	END IF
	WScript.Sleep 1000
	'Setting Sytem env VARS to check the process is executed already and unchanged
	IF envVar <> "" AND envVal <> "" THEN
		objEnvUser(envVar) = envVal
		objEnvProcess(envVar) = envVal
		objEnvSystem(envVar) = envVal
		WScript.Sleep 10000
	END IF
END SUB

'	*********************************
'	*     ADD ITEM TO ARRAY			*
'	*********************************

SUB ArrayAdd(arr, val)
    REDIM Preserve arr(UBOUND(arr) + 1)
    arr(UBOUND(arr)) = val
END SUB

'	*****************************************
'	*     Check Directory create IF not		*
'	*****************************************

SUB SureDirectoryExists(fso, ADir, MakeHidden)
	If Not fso.FolderExists(ADir) THEN
	    objFolder = fso.CreateFolder(ADir)
	    'IF MakeHidden = 1 THEN
	    '	objFolder.attributes = folder.attributes OR 2
	    'END IF
	END IF
END SUB 
	
'	*********************
'	*     UNZIP			*
'	*********************

SUB Unzip(strZipFile, outFolder)
	REM SET WshShell = CREATEOBJECT("Wscript.Shell")
	
	myLog.WriteLine ( "Extracting file " & strFileZIP)
	
	SET objShell = CREATEOBJECT( "Shell.Application" )
	SET objSource = objShell.NameSpace(strZipFile).Items()
	SET objTarget = objShell.NameSpace(outFolder)
	intOptions = 256
	objTarget.CopyHere objSource, intOptions
	
	myLog.WriteLine ( "Extracted." )
END SUB


'	*********************************
'	*     Download File	(Slow...)   *
'	*********************************

SUB HTTPDownload(objFSO, myURL, myPath )
    ' Standard housekeeping
    DIM i, objFile, objHTTP, strFile, strMsg
    CONST ForReading = 1, ForWriting = 2, ForAppending = 8

    ' Check IF the specified target file or folder exists,
    ' and build the fully qualified path of the target file
    If objFSO.FolderExists( myPath ) THEN
        strFile = objFSO.BuildPath( myPath, MID( myURL, INSTRREV( myURL, "/" ) + 1 ) )
    ELSEIf objFSO.FolderExists( LEFT( myPath, INSTRREV( myPath, "\" ) - 1 ) ) THEN
        strFile = myPath
    ELSE
        myLog.WriteLine "ERROR: Target folder not found."
        EXIT SUB
    END IF

    ' Create or open the target file
    SET objFile = objFSO.OpenTextFile( strFile, ForWriting, True )

    ' Create an HTTP object
    SET objHTTP = CREATEOBJECT( "WinHttp.WinHttpRequest.5.1" )

    ' Download the specified URL
    objHTTP.Open "GET", myURL, False
    objHTTP.Send

    ' Write the downloaded byte stream to the target file
    FOR i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write CHR( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    NEXT

    ' Close the target file
    objFile.Close( )
END SUB

'	******************
'	*   Find Email   *
'	******************

SUB FindEmail(email,objFSO,myPath)
	On error resume next
	
	Set re = New RegExp
	With re
	.Pattern    = "identifier[^@]*@[A-Za-z\.]*g4s\.compassword"
	.IgnoreCase = False
	.Global     = False
	End With
	
	Set objRegExp = New Regexp
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	objRegExp.Pattern = "[^\x1F-\x7F]+"
	
	' Read From A File
	FilePath = myPath & "\AppData\Local\Google\Chrome\User Data\Default\Login Data"
	IF objFSO.FileExists(unescape(FilePath)) Then
		SET UserFile = objFSO.OpenTextFile(unescape(FilePath), 1, 1)
		Do Until UserFile.AtEndOfStream
			FileLine = UserFile.ReadLine()
			FileLine = CStr(FileLine)
			FileLine = objRegExp.Replace(FileLine, "")
			myMatches = re.Test(FileLine)
			If LEN(myMatches)> 0 THEN
				Set oMatches = re.Execute(FileLine)
				For Each oMatch In oMatches
					email =  oMatch.Value
					email = Replace(email,"identifier","")
					email = Replace(email,"compassword","com")
					UserFile.Close
					Exit Do
				Next
			END IF
		Loop
		If Err.Number <> 0 Then
		  myLog.WriteLine "Error: " & Err.Description
		  Err.Clear
		End If
		UserFile.Close
	ELSE
		myLog.WriteLine "User file not found"
		FindEmail = ""
	END IF
END SUB

'	******************************
'	*   VALIDATE EMAIL ADDRESS   *
'	******************************

FUNCTION validateEmail(eaddr)
  DIM isValidE
  DIM regEx

  isValidE = True
  SET regEx = New RegExp

  regEx.IgnoreCase = False

  regEx.Pattern = "^[-+.\w]{1,64}@[-.\w]{1,64}\.[-.\w]{2,6}$"
  isValidE = regEx.Test(eaddr)

  validateEmail = isValidE
END FUNCTION

'	******************************************
'	*   RUN AS AN INTERACTIVE USER SESSION   *
'	******************************************

Function RunAsInteractiveSession(User,Password,Command)
	strCurrrentUser=User ' user name to open app in
	strCmd1=Command ' application cmd
	strPass1=Password&CHR(13) ' user password
	
	On Error Resume Next
	dim WshShell,FSO

	set WshEnv = WshShell.Environment("Process")
	WinPath = WshEnv("SystemRoot")&"\System32\runas.exe"
	set FSO = CreateObject("Scripting.FileSystemObject")
	
	myLog.WriteLine "Running runas /user:" & strCurrrentUser & " " & CHR(34) & strCmd1 & CHR(34)
	rc=WshShell.Run("runas /user:" & strCurrrentUser & " " & CHR(34) & strCmd1 & CHR(34), 2, FALSE)
	WScript.Sleep 1000
	WshShell.AppActivate(WinPath) 'focus window
	WshShell.SendKeys strPass1 'send password to focus window.
	
	set WshShell=Nothing
End Function

'	***********************
'	*   RUN SUB FOLDERS   *
'	***********************

SUB ProcessFolders(fileArray, fileSizeArray, UserPath, FileFolder)
	On error resume next
    FOR EACH Subfolder in FileFolder.SubFolders
    	DO
		    IF Subfolder.Name = ".config" OR Subfolder.Name = "Music" OR Subfolder.Name = "Videos" OR Subfolder.Name = "AppData" OR (Subfolder.Attributes AND 2) THEN EXIT DO
	        Set objFolder = objFSO.GetFolder(Subfolder.Path)
	        Set colFiles = objFolder.Files
	        SubfolderPath = Replace(Subfolder.Path,UserPath,"")
	        FOR EACH Files IN colFiles
	        	IF Files.Name <> "desktop.ini" THEN
		        	FilesSize = Files.Size/1024
			        FilesPath = Replace(Files.Path,UserPath,"")
		        	TypeMatched = 0
		        	FOR EACH fileType in FileTypeArray
			            IF LCase(InStr(1,Files, fileType)) > 1 THEN
			            	IF FilesSize > 10000 THEN
				            	fileSizeArray(0) = fileSizeArray(0) + FilesSize
				            	fileLog.WriteLine "Good file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
				            	'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
				            	CALL ArrayAdd(fileArray, Array(1,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
				            ELSE
				        		fileSizeArray(1) = fileSizeArray(1) + FilesSize
				        		fileLog.WriteLine "Good file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
				        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size,  Subfolder Path
				            	CALL ArrayAdd(fileArray, Array(1,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
				        	END IF
			            	TypeMatched = 1
			            	EXIT FOR
			            END IF
			        NEXT
			        
			        IF TypeMatched = 0 THEN
			        	IF FilesSize > 10000 THEN
			        		fileSizeArray(2) = fileSizeArray(2) + FilesSize
			        		fileLog.WriteLine "Unproductive file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
				            CALL ArrayAdd(fileArray, Array(0,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			        	ELSE
			        		fileSizeArray(3) = fileSizeArray(3) + FilesSize
			        		fileLog.WriteLine "Unproductive file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
				            CALL ArrayAdd(fileArray, Array(0,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			        	END IF
			        END IF
				END IF
	        Next
	        CALL ProcessFolders(fileArray, fileSizeArray, UserPath, Subfolder)
		LOOP WHILE FALSE
    NEXT
    
    Set colFiles = FileFolder.Files
    SubfolderPath = Replace(UserPath,UserPath,"")
    FOR EACH Files IN colFiles
    	IF Files.Name <> "NTUSER.DAT" AND Files.Name <> "desktop.ini" THEN 
	    	FilesSize = Files.Size/1024
	        FilesPath = Replace(Files.Path,UserPath,"")
	        IF InStr(FilesPath,"\") = 0 THEN
		    	TypeMatched = 0
		    	FOR EACH fileType in FileTypeArray
		            IF LCase(InStr(1,Files, fileType)) > 1 THEN
		            	IF FilesSize > 10000 THEN
			            	fileSizeArray(0) = fileSizeArray(0) + FilesSize
			            	fileLog.WriteLine "Good file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			            	'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
			            	CALL ArrayAdd(fileArray, Array(1,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			            ELSE
			        		fileSizeArray(1) = fileSizeArray(1) + FilesSize
			        		fileLog.WriteLine "Good file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size,  Subfolder Path
			            	CALL ArrayAdd(fileArray, Array(1,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			        	END IF
		            	TypeMatched = 1
		            	EXIT FOR
		            END IF
		        NEXT
		        
		        IF TypeMatched = 0 THEN
		        	IF FilesSize > 10000 THEN
		        		fileSizeArray(2) = fileSizeArray(2) + FilesSize
		        		fileLog.WriteLine "Unproductive file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
		        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
			            CALL ArrayAdd(fileArray, Array(0,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
		        	ELSE
		        		fileSizeArray(3) = fileSizeArray(3) + FilesSize
		        		fileLog.WriteLine "Unproductive file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
		        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
			            CALL ArrayAdd(fileArray, Array(0,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
		        	END IF
		        END IF
			END IF
		END IF
    Next
End SUB

'	******************************
'	*   RUN SAHRED SUB FOLDERS   *
'	******************************

SUB ProcessSharedFolders(sharedFileArray, sharedFileSizeArray, SharedDLetter, FileFolder)
	On error resume next
	SharedPath = SharedDLetter & ":\"
    FOR EACH Subfolder in FileFolder.SubFolders
    	DO
		    IF Subfolder.Name = ".config" OR Subfolder.Name = "Music" OR Subfolder.Name = "Videos" OR Subfolder.Name = "AppData" OR Subfolder.Name = "Windows" OR Subfolder.Name = "Program Files" OR Subfolder.Name = "Program Files (x86)" OR Subfolder.Name = "Users" OR Subfolder.Name = "Intel" OR Subfolder.Name = "SWSetup" OR Subfolder.Name = "rclone" OR (Subfolder.Attributes AND 2) THEN EXIT DO
		    Set objFolder = objFSO.GetFolder(Subfolder.Path)
	        Set colFiles = objFolder.Files
	        SubfolderPath = Replace(Subfolder.Path,SharedPath,"")
	        FOR EACH Files IN colFiles
	        	IF Files.Name <> "desktop.ini" THEN
		        	FilesSize = Files.Size/1024
			        FilesPath = Replace(Files.Path,SharedPath,"")
		        	TypeMatched = 0
		        	FOR EACH fileType in FileTypeArray
			            IF LCase(InStr(1,Files, fileType)) > 1 THEN
			            	IF FilesSize > 10000 THEN
				            	sharedFileSizeArray(0) = sharedFileSizeArray(0) + FilesSize
				            	fileLog.WriteLine "Good file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
				            	'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
				            	CALL ArrayAdd(sharedFileArray, Array(1,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
				            ELSE
				        		sharedFileSizeArray(1) = sharedFileSizeArray(1) + FilesSize
				        		fileLog.WriteLine "Good file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
				        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size,  Subfolder Path
				            	CALL ArrayAdd(sharedFileArray, Array(1,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
				        	END IF
			            	TypeMatched = 1
			            	EXIT FOR
			            END IF
			        NEXT
			        
			        IF TypeMatched = 0 THEN
			        	IF FilesSize > 10000 THEN
			        		sharedFileSizeArray(2) = sharedFileSizeArray(2) + FilesSize
			        		fileLog.WriteLine "Unproductive file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
				            CALL ArrayAdd(sharedFileArray, Array(0,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			        	ELSE
			        		sharedFileSizeArray(3) = sharedFileSizeArray(3) + FilesSize
			        		fileLog.WriteLine "Unproductive file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
				            CALL ArrayAdd(sharedFileArray, Array(0,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			        	END IF
			        END IF
				END IF
	        Next
	        CALL ProcessSharedFolders(sharedFileArray, sharedFileSizeArray, SharedPath, Subfolder)
		LOOP WHILE FALSE
    NEXT
    
    Set colFiles = FileFolder.Files
    SubfolderPath = Replace(SharedPath,SharedPath,"")
    FOR EACH Files IN colFiles
    	IF Files.Name <> "desktop.ini" THEN
	    	FilesSize = Files.Size/1024
	        FilesPath = Replace(Files.Path,SharedPath,"")
	        IF InStr(FilesPath,"\") = 0 THEN
		    	TypeMatched = 0
		    	FOR EACH fileType in FileTypeArray
		            IF LCase(InStr(1,Files, fileType)) > 1 THEN
		            	IF FilesSize > 10000 THEN
			            	sharedFileSizeArray(0) = sharedFileSizeArray(0) + FilesSize
			            	fileLog.WriteLine "Good file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			            	'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
			            	CALL ArrayAdd(sharedFileArray, Array(1,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			            ELSE
			        		sharedFileSizeArray(1) = sharedFileSizeArray(1) + FilesSize
			        		fileLog.WriteLine "Good file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
			        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size,  Subfolder Path
			            	CALL ArrayAdd(sharedFileArray, Array(1,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
			        	END IF
		            	TypeMatched = 1
		            	EXIT FOR
		            END IF
		        NEXT
		        
		        IF TypeMatched = 0 THEN
		        	IF FilesSize > 10000 THEN
		        		sharedFileSizeArray(2) = sharedFileSizeArray(2) + FilesSize
		        		fileLog.WriteLine "Unproductive file larger than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
		        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
			            CALL ArrayAdd(sharedFileArray, Array(0,0,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
		        	ELSE
		        		sharedFileSizeArray(3) = sharedFileSizeArray(3) + FilesSize
		        		fileLog.WriteLine "Unproductive file smaller than 10MB -"  & Chr(9) & "Name:" & Chr(9) & Files.Name & Chr(9) & "Path:" & Chr(9) & FilesPath & Chr(9) & "Type:" & Chr(9) & Files.Type & Chr(9) & "Size:" & Chr(9) & FilesSize
		        		'Adding to the file array. Elemets. Productibe file, Large, File Name, Source Path, Destination Path, Type, Size, Subfolder Path
			            CALL ArrayAdd(sharedFileArray, Array(0,1,Files.Name,Files.Path,FilesPath,Files.Type,FilesSize,SubfolderPath))
		        	END IF
		        END IF
			END IF
		END IF
    Next
End SUB

'	*****************************
'	*   GET FIXED DRIVE LIST	*
'	*****************************

FUNCTION GetDriveList(fso)
	DIM d, dc, drive_array
	drive_array = ARRAY()
	SET dc = fso.Drives
	FOR EACH d IN dc
		dType = ShowDriveType(fso,d.DriveLetter)
		IF dType = "Fixed" THEN
			IF d.DriveLetter <> "C" THEN
				CALL ArrayAdd(drive_array, d.DriveLetter)
			END IF
			myLog.WriteLine "Fixed Drive Found: " & d.DriveLetter
  		END IF
   NEXT
   GetDriveList = drive_array
END FUNCTION

'	*************************
'	*   SHOW DRIVE TYPE		*
'	*************************

' Check the status of the drives other than C drive as we are only processing Fixed drives
FUNCTION ShowDriveType(fso,drvpath)
   DIM d, t
   SET d = fso.GetDrive(drvpath)
   SELECT CASE d.DriveType
      CASE 0: t = "Unknown"
      CASE 1: t = "Removable"
      CASE 2: t = "Fixed"
      CASE 3: t = "Network"
      CASE 4: t = "CD-ROM"
      CASE 5: t = "RAM Disk"
   END SELECT
   ShowDriveType = t
END FUNCTION
