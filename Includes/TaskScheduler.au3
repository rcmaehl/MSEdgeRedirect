#include-once
#include <Date.au3>
#include <Array.au3>
; #include <WinAPIProc.au3>  REMOVE THIS COMMENT TO USE FUNCTION _TS_IsRunFromTaskScheduler - EXPERIMENTAL!

; #INDEX# =======================================================================================================================
; Title .........: Microsoft Task Scheduler Function Library
; AutoIt Version : 3.3.14.5
; UDF Version ...: 1.6.0.0
; Language ......: English
; Description ...: A collection of functions to access and manipulate the Microsoft Tasks Scheduler Service
; Author(s) .....: water
; Modified.......: 20211119 (YYYYMMDD)
; Contributors ..: AdamUL, allow2010
; Resources .....: https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-reference
;                  https://docs.microsoft.com/de-de/windows/win32/taskschd/task-scheduler-error-and-success-constants
;                  https://www.experts-exchange.com/articles/11326/VBScript-and-Task-Scheduler-2-0-Listing-Scheduled-Tasks.html
; Links .........: Forum:    https://www.autoitscript.com/forum/topic/200068-creating-a-brushed-up-task-scheduler-udf
;                  Download: https://www.autoitscript.com/forum/files/file/487-task-scheduler/
; ===============================================================================================================================

#Region #VARIABLES#
; #VARIABLES# ===================================================================================================================
Global $__iTS_Debug = 0                              ; Debug level:
;                                                      0 = no debug information
;                                                      1 = Debug info to console
;                                                      2 = Debug info to MsgBox
;                                                      3 = Debug info to File
Global $__sTS_DebugFile = @ScriptDir & "\TaskScheduler_Debug.txt" ; Debug file when $__iTS_Debug is set to 3
Global $__oTS_Error                                  ; COM Error handler
; ===============================================================================================================================
#EndRegion #VARIABLES#

#Region #CONSTANTS#
; #CONSTANTS# ===================================================================================================================
; TASK_ACTION_TYPE constants define the type of Actions a Task can perform
; https://docs.microsoft.com/de-de/windows/win32/taskschd/action-type
Global Const $TASK_ACTION_EXEC = 0                   ; Perform a command-line operation e.g. run a script, launch an executable, or, if the name of a document
;                                                      is provided, open a the application with the document
Global Const $TASK_ACTION_COM_HANDLER = 5            ; Fires a handler
Global Const $TASK_ACTION_SEND_EMAIL = 6             ; Sends an email message (no longer supported)
Global Const $TASK_ACTION_SHOW_MESSAGE = 7           ; Show a message box (no longer supported)

; TASK_COMPATIBILITY constants define the versions of Task Scheduler or the AT command the Task is compatible with
; https://docs.microsoft.com/de-de/windows/win32/taskschd/tasksettings-compatibility
Global Const $TASK_COMPATIBILITY_AT = 0              ; The Task is compatible with the AT command
Global Const $TASK_COMPATIBILITY_V1 = 1              ; The Task is compatible with Task Scheduler 1.0
Global Const $TASK_COMPATIBILITY_V2 = 2              ; The Task is compatible with Task Scheduler 2.0

; TASK_CREATE constants define how the Task Scheduler service creates, updates, or disables the Task
; https://docs.microsoft.com/de-de/windows/win32/taskschd/taskfolder-registertaskdefinition
Global Const $TASK_VALIDATE_ONLY = 1                 ; The Task Scheduler checks the syntax of the XML that describes the Task but does not register the Task
;                                                      This constant cannot be combined with the TASK_CREATE, TASK_UPDATE, or TASK_CREATE_OR_UPDATE values
Global Const $TASK_CREATE = 2                        ; The Task Scheduler registers the Task as a new Task
Global Const $TASK_UPDATE = 4                        ; The Task Scheduler registers the Task as an updated version of an existing Task
;                                                      When a Task with a registration Trigger is updated, the Task will execute after the update occurs
Global Const $TASK_CREATE_OR_UPDATE = 6              ; The Task Scheduler either registers the Task as a new Task or as an updated version if the Task already exists
;                                                      Equivalent to TASK_CREATE | TASK_UPDATE
Global Const $TASK_DISABLE = 8                       ; The Task Scheduler disables the existing Task
Global Const $TASK_DONT_ADD_PRINCIPAL_ACE = 16       ; The Task Scheduler is prevented from adding the allow access-control entry (ACE) for the context principal
;                                                      When the TaskFolder.RegisterTaskDefinition function is called with this flag to update a Task, the Task Scheduler
;                                                      service does not add the ACE for the new context principal and does not remove the ACE from the old context principal
Global Const $TASK_IGNORE_REGISTRATION_TRIGGERS = 32 ; The Task Scheduler creates the Task, but ignores the registration Triggers in the Task.
;                                                      By ignoring the registration Triggers, the Task will not execute when it is registered unless a time-based Trigger
;                                                      causes it to execute on registration

; TASK_INSTANCES_POLICY constants define how the Task Scheduler handles existing instances of the Task when it starts a new instance of the Task
; https://docs.microsoft.com/de-de/windows/win32/taskschd/tasksettings-multipleinstances
Global Const $TASK_INSTANCES_PARALLEL = 0            ; Starts a new instance while an existing instance of the Task is running
Global Const $TASK_INSTANCES_QUEUE = 1               ; Starts a new instance of the Task after all other instances of the Task are complete
Global Const $TASK_INSTANCES_IGNORE_NEW = 2          ; Does not start a new instance if an existing instance of the Task is running
Global Const $TASK_INSTANCES_STOP_EXISTING = 3       ; Stops an existing instance of the Task before it starts new instance.

; TASK_LOGON_TYPE constants define what logon technique is required to run a Task
; https://docs.microsoft.com/de-de/windows/win32/taskschd/taskfolder-registertaskdefinition
Global Const $TASK_LOGON_NONE = 0                    ; The logon method is not specified. Used for non-NT credentials
Global Const $TASK_LOGON_PASSWORD = 1                ; Use a password for logging on the user. The password must be supplied at registration time
Global Const $TASK_LOGON_S4U = 2                     ; Use an existing interactive token to run a Task. The user must log on using a service for user (S4U) logon
;                                                      When an S4U logon is used, no password is stored by the system and there is no access to either the network or to
;                                                      encrypted files
Global Const $TASK_LOGON_INTERACTIVE_TOKEN = 3       ; User must already be logged on. The Task will be run only in an existing interactive session
Global Const $TASK_LOGON_GROUP = 4                   ; Group activation. The groupId field specifies the group
Global Const $TASK_LOGON_SERVICE_ACCOUNT = 5         ; Indicates that a Local System, Local Service, or Network Service account is being used as a security context to run the Task
Global Const $TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD = 6 ; First use the interactive token. If the user is not logged on (no interactive token is available), then
;                                                      the password is used. The password must be specified when a Task is registered
;                                                      This flag is not recommended for new Tasks because it is less reliable than TASK_LOGON_PASSWORD

; TASK_RUNLEVEL_TYPE constants define LUA elevation flags that specify the privilege level the Task will be run with
; https://docs.microsoft.com/de-de/windows/win32/taskschd/principal-runlevel
Global Const $TASK_RUNLEVEL_LUA = 0                  ; Tasks will be run with the least privileges (LUA)
Global Const $TASK_RUNLEVEL_HIGHEST = 1              ; Tasks will be run with the highest privileges

; TASK_SESSION_STATE_CHANGE_TYPE constants define what kind of Terminal Server session state change you can use to Trigger a Task to start
; https://docs.microsoft.com/de-de/windows/win32/taskschd/sessionstatechangetrigger-statechange
Global Const $TASK_CONSOLE_CONNECT = 1               ; Terminal Server console connection state change. For example, when you connect to a user session on the local computer by switching users on the computer
Global Const $TASK_CONSOLE_DISCONNECT = 2            ; Terminal Server console disconnection state change. For example, when you disconnect to a user session on the local computer by switching users on the computer
Global Const $TASK_REMOTE_CONNECT = 3                ; Terminal Server remote connection state change. For example, when a user connects to a user session by using the Remote Desktop Connection program from a remote computer.
Global Const $TASK_REMOTE_DISCONNECT = 4             ; Terminal Server remote disconnection state change. For example, when a user disconnects from a user session while using the Remote Desktop Connection program from a remote computer
Global Const $TASK_SESSION_LOCK = 7                  ; Terminal Server session locked state change. For example, this state change causes the Task to run when the computer is locked
Global Const $TASK_SESSION_UNLOCK = 8                ; Terminal Server session unlocked state change. For example, this state change causes the Task to run when the computer is unlocked

; TASK_STATE constants define the operational state of a Task
; https://docs.microsoft.com/en-us/windows/win32/taskschd/registeredtask-state
Global Const $TASK_STATE_UNKNOWN = 0                 ; The state of the Task is unknown
Global Const $TASK_STATE_DISABLED = 1                ; The Task is registered but is disabled and no instances of the Task are queued or running
;                                                      The Task cannot be run until it is enabled
Global Const $TASK_STATE_QUEUED = 2                  ; Instances of the Task are queued
Global Const $TASK_STATE_READY = 3                   ; The Task is ready to be executed, but no instances are queued or running
Global Const $TASK_STATE_RUNNING = 4                 ; One or more instances of the Task is running

; TASK_TRIGGER_TYPE2 constants define the type of Triggers that can be used by Tasks
; https://docs.microsoft.com/de-de/windows/win32/taskschd/trigger-type
Global Const $TASK_TRIGGER_EVENT = 0                 ; Starts the Task when a specific event occurs
Global Const $TASK_TRIGGER_TIME = 1                  ; Starts the Task at a specific time of day
Global Const $TASK_TRIGGER_DAILY = 2                 ; Starts the Task daily
Global Const $TASK_TRIGGER_WEEKLY = 3                ; Starts the Task weekly
Global Const $TASK_TRIGGER_MONTHLY = 4               ; Starts the Task monthly
Global Const $TASK_TRIGGER_MONTHLYDOW = 5            ; Starts the Task every month on a specific day of the week
Global Const $TASK_TRIGGER_IDLE = 6                  ; Starts the Task when the computer goes into an idle state
Global Const $TASK_TRIGGER_REGISTRATION = 7          ; Starts the Task when the Task is registered
Global Const $TASK_TRIGGER_BOOT = 8                  ; Starts the Task when the computer boots
Global Const $TASK_TRIGGER_LOGON = 9                 ; Starts the Task when a specific user logs on
Global Const $TASK_TRIGGER_SESSION_STATE_CHANGE = 11 ; Triggers the Task when a specific session state changes
; ===============================================================================================================================
#EndRegion #CONSTANTS#

; #CURRENT# =====================================================================================================================
;_TS_Open
;_TS_Close
;_TS_ErrorNotify
;_TS_ErrorText
;_TS_ActionCreate
;_TS_ActionDelete
;_TS_FolderCreate
;_TS_FolderDelete
;_TS_FolderExists
;_TS_FolderGet
;_TS_IsRunFromTaskScheduler - EXPERIMENTAL!
;_TS_RunningTaskList
;_TS_TaskCopy
;_TS_TaskCreate
;_TS_TaskDelete
;_TS_TaskExists
;_TS_TaskExportXML
;_TS_TaskGet
;_TS_TaskImportXML
;_TS_TaskList
;_TS_TaskListHeader
;_TS_TaskPropertiesGet
;_TS_TaskPropertiesSet
;_TS_TaskRegister
;_TS_TaskRun
;_TS_TaskStop
;_TS_TaskUpdate
;_TS_TaskValidate
;_TS_TriggerCreate
;_TS_TriggerDelete
;_TS_VersionInfo
;_TS_Wrapper_ActionCreate
;_TS_Wrapper_PrincipalSet
;_TS_Wrapper_TaskCreate
;_TS_Wrapper_TaskRegister
;_TS_Wrapper_TriggerDateTime
;_TS_Wrapper_TriggerLogon
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
;__TS_ErrorHandler
;__TS_TaskListWrite
;__TS_PropertyGetWrite
;__TS_ConvertDaysOfMonth
;__TS_ConvertDaysOfWeek
;__TS_ConvertMonthsOfYear
;__TS_ConvertWeeksOfMonth
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _TS_Open
; Description ...: Opens a connection to the Microsoft Task Scheduler Service on the local or a remote computer.
; Syntax.........: _TS_Open([$sComputer[, $sUser[, $sDomain[, $sPassword]]]])
; Parameters ....: $sComputer - [optional] The name of the computer that you want to connect to. If this parameter is empty, the function will connect to the local computer.
;                  $sUser     - [optional] The user name that is used during the connection to $sComputer. If the user is not specified, then the current token is used.
;                  $sDomain   - [optional] The domain of the user specified in the user parameter.
;                  $sPassword - [optional] The password that is used to connect to $sComputer. If user name and password are not specified, the current token is used.
; Return values .: Success - Object of the Task Scheduler Service
;                  Failure - Returns 0 and sets @error:
;                  |101 - Error creating the COM error handler. @extended is set to the error code returned by _TS_ErrorNotify
;                  |102 - Error creating the Task Scheduler Service. @extended is set to the COM error code
;                  |103 - Error connecting to the Task Scheduler Service. @extended is set to the COM error code:
;                  |      0x80070005 - Access is denied to connect to the Task Scheduler service.
;                  |      0x8007000e - The application does not have enough memory to complete the operation or
;                  |                   the user, password, or domain has at least one null and one non-null value.
;                  |      53         - This error is returned in the following situations:
;                  |                   The computer name specified in the serverName parameter does not exist.
;                  |                   When you are trying to connect to a Windows Server 2003 or Windows XP computer, and the remote computer does not have the
;                  |                       File and Printer Sharing firewall exception enabled or the Remote Registry service is not running.
;                  |                   When you are trying to connect to a Windows Vista computer, and the remote computer does not have the
;                  |                       Remote Scheduled Tasks Management firewall exception enabled and the File and Printer Sharing firewall exception enabled, or the Remote Registry service is not running.
;                  |      50         - The user, password, or domain parameters cannot be specified when connecting to a remote Windows XP or Windows Server 2003 computer from a Windows Vista computer.
; Author ........: water
; Modified ......:
; Remarks .......: $sComputer can be specified as name or IP-Address. Name is network\computername or \\computername
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Open($sComputer = Default, $sUser = Default, $sDomain = Default, $sPassword = Default)
	Local $oTS_Service
	_TS_ErrorNotify(4)
	If @error Then Return SetError(101, @error, 0)
	$oTS_Service = ObjCreate("Schedule.Service")
	If @error Then Return SetError(102, @error, 0)
	$oTS_Service.Connect($sComputer, $sUser, $sDomain, $sPassword)
	If @error Then Return SetError(103, @error, 0)
	Return $oTS_Service
EndFunc   ;==>_TS_Open

; #FUNCTION# ====================================================================================================================
; Name ..........: _TS_Close
; Description ...: Closes the connection to the Microsoft Task Scheduler Service.
; Syntax.........: _TS_Close([$oService = 0])
; Parameters ....: $oService - [optional] Object of the Task Scheduler Service created by _TS_Open
; Return values .: Success - 1
;                  Failure - None
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Close(ByRef $oService)
	$__iTS_Debug = 0
	$__sTS_DebugFile = @ScriptDir & "\TaskScheduler_Debug.txt"
	$__oTS_Error = 0
	If IsObj($oService) Then $oService = 0
	Return 1
EndFunc   ;==>_TS_Close

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_ErrorNotify
; Description ...: Sets or queries the debug level.
; Syntax.........: _TS_ErrorNotify($iDebug[, $sDebugFile = @ScriptDir & "\TaskScheduler_Debug.txt"])
; Parameters ....: $iDebug     - Debug level. Allowed values are:
;                  |-1 - Return the current settings
;                  |0  - Disable debugging
;                  |1  - Enable debugging. Output the debug info to the console
;                  |2  - Enable Debugging. Output the debug info to a MsgBox
;                  |3  - Enable Debugging. Output the debug info to a file defined by $sDebugFile
;                  |4  - Enable Debugging. The COM errors will be handled (the script no longer crashes) without any output
;                  $sDebugFile - [optional] File to write the debugging info to if $iDebug = 3 (Default = @ScriptDir & "TaskScheduler_Debug.txt")
; Return values .: Success - Depends on the value set for $iDebug.
;                  For $iDebug >= 0: 1, sets @extended to:
;                  |0 - The COM error handler for this UDF was already active
;                  |1 - A COM error handler has been initialized for this UDF
;                  For $iDebug = -1: A one based one-dimensional array with the following elements:
;                  |1 - Debug level. Value from 0 to 3. Check parameter $iDebug for details
;                  |2 - Debug file. File to write the debugging info to as defined by parameter $sDebugFile
;                  |3 - True if the COM error handler has been defined for this UDF. False if debugging is set off or a COM error handler was already defined
;                  Failure - 0, sets @error to:
;                  |201 - $iDebug is not an integer or < -1 or > 4
;                  |202 - Installation of the custom error handler failed. @extended is set to the error code returned by ObjEvent
;                  |203 - COM error handler already set to another function
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_ErrorNotify($iDebug, $sDebugFile = "")
	If $sDebugFile = Default Then $sDebugFile = ""
	If Not IsInt($iDebug) Or $iDebug < -1 Or $iDebug > 4 Then Return SetError(201, 0, 0)
	If $sDebugFile = "" Then $sDebugFile = @ScriptDir & "\TaskScheduler_Debug.txt"
	Switch $iDebug
		Case -1
			Local $avDebug[4] = [3]
			$avDebug[1] = $__iTS_Debug
			$avDebug[2] = $__sTS_DebugFile
			$avDebug[3] = IsObj($__oTS_Error)
			Return $avDebug
		Case 0
			$__iTS_Debug = 0
			$__sTS_DebugFile = ""
			$__oTS_Error = 0
		Case Else
			$__iTS_Debug = $iDebug
			$__sTS_DebugFile = $sDebugFile
			; A COM error handler will be initialized only if one does not exist
			If ObjEvent("AutoIt.Error") = "" Then
				$__oTS_Error = ObjEvent("AutoIt.Error", "__TS_ErrorHandler") ; Creates a custom error handler
				If @error <> 0 Then Return SetError(202, @error, 0)
				Return SetError(0, 1, 1)
			ElseIf ObjEvent("AutoIt.Error") = "__TS_ErrorHandler" Then
				Return SetError(0, 0, 1) ; COM error handler already set by a call to this function
			Else
				Return SetError(203, 0, 0) ; COM error handler already set to another function
			EndIf
	EndSwitch
	Return 1
EndFunc   ;==>_TS_ErrorNotify

; #FUNCTION# ====================================================================================================================
; Name ..........: _TS_ErrorText
; Description ...: Returns the full error message for an UDF function or the error/success message for the Task Scheduler API.
; Syntax.........: _TS_ErrorMessage($iErrorNumber[, $bPrefix = True])
; Parameters ....: $iErrorNumber - Integer of the error returned by a function of this UDF or the Task Scheduler API
;                  $bPrefix      - [optional] True prefixes the error with functionname/errorname and errornumber (default = True)
; Return values .: Success - If $bPrefix = True: FunctionName (ErrorNumber): ErrorText, else: ErrorText
;                  Failure - Returns "" and sets @error:
;                  |301 - Specified error number could not be found
; Author ........: water
; Modified ......:
; Remarks .......: The Task Scheduler APIs return error/success as an HRESULT value e.g:
;                    $bPrefix = True: SCHED_S_TASK_NOT_SCHEDULED (0x00041305): One or more of the properties that are needed to run this task on a schedule have not been set.
;                    $bPrefix = False: One or more of the properties that are needed to run this task on a schedule have not been set.
;                  "S" as the second token of the APIs errorname denotes a success value, "E" an error value.
;+
;                  $bPrefix = True: ErrorNumber will be returned in hex notation (e.g. 0x00041305) if $iErrorNumber < 0
; Related .......:
; Link ..........: https://docs.microsoft.com/de-de/windows/win32/taskschd/task-scheduler-error-and-success-constants
; Example .......: Yes
; ===============================================================================================================================
Func _TS_ErrorText($iErrorNumber, $bPrefix = True)
	If $bPrefix = Default Then $bPrefix = True
	Local $aErrorMessages[][] = [ _
			["_TS_Open", 101, "Error creating the COM error handler. @extended is set to the error code returned by _TS_ErrorNotify"], _
			["_TS_Open", 102, "Error creating the Task Scheduler Service. @extended is set to the COM error code"], _
			["_TS_Open", 103, "Error connecting to the Task Scheduler Service. @extended is set to the COM error code:" & @CRLF & _
			"  0x80070005 - Access is denied to connect to the Task Scheduler service." & @CRLF & _
			"  0x8007000e - The application does not have enough memory to complete the operation or" & @CRLF & _
			"               the user, password, or domain has at least one null and one non-null value." & @CRLF & _
			"  53         - This error is returned in the following situations:" & @CRLF & _
			"               The computer name specified in the serverName parameter does not exist." & @CRLF & _
			"               When you are trying to connect to a Windows Server 2003 or Windows XP computer, and the remote computer does not have the" & @CRLF & _
			"               File and Printer Sharing firewall exception enabled or the Remote Registry service is not running." & @CRLF & _
			"               When you are trying to connect to a Windows Vista computer, and the remote computer does not have theRemote Scheduled Tasks Management firewall exception enabled " & @CRLF & _
			"               and the File and Printer Sharing firewall exception enabled, or the Remote Registry service is not running." & @CRLF & _
			"  50         - The user, password, or domain parameters cannot be specified when connecting to a remote Windows XP or Windows Server 2003 computer from a Windows Vista computer."], _
			["_TS_ErrorNotify", 201, "Return values .: Success (for $iDebug => 0) - 1, sets @extended to:" & @CRLF & _
			"0 - The COM error handler for this UDF was already active" & @CRLF & _
			"1 - A COM error handler has been initialized for this UDF" & @CRLF & _
			"Success (for $iDebug = -1) - one based one-dimensional array with the following elements:" & @CRLF & _
			"1 - Debug level. Value from 0 to 3. Check parameter $iDebug for details" & @CRLF & _
			"2 - Debug file. File to write the debugging info to as defined by parameter $sDebugFile" & @CRLF & _
			"3 - True if the COM error handler has been defined for this UDF. False if debugging is set off or a COM error handler was already defined" & @CRLF & _
			"Failure - 0, sets @error to:" & @CRLF & _
			"201 - $iDebug is not an integer or < -1 or > 4" & @CRLF & _
			"202 - Installation of the custom error handler failed. @extended is set to the error code returned by ObjEvent" & @CRLF & _
			"203 - COM error handler already set to another function"], _
			["_TS_Errortext", 301, "Specified error number could not be found"], _
			["_TS_ActionCreate", 401, "$oTaskDefinition isn't an object or not a Task Definition object"], _
			["_TS_ActionCreate", 402, "The Action could not be created. @extended is set to the COM error code"], _
			["_TS_ActionDelete", 502, "The Action could not be deleted. @extended is set to the COM error code"], _
			["_TS_ActionDelete", 503, "The Actions could not be deleted. @extended is set to the COM error code"], _
			["_TS_ActionDelete", 504, "Either $iIndex or $sID has to be specified when $bDeleteAll is set to False"], _
			["_TS_FolderCreate", 601, "Error accessing the TaskFolder collection. @extended is set to the COM error code"], _
			["_TS_FolderCreate", 602, "Specified $sFolder already exists"], _
			["_TS_FolderCreate", 603, "Error creating the specified TaskFolder. @extended is set to the COM error code"], _
			["_TS_FolderDelete", 701, "Error accessing the TaskFolder collection. @extended is set to the COM error code"], _
			["_TS_FolderDelete", 702, "Specified $sFolder does not exist. @extended is set to the COM error code"], _
			["_TS_FolderDelete", 703, "You can't delete a Folder before all subfolders have been deleted. @extended is set to the COM error code"], _
			["_TS_FolderDelete", 704, "Error deleting the specified TaskFolder. @extended is set to the COM error code"], _
			["_TS_FolderExists", 801, "Error accessing the parent Folder of $sFolder (GetFolder). @extended is set to the COM error code"], _
			["_TS_FolderExists", 802, "Error accessing the Taskfolder collection (GetFolders). @extended is set to the COM error code"], _
			["_TS_FolderGet", 901, "Error accessing the specified Folder. @extended is set to the COM error code"], _
			["_TS_RunningTaskList", 1001, "Error retrieving the RunningTasks collection. @extended is set to the COM error code"], _
			["_TS_TaskCopy", 1101, "Error returned when reading the source Task. @extended is set to the error returned by _TS_TaskExportXML"], _
			["_TS_TaskCopy", 1102, "Error returned when creating the target Task. @extended is set to the error code returned by _TS_TaskImportXML"], _
			["_TS_TaskCreate", 1201, "Error creating the Task Definition. @extended is set to the COM error code"], _
			["_TS_TaskDelete", 1301, "The specified Task Folder does not exist. @extended is set to the @error returned by _TS_FolderExists"], _
			["_TS_TaskDelete", 1302, "The specified Task does not exist. @extended is set to the @error returned by _TS_TaskExists"], _
			["_TS_TaskDelete", 1303, "Error accessing the Task Folder. @extended is set to the @error returned by _TS_FolderGet"], _
			["_TS_TaskDelete", 1304, "Error deleting the Task. @extended is set to the COM error code"], _
			["_TS_TaskExists", 1401, "Error accessing the specified Taskfolder. @extended is set to the COM error code"], _
			["_TS_TaskExists", 1402, "Error retrieving the Tasks collection. @extended is set to the COM error code"], _
			["_TS_TaskExportXML", 1501, "Error returned by _TS_TaskGet. @extended is set to the COM error code"], _
			["_TS_TaskExportXML", 1502, "Error opening the output file."], _
			["_TS_TaskGet", 1601, "Error accessing the specified Taskfolder. @extended is set to the COM error code"], _
			["_TS_TaskGet", 1602, "Error retrieving the Tasks collection. @extended is set to the COM error code"], _
			["_TS_TaskGet", 1603, "Task with the specified name could not be found in the Task Folder"], _
			["_TS_TaskImportXML", 1702, "Error creating a new Task Definition. @extended is set to the COM error code"], _
			["_TS_TaskImportXML", 1703, "Error creating a XML from the passed array. @extended is set to the COM error code"], _
			["_TS_TaskImportXML", 1704, "Error creating a XML from the passed file. @extended is set to the COM error code"], _
			["_TS_TaskList", 1801, "Error accessing the specified Taskfolder. @extended is set to the COM error code"], _
			["_TS_TaskPropertiesGet", 1901, "Error returned by _TS_TaskGet. @extended is set to the COM error code. Most probably the Task could not be found"], _
			["_TS_TaskPropertiesSet", 2001, "Unsupported or invalid Task Scheduler COM object"], _
			["_TS_TaskPropertiesSet", 2002, "Unsupported or invalid property name. @extended is set to the zero based index of the property in error"], _
			["_TS_TaskPropertiesSet", 2003, "The row in $aProperties does not have the required format: 'object name|property name|property value'. @extended is set to the index of the row in error."], _
			["_TS_TaskPropertiesSet", 2004, "$oObject is invalid. Must be: TaskDefinition, RegisteredTask, Trigger or Action"], _
			["_TS_TaskRegister", 2101, "Parameter $oService is not an object or not an ITaskService object"], _
			["_TS_TaskRegister", 2102, "$sFolder does not exist or an error occurred in _TS_FolderExists. @extended is set to the COM error (if any)"], _
			["_TS_TaskRegister", 2103, "Task exists which is incompatible with flags $TASK_CREATE, $TASK_DISABLE and $TASK_CREATE or" & @CRLF & _
			"Task does not exist which is incompatible with flags $TASK_UPDATE And $TASK_DONT_ADD_PRINCIPAL_ACE"], _
			["_TS_TaskRegister", 2104, "Parameter $oTaskDefinition is not an object or not an ITaskDefinition object"], _
			["_TS_TaskRegister", 2105, "Error accessing $sFolder using _TS_FolderGet. @extended is set to the COM error"], _
			["_TS_TaskRegister", 2106, "Error creating the Task. @extended is set to the COM error"], _
			["_TS_TaskRun", 2201, "The Task does not exist. @extended is set to the COM error code returned by _TS_TaskGet"], _
			["_TS_TaskRun", 2202, "Error starting the Task. @extended is set to the COM error"], _
			["_TS_TaskStop", 2301, "The Task does not exist. @extended is set to the COM error code returned by _TS_TaskGet"], _
			["_TS_TaskStop", 2302, "Error stopping the Task. @extended is set to the COM error"], _
			["_TS_TaskValidate", 2401, "The Task does not exist. @extended is set to the COM error code returned by _TS_TaskGet"], _
			["_TS_TaskValidate", 240101, "You have to define at least one Action"], _
			["_TS_TaskValidate", 240102, "Action type is unsupported"], _
			["_TS_TaskValidate", 240103, "Action ID has to be unique"], _
			["_TS_TaskValidate", 240501, "You should at least define one Trigger"], _
			["_TS_TaskValidate", 245001, "Make sure to provide Userid and Password for the selected logon type"], _
			["_TS_TriggerCreate", 2501, "$oTaskDefinition isn't an object or not a Task Definition object"], _
			["_TS_TriggerCreate", 2502, "The Trigger could not be created. @extended is set to the COM error code"], _
			["_TS_TriggerDelete", 2602, "The Trigger could not be deleted. @extended is set to the COM error code"], _
			["_TS_TriggerDelete", 2603, "The Triggers could not be deleted. @extended is set to the COM error code"], _
			["_TS_TriggerDelete", 2604, "Either $iIndex or $sID has to be specified when $bDeleteAll is set to False"], _
			["_TS_Wrapper_ActionCreate", 2701, "Error returned when accessing the Actions collection. @extended is set to the COM error code"], _
			["_TS_Wrapper_ActionCreate", 2702, "Error returned when creating the Action object. @extended is set to the COM error code"], _
			["_TS_Wrapper_ActionCreate", 2703, "Parameter $oTaskDefinition is not an object or not an ITaskDefinition object"], _
			["_TS_Wrapper_PrincipalSet", 2801, "Error creating the Task Definition. @extended is set to the COM error code"], _
			["_TS_Wrapper_PrincipalSet", 2802, "Parameter $oTaskDefinition is not an object or not an ITaskDefinition object"], _
			["_TS_Wrapper_TaskCreate", 2901, "Error creating the Task Definition. @extended is set to the COM error code"], _
			["_TS_Wrapper_TaskCreate", 2902, "Error setting property Date. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TaskCreate", 2903, "Parameter $oService is not an object or not an ITaskService object"], _
			["_TS_Wrapper_TriggerDateTime", 3101, "Invalid $iTriggerType specified. Has to be $TASK_TRIGGER_TIME, $TASK_TRIGGER_DAILY or $TASK_TRIGGER_WEEKLY"], _
			["_TS_Wrapper_TriggerDateTime", 3102, "Error returned when creating the Trigger object. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3103, "Error setting property StartBoundary. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3104, "Error setting property EndBoundary. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3105, "Error setting property ExecutionTimeLimit. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3106, "Error setting property DaysOfWeek. Please check the correct format. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3107, "Error setting property Weeksinterval. Please check the correct format. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3108, "Error setting property DaysInterval. Please check the correct format. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerDateTime", 3109, "Parameter $oTaskDefinition is not an object or not an ITaskDefinition object"], _
			["_TS_Wrapper_TriggerDateTime", 3110, "Error returned when accessing the Triggers collection. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerLogon", 3201, "Error returned when creating the Trigger object. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerLogon", 3202, "Error setting property StartBoundary. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerLogon", 3203, "Error setting property EndBoundary. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerLogon", 3204, "Error setting property ExecutionTimeLimit. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerLogon", 3205, "Error setting property Delay. Please check the correct format as described above. @extended is set to the COM error code"], _
			["_TS_Wrapper_TriggerLogon", 3206, "Parameter $oTaskDefinition is not an object or not an ITaskDefinition object"], _
			["_TS_Wrapper_TriggerLogon", 3207, "Error returned when accessing the Triggers collection. @extended is set to the COM error code"], _
			["_TS_TaskUpdate", 3301, "Parameter $oService is not an object or not an ITaskService object"], _
			["_TS_TaskUpdate", 3302, "Parameter $oTask is not an object or not an IRegisteredTask object"], _
			["_TS_TaskUpdate", 3303, "Error accessing $sFolder using _TS_FolderGet. @extended is set to the COM error"], _
			["_TS_TaskUpdate", 3304, "Error updating the Task. @extended is set to the COM error"], _
			["_TS_TaskUpdate", 3305, "Parameter $oTaskDefinition is not an object or not an ITaskDefinition object"], _
			["_TS_IsRunFromTaskScheduler", 3401, "_TS_RunningTaskList returned an error. @extended has been set to this error code"], _
			["_TS_IsRunFromTaskScheduler", 3402, "_WinAPI_GetParentProcess returned an error. @extended has been set to this error code"], _
			["SCHED_S_TASK_READY", 0x00041300, "The task is ready to run at its next scheduled time."], _
			["SCHED_S_TASK_RUNNING", 0x00041301, "The task is currently running."], _
			["SCHED_S_TASK_DISABLED", 0x00041302, "The task will not run at the scheduled times because it has been disabled."], _
			["SCHED_S_TASK_HAS_NOT_RUN", 0x00041303, "The task has not yet run."], _
			["SCHED_S_TASK_NO_MORE_RUNS", 0x00041304, "There are no more runs scheduled for this task."], _
			["SCHED_S_TASK_NOT_SCHEDULED", 0x00041305, "One or more of the properties that are needed to run this task on a schedule have not been set."], _
			["SCHED_S_TASK_TERMINATED", 0x00041306, "The last run of the task was terminated by the user."], _
			["SCHED_S_TASK_NO_VALID_TRIGGERS", 0x00041307, "Either the task has no triggers or the existing triggers are disabled or not set."], _
			["SCHED_S_EVENT_TRIGGER", 0x00041308, "Event triggers do not have set run times."], _
			["SCHED_E_TRIGGER_NOT_FOUND", 0x80041309, "A task's trigger is not found."], _
			["SCHED_E_TASK_NOT_READY", 0x8004130A, "One or more of the properties required to run this task have not been set."], _
			["SCHED_E_TASK_NOT_RUNNING", 0x8004130B, "There is no running instance of the task."], _
			["SCHED_E_SERVICE_NOT_INSTALLED", 0x8004130C, "The Task Scheduler service is not installed on this computer."], _
			["SCHED_E_CANNOT_OPEN_TASK", 0x8004130D, "The task object could not be opened."], _
			["SCHED_E_INVALID_TASK", 0x8004130E, "The object is either an invalid task object or is not a task object."], _
			["SCHED_E_ACCOUNT_INFORMATION_NOT_SET", 0x8004130F, "No account information could be found in the Task Scheduler security database for the task indicated."], _
			["SCHED_E_ACCOUNT_NAME_NOT_FOUND", 0x80041310, "Unable to establish existence of the account specified."], _
			["SCHED_E_ACCOUNT_DBASE_CORRUPT", 0x80041311, "Corruption was detected in the Task Scheduler security database; the database has been reset."], _
			["SCHED_E_NO_SECURITY_SERVICES", 0x80041312, "Task Scheduler security services are available only on Windows NT."], _
			["SCHED_E_UNKNOWN_OBJECT_VERSION", 0x80041313, "The task object version is either unsupported or invalid."], _
			["SCHED_E_UNSUPPORTED_ACCOUNT_OPTION", 0x80041314, "The task has been configured with an unsupported combination of account settings and run time options."], _
			["SCHED_E_SERVICE_NOT_RUNNING", 0x80041315, "The Task Scheduler Service is not running."], _
			["SCHED_E_UNEXPECTEDNODE", 0x80041316, "The task XML contains an unexpected node."], _
			["SCHED_E_NAMESPACE", 0x80041317, "The task XML contains an element or attribute from an unexpected namespace."], _
			["SCHED_E_INVALIDVALUE", 0x80041318, "The task XML contains a value which is incorrectly formatted or out of range."], _
			["SCHED_E_MISSINGNODE", 0x80041319, "The task XML is missing a required element or attribute."], _
			["SCHED_E_MALFORMEDXML", 0x8004131A, "The task XML is malformed."], _
			["SCHED_S_SOME_TRIGGERS_FAILED", 0x0004131B, "The task is registered, but not all specified triggers will start the task."], _
			["SCHED_S_BATCH_LOGON_PROBLEM", 0x0004131C, "The task is registered, but may fail to start. Batch logon privilege needs to be enabled for the task principal."], _
			["SCHED_E_TOO_MANY_NODES", 0x8004131D, "The task XML contains too many nodes of the same type."], _
			["SCHED_E_PAST_END_BOUNDARY", 0x8004131E, "The task cannot be started after the trigger end boundary."], _
			["SCHED_E_ALREADY_RUNNING", 0x8004131F, "An instance of this task is already running."], _
			["SCHED_E_USER_NOT_LOGGED_ON", 0x80041320, "The task will not run because the user is not logged on."], _
			["SCHED_E_INVALID_TASK_HASH", 0x80041321, "The task image is corrupt or has been tampered with."], _
			["SCHED_E_SERVICE_NOT_AVAILABLE", 0x80041322, "The Task Scheduler service is not available."], _
			["SCHED_E_SERVICE_TOO_BUSY", 0x80041323, "The Task Scheduler service is too busy to handle your request. Please try again later."], _
			["SCHED_E_TASK_ATTEMPTED", 0x80041324, "The Task Scheduler service attempted to run the task, but the task did not run due to one of the constraints in the task definition."], _
			["SCHED_S_TASK_QUEUED", 0x00041325, "The Task Scheduler service has asked the task to run."], _
			["SCHED_E_TASK_DISABLED", 0x80041326, "The task is disabled."], _
			["SCHED_E_TASK_NOT_V1_COMPAT", 0x80041327, "The task has properties that are not compatible with earlier versions of Windows."], _
			["SCHED_E_START_ON_DEMAND", 0x80041328, "The task settings do not allow the task to start on demand."], _
			["XXXXXXXX", 9999, "Dummy error message!"]]
	; Aliases for duplicate error messages
	If $iErrorNumber = 202 Then $iErrorNumber = 201
	If $iErrorNumber = 203 Then $iErrorNumber = 201

	For $i = 0 To UBound($aErrorMessages, 1) - 1
		If $aErrorMessages[$i][1] = $iErrorNumber Then
			If $bPrefix Then
				If $aErrorMessages[$i][1] < 0 Then
					Return $aErrorMessages[$i][0] & " (0x" & Hex($aErrorMessages[$i][1]) & "): " & $aErrorMessages[$i][2]
				Else
					Return $aErrorMessages[$i][0] & " (" & $aErrorMessages[$i][1] & "): " & $aErrorMessages[$i][2]
				EndIf
			EndIf
			Return $aErrorMessages[$i][2]
		EndIf
	Next
	Return SetError(301, 0, "") ; Error message not found
EndFunc   ;==>_TS_ErrorText

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_ActionCreate
; Description ...: Create a new Action object for a new or already Registered Task.
; Syntax.........: _TS_ActionCreate($oTaskDefinition, $iActionType[, $sID = ""])
; Parameters ....: $oTaskDefinition - Task Definition object to add this Action to.
;                  $iActionType     - The Action type to be created. Can be $TASK_ACTION_EXEC or $TASK_ACTION_COM_HANDLER
;                                     of the TASK_ACTION_TYPE enumeration.All other types are no longer supported by MS.
;                  $sID             - [optional] ID for easier access to the object
; Return values .: Success - Action object
;                  Failure - Returns 0 and sets @error:
;                  |401 - $oTaskDefinition isn't an object or not a Task Definition object
;                  |402 - The Action could not be created. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_ActionCreate($oTaskDefinition, $iActionType, $sID = "")
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(401, @error, 0)
	Local $oAction = $oTaskDefinition.Actions.Create($iActionType)
	If @error Then Return SetError(402, @error, 0)
	If $sID <> "" Then $oAction.ID = $sID
	Return $oAction
EndFunc   ;==>_TS_ActionCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_ActionDelete
; Description ...: Delete a single or all Action objects by ID or index.
; Syntax.........: _TS_ActionDelete($oTaskDefinition, $iIndex[, $sID = ""[, $bDeleteAll = False]])
; Parameters ....: $oTaskDefinition - Task Definition object of a new or Registered Task
;                  $iIndex          - Delete the Action with the specified index (one based)
;                  $sID             - [optional] Deletes all Actions with the same ID (default = "")
;                  $bDeleteAll      - [optional] Removes all Actions (default = False)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |502 - The Action could not be deleted. @extended is set to the COM error code
;                  |503 - The Actions could not be deleted. @extended is set to the COM error code
;                  |504 - Either $iIndex or $sID has to be specified when $bDeleteAll is set to False
; Author ........: water
; Modified.......:
; Remarks .......: Set one of this three parameters to delete specific or all Actions: $iIndex, $sID, $bDeleteAll.
;                  The parameters will be processed in the following sequence:
;                  If $iIndex > 0 then delete by index, else if $sID <> "" then delete by ID, else if $bDeleteAll is True then delete all Actions
;+
;                  When you used $bDeleteAll = True you have to create at least one Action object as you can't register a Task without Actions.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_ActionDelete($oTaskDefinition, $iIndex, $sID = "", $bDeleteAll = False)
	If $iIndex = Default Then $iIndex = 0
	If $sID = Default Then $sID = ""
	If $bDeleteAll = Default Then $bDeleteAll = False
	If $iIndex = 0 And $sID = "" And $bDeleteAll = False Then Return SetError(504, 0, 0)
	If $iIndex > 0 Then
		$oTaskDefinition.Actions.Remove($iIndex)
		If @error Then Return SetError(502, @error, 0)
	ElseIf $sID <> "" Then
		For $i = 1 To $oTaskDefinition.Actions.Count
			If $oTaskDefinition.Actions.Item($i).ID = $sID Then
				$oTaskDefinition.Actions.Remove($i)
				If @error Then Return SetError(502, @error, 0)
				ExitLoop
			EndIf
		Next
	ElseIf $bDeleteAll Then
		$oTaskDefinition.Actions.Clear()
		If @error Then Return SetError(503, @error, 0)
	EndIf
	Return 1
EndFunc   ;==>_TS_ActionDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_FolderCreate
; Description ...: Creates a Task Folder.
; Syntax.........: _TS_FolderCreate($oService, $sFolder)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $sFolder  - The name that is used to identify the Folder. It is created on the root Folder
;                  If "FolderName\SubFolder1\SubFolder2" is specified, the entire Folder tree will be created if the Folders do not exist
; Return values .: Success - Object of the created Task Folder
;                  Failure - Returns 0 and sets @error:
;                  |601 - Error accessing the TaskFolder collection. @extended is set to the COM error code
;                  |602 - Specified $sFolder already exists
;                  |603 - Error creating the specified TaskFolder. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: $sFolder has always to start at the root Folder (means you have to specify all the Folders from the root down even when they already exist)
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_FolderCreate($oService, $sFolder)
	Local $oFolder = $oService.GetFolder("\")
	If @error Then Return SetError(601, @error, 0)
	$oService.GetFolder($sFolder)
	If @error = 0 Then Return SetError(602, 0, 0)
	Local $oFolderCreated = $oFolder.CreateFolder($sFolder)
	If @error Then Return SetError(603, @error, 0)
	Return $oFolderCreated
EndFunc   ;==>_TS_FolderCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_FolderDelete
; Description ...: Deletes a Task Folder.
; Syntax.........: _TS_FolderDelete($oService, $sFolder)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $sFolder  - The absolute path to the Folder to be deleted
;                    e.g. \Folder-Level1\Folder-Level2. No trailing backslash allowed.
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |701 - Error accessing the TaskFolder collection. @extended is set to the COM error code
;                  |702 - Specified $sFolder does not exist. @extended is set to the COM error code
;                  |703 - You can't delete a Folder before all subfolders have been deleted. @extended is set to the COM error code
;                  |704 - Error deleting the specified TaskFolder. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: Before you can delete a Folder you have to delete all subfolders
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_FolderDelete($oService, $sFolder)
	Local $oRoot = $oService.GetFolder("\")
	If @error Then Return SetError(701, @error, 0)
	Local $oFolder = $oService.GetFolder($sFolder)
	If @error Then Return SetError(702, @error, 0)
	Local $oSubFolders = $oFolder.GetFolders(0)
	If $oSubFolders.Count > 0 Then Return SetError(703, @error, 0)
	$oRoot.DeleteFolder($sFolder, 0) ; Parameter 2 is mandatory but not supported
	If @error Then Return SetError(704, @error, 0)
	Return 1
EndFunc   ;==>_TS_FolderDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_FolderExists
; Description ...: Checks if a Task Folder exists.
; Syntax.........: _TS_FolderExists($oService, $sFolder)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $sFolder  - The absolute path to the Folder to be checked
;                    e.g. \Folder-Level1\Folder-Level2. No trailing backslash allowed.
; Return values .: Success - 1 when Folder was found and 0 when Folder was not found
;                  Failure - Returns 0 and sets @error:
;                  |801 - Error accessing the parent Folder of $sFolder (GetFolder). @extended is set to the COM error code
;                  |802 - Error accessing the Taskfolder collection (GetFolders). @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_FolderExists($oService, $sFolder)
	Local $oFolder, $oFolders, $sFolder1, $sFolder2, $iPos
	If $sFolder = "\" Then Return 1
	$iPos = StringInStr($sFolder, "\", $STR_NOCASESENSE, -1)
	$sFolder1 = ($iPos = 1) ? "\" : (StringLeft($sFolder, $iPos - 1))
	$sFolder2 = StringMid($sFolder, $iPos + 1)
	$oFolder = $oService.GetFolder($sFolder1)
	If @error Then Return SetError(801, @error, 0)
	$oFolders = $oFolder.GetFolders(0)
	If @error Then Return SetError(802, @error, 0)
	For $oFolder In $oFolders
		If $oFolder.Name = $sFolder2 Then Return 1
	Next
	Return 0
EndFunc   ;==>_TS_FolderExists

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_FolderGet
; Description ...: Returns the object of the specified Folder.
; Syntax.........: _TS_FolderGet($oService, $sFolder)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $sFolder  - The absolute path to the Folder to be processed
;                    e.g. \Folder-Level1\Folder-Level2. No trailing backslash allowed.
; Return values .: Success - Object of the specified Folder
;                  Failure - Returns 0 and sets @error:
;                  |901 - Error accessing the specified Folder. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_FolderGet($oService, $sFolder)
	Local $oFolder = $oService.GetFolder($sFolder)
	If @error Then Return SetError(901, @error, 0)
	Return $oFolder
EndFunc   ;==>_TS_FolderGet

#cs REMOVE THIS COMMENT TO USE THIS FUNCTION
; #FUNCTION# ====================================================================================================================
; Name ..........: _TS_IsRunFromTaskScheduler
; Description ...: Check if process is run by Task Scheduler. - EXPERIMENTAL!
; Syntax ........: _TS_IsRunFromTaskScheduler($oService[, $iPID = @AutoItPID[, $bCheckPath = False]])
; Parameters ....: $oService   - Task Scheduler Service object as returned by _TS_Open
;                  $iPID       - [optional] ProcessID of the process to check (default = @AutoItPID (current process))
;                  $bCheckPath - [optional] Boolen value that specifies whether to check the path as well, not just the PID (default = False)
; Return values .: Success - True or False
;                  Failure - Returns 0 and sets @error:
;                  |3401 - _TS_RunningTaskList returned an error. @extended has been set to this error code.
;                  |3402 - _WinAPI_GetParentProcess returned an error. @extended has been set to this error code.
; Author ........: mLipok
; Modified ......: water
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _TS_IsRunFromTaskScheduler($oService, $iPID = @AutoItPID, $bCheckPath = False)
	If $iPID = Default Then $iPID = @AutoItPID
	If $bCheckPath = Default Then $bCheckPath = False
	Local $aRunningTasks = _TS_RunningTaskList($oService, 1)
	If @error Then Return SetError(3401, @error, 0)
	Local $iParentPID = _WinAPI_GetParentProcess($iPID)
	If @error Then Return SetError(3402, @error, 0)
	Local $bTestingPath = False
	For $i = 0 To UBound($aRunningTasks) - 1
		$bTestingPath = ($bCheckPath = False) Or ($aRunningTasks[$i][0] = @ScriptFullPath)
		If $bTestingPath And ($aRunningTasks[$i][1] = $iPID Or $aRunningTasks[$i][1] = $iParentPID) Then Return True
	Next
	Return False
EndFunc   ;==>_TS_IsRunFromTaskScheduler
#ce REMOVE THIS COMMENT TO USE THIS FUNCTION

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_RunningTaskList
; Description ...: Returns a list of all currently running Tasks.
; Syntax.........: _TS_RunningTaskList($oService[, $iShowHidden = 0])
; Parameters ....: $oService    - Task Scheduler Service object as returned by _TS_Open
;                  $iShowHidden - [optional] Returns hidden Tasks as well when set to 1 (default = 0)
; Return values .: Success - two-dimensional zero based array with the following information:
;                  |0 - CurrentAction: Name of the current action that the running task is performing
;                  |1 - EnginePID: Process ID for the engine (process) which is running the task
;                  |2 - InstanceGuid: GUID identifier for this instance of the task
;                  |3 - Name: Name of the task
;                  |4 - Path: Path to where the task is stored
;                  |5 - State: State of the running task (usually $TASK_STATE_RUNNING)
;                  Failure - Returns "" and sets @error:
;                  |1001 - Error retrieving the RunningTasks collection. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_RunningTaskList($oService, $iShowHidden = 0)
	If $iShowHidden = Default Then $iShowHidden = 0
	Local $oRunningTasks = $oService.GetRunningTasks($iShowHidden)
	If @error Then Return SetError(1001, @error, "")
	Local $iTaskCount = $oRunningTasks.Count, $iIndex = 0, $aRunningTasks[$iTaskCount][6]
	For $oRunningTask In $oRunningTasks
		$aRunningTasks[$iIndex][0] = $oRunningTask.CurrentAction
		$aRunningTasks[$iIndex][1] = $oRunningTask.EnginePID
		$aRunningTasks[$iIndex][2] = $oRunningTask.InstanceGUID
		$aRunningTasks[$iIndex][3] = $oRunningTask.Name
		$aRunningTasks[$iIndex][4] = $oRunningTask.Path
		$aRunningTasks[$iIndex][5] = $oRunningTask.State
		$iIndex += 1
	Next
	Return $aRunningTasks
EndFunc   ;==>_TS_RunningTaskList

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskCopy
; Description ...: Copies the definition of an existing Task to a new Task Definition object.
; Syntax.........: _TS_TaskCopy($oService, $sSourceTaskPath)
; Parameters ....: $oService        - Task Scheduler Service object as returned by _TS_Open
;                  $sSourceTaskPath - Task Folder(s) and Task name of the source Task e.g. \folder1\folder1-1\source-task-name
; Return values .: Success - Task Definition object
;                  Failure - Returns 0 and sets @error:
;                  |1101 - Error returned when reading the source Task. @extended is set to the error returned by _TS_TaskExportXML
;                  |1102 - Error returned when creating the target Task. @extended is set to the error code returned by _TS_TaskImportXML
; Author ........: water
; Modified.......:
; Remarks .......: You can modify the Task Definition as needed before calling _TS_TaskRegister to create a new Task in the same or a different folder.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskCopy($oService, $sSourceTaskPath)
	Local $aSourceTaskXML = _TS_TaskExportXML($oService, $sSourceTaskPath)
	If @error Then Return SetError(1101, @error, 0)
	Local $oTaskDefinition = _TS_TaskImportXML($oService, 2, $aSourceTaskXML)
	If @error Then Return SetError(1102, @error, 0)
	Return $oTaskDefinition
EndFunc   ;==>_TS_TaskCopy

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskCreate
; Description ...: Create a Task Definition object so you can then set all needed properties using _TS_TaskPropertiesSet.
; Syntax.........: _TS_TaskCreate($oService)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
; Return values .: Success - Task Definition object
;                  Failure - Returns 0 and sets @error:
;                  |1201 - Error creating the Task Definition. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskCreate($oService)
	; Set the TaskDefinition object. The flags parameter is 0 because it is not supported
	Local $oTaskDefinition = $oService.NewTask(0)
	If @error Then Return SetError(1201, @error, 0)
	Return $oTaskDefinition
EndFunc   ;==>_TS_TaskCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskDelete
; Description ...: Delete a Task from the specified Task Folder.
; Syntax.........: _TS_TaskDelete($oService, $sTaskPath)
; Parameters ....: $oService  - Task Scheduler Service object as returned by _TS_Open
;                  $sTaskPath - Task path (Folder plus Task name) to be deleted e.g. \folder1\folder1-1\task-name
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |1301 - The specified Task Folder does not exist. @extended is set to the @error returned by _TS_FolderExists
;                  |1302 - The specified Task does not exist. @extended is set to the @error returned by _TS_TaskExists
;                  |1303 - Error accessing the Task Folder. @extended is set to the @error returned by _TS_FolderGet
;                  |1304 - Error deleting the Task. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskDelete($oService, $sTaskPath)
	Local $iPos = StringInStr($sTaskPath, "\", $STR_NOCASESENSE, -1)
	Local $sFolder = ($iPos = 1) ? "\" : (StringLeft($sTaskPath, $iPos - 1))
	Local $sTask = StringMid($sTaskPath, $iPos + 1)
	If Not _TS_FolderExists($oService, $sFolder) Then Return SetError(1301, @error, 0)
	If Not _TS_TaskExists($oService, $sTaskPath) Then Return SetError(1302, @error, 0)
	Local $oFolder = _TS_FolderGet($oService, $sFolder)
	If @error Then Return SetError(1303, @error, 0)
	$oFolder.DeleteTask($sTask, 0)
	If @error Then Return SetError(1304, @error, 0)
	Return 1
EndFunc   ;==>_TS_TaskDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskExists
; Description ...: Checks if a Task exists.
; Syntax.........: _TS_TaskExists($oService, $sTaskPath)
; Parameters ....: $oService  - Task Scheduler Service object as returned by _TS_Open
;                  $sTaskPath - Task path (Folder plus Task name) to process
; Return values .: Success - Returns 1 when the Task was found and 0 when the Task was not found
;                  Failure - Returns 0 and sets @error:
;                  |1401 - Error accessing the specified Taskfolder. @extended is set to the COM error code
;                  |1402 - Error retrieving the Tasks collection. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: This function even checks hidden Tasks
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskExists($oService, $sTaskPath)
	Local $sFolder, $oFolder, $sTaskName, $oTask, $oTasks, $iPos, $bHidden = 1
	$iPos = StringInStr($sTaskPath, "\", $STR_NOCASESENSE, -1)
	$sFolder = ($iPos = 1) ? "\" : (StringLeft($sTaskPath, $iPos - 1))
	$sTaskName = StringMid($sTaskPath, $iPos + 1)
	$oFolder = $oService.GetFolder($sFolder)
	If @error Then Return SetError(1401, @error, 0)
	$oTasks = $oFolder.GetTasks($bHidden)
	If @error Then Return SetError(1402, @error, 0)
	For $oTask In $oTasks
		If $oTask.Name = $sTaskName Then Return 1
	Next
	Return 0
EndFunc   ;==>_TS_TaskExists

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskExportXML
; Description ...: Returns the XML representation of a single Task and writes it to a file or returns it in an array.
; Syntax.........: _TS_TaskExportXML($oService, $sTaskPath[, $sXMLOutput = Default])
; Parameters ....: $oService   - Task Scheduler Service object as returned by _TS_Open
;                  $sTaskPath  - Task path (Folder plus Task name) to process
;                  $sXMLOutput - [optional] Destination to export the XML to. If not specified the XML gets returned as an array (default)
; Return values .: Success - 1 if written to a file, one-dimensional zero based array holding the XML.
;                  Failure - Returns 0 and sets @error:
;                  |1501 - Error returned by _TS_TaskGet. @extended is set to the COM error code
;                  |1502 - Error opening the output file.
; Author ........: water
; Modified.......:
; Remarks .......: An existing file will be overwritten.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskExportXML($oService, $sTask, $sXMLOutput = Default)
	Local $oTask = _TS_TaskGet($oService, $sTask)
	If @error Then Return SetError(1501, @error, 0)
	If $sXMLOutput = Default Then $sXMLOutput = ""
	If $sXMLOutput = "" Then
		Return StringSplit($oTask.XML, @CRLF, BitOR($STR_ENTIRESPLIT, $STR_NOCOUNT))
	Else
		Local $hXMLOutput = FileOpen($sXMLOutput, $FO_OVERWRITE)
		If $hXMLOutput = -1 Then Return SetError(1502, @error, 0)
		FileWrite($hXMLOutput, $oTask.XML)
		FileClose($hXMLOutput)
		Return 1
	EndIf
EndFunc   ;==>_TS_TaskExportXML

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskGet
; Description ...: Returns the Task or Task Definition object of the specified Task.
; Syntax.........: _TS_TaskGet($oService, $sTaskPath[, $bReturnFlag = 0])
; Parameters ....: $oService    - Task Scheduler Service object as returned by _TS_Open
;                  $sTaskPath   - Task path (Folder plus Task name) to process
;                  $bReturnFlag - [optional] Set to 1 if the function should return the Task.Definition object (default = Task object)
; Return values .: Success - Object of the Task
;                  Failure - Returns 0 and sets @error:
;                  |1601 - Error accessing the specified Taskfolder. @extended is set to the COM error code
;                  |1602 - Error retrieving the Tasks collection. @extended is set to the COM error code
;                  |1603 - Task with the specified name could not be found in the Task Folder
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskGet($oService, $sTaskPath, $bReturnFlag = 0)
	Local $sFolder, $oFolder, $sTaskName, $oTask, $oTasks, $iPos, $bHidden = 1
	If $bReturnFlag = Default Then $bReturnFlag = 0
	$iPos = StringInStr($sTaskPath, "\", $STR_NOCASESENSE, -1)
	$sTaskName = StringMid($sTaskPath, $iPos + 1)
	$sFolder = ($iPos = 1) ? "\" : (StringLeft($sTaskPath, $iPos - 1))
	$oFolder = $oService.GetFolder($sFolder)
	If @error Then Return SetError(1601, @error, 0)
	$oTasks = $oFolder.GetTasks($bHidden)
	If @error Then Return SetError(1602, @error, 0)
	For $oTask In $oTasks
		If $oTask.Name = $sTaskName Then
			If $bReturnFlag = 0 Then
				Return $oTask
			Else
				Return $oTask.Definition
			EndIf
		EndIf
	Next
	Return SetError(1603, 0, 0)
EndFunc   ;==>_TS_TaskGet

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskImportXML
; Description ...: Imports a Task from an XML file or an array and returns a Task Definition object.
; Syntax.........: _TS_TaskImportXML($oService, $iInputType, $vXMLInput = Default)
; Parameters ....: $oService   - Task Scheduler Service object as returned by _TS_Open
;                  $iInputType - 1 = $vXMLInput is a file to import the XML from, 2 = $vXMLInput is a XML string or a 1D XML array
;                  $vXMLInput  - Input to import the XML from. Is either a string, an array or a file to import the XML from
; Return values .: Success - Object of the created Task Definition
;                  Failure - Returns 0 and sets @error:
;                  |1702 - Error creating a new Task Definition. @extended is set to the COM error code
;                  |1703 - Error creating a XML from the passed array. @extended is set to the COM error code
;                  |1704 - Error creating a XML from the passed file. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: You can modify the Task Definition as needed before calling _TS_TaskRegister to create a new Task.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskImportXML($oService, $iInputType, $vXMLInput)
	; Create a Task Definition object. The flags parameter is 0 because it is not supported
	Local $oTaskDefinition = $oService.NewTask(0)
	If @error Then Return SetError(1702, @error, 0)
	If $iInputType = 1 Then
		Local $sXML = FileRead($vXMLInput)
		$oTaskDefinition.XMLtext = $sXML
		If @error Then Return SetError(1704, @error, 0)
	Else
		If IsArray($vXMLInput) Then $vXMLInput = _ArrayToString($vXMLInput, @CRLF)
		$oTaskDefinition.XMLText = $vXMLInput
		If @error Then Return SetError(1703, @error, 0)
	EndIf
	Return $oTaskDefinition
EndFunc   ;==>_TS_TaskImportXML

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskList
; Description ...: Returns a list of all Tasks in a given Folder and all subfolders.
; Syntax.........: _TS_TaskList($oService[, $vFolder = "\"[, $iShowHidden = 0[, $iShowDisabled = 1[, $iShowMS = 1[, $iProperties = 0[, $bReadable = True]]]]]])
; Parameters ....: $oService      - Task Scheduler Service object as returned by _TS_Open
;                  $vFolder       - [optional] Task Folder to process either as string Folderpath (e.g. "\test\test2") or Folder object (default = "\" means: root Folder)
;                  $iShowHidden   - [optional] Returns hidden Tasks as well when set to 1 (default = 0)
;                  $iShowDisabled - [optional] Returns disabled Tasks as well when set to 1 (default = 1)
;                  $iShowMS       - [optional] Returns Tasks in Microsoft Folders as well when set to 1 (default = 1)
;                  $iProperties   - [optional] A bitwise mask that indicates the properties to be returned (default = 0 = all properties) e.g. 5 returns Task name and State.
;                  |Possible values (to select all 25 possible columns use the maximum value of 2^25 - 1 or 0 (which gets translated to 2^25 - 1):
;                  |        1 (2^0)  - Task name
;                  |        2 (2^1)  - Task Folder
;                  |        4 (2^2)  - State
;                  |        8 (2^3)  - Hidden
;                  |       16 (2^4)  - Last Task result
;                  |       32 (2^5)  - Last run
;                  |       64 (2^6)  - Next run
;                  |      128 (2^7)  - Missed runs
;                  |      256 (2^8)  - Allow demand start
;                  |      512 (2^9)  - AllowHardTerminate
;                  |     1024 (2^10) - DeleteExpiredTaskAfter
;                  |     2048 (2^11) - DisallowStartIfOnBatteries
;                  |     4096 (2^12) - ExecutionTimeLimit
;                  |     8192 (2^13) - MultipleInstances
;                  |    16384 (2^14) - Priority
;                  |    32768 (2^15) - RestartCount
;                  |    65536 (2^16) - RestartInterval
;                  |   131072 (2^17) - RunOnlyIfIdle
;                  |   262144 (2^18) - RunOnlyIfNetworkAvailable
;                  |   524288 (2^19) - StartWhenAvailable
;                  |  1048576 (2^20) - StopIfGoingOnBatteries
;                  |  2097152 (2^21) - WakeToRun
;                  |  4194304 (2^22) - Author
;                  |  8388608 (2^23) - Date
;                  | 16777216 (2^24) - Description
;                  $bReadable     - [optional] True translates some values (e.g. State to text, date/time to readable format and suppresses empty values).
; Return values .: Success - two-dimensional zero based array with the information requested by parameter $iProperties (see there)
;                  Failure - Returns "" and sets @error:
;                  |1801 - Error accessing the specified Taskfolder. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskList($oService, $vFolder = "\", $iShowHidden = 0, $iShowDisabled = 1, $iShowMS = 1, $iProperties = 0, $bReadable = True)
	Local $oFolder, $oFolders, $oTasks, $oTask, $oSettings, $oRegistrationInfo, $aTemp, $iRow = 0, $iColumn, $iPos, $sFolder, $iTemp
	Local Static $iMaxcolumn = 0
	If $vFolder = Default Then $vFolder = "\"
	If $iShowHidden = Default Then $iShowHidden = 0
	If $iShowDisabled = Default Then $iShowDisabled = 1
	If $iShowMS = Default Then $iShowMS = 1
	If $iProperties = Default Or $iProperties = 0 Then $iProperties = 2 ^ 25 - 1 ; Set all bits to 1
	If $bReadable = Default Then $bReadable = True
	If Not IsObj($vFolder) Then
		$vFolder = $oService.GetFolder($vFolder)
		If @error Then Return SetError(1801, @error, "")
	EndIf
	; Calculate number of columns to be used (depends on $iProperties)
	If $iMaxcolumn = 0 Then
		For $i = 0 To 59 ; 60 Bit equals 0xFFFFFFFFFFFFFFF = the maximum value for $iProperties
			If BitAND($iProperties, 2 ^ $i) = 2 ^ $i Then $iMaxcolumn = $iMaxcolumn + 1
		Next
	EndIf
	Local $aTasks[0][$iMaxcolumn]
	If $iShowMS = 0 And $vFolder.Path = "\Microsoft" Then Return $aTasks
	; Get all Tasks in the Folder
	$oTasks = $vFolder.GetTasks($iShowHidden)
	If @error = 0 Then
		Local $iTaskCount = $oTasks.Count, $aTasks[$iTaskCount][$iMaxcolumn]
		For $i = 1 To $iTaskCount
			$oTask = $oTasks($i)
			$oSettings = $oTask.Definition.Settings
			$oRegistrationInfo = $oTask.Definition.RegistrationInfo
			$iColumn = 0
			If $iShowDisabled = 0 And $oTask.State = $TASK_STATE_DISABLED Then ContinueLoop
			If BitAND($iProperties, 2 ^ 0) = 2 ^ 0 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oTask.Name)
			If BitAND($iProperties, 2 ^ 1) = 2 ^ 1 Then
				$iPos = StringInStr($oTask.Path, "\", $STR_NOCASESENSE, -1)
				$sFolder = ($iPos = 1) ? ("\") : (StringLeft($oTask.Path, $iPos - 1))
				__TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $sFolder)
			EndIf
			If BitAND($iProperties, 2 ^ 2) = 2 ^ 2 Then
				$iTemp = $oTask.State
				If $bReadable Then
					If $iTemp = 0 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, "Unknown")
					If $iTemp = 1 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, "Disabled")
					If $iTemp = 2 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, "Queued")
					If $iTemp = 3 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, "Ready")
					If $iTemp = 4 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, "Running")
				Else
					__TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $iTemp)
				EndIf
			EndIf
			If BitAND($iProperties, 2 ^ 3) = 2 ^ 3 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.Hidden)
			If BitAND($iProperties, 2 ^ 4) = 2 ^ 4 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oTask.LastTaskResult)
			If BitAND($iProperties, 2 ^ 5) = 2 ^ 5 Then
				If $bReadable Then
					__TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, StringRegExpReplace($oTask.LastRunTime, "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", "$1/$2/$3 $4:$5:$6")) ; YYYY/MM/DD HH:MM:SS
				Else
					__TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oTask.LastRunTime)
				EndIf
			EndIf
			If BitAND($iProperties, 2 ^ 6) = 2 ^ 6 Then
				If $bReadable Then
					__TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, StringRegExpReplace($oTask.NextRunTime, "(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", "$1/$2/$3 $4:$5:$6")) ; YYYY/MM/DD HH:MM:SS
				Else
					__TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oTask.NextRunTime)
				EndIf
			EndIf
			If BitAND($iProperties, 2 ^ 7) = 2 ^ 7 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oTask.NumberOfMissedRuns)
			If BitAND($iProperties, 2 ^ 8) = 2 ^ 8 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.AllowDemandStart)
			If BitAND($iProperties, 2 ^ 9) = 2 ^ 9 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.AllowHardTerminate)
			If BitAND($iProperties, 2 ^ 10) = 2 ^ 10 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.DeleteExpiredTaskAfter)
			If BitAND($iProperties, 2 ^ 11) = 2 ^ 11 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.DisallowStartIfOnBatteries)
			If BitAND($iProperties, 2 ^ 12) = 2 ^ 12 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.ExecutionTimeLimit)
			If BitAND($iProperties, 2 ^ 13) = 2 ^ 13 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.MultipleInstances)
			If BitAND($iProperties, 2 ^ 14) = 2 ^ 14 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.Priority)
			If BitAND($iProperties, 2 ^ 15) = 2 ^ 15 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.RestartCount)
			If BitAND($iProperties, 2 ^ 16) = 2 ^ 16 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.RestartInterval)
			If BitAND($iProperties, 2 ^ 17) = 2 ^ 17 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.RunOnlyIfIdle)
			If BitAND($iProperties, 2 ^ 18) = 2 ^ 18 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.RunOnlyIfNetworkAvailable)
			If BitAND($iProperties, 2 ^ 19) = 2 ^ 19 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.StartWhenAvailable)
			If BitAND($iProperties, 2 ^ 20) = 2 ^ 20 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.StopIfGoingOnBatteries)
			If BitAND($iProperties, 2 ^ 21) = 2 ^ 21 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oSettings.WakeToRun)
			If BitAND($iProperties, 2 ^ 22) = 2 ^ 22 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oRegistrationInfo.Author)
			If BitAND($iProperties, 2 ^ 23) = 2 ^ 23 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oRegistrationInfo.Date)
			If BitAND($iProperties, 2 ^ 24) = 2 ^ 24 Then __TS_TaskListWrite($aTasks, $iRow, $iColumn, $bReadable, $oRegistrationInfo.Description)
			$iRow = $iRow + 1
		Next
	EndIf
	ReDim $aTasks[$iRow][$iMaxcolumn]
	; Get all Folders and call this function recursively
	$oFolders = $vFolder.GetFolders(0)
	For $oFolder In $oFolders
		$aTemp = _TS_TaskList($oService, $oFolder, $iShowHidden, $iShowDisabled, $iShowMS, $iProperties, $bReadable)
		If @error Then Return SetError(@error, @extended, "")
		_ArrayConcatenate($aTasks, $aTemp)
	Next
	Return $aTasks
EndFunc   ;==>_TS_TaskList

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskListHeader
; Description ...: Returns the header line for function _TS_TaskList so you can pass this to _ArrayDisplay as parameter $sHeader.
; Syntax.........: _TS_TaskListHeader([$iProperties = 0])
; Parameters ....: $iProperties - [optional] A bitwise mask that indicates the properties to be returned.
; Return values .: Success - String with column headers separated by | (pipe symbol)
;                  Failure - None
; Author ........: water
; Modified.......:
; Remarks .......: For possible values for parameter $iProperties please see function _TS_TaskList.
;                  For an example please see the example script for _TS_TaskList.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskListHeader($iProperties = 0)
	Local $sHeader
	If $iProperties = Default Or $iProperties = 0 Then $iProperties = 2 ^ 25 - 1 ; Set all bits to 1
	If BitAND($iProperties, 2 ^ 0) = 2 ^ 0 Then $sHeader &= "|Task Name"
	If BitAND($iProperties, 2 ^ 1) = 2 ^ 1 Then $sHeader &= "|Task Folder"
	If BitAND($iProperties, 2 ^ 2) = 2 ^ 2 Then $sHeader &= "|State"
	If BitAND($iProperties, 2 ^ 3) = 2 ^ 3 Then $sHeader &= "|Hidden"
	If BitAND($iProperties, 2 ^ 4) = 2 ^ 4 Then $sHeader &= "|Last Task Result"
	If BitAND($iProperties, 2 ^ 5) = 2 ^ 5 Then $sHeader &= "|Last Run"
	If BitAND($iProperties, 2 ^ 6) = 2 ^ 6 Then $sHeader &= "|Next Run"
	If BitAND($iProperties, 2 ^ 7) = 2 ^ 7 Then $sHeader &= "|Missed Runs"
	If BitAND($iProperties, 2 ^ 8) = 2 ^ 8 Then $sHeader &= "|Allow Demand Start"
	If BitAND($iProperties, 2 ^ 9) = 2 ^ 9 Then $sHeader &= "|Allow Hard Terminate"
	If BitAND($iProperties, 2 ^ 10) = 2 ^ 10 Then $sHeader &= "|Delete Expired Task After"
	If BitAND($iProperties, 2 ^ 11) = 2 ^ 11 Then $sHeader &= "|Disallow Start If On Batteries"
	If BitAND($iProperties, 2 ^ 12) = 2 ^ 12 Then $sHeader &= "|Execution Time Limit"
	If BitAND($iProperties, 2 ^ 13) = 2 ^ 13 Then $sHeader &= "|Multiple Instances"
	If BitAND($iProperties, 2 ^ 14) = 2 ^ 14 Then $sHeader &= "|Priority"
	If BitAND($iProperties, 2 ^ 15) = 2 ^ 15 Then $sHeader &= "|Restart Count"
	If BitAND($iProperties, 2 ^ 16) = 2 ^ 16 Then $sHeader &= "|Restart Interval"
	If BitAND($iProperties, 2 ^ 17) = 2 ^ 17 Then $sHeader &= "|Run Only If Idle"
	If BitAND($iProperties, 2 ^ 18) = 2 ^ 18 Then $sHeader &= "|Run Only If Network Available"
	If BitAND($iProperties, 2 ^ 19) = 2 ^ 19 Then $sHeader &= "|Start When Available"
	If BitAND($iProperties, 2 ^ 20) = 2 ^ 20 Then $sHeader &= "|Stop If Going On Batteries"
	If BitAND($iProperties, 2 ^ 21) = 2 ^ 21 Then $sHeader &= "|Wake To Run"
	If BitAND($iProperties, 2 ^ 22) = 2 ^ 22 Then $sHeader &= "|Author"
	If BitAND($iProperties, 2 ^ 23) = 2 ^ 23 Then $sHeader &= "|Date"
	If BitAND($iProperties, 2 ^ 24) = 2 ^ 24 Then $sHeader &= "|Description"
	If StringLeft($sHeader, 1) = "|" Then $sHeader = StringMid($sHeader, 2)
	Return $sHeader
EndFunc   ;==>_TS_TaskListHeader

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskPropertiesGet
; Description ...: Lists all or specified properties of a Task or Task Definition and returns an array or string or writes the properties to the console.
; Syntax.........: _TS_TaskPropertiesGet($oService, $vTask[, $iFormat = 1[, $bIgnoreNoValues = False[, $sQuerySection = ""[, $sQueryProperties = ""]]]])
; Parameters ....: $oService         - Task Scheduler Service object as returned by _TS_Open
;                  $vTask            - Path and name or object of the Registered Task to process or a Task Definition object
;                  $iFormat          - Format of the output. Can be one of the following values
;                  |1 - User friendly format (default). Please see Remarks
;                  |2 - Format you can use as input to _TS_TaskPropertiesSet (just the content of the array)
;                  |3 - Format you can use as input to _TS_TaskPropertiesSet (full AutoIt syntax to define the array - without XML and written to the console)
;                  $bIgnoreNoValues  - [optional] If set to True properties without a value do not get returned (default = False)
;                  $sQuerySection    - [optional] Name of the Scheduler object to retrieve the properties from. If set to "" all objects will be retrieved (default = "")
;                  $sQueryProperties - [optional] Comma separated list of properties to retrieve from $sSection. If set to "" all properties will be retrieved (default = "")
; Return values .: Success - For $iFormat=1 or 2: Zero based two-dimensional array with the following information. Please see Remarks as well.
;                  | 0 - Section related to a COM object
;                  | 1 - Property name
;                  | 2 - Property value
;                  | 3 - Comment
;                  Success - For $iFormat=3: Writes the AutoIt array definition to the console
;                  Failure - Returns "" and sets @error
;                  |1901 - Error returned by _TS_TaskGet. @extended is set to the COM error code. Most probably the Task could not be found
; Author ........: water
; Modified.......:
; Remarks .......: For $iFormat = 1 if you only request a single property you will get a string holding the value of the property. Else you get an array as described above.
;                  All data returned by the function is in the format as retrieved from the Taskscheduler object.
;                  e.g. LogonType is of type Integer, UserId is returned as String.
; Related .......:
; Link ..........: https://www.experts-exchange.com/articles/11326/VBScript-and-Task-Scheduler-2-0-Listing-Scheduled-Tasks.html
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskPropertiesGet($oTaskService, $vTask, $iFormat = Default, $bIgnoreNoValues = Default, $sQuerySection = Default, $sQueryProperties = Default)
	Local $oTask, $sTaskState, $sObjectType
	Local $oTaskDefinition, $cActions, $oTaskAction
	Local $oPrincipal, $oRegistrationInfo, $oTaskSettings, $oIdleSettings, $oTaskNetworkSettings
	Local $cTaskTriggers, $oTaskTrigger, $oTaskRepetition, $cAttachments, $cHeaderfields
	Local $iIndex = 0, $sSection
	If $iFormat = Default Then $iFormat = 1
	If $bIgnoreNoValues = Default Then $bIgnoreNoValues = False
	If $sQuerySection = Default Then $sQuerySection = ""
	If $sQueryProperties = Default Then $sQueryProperties = ""
	If $iFormat = 1 Then
		Local $aProperties[1000][4]
	Else
		Local $aProperties[1000]
	EndIf
	If IsObj($vTask) Then
		$oTask = $vTask ; Dummy. Just to make sure $oTask is set. Else the function would crash for a Task Definition
		If ObjName($vTask) = "IRegisteredTask" Then ; Registered Task object
			$sObjectType = "Task"
		Else ; Task Definition object
			$oTaskDefinition = $vTask
			$sObjectType = "Task Definition"
		EndIf
	Else
		$oTask = _TS_TaskGet($oTaskService, $vTask)
		If @error Then Return SetError(1901, @error, "")
		$sObjectType = "Task"
	EndIf
	With $oTask
		If $sObjectType = "Task" Then
			$sSection = "TASK"
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Name", .Name)
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Enabled", .Enabled)
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LastRunTime", .LastRunTime)
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LastTaskResult", .LastTaskResult)
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "NextRunTime", .NextRunTime)
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "NumberOfMissedRuns", .NumberOfMissedRuns)
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Path", .Path)
			Switch (.State)
				Case $TASK_STATE_UNKNOWN
					$sTaskState = "Unknown"
				Case $TASK_STATE_DISABLED
					$sTaskState = "Disabled"
				Case $TASK_STATE_QUEUED
					$sTaskState = "Queued"
				Case $TASK_STATE_QUEUED
					$sTaskState = "Ready"
				Case $TASK_STATE_RUNNING
					$sTaskState = "Running"
			EndSwitch
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "State", .State, $sTaskState)
			If $iFormat <> 3 Then __TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "XML", .XML)
			$oTaskDefinition = .Definition
		EndIf
		$sSection = "DEFINITION"
		With $oTaskDefinition
			__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Data", .Data)
			$cActions = $oTaskDefinition.Actions
			If IsObj($cActions) Then
				For $oTaskAction In $cActions
					With $oTaskAction
						$sSection = "ACTIONS"
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ID", .Id)
						Switch (.Type)
							Case $TASK_ACTION_EXEC
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Execute / Command Line Operation")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Arguments", .Arguments)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Path", .Path)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "WorkingDirectory", .WorkingDirectory)
							Case $TASK_ACTION_COM_HANDLER
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Handler")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ClassId", .ClassId)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Data", .Data)
							Case $TASK_ACTION_SEND_EMAIL
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Email Message")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "From", .From)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ReplyTo", .ReplyTo)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "To", .To)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Cc", .Cc)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Bcc", .Bcc)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Subject", .Subject)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Body", .Body)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Server", .Server)
								; ==> CHECK					$cAttachments = .Attachments
								$sSection = "ATTACHMENTS"
								For $sAttachment In $cAttachments
									__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Attachment", $sAttachment)
								Next
								$cHeaderfields = .HeaderFields
								$sSection = "HEADERFIELDS"
								For $oHeaderPair In $cHeaderfields
									__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "HeaderName|Value", $oHeaderPair.Name & "|" & $oHeaderPair.Value)
								Next
							Case $TASK_ACTION_SHOW_MESSAGE
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Message Box")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Title", .Title)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MessageBody", .MessageBody)
						EndSwitch
					EndWith ; oTaskAction
				Next ; objTaskAction
			EndIf
			$oPrincipal = .Principal
			If IsObj($oPrincipal) Then
				$sSection = "PRINCIPAL"
				With $oPrincipal
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ID", .Id)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DisplayName", .DisplayName)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "GroupId", .GroupId)
					Switch (.LogonType)
						Case $TASK_LOGON_NONE
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "None")
						Case $TASK_LOGON_PASSWORD
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "Password")
						Case $TASK_LOGON_S4U
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "Service 4 Users")
						Case $TASK_LOGON_INTERACTIVE_TOKEN
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "Interactive (User must be logged in)")
						Case $TASK_LOGON_GROUP
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "Group")
						Case $TASK_LOGON_SERVICE_ACCOUNT
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "Local $Service/System or Network Service")
						Case $TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "LogonType", .LogonType, "Interactive Token then Try Password")
					EndSwitch
					Switch (.RunLevel)
						Case $TASK_RUNLEVEL_LUA
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RunLevel", .RunLevel, "Least Privileges (LUA)")
						Case $TASK_RUNLEVEL_HIGHEST
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RunLevel", .RunLevel, "Highest Privileges")
					EndSwitch
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "UserId", .UserId)
				EndWith ; oPrincipal
			EndIf
			$oRegistrationInfo = .RegistrationInfo
			If IsObj($oRegistrationInfo) Then
				$sSection = "REGISTRATIONINFO"
				With $oRegistrationInfo
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Author", .Author)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Date", .Date)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Description", .Description)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Date", .Date)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Documentation", .Documentation)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "SecurityDescriptor", .SecurityDescriptor)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Source", .Source)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "URI", .URI)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Version", .Version)
				EndWith ; oRegistrationInfo
			EndIf
			$oTaskSettings = .Settings
			If IsObj($oTaskSettings) Then
				$sSection = "SETTINGS"
				With $oTaskSettings
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "AllowDemandStart", .AllowDemandStart)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "AllowHardTerminate", .AllowHardTerminate)
					Switch (.Compatibility)
						Case $TASK_COMPATIBILITY_AT
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Compatibility", .Compatibility, "compatible with the AT command")
						Case $TASK_COMPATIBILITY_V1
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Compatibility", .Compatibility, "compatible with Task Scheduler 1.0")
						Case $TASK_COMPATIBILITY_V2
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Compatibility", .Compatibility, "compatible with Task Scheduler 2.0 (Windows Vista / Windows 2008)")
						Case 3             ; Not Documented
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Compatibility", .Compatibility, "compatible with Task Scheduler 2.0 (Windows 7 / Windows 2008 R2)")
					EndSwitch
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DeleteExpiredTaskAfter", .DeleteExpiredTaskAfter)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DisallowStartIfOnBatteries", .DisallowStartIfOnBatteries)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Enabled", .Enabled)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ExecutionTimeLimit", .ExecutionTimeLimit)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Hidden", .Hidden)
					Switch (.MultipleInstances)
						Case $TASK_INSTANCES_PARALLEL
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MultipleInstances", .MultipleInstances, "Run in parallel")
						Case $TASK_INSTANCES_QUEUE
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MultipleInstances", .MultipleInstances, "Add to queue")
						Case $TASK_INSTANCES_IGNORE_NEW
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MultipleInstances", .MultipleInstances, "Ignore new")
						Case $TASK_INSTANCES_STOP_EXISTING
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MultipleInstances", .MultipleInstances, "Stop existing task")
					EndSwitch
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Priority", .Priority, "(0=High / 10=Low)")
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RestartCount", .RestartCount)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RestartInterval", .RestartInterval)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RunOnlyIfIdle", .RunOnlyIfIdle)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RunOnlyIfNetworkAvailable", .RunOnlyIfNetworkAvailable)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StartWhenAvailable", .StartWhenAvailable)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StopIfGoingOnBatteries", .StopIfGoingOnBatteries)
					__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "WakeToRun", .WakeToRun)
					$oIdleSettings = .IdleSettings
					$sSection = "IDLESETTINGS"
					With $oIdleSettings
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "IdleDuration", .IdleDuration)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RestartOnIdle", .RestartOnIdle)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StopOnIdleEnd", .StopOnIdleEnd)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "WaitTimeout", .WaitTimeout)
					EndWith ; oIdleSettings
					$oTaskNetworkSettings = .NetworkSettings
					$sSection = "NETWORKSETTINGS"
					With $oTaskNetworkSettings
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ID", .Id)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Name", .Name)
					EndWith ; oTaskNetworkSettings
				EndWith ; oTaskSettings
			EndIf
			$cTaskTriggers = .Triggers
			If IsObj($cTaskTriggers) Then
				For $oTaskTrigger In $cTaskTriggers
					$sSection = "TRIGGERS"
					With $oTaskTrigger
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Enabled", .Enabled)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Id", .Id)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StartBoundary", .StartBoundary)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "EndBoundary", .EndBoundary)
						__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "ExecutionTimeLimit", .ExecutionTimeLimit)
						Switch (.Type)
							Case $TASK_TRIGGER_EVENT
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, " Event")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Delay", .Delay)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Subscription", .Subscription)
							Case $TASK_TRIGGER_TIME
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, " Time")
							Case $TASK_TRIGGER_DAILY
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, " Daily")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DaysInterval", .DaysInterval)
							Case $TASK_TRIGGER_WEEKLY
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, " Weekly")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "WeeksInterval", .WeeksInterval)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DaysOfWeek", .DaysOfWeek, "=" & __TS_ConvertDaysOfWeek(.DaysOfWeek))
							Case $TASK_TRIGGER_MONTHLY
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Monthly")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DaysOfMonth", .DaysOfMonth, "=" & __TS_ConvertDaysOfMonth(.DaysOfMonth))
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MonthsOfYear", .MonthsOfYear, "=" & __TS_ConvertMonthsOfYear(.MonthsOfYear))
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RandomDelay", .RandomDelay)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RunOnLastDayOfMonth", .RunOnLastDayOfMonth)
							Case $TASK_TRIGGER_MONTHLYDOW
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Monthly on Specific Day")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "DaysOfWeek", .DaysOfWeek, "=" & __TS_ConvertDaysOfWeek(.DaysOfWeek))
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "MonthsOfYear", .MonthsOfYear, "=" & __TS_ConvertMonthsOfYear(.MonthsOfYear))
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RandomDelay", .RandomDelay)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "RunOnLastWeekOfMonth", .RunOnLastWeekOfMonth)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "WeeksOfMonth", .WeeksOfMonth, "=" & __TS_ConvertWeeksOfMonth(.WeeksOfMonth))
							Case $TASK_TRIGGER_IDLE
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "When computer is idle")
							Case $TASK_TRIGGER_REGISTRATION
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "When task is registered")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Delay", .Delay)
							Case $TASK_TRIGGER_BOOT
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Boot")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Delay", .Delay)
							Case $TASK_TRIGGER_LOGON
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Logon")
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Delay", .Delay)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "UserId", .UserId)
							Case $TASK_TRIGGER_SESSION_STATE_CHANGE
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Type", .Type, "Session State Change")
								Switch (.StateChange)
									Case 0
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "None")
									Case $TASK_CONSOLE_CONNECT
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "User session connect to local computer")
									Case $TASK_CONSOLE_DISCONNECT
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "User session disconnect from local computer")
									Case $TASK_REMOTE_CONNECT
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "User session connect to remote computer")
									Case $TASK_REMOTE_DISCONNECT
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "User session disconnect from remote computer")
									Case $TASK_SESSION_LOCK
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "On workstation lock")
									Case $TASK_SESSION_UNLOCK
										__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StateChange", .StateChange, "On workstation unlock")
								EndSwitch
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Delay", .Delay)
								__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "UserId", .UserId)
						EndSwitch
						$oTaskRepetition = .Repetition
						$sSection = "REPETITION"
						With $oTaskRepetition
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Duration", .Duration)
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "Interval", .Interval)
							__TS_PropertyGetWrite($aProperties, $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, "StopAtDurationEnd", .StopAtDurationEnd)
						EndWith ; oTaskRepetition
					EndWith ; oTaskTrigger
				Next ; oTaskTrigger
			EndIf
		EndWith ; oTaskDefinition
	EndWith ; oTask
	If $iFormat = 1 Then
		ReDim $aProperties[$iIndex][4]
		; Return a string if just a single property has been queried
		If $sQueryProperties <> "" And StringInStr($sQueryProperties, ",") = 0 Then
			If $iIndex = 0 Then
				$aProperties = ""
			Else
				$aProperties = $aProperties[0][2]
			EndIf
		EndIf
	Else
		ReDim $aProperties[$iIndex]
	EndIf
	If $iFormat = 3 Then
		Local $iLastIndex = UBound($aProperties) - 1
		ConsoleWrite("Global $aProperties[] = [ _" & @CRLF)
		For $i = 0 To $iLastIndex
			$aProperties[$i] = StringReplace($aProperties[$i], @LF, '" & @CRLF & "')
			If $i = $iLastIndex Then
				ConsoleWrite('"' & $aProperties[$i] & '" _' & @CRLF)
			Else
				ConsoleWrite('"' & $aProperties[$i] & '", _' & @CRLF)
			EndIf
		Next
		ConsoleWrite('"]' & @CRLF)
		Return 1
	EndIf
	Return $aProperties
EndFunc   ;==>_TS_TaskPropertiesGet

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskPropertiesSet
; Description ...: Sets the properties of a Tasks object (Task Definition, RegisteredTask, Action or Trigger).
; Syntax.........: _TS_TaskPropertiesSet(ByRef $oObject, $aProperties)
; Parameters ....: $oObject     - Object of a Task Definition, RegisteredTask, Action or Trigger
;                  $aProperties - one-dimensional zero based array in the following format: "object name|property name|property value"
;                  |Name of the object to process. Valid are: TASK, DEFINITION, PRINCIPAL, REGISTRATIONINFO, SETTINGS, IDLESETTINGS, NETWORKSETTINGS, TRIGGERS, REPETITION and ACTIONS
;                  |Name of the property to set
;                  |Value for the property to set
; Return values .: Success - 1
;                  Failure - 0, sets @error to:
;                  |2001 - Unsupported or invalid Task Scheduler COM object
;                  |2002 - Unsupported or invalid property name. @extended is set to the zero based index of the property in error
;                  |2003 - The row in $aProperties does not have the required format: "object name|property name|property value". @extended is set to the index of the row in error.
;                  |2004 - $oObject is invalid. Must be: TaskDefinition, RegisteredTask, Trigger or Action
; Author ........: water
; Modified.......:
; Remarks .......: Sections that are not valid for the passed object are ignored!
;                  E.g. The "Task" section and its properties are only valid for a RegisteredTask object
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskPropertiesSet(ByRef $oObject, $aProperties)
	Local Const $ObjectType_RegisteredTask = 1
	Local Const $ObjectType_TaskDefinition = 2
	Local Const $ObjectType_Trigger = 3
	Local Const $ObjectType_Action = 4
	Local $iObjectType = 0
	Select
		Case ObjName($oObject) = "IRegisteredTask"
			Local $oTask = $oObject
			$iObjectType = $ObjectType_RegisteredTask
		Case ObjName($oObject) = "ITaskDefinition"
			Local $oTaskDefinition = $oObject, $oPrincipal = $oObject.Principal, $oRegistrationInfo = $oObject.RegistrationInfo
			Local $oSettings = $oObject.Settings, $oIdleSettings = $oSettings.IdleSettings, $oNetworkSettings = $oSettings.NetworkSettings
			$iObjectType = $ObjectType_TaskDefinition
		Case StringInStr(ObjName($oObject), "Trigger") > 0
			Local $oTrigger = $oObject
			$iObjectType = $ObjectType_Trigger
		Case StringInStr(ObjName($oObject), "Action") > 0
			Local $oAction = $oObject
			$iObjectType = $ObjectType_Action
		Case Else
			Return SetError(2004, 0, 0)
	EndSelect
	Local $aTemp
	For $i = 0 To UBound($aProperties) - 1
		If $aProperties[$i] = "" Then ContinueLoop
		$aTemp = StringSplit($aProperties[$i], "|", $STR_NOCOUNT)
		If @error Or UBound($aTemp, 1) <> 3 Then Return SetError(2003, $i, 0)
		Switch $aTemp[0]
			Case "Task"
				If $iObjectType <> $ObjectType_RegisteredTask Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "LastRunTime", "LastTaskResult", "Name", "NextRunTime", "NumberOfMissedRuns", "Path", "State", "XML"     ; Ignore read only properties
					Case "Enabled"
						$oTask.Enabled = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "Definition"
				If $iObjectType <> $ObjectType_TaskDefinition Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "Data"
						$oTaskDefinition.Data = $aTemp[2]
					Case "XmlText"
						$oTaskDefinition.XmlText = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "Principal"
				If $iObjectType <> $ObjectType_TaskDefinition Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "DisplayName"
						$oPrincipal.DisplayName = $aTemp[2]
					Case "GroupId"
						$oPrincipal.GroupId = $aTemp[2]
					Case "Id"
						$oPrincipal.Id = $aTemp[2]
					Case "LogonType"
						$oPrincipal.LogonType = $aTemp[2]
					Case "RunLevel"
						$oPrincipal.RunLevel = $aTemp[2]
					Case "UserId"
						$oPrincipal.UserId = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "RegistrationInfo"
				If $iObjectType <> $ObjectType_TaskDefinition Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "Author"
						$oRegistrationInfo.Author = $aTemp[2]
					Case "Date"
						$oRegistrationInfo.Date = $aTemp[2]
					Case "Description"
						$oRegistrationInfo.Description = $aTemp[2]
					Case "Documentation"
						$oRegistrationInfo.Documentation = $aTemp[2]
					Case "SecurityDescriptor"
						$oRegistrationInfo.SecurityDescriptor = $aTemp[2]
					Case "Source"
						$oRegistrationInfo.Source = $aTemp[2]
					Case "URI"
						$oRegistrationInfo.URI = $aTemp[2]
					Case "Version"
						$oRegistrationInfo.Version = $aTemp[2]
					Case "XmlText"
						$oRegistrationInfo.XmlText = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "Settings"
				If $iObjectType <> $ObjectType_TaskDefinition Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "AllowDemandStart"
						$oSettings.AllowDemandStart = $aTemp[2]
					Case "AllowHardTerminate"
						$oSettings.AllowHardTerminate = $aTemp[2]
					Case "Compatibility"
						$oSettings.Compatibility = $aTemp[2]
					Case "DeleteExpiredTaskAfter"
						$oSettings.DeleteExpiredTaskAfter = $aTemp[2]
					Case "DisallowStartIfOnBatteries"
						$oSettings.DisallowStartIfOnBatteries = $aTemp[2]
					Case "Enabled"
						$oSettings.Enabled = $aTemp[2]
					Case "ExecutionTimeLimit"
						$oSettings.ExecutionTimeLimit = $aTemp[2]
					Case "Hidden"
						$oSettings.Hidden = $aTemp[2]
					Case "MultipleInstances"
						$oSettings.MultipleInstances = $aTemp[2]
					Case "Priority"
						$oSettings.Priority = $aTemp[2]
					Case "RestartCount"
						$oSettings.RestartCount = $aTemp[2]
					Case "RestartInterval"
						$oSettings.RestartInterval = $aTemp[2]
					Case "RunOnlyIfIdle"
						$oSettings.RunOnlyIfIdle = $aTemp[2]
					Case "RunOnlyIfNetworkAvailable"
						$oSettings.RunOnlyIfNetworkAvailable = $aTemp[2]
					Case "StartWhenAvailable"
						$oSettings.StartWhenAvailable = $aTemp[2]
					Case "StopIfGoingOnBatteries"
						$oSettings.StopIfGoingOnBatteries = $aTemp[2]
					Case "WakeToRun"
						$oSettings.WakeToRun = $aTemp[2]
					Case "XmlText"
						$oSettings.XmlText = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "IdleSettings"
				If $iObjectType <> $ObjectType_TaskDefinition Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "IdleDuration"
						$oIdleSettings.IdleDuration = $aTemp[2]
					Case "RestartOnIdle"
						$oIdleSettings.RestartOnIdle = $aTemp[2]
					Case "StopOnIdleEnd"
						$oIdleSettings.StopOnIdleEnd = $aTemp[2]
					Case "WaitTimeout"
						$oIdleSettings.WaitTimeout = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "NetworkSettings"
				If $iObjectType <> $ObjectType_TaskDefinition Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "Id"
						$oNetworkSettings.Id = $aTemp[2]
					Case "Name"
						$oNetworkSettings.Name = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "Triggers"
				If $iObjectType <> $ObjectType_Trigger Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "Type"     ; Ignore read only properties
					Case "Enabled"
						$oTrigger.Enabled = $aTemp[2]
					Case "EndBoundary"
						$oTrigger.EndBoundary = $aTemp[2]
					Case "ExecutionTimeLimit"
						$oTrigger.ExecutionTimeLimit = $aTemp[2]
					Case "Id"
						$oTrigger.Id = $aTemp[2]
					Case "StartBoundary"
						$oTrigger.StartBoundary = $aTemp[2]
					Case "Delay"
						$oTrigger.Delay = $aTemp[2]
					Case "DaysInterval"
						$oTrigger.DaysInterval = $aTemp[2]
					Case "RandomDelay"
						$oTrigger.RandomDelay = $aTemp[2]
					Case "Subscription"
						$oTrigger.Subscription = $aTemp[2]
					Case "UserID"
						$oTrigger.UserId = $aTemp[2]
					Case "DaysOfWeek"
						$oTrigger.DaysOfWeek = $aTemp[2]
					Case "MonthsOfYear"
						$oTrigger.MonthsOfYear = $aTemp[2]
					Case "RunOnLastWeekOfMonth"
						$oTrigger.RunOnLastWeekOfMonth = $aTemp[2]
					Case "WeeksOfMonth"
						$oTrigger.WeeksOfMonth = $aTemp[2]
					Case "DaysOfMonth"
						$oTrigger.DaysOfMonth = $aTemp[2]
					Case "MonthsOfYear"
						$oTrigger.MonthsOfYear = $aTemp[2]
					Case "RunOnLastDayOfMonth"
						$oTrigger.RunOnLastDayOfMonth = $aTemp[2]
					Case "StateChange"
						$oTrigger.StateChange = $aTemp[2]
					Case "WeeksInterval"
						$oTrigger.WeeksInterval = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "Repetition"
				If $iObjectType <> $ObjectType_Trigger Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "Duration"
						$oTrigger.Repetition.Duration = $aTemp[2]
					Case "Interval"
						$oTrigger.Repetition.Interval = $aTemp[2]
					Case "StopAtDurationEnd"
						$oTrigger.Repetition.StopAtDurationEnd = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case "Actions"
				If $iObjectType <> $ObjectType_Action Then ContinueLoop     ; Ignore this section when the properties are invalid for the object
				Switch $aTemp[1]
					Case "Type"     ; Ignore read only properties
					Case "Id"
						$oAction.Id = $aTemp[2]
					Case "ClassId"
						$oAction.ClassId = $aTemp[2]
					Case "Data"
						$oAction.Data = $aTemp[2]
					Case "Arguments"
						$oAction.Arguments = $aTemp[2]
					Case "Path"
						$oAction.Path = $aTemp[2]
					Case "WorkingDirectory"
						$oAction.WorkingDirectory = $aTemp[2]
					Case Else
						Return SetError(2002, $i, 0)
				EndSwitch
			Case Else
				Return SetError(2001, 0, 0)
		EndSwitch
	Next
	Return 1
EndFunc   ;==>_TS_TaskPropertiesSet

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskRegister
; Description ...: Register or update a Task.
; Syntax.........: _TS_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition, $sUserId = "", $sPassword = "", $iLogonType = Default, $iCreateFlag = $TASK_CREATE)
; Parameters ....: $oService        - Task Scheduler Service object as returned by _TS_Open
;                  $sFolder         - Folder where the Task should be created
;                  $sName           - Name of the Task
;                  $oTaskDefinition - Task Definition object as created by _TS_TaskCreate and filled by _TS_TaskPropertiesSet
;                  $sUserId         - [optional] The user credentials that are used to register the Task. If present, these credentials
;                                     take priority over the credentials specified in the Task Definition object pointed to by the definition parameter
;                  $sPassword       - [optional] The password for the UserId that is used to register the Task. When the TASK_LOGON_SERVICE_ACCOUNT logon type
;                                     is used, the password must be an empty value such as NULL or ""
;                  $iLogonType      - [optional] Can be any of the TASK_LOGON_TYPE constants enumeration. For the default please check Remarks
;                  $iCreateFlag     - [optional] Defines if to create or update the task. Can be any of the TASK_CREATE constants enumeration. Default is $TASK_CREATE
; Return values .: Success - Task object
;                  Failure - Returns 0 and sets @error:
;                  |2101 - Parameter $oService is not an object or not an ITaskService object
;                  |2102 - $sFolder does not exist or an error occurred in _TS_FolderExists. @extended is set to the COM error (if any)
;                  |2103 - Task exists which is incompatible with flags $TASK_CREATE, $TASK_DISABLE and $TASK_CREATE or
;                  |       Task does not exist which is incompatible with flags $TASK_UPDATE And $TASK_DONT_ADD_PRINCIPAL_ACE
;                  |2104 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
;                  |2105 - Error accessing $sFolder using _TS_FolderGet. @extended is set to the COM error
;                  |2106 - Error creating the Task. @extended is set to the COM error
; Author ........: water
; Modified.......:
; Remarks .......: If the logon type has been set in the Principal sub-object then $TASK_LOGON_NONE is the default to not overwrite the existing setting.
;                  Else $TASK_LOGON_INTERACTIVE_TOKEN will be used as default.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition, $sUserId = Default, $sPassword = Default, $iLogonType = Default, $iCreateFlag = Default)
	If $sUserId = Default Then $sUserId = ""
	If $sPassword = Default Then $sPassword = ""
	If $iLogonType = Default Then
		If $oTaskDefinition.Principal.LogonType <> 0 Then
			$iLogonType = $TASK_LOGON_NONE
		Else
			$iLogonType = $TASK_LOGON_INTERACTIVE_TOKEN
		EndIf
	EndIf
	If $iCreateFlag = Default Then $iCreateFlag = $TASK_CREATE
	If Not IsObj($oService) Or ObjName($oService) <> "ITaskService" Then Return SetError(2101, 0, 0)
	If Not _TS_FolderExists($oService, $sFolder) Then Return SetError(2102, @error, 0)
	If ($iCreateFlag = $TASK_CREATE Or $iCreateFlag = $TASK_DISABLE Or $iCreateFlag = $TASK_IGNORE_REGISTRATION_TRIGGERS) And _
			_TS_TaskExists($oService, $sFolder & "\" & $sName) Then
		Return SetError(2103, @error, 0)
	ElseIf ($iCreateFlag = $TASK_UPDATE Or $iCreateFlag = $TASK_DONT_ADD_PRINCIPAL_ACE) And Not _TS_TaskExists($oService, $sFolder & "\" & $sName) Then
		Return SetError(2103, @error, 0)
	EndIf
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(2104, 0, 0)
	; Register (create) the Task
	Local $oFolder = _TS_FolderGet($oService, $sFolder)
	If @error Then Return SetError(2105, @error, 0)
	Local $oTask = $oFolder.RegisterTaskDefinition($sName, $oTaskDefinition, $iCreateFlag, $sUserId, $sPassword, $iLogonType)
	If @error Then Return SetError(2106, @error, 0)
	Return $oTask
EndFunc   ;==>_TS_TaskRegister

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskRun
; Description ...: Run the Registered Task immediately.
; Syntax.........: _TS_TaskRun($oService, $vTask)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $vTask    - Registered Task to run. Can be the object or a string e.g. "\folder\task"
; Return values .: Success - A RunningTask Object that defines the new instance of the task.
;                  Failure - Returns 0 and sets @error
;                  |2201 - The Task does not exist. @extended is set to the COM error code returned by _TS_TaskGet
;                  |2202 - Error starting the Task. @extended is set to the COM error
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskRun($oService, $vTask)
	Local $oTask, $oRunningTask
	; Check if the Task exists when specified as \Folder\Task
	If Not IsObj($vTask) Then
		$oTask = _TS_TaskGet($oService, $vTask)
		If @error Then Return SetError(2201, @error, 0)
	Else
		$oTask = $vTask
	EndIf
	$oRunningTask = $oTask.Run(Null)
	If @error Then Return SetError(2202, @error, 0)
	Return $oRunningTask
EndFunc   ;==>_TS_TaskRun

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskStop
; Description ...: Stops all instances of the Registered Task immediately.
; Syntax.........: _TS_TaskStop($oService, $vTask)
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $vTask    - Registered Task to stop. Can be the object or a string e.g. "\folder\task"
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error
;                  |2301 - The Task does not exist. @extended is set to the COM error code returned by _TS_TaskGet
;                  |2302 - Error stopping the Task. @extended is set to the COM error
; Author ........: water
; Modified.......:
; Remarks .......: The function stops all instances of the task.
;                  System account users can stop a task, users with Administrator group privileges can stop a task,
;                  and if a user has rights to execute and read a task, then the user can stop the task.
;                  A user can stop the task instances that are running under the same credentials as the user account.
;                  In all other cases, the user is denied access to stop the task.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskStop($oService, $vTask)
	Local $oTask
	; Check if the Task exists when specified as \Folder\Task
	If Not IsObj($vTask) Then
		$oTask = _TS_TaskGet($oService, $vTask)
		If @error Then Return SetError(2301, @error, 0)
	Else
		$oTask = $vTask
	EndIf
	; Stop the task. The flags parameter is 0 because it is not supported
	$oTask.Stop(0)
	If @error Then Return SetError(2302, @error, 0)
	Return 1
EndFunc   ;==>_TS_TaskStop

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskUpdate
; Description ...: Update a Task.
; Syntax.........: _TS_TaskUpdate($oService, $oTask, $oTaskDefinition)
; Parameters ....: $oService        - Task Scheduler Service object as returned by _TS_Open
;                  $oTask           - Registered Task object
;                  $oTaskDefinition - TaskDefinition object of the Registered Task
; Return values .: Success - Task object
;                  Failure - Returns 0 and sets @error:
;                  |3301 - Parameter $oService is not an object or not an ITaskService object
;                  |3302 - Parameter $oTask is not an object or not an IRegisteredTask object
;                  |3303 - Error accessing $sFolder using _TS_FolderGet. @extended is set to the COM error
;                  |3304 - Error updating the Task. @extended is set to the COM error
;                  |3305 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskUpdate($oService, $oTask, $oTaskDefinition)
	If Not IsObj($oService) Or ObjName($oService) <> "ITaskService" Then Return SetError(3301, 0, 0)
	If Not IsObj($oTask) Or ObjName($oTask) <> "IRegisteredTask" Then Return SetError(3302, 0, 0)
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(3305, 0, 0)
	Local $sTaskPath = $oTask.Path
	Local $iPos = StringInStr($sTaskPath, "\", $STR_NOCASESENSE, -1)
	Local $sFolder = ($iPos = 1) ? "\" : (StringLeft($sTaskPath, $iPos - 1))
	Local $sTask = StringMid($sTaskPath, $iPos + 1)
	Local $oFolder = _TS_FolderGet($oService, $sFolder)
	If @error Then Return SetError(3303, @error, 0)
	$oFolder.RegisterTaskDefinition($sTask, $oTaskDefinition, $TASK_UPDATE, "", "", $TASK_LOGON_NONE)
	If @error Then Return SetError(3304, @error, 0)
	Return $oTask
EndFunc   ;==>_TS_TaskUpdate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TaskValidate
; Description ...: Validate the Task Definition.
; Syntax.........: _TS_TaskValdiate()
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $vTask    - Object to validate. Can be a Task Definition or a Registered Task object
;                              The Registered Task can be specified as a string as well e.g. "\folder\task"
; Return values .: Success - Returns a zero-based 2D array holding the following information:
;                  |0 - Errornumber as described in section Remarks
;                  |1 - Severity: I (Information), W (Warning, E (Error)
;                  Failure - Returns "" and sets @error
;                  |2401 - The Task does not exist. @extended is set to the COM error code returned by _TS_TaskGet
;                  |240101 - You have to define at least one Action
;                  |240102 - Action type is unsupported
;                  |240103 - Action ID has to be unique
;                  |240501 - You should at least define one Trigger
;                  |245001 - Make sure to provide Userid and Password for the selected logon type
; Author ........: water
; Modified.......:
; Remarks .......: This function is necessary because the validate function of the Registration method does not return any meaningful results.
;                  The function first checks the integrity of each object (triggers, actions ...) then the task as a whole.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TaskValidate($oService, $vTask)
	Local $iIndex = 0, $aCheckResult[1000][7]     ; Errornumber, Severity, ErrorText, ObjectType, ObjectID, ObjectIndex, Comment
	Local $oTaskDefinition
	; Check if the Task exists when specified as \Folder\Task
	If Not IsObj($vTask) Then
		$oTaskDefinition = _TS_TaskGet($oService, $vTask)
		If @error Then Return SetError(2401, @error, 0)
	Else
		$oTaskDefinition = $vTask
	EndIf

	; **********************************
	; Validate each collection or object
	; **********************************
	; Actions - ErrorNumber = 2401nn
	Local $sIDs = ""
	; # of actions > 0
	If $oTaskDefinition.Actions.Count() = 0 Then __TS_TaskValidateWrite($aCheckResult, $iIndex, 240101, "E", "ACTIONS")
	For $i = 1 To $oTaskDefinition.Actions.Count()
		; Check for unsupported Action types
		If $oTaskDefinition.Actions($i).Type <> $TASK_ACTION_EXEC Then __TS_TaskValidateWrite($aCheckResult, $iIndex, 240102, "E", "ACTIONS", $oTaskDefinition.Actions($i).ID, $i, _
				"Current Action type is " & $oTaskDefinition.Actions($i).Type)
		; Check for unique Action IDs
		If StringInStr($sIDs, "|" & $oTaskDefinition.Actions($i).ID & "|") > 0 Then
			__TS_TaskValidateWrite($aCheckResult, $iIndex, 240103, "E", "ACTIONS", $oTaskDefinition.Actions($i).ID, $i, "Current Action ID is " & $oTaskDefinition.Actions($i).ID)
		Else
			$sIDs = $sIDs & "|" & $oTaskDefinition.Actions($i).ID & "|"
		EndIf
	Next

	; Principal - ErrorNumber = 2402nn

	; RegistrationInfo - ErrorNumber = 2403nn

	; Settings - ErrorNumber = 2404nn

	; Triggers - ErrorNumber = 2405nn
	If $oTaskDefinition.Triggers.Count() = 0 Then __TS_TaskValidateWrite($aCheckResult, $iIndex, 240501, "W", "TRIGGERS")

	; *********************************************************************************
	; Check the Task as a whole and the relation between objects - ErrorNumber = 2450nn
	; *********************************************************************************
	If $oTaskDefinition.Principal.LogonType = $TASK_LOGON_PASSWORD Or $oTaskDefinition.Principal.LogonType = $TASK_LOGON_INTERACTIVE_TOKEN Or _
			$oTaskDefinition.Principal.LogonType = $TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD Then __TS_TaskValidateWrite($aCheckResult, $iIndex, 245001, "I", "PRINCIPAL")

	ReDim $aCheckResult[$iIndex][7]
	Return $aCheckResult
EndFunc   ;==>_TS_TaskValidate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TriggerCreate
; Description ...: Create a new Trigger object for a new or already Registered Task.
; Syntax.........: _TS_TriggerCreate(($oTaskDefinition, $iTriggerType[, $sId = ""])
; Parameters ....: $oTaskDefinition - Task Definition object to add this Trigger to.
;                  $iTriggerType    - Type of Trigger to use. Can be any of the TASK_TRIGGER_TYPE2 enumeration
;                  $sID             - [optional] ID for easier access to the Trigger
; Return values .: Success - Object of the created Trigger
;                  Failure - Returns 0 and sets @error
;                  |2501 - $oTaskDefinition isn't an object or not a Task Definition object
;                  |2502 - The Trigger could not be created. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TriggerCreate($oTaskDefinition, $iTriggerType, $sID = "")
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(2501, @error, 0)
	Local $oTrigger = $oTaskDefinition.Triggers.Create($iTriggerType)
	If @error Then Return SetError(2502, @error, 0)
	If $sID <> "" Then $oTrigger.ID = $sID
	Return $oTrigger
EndFunc   ;==>_TS_TriggerCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_TriggerDelete
; Description ...: Delete a single or all Trigger objects by ID or index.
; Syntax.........: _TS_TriggerDelete($oTaskDefinition, $iIndex[, $sID = ""[, $bDeleteAll = False]])
; Parameters ....: $oTaskDefinition - Task Definition object of a new or Registered Task
;                  $iIndex          - Delete the Trigger with the specified index (one based)
;                  $sID             - [optional] Deletes all Triggers with the same ID (default = "")
;                  $bDeleteAll      - [optional] Removes all Triggers (default = False)
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |2602 - The Trigger could not be deleted. @extended is set to the COM error code
;                  |2603 - The Triggers could not be deleted. @extended is set to the COM error code
;                  |2604 - Either $iIndex or $sID has to be specified when $bDeleteAll is set to False
; Author ........: water
; Modified.......:
; Remarks .......: Set one of this three parameters to delete specific or all Triggers: $iIndex, $sID, $bDeleteAll.
;                  The parameters will be processed in the following sequence:
;                  If $iIndex > 0 then delete by index, else if $sID <> "" then delete by ID, else if $bDeleteAll is True then delete all Triggers
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_TriggerDelete($oTaskDefinition, $iIndex, $sID = "", $bDeleteAll = False)
	If $iIndex = Default Then $iIndex = 0
	If $sID = Default Then $sID = ""
	If $bDeleteAll = Default Then $bDeleteAll = False
	If $iIndex = 0 And $sID = "" And $bDeleteAll = False Then Return SetError(2604, 0, 0)
	If $iIndex > 0 Then
		$oTaskDefinition.Triggers.Remove($iIndex)
		If @error Then Return SetError(2602, @error, 0)
	ElseIf $sID <> "" Then
		For $i = 1 To $oTaskDefinition.Triggers.Count
			If $oTaskDefinition.Triggers.Item($i).ID = $sID Then
				$oTaskDefinition.Triggers.Remove($i)
				If @error Then Return SetError(2602, @error, 0)
				ExitLoop
			EndIf
		Next
	ElseIf $bDeleteAll Then
		$oTaskDefinition.Triggers.Clear()
		If @error Then Return SetError(2603, @error, 0)
	EndIf
	Return 1
EndFunc   ;==>_TS_TriggerDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_VersionInfo
; Description ...: Returns an array of information about the UDF.
; Syntax.........: _TS_VersionInfo()
; Parameters ....: None
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - Release Type (T=Test or V=Production)
;                  |2 - Major Version
;                  |3 - Minor Version
;                  |4 - Sub Version
;                  |5 - Release Date (YYYYMMDD)
;                  |6 - AutoIt version required
;                  |7 - List of authors separated by ","
;                  |8 - List of contributors separated by ","
;                  Failure - None
; Author ........: water
; Modified.......:
; Remarks .......: Based on function _IE_VersionInfo written bei Dale Hohm
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_VersionInfo()
	Local $aVersionInfo[9] = [8, "V", 1, 6, 0.0, "20211119", "3.3.14.5", "water", "allow2010, AdmUL, water"]
	Return $aVersionInfo
EndFunc   ;==>_TS_VersionInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_Wrapper_ActionCreate
; Description ...: Create the Actions to be executed when the Task runs.
; Syntax.........: _TS_Wrapper_ActionCreate($oTaskDefinition, $sPath, $sWorkingDirectory = "", $sArguments = "")
; Parameters ....: $oTaskDefinition   - Task Definition object as returned by _TS_Wrapper_TaskCreate
;                  $sPath             - Path to an executable file
;                  $sWorkingDirectory - [optional] Directory that contains either the executable file or the files that are used by the executable file
;                  $sArguments        - [optional] Arguments pased to the executable file
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |2701 - Error returned when accessing the Actions collection. @extended is set to the COM error code
;                  |2702 - Error returned when creating the Action object. @extended is set to the COM error code
;                  |2703 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
; Author ........: water
; Modified.......:
; Remarks .......: Populates the Action object
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Wrapper_ActionCreate($oTaskDefinition, $sPath, $sWorkingDirectory = "", $sArguments = "")
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(2703, 0, 0)
	Local $oActions = $oTaskDefinition.Actions
	If @error Then Return SetError(2701, @error, 0)
	$oActions.Context = "Author"
	Local $oAction = $oActions.Create($TASK_ACTION_EXEC)
	If @error Then Return SetError(2702, @error, 0)
	$oAction.Path = $sPath
	$oAction.WorkingDirectory = $sWorkingDirectory
	$oAction.Arguments = $sArguments
	Return 1
EndFunc   ;==>_TS_Wrapper_ActionCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_Wrapper_PrincipalSet
; Description ...: Set the Principal properties of a Task Definition object.
; Syntax.........: _TS_Wrapper_PrincipalSet($oTaskDefinition, $iLogonType[, $iRunLevel = $TASK_RUNLEVEL_LUA[, $sUserId = "", $sGroupId = ""]])
; Parameters ....: $oTaskDefinition - Task Definition object as returned by _TS_Wrapper_TaskCreate
;                  $iLogonType      - Sets the security logon method that is required to run the Tasks that are associated with the principal.
;                                     Can be any of the TASK_LOGON_TYPE constants.
;                  $iRunLevel       - [optional] Sets the identifier that is used to specify the privilege level that is required to run the Tasks that are associated with the principal.
;                                     Can be any of the TASK_RUNLEVEL_TYPE constants (default = $TASK_RUNLEVEL_LUA (Tasks will be run with the least privileges (LUA))
;                  $sUserId         - [optional] Sets the user identifier that is required to run the Tasks that are associated with the principal.
;                  $sGroupId        - [optional] Sets the group identifier that is required to run the Tasks that are associated with the principal.
; Return values .: Success - 1
;                  Failure - Returns 0 and sets @error:
;                  |2801 - Error creating the Task Definition. @extended is set to the COM error code
;                  |2802 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
; Author ........: water
; Modified.......:
; Remarks .......: Populates the RegistrationInfo object.
;                  Either set $sUserId or $sGroupId - if any
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Wrapper_PrincipalSet($oTaskDefinition, $iLogonType, $iRunLevel = $TASK_RUNLEVEL_LUA, $sUserId = "", $sGroupId = "")
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(2802, 0, 0)
	If $iRunLevel = Default Then $iRunLevel = $TASK_RUNLEVEL_LUA
	If $sUserId = Default Then $sUserId = ""
	If $sGroupId = Default Then $sGroupId = ""
	;	If $sUserId = "" And $sGroupId = "" Then $sGroupId = "S-1-5-32-545"
	;	If $sUserId <> "" And $sGroupId <> "" Then $sGroupId = ""
	; Set the Principal object
	$oTaskDefinition.Principal.LogonType = $iLogonType
	$oTaskDefinition.Principal.RunLevel = $iRunLevel
	$oTaskDefinition.Principal.Id = "Author"
	If $sUserId <> "" Then $oTaskDefinition.Principal.UserId = $sUserId
	If $sGroupId <> "" Then $oTaskDefinition.Principal.GroupId = $sGroupId
	Return 1
EndFunc   ;==>_TS_Wrapper_PrincipalSet

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_Wrapper_TaskCreate
; Description ...: Create a Task Definition object so you can then set all needed properties using other Wrapper functions.
; Syntax.........: _TS_Wrapper_TaskCreate($oService[, $sDescription = ""[, $sDocumentation = ""[, $sAuthor = @UserName[, $sDate = _NowCalc()]]]])
; Parameters ....: $oService - Task Scheduler Service object as returned by _TS_Open
;                  $sDescription   - [optional] Describe the Task in a single line
;                  $sDocumentation - [optional] Extended documentation of the Task
;                  $sAuthor        - [optional] Author of the Task. If not specified @UserName is used (default)
;                  $sDate          - [optional] Date when the Task was created. If not specified _Now() is usded (default)
;                                    Format has to be: YYYY-MM-DD-THH:MM:SS
; Return values .: Success - Task Definition object
;                  Failure - Returns 0 and sets @error:
;                  |2901 - Error creating the Task Definition. @extended is set to the COM error code
;                  |2902 - Error setting property Date. Please check the correct format as described above. @extended is set to the COM error code
;                  |2903 - Parameter $oService is not an object or not an ITaskService object
; Author ........: water
; Modified.......:
; Remarks .......: Creates the Task Definition object and populates the RegistrationInfo object
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Wrapper_TaskCreate($oService, $sDescription = "", $sDocumentation = "", $sAuthor = Default, $sDate = Default)
	If Not IsObj($oService) Or ObjName($oService) <> "ITaskService" Then Return SetError(2903, 0, 0)
	If $sAuthor = Default Or $sAuthor = "" Then $sAuthor = @UserName
	If $sDate = Default Or $sDate = "" Then
		$sDate = StringReplace(_NowCalc(), "/", "-")
		$sDate = StringReplace($sDate, " ", "T")
	EndIf
	; Set the TaskDefinition object. The flags parameter is 0 because it is not supported
	Local $oTaskDefinition = $oService.NewTask(0)
	If @error Then Return SetError(2901, @error, 0)
	; Set the Registrationinfo object
	$oTaskDefinition.RegistrationInfo.Description = $sDescription
	$oTaskDefinition.RegistrationInfo.Documentation = $sDocumentation
	$oTaskDefinition.RegistrationInfo.Author = @UserName
	$oTaskDefinition.RegistrationInfo.Date = $sDate
	If @error Then Return SetError(2902, @error, 0)
	Return $oTaskDefinition
EndFunc   ;==>_TS_Wrapper_TaskCreate

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_Wrapper_TaskRegister
; Description ...: Alias for _TS_TaskRegister: Register a Task in the specified Folder.
; Syntax.........: _TS_Wrapper_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition[, $sUserId = ""[, $sPassword = ""[, $iLogonType = Default]]])
; Parameters ....: $oService        - Task Scheduler Service object as returned by _TS_Open
;                  $sFolder         - Folder where the Task should be created
;                  $sName           - Name of the Task
;                  $oTaskDefinition - Task Definition object as created by _TS_TaskCreate and filled by _TS_TaskPropertiesSet
;                  $sUserId         - [optional] The user credentials that are used to register the Task. If present, these credentials
;                                     take priority over the credentials specified in the Task Definition object pointed to by the definition parameter
;                  $sPassword       - [optional] The password for the UserId that is used to register the Task. When the TASK_LOGON_SERVICE_ACCOUNT logon type
;                                     is used, the password must be an empty value such as NULL or ""
;                  $iLogonType      - [optional] Can be any of the TASK_LOGON_TYPE constants enumeration. For the default please check Remarks
;                  $iCreateFlag     - [optional] Defines if to create or update the task. Can be any of the TASK_CREATE constants enumeration. Default is $TASK_CREATE
; Return values .: Success - Task object
;                  Failure - Returns 0 and sets @error:
;                  |2101 - Parameter $oService is not an object or not an ITaskService object
;                  |2102 - $sFolder does not exist or an error occurred in _TS_FolderExists. @extended is set to the COM error (if any)
;                  |2103 - Task exists which is incompatible with flags $TASK_CREATE, $TASK_DISABLE and $TASK_CREATE or
;                  |       Task does not exist which is incompatible with flags $TASK_UPDATE And $TASK_DONT_ADD_PRINCIPAL_ACE
;                  |2104 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
;                  |2105 - Error accessing $sFolder using _TS_FolderGet. @extended is set to the COM error
;                  |2106 - Error creating the Task. @extended is set to the COM error
; Author ........: water
; Modified.......:
; Remarks .......: This function is a copy of function _TS_TaskRegister added for completeness.
;                  The Return Values are identical to those of _TS_TaskRegister.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Wrapper_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition, $sUserId = Default, $sPassword = Default, $iLogonType = Default)
	Local $vReturnValue = _TS_TaskRegister($oService, $sFolder, $sName, $oTaskDefinition, $sUserId, $sPassword, $iLogonType)
	Return SetError(@error, @extended, $vReturnValue)
EndFunc   ;==>_TS_Wrapper_TaskRegister

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_Wrapper_TriggerDateTime
; Description ...: Creates a date/time based Trigger.
; Syntax.........: _TS_Wrapper_TriggerDateTime($oTaskDefinition, $iTriggerType, $iDoW, $iInterval, $sStart[, $sEnd = ""])
; Parameters ....: $oTaskDefinition - Task Definition object as returned by _TS_Wrapper_TaskCreate
;                  $iTriggerType    - Type of Trigger to use. Only supports $TASK_TRIGGER_TIME, $TASK_TRIGGER_DAILY and $TASK_TRIGGER_WEEKLY
;                  $iDoW            - A bitwise mask that indicates the days of the week on which the Task runs. Possible values:
;                  |Sunday - 1
;                  |Monday - 2
;                  |Tuesday - 4
;                  |Wednesday - 8
;                  |Thursday - 16
;                  |Friday - 32
;                  |Saturday - 64
;                  |10 means: Run the schedule on Monday and Wednesday
;                  $iInterval       - The interval between the days ($TASK_TRIGGER_DAILY) or weeks ($TASK_TRIGGER_WEEKLY) in the schedule
;                  $sStart          - The date and time when the Trigger is activated. Format: YYYY-MM-DDTHH:MM:SS(+-)HH:MM. See Remarks
;                  $sEnd            - [optional] The date and time when the Trigger is deactivated. Format: YYYY-MM-DDTHH:MM:SS(+-)HH:MM. See Remarks
; Return values .: Success - Object of the created schedule
;                  Failure - Returns 0 and sets @error
;                  |3101 - Invalid $iTriggerType specified. Has to be $TASK_TRIGGER_TIME, $TASK_TRIGGER_DAILY or $TASK_TRIGGER_WEEKLY
;                  |3102 - Error returned when creating the Trigger object. @extended is set to the COM error code
;                  |3103 - Error setting property StartBoundary. Please check the correct format as described above. @extended is set to the COM error code
;                  |3104 - Error setting property EndBoundary. Please check the correct format as described above. @extended is set to the COM error code
;                  |3105 - Error setting property ExecutionTimeLimit. Please check the correct format as described above. @extended is set to the COM error code
;                  |3106 - Error setting property DaysOfWeek. Please check the correct format. @extended is set to the COM error code
;                  |3107 - Error setting property Weeksinterval. Please check the correct format. @extended is set to the COM error code
;                  |3108 - Error setting property DaysInterval. Please check the correct format. @extended is set to the COM error code
;                  |3109 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
;                  |3110 - Error returned when accessing the Triggers collection. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: The (+-)HH:MM section describes the time zone as a certain number of hours ahead or behind Coordinated Universal Time (Greenwich Mean Time).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Wrapper_TriggerDateTime($oTaskDefinition, $iTriggerType, $iDoW, $iInterval, $sStart, $sEnd = "")
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(3109, 0, 0)
	If $iTriggerType <> $TASK_TRIGGER_TIME And $iTriggerType <> $TASK_TRIGGER_DAILY And $iTriggerType <> $TASK_TRIGGER_WEEKLY Then Return SetError(3101, 0, 0)
	; Create a time-based Trigger
	Local $oTriggers = $oTaskDefinition.Triggers
	If @error Then Return SetError(3110, @error, 0)
	Local $oTrigger = $oTriggers.Create($iTriggerType)
	If @error Then Return SetError(3102, @error, 0)
	; Trigger variables that define when the Trigger is active
	$oTrigger.StartBoundary = $sStart
	If @error Then Return SetError(3103, @error, 0)
	$oTrigger.EndBoundary = $sEnd
	If @error Then Return SetError(3104, @error, 0)
	$oTrigger.ExecutionTimeLimit = "PT5M"     ; Five minutes
	If @error Then Return SetError(3105, @error, 0)
	$oTrigger.Id = "_TS_Wrapper_"     ; TriggerDateTime_1"
	$oTrigger.Enabled = True
	If $iTriggerType = $TASK_TRIGGER_DAILY Then
		$oTrigger.DaysInterval = $iInterval
		If @error Then Return SetError(3108, @error, 0)
	EndIf
	If $iTriggerType = $TASK_TRIGGER_WEEKLY Then
		$oTrigger.Daysofweek = $iDoW
		If @error Then Return SetError(3106, @error, 0)
		$oTrigger.WeeksInterval = $iInterval
		If @error Then Return SetError(3107, @error, 0)
	EndIf
	Return $oTrigger
EndFunc   ;==>_TS_Wrapper_TriggerDateTime

; #FUNCTION# ====================================================================================================================
; Name...........: _TS_Wrapper_TriggerLogon
; Description ...: Creates a logon Trigger.
; Syntax.........: _TS_Wrapper_TriggerLogon($oTriggers, $iDelay, $sStart[, $sEnd = ""])
; Parameters ....: $oTaskDefinition - Task Definition object as returned by _TS_Wrapper_TaskCreate
;                  $iDelay          - Value in minutes indicating the time between the users logon and the start of the Task
;                  $sStart          - The date and time when the Trigger is activated. Format: YYYY-MM-DDTHH:MM:SS(+-)HH:MM. See Remarks
;                  $sEnd            - [optional] The date and time when the Trigger is deactivated. Format: YYYY-MM-DDTHH:MM:SS(+-)HH:MM. See Remarks
; Return values .: Success - Object of the created schedule
;                  Failure - Returns 0 and sets @error
;                  |3201 - Error returned when creating the Trigger object. @extended is set to the COM error code
;                  |3202 - Error setting property StartBoundary. Please check the correct format as described above. @extended is set to the COM error code
;                  |3203 - Error setting property EndBoundary. Please check the correct format as described above. @extended is set to the COM error code
;                  |3204 - Error setting property ExecutionTimeLimit. Please check the correct format as described above. @extended is set to the COM error code
;                  |3205 - Error setting property Delay. Please check the correct format as described above. @extended is set to the COM error code
;                  |3206 - Parameter $oTaskDefinition is not an object or not an ITaskDefinition object
;                  |3207 - Error returned when accessing the Triggers collection. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: The (+-)HH:MM section describes the time zone as a certain number of hours ahead or behind Coordinated Universal Time (Greenwich Mean Time).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _TS_Wrapper_TriggerLogon($oTaskDefinition, $iDelay, $sStart, $sEnd = "")
	If Not IsObj($oTaskDefinition) Or ObjName($oTaskDefinition) <> "ITaskDefinition" Then Return SetError(3206, 0, 0)
	; Create a logon Trigger
	Local $oTriggers = $oTaskDefinition.Triggers
	If @error Then Return SetError(3207, @error, 0)
	Local $oTrigger = $oTriggers.Create($TASK_TRIGGER_LOGON)
	If @error Then Return SetError(3201, @error, 0)
	; Trigger variables that define when the Trigger is active
	$oTrigger.StartBoundary = $sStart
	If @error Then Return SetError(3202, @error, 0)
	$oTrigger.EndBoundary = $sEnd
	If @error Then Return SetError(3203, @error, 0)
	$oTrigger.ExecutionTimeLimit = "PT5M"     ; Five minutes
	If @error Then Return SetError(3204, @error, 0)
	$oTrigger.Id = "_TS_Wrapper_TriggerLogon_1"
	$oTrigger.Delay = "PT" & $iDelay & "M"     ; n minutes
	If @error Then Return SetError(3205, @error, 0)
	$oTrigger.Enabled = True
	$oTrigger.UserId = @UserName
	Return $oTrigger
EndFunc   ;==>_TS_Wrapper_TriggerLogon

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_ErrorHandler
; Description ...: Called if an ObjEvent error occurs.
; Syntax.........: __TS_ErrorHandler()
; Parameters ....: None
; Return values .: @error is set to the COM error by AutoIt
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __TS_ErrorHandler()
	Local $bHexNumber = Hex($__oTS_Error.number, 8)
	Local $aVersionInfo = _TS_VersionInfo()
	Local $sError = "COM Error Encountered in " & @ScriptName & @CRLF & _
			"TS UDF version = " & $aVersionInfo[2] & "." & $aVersionInfo[3] & "." & $aVersionInfo[4] & @CRLF & _
			"@AutoItVersion = " & @AutoItVersion & @CRLF & _
			"@AutoItX64 = " & @AutoItX64 & @CRLF & _
			"@Compiled = " & @Compiled & @CRLF & _
			"@OSArch = " & @OSArch & @CRLF & _
			"@OSVersion = " & @OSVersion & @CRLF & _
			"Scriptline = " & $__oTS_Error.scriptline & @CRLF & _
			"NumberHex = " & $bHexNumber & @CRLF & _
			"Number = " & $__oTS_Error.number & @CRLF & _
			"WinDescription = " & StringStripWS($__oTS_Error.WinDescription, $STR_STRIPTRAILING) & @CRLF & _
			"Description = " & StringStripWS($__oTS_Error.Description, $STR_STRIPTRAILING) & @CRLF & _
			"Source = " & $__oTS_Error.Source & @CRLF & _
			"HelpFile = " & $__oTS_Error.HelpFile & @CRLF & _
			"HelpContext = " & $__oTS_Error.HelpContext & @CRLF & _
			"LastDllError = " & $__oTS_Error.LastDllError
	If $__iTS_Debug > 0 Then
		If $__iTS_Debug = 1 Then ConsoleWrite($sError & @CRLF & "========================================================" & @CRLF)
		If $__iTS_Debug = 2 Then MsgBox(64, "TS UDF - Debug Info", $sError)
		If $__iTS_Debug = 3 Then FileWrite($__sTS_DebugFile, @YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & @CRLF & _
				"-------------------" & @CRLF & $sError & @CRLF & "========================================================" & @CRLF)
	EndIf
EndFunc   ;==>__TS_ErrorHandler

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_TaskListWrite
; Description ...: Helper function for _TS_TaskList to write a record to an array.
; Syntax.........: __TS_TaskListWrite(ByRef $aTasks, $iRow, ByRef $iColumn, $bReadable, $vValue)
; Parameters ....: $aTasks    - Array to write the values to
;                  $iRow      - Index of the row in the array to write to
;                  $iColumn   - Index of the column in the array to write to
;                  $bReadable - Format of the array. True = User friendly format. See function _TS_TaskList
;                  $vValue    - Data to write to the array
; Return values .: None
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __TS_TaskListWrite(ByRef $aTasks, $iRow, ByRef $iColumn, $bReadable, $vValue)
	If $bReadable Then
		$aTasks[$iRow][$iColumn] = ($vValue == "19991130000000" Or $vValue == "18991230000000" Or $vValue == "1999/11/30 00:00:00" Or $vValue == "1899/12/30 00:00:00") _
				 ? "" : ($vValue)     ; Drop date values meaning "never used", "never run"
	Else
		$aTasks[$iRow][$iColumn] = $vValue
	EndIf
	$iColumn = $iColumn + 1
EndFunc   ;==>__TS_TaskListWrite

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_PropertyGetWrite
; Description ...: Writes a Task property to a 1D or 2D array.
; Syntax.........: __TS_PropertyGetWrite(ByRef $aProperties, ByRef $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $sProperty, $sValue[, $sComment = ""])
; Parameters ....: $aProperties      - Array to write the values to
;                  $iIndex           - Index of the row in the array to write to
;                  $sSection         - Name of the Task Scheduler COM object the property belongs to
;                  $sQuerySection    - Name of the Task Scheduler COM object to process
;                  $sQueryProperties - Comma separated list of properties to return
;                  $iFormat          - Format of the array. 1 = User friendly format (2D array), 2 = Format you can use as input to _TS_TaskPropertiesSet (1D array)
;                  $sProperty        - Name of the property
;                  $sValue           - Value of the property
;                  $sComment         - [optional] Comment
; Return values .: None
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __TS_PropertyGetWrite(ByRef $aProperties, ByRef $iIndex, $sSection, $sQuerySection, $sQueryProperties, $iFormat, $bIgnoreNoValues, $sProperty, $sValue, $sComment = "")
	If $bIgnoreNoValues And $sValue = "" Then Return
	If $sQuerySection <> "" And $sSection <> $sQuerySection Then Return
	If $sQueryProperties <> "" And StringInStr("," & $sQueryProperties & ",", "," & $sProperty & ",") = 0 Then Return
	If $iFormat = 1 Then
		$aProperties[$iIndex][0] = $sSection
		$aProperties[$iIndex][1] = $sProperty
		$aProperties[$iIndex][2] = $sValue
		$aProperties[$iIndex][3] = $sComment
		$iIndex = $iIndex + 1
		If Mod($iIndex, 1000) = 0 Then ReDim $aProperties[$iIndex + 1000][4]
	Else
		$aProperties[$iIndex] = $sSection & "|" & $sProperty & "|" & $sValue
		$iIndex = $iIndex + 1
		If Mod($iIndex, 1000) = 0 Then ReDim $aProperties[$iIndex + 1000][4]
	EndIf
	Return
EndFunc   ;==>__TS_PropertyGetWrite

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_TaskValidateWrite
; Description ...: Writes a validate message to a 2D array.
; Syntax.........: __TS_TaskValidateWrite(ByRef $aCheck, ByRef $iIndex, $iErrorNumber, $sSeverity[, $sObjectType[, $sObjectID[, $sObjectIndex[, $sComment = ""]]]])
; Parameters ....: $aCheck       - Array to write the messages to
;                  $iIndex       - Index of the row in the array to write to
;                  $iErrorNumber - Number of the validation message
;                  $sSeverity    - Type of the error: I - Information, W - Warning, E - Error
;                  $sObjectType  - [optional] Name of the sub-object e.g. ACTIONS, SETTINGS, TRIGGERS etc.
;                  $sObjectID    - [optional] ID of the sub-object in the collection e.g. ACTIONS, TRIGGERS
;                  $sObjectIndex - [optional] Index (one-based) of the sub-object in the collection
;                  $sComment     - [optional] Comment
; Return values .: Success - zero-based two dimensional array with the following elements
;                  |0 - See $iErrorNumber in section Parameters
;                  |1 - See $sSeverity in section Parameters
;                  |2 - Error Text
;                  |3 - See $sObjectType in section Parameters
;                  |4 - See $sObjectID in section Parameters
;                  |5 - See $sObjectIndex in section Parameters
;                  |6 - See $sComment in section Parameters
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __TS_TaskValidateWrite(ByRef $aCheck, ByRef $iIndex, $iErrorNumber, $sSeverity, $sObjectType = "", $sObjectID = "", $iObjectIndex = "", $sComment = "")
	$aCheck[$iIndex][0] = $iErrorNumber
	$aCheck[$iIndex][1] = $sSeverity
	$aCheck[$iIndex][2] = _TS_ErrorText($iErrorNumber, False)
	$aCheck[$iIndex][3] = $sObjectType
	$aCheck[$iIndex][4] = $sObjectID
	$aCheck[$iIndex][5] = $iObjectIndex
	$aCheck[$iIndex][6] = $sComment
	$iIndex = $iIndex + 1
	If Mod($iIndex, 1000) = 0 Then ReDim $aCheck[$iIndex + 1000][7]
	Return
EndFunc   ;==>__TS_TaskValidateWrite

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_ConvertDaysOfMonth
; Description ...: Translates the DaysOfMonth values of a monthly Trigger to readable text
; Syntax.........: __TS_ConvertDaysOfMonth($dBitValue)
; Parameters ....: $dBitValue - Bitvalue to translate
; Return values .: Success - String holding all set bits as day-of-month (number) separated by ", "
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/de-de/windows/win32/taskschd/monthlytrigger-daysofmonth
; Example .......:
; ===============================================================================================================================
Func __TS_ConvertDaysOfMonth($dBitValue)
	Local $sMsg = ""
	If BitAND($dBitValue, 0x01) Then $sMsg = "1"
	If BitAND($dBitValue, 0x02) Then $sMsg = $sMsg & ", 2"
	If BitAND($dBitValue, 0x04) Then $sMsg = $sMsg & ", 3"
	If BitAND($dBitValue, 0x08) Then $sMsg = $sMsg & ", 4"
	If BitAND($dBitValue, 0x10) Then $sMsg = $sMsg & ", 5"
	If BitAND($dBitValue, 0x20) Then $sMsg = $sMsg & ", 6"
	If BitAND($dBitValue, 0x40) Then $sMsg = $sMsg & ", 7"
	If BitAND($dBitValue, 0x80) Then $sMsg = $sMsg & ", 8"
	If BitAND($dBitValue, 0x100) Then $sMsg = $sMsg & ", 9"
	If BitAND($dBitValue, 0x200) Then $sMsg = $sMsg & ", 10"
	If BitAND($dBitValue, 0x400) Then $sMsg = $sMsg & ", 11"
	If BitAND($dBitValue, 0x800) Then $sMsg = $sMsg & ", 12"
	If BitAND($dBitValue, 0x1000) Then $sMsg = $sMsg & ", 13"
	If BitAND($dBitValue, 0x2000) Then $sMsg = $sMsg & ", 14"
	If BitAND($dBitValue, 0x4000) Then $sMsg = $sMsg & ", 15"
	If BitAND($dBitValue, 0x8000) Then $sMsg = $sMsg & ", 16"
	If BitAND($dBitValue, 0x10000) Then $sMsg = $sMsg & ", 17"
	If BitAND($dBitValue, 0x20000) Then $sMsg = $sMsg & ", 18"
	If BitAND($dBitValue, 0x40000) Then $sMsg = $sMsg & ", 19"
	If BitAND($dBitValue, 0x80000) Then $sMsg = $sMsg & ", 20"
	If BitAND($dBitValue, 0x100000) Then $sMsg = $sMsg & ", 21"
	If BitAND($dBitValue, 0x200000) Then $sMsg = $sMsg & ", 22"
	If BitAND($dBitValue, 0x400000) Then $sMsg = $sMsg & ", 23"
	If BitAND($dBitValue, 0x800000) Then $sMsg = $sMsg & ", 24"
	If BitAND($dBitValue, 0x1000000) Then $sMsg = $sMsg & ", 25"
	If BitAND($dBitValue, 0x2000000) Then $sMsg = $sMsg & ", 26"
	If BitAND($dBitValue, 0x4000000) Then $sMsg = $sMsg & ", 27"
	If BitAND($dBitValue, 0x8000000) Then $sMsg = $sMsg & ", 28"
	If BitAND($dBitValue, 0x10000000) Then $sMsg = $sMsg & ", 29"
	If BitAND($dBitValue, 0x20000000) Then $sMsg = $sMsg & ", 30"
	If BitAND($dBitValue, 0x40000000) Then $sMsg = $sMsg & ", 31"
	If BitAND($dBitValue, 0x80000000) Then $sMsg = $sMsg & ", LAST"
	If StringLeft($sMsg, 2) = ", " Then $sMsg = StringMid($sMsg, 3)     ; Remove leading ", "
	Return $sMsg
EndFunc   ;==>__TS_ConvertDaysOfMonth

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_ConvertDaysOfWeek
; Description ...: Translates the DaysOfWeek values of a weekly or monthly-day-of-week Trigger to readable text
; Syntax.........: __TS_ConvertDaysOfWeek($dBitValue)
; Parameters ....: $dBitValue - Bitvalue to translate
; Return values .: Success - String holding all set bits as day-of-week (text) separated by ", "
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/de-de/windows/win32/taskschd/weeklytrigger-daysofweek
; Example .......:
; ===============================================================================================================================
Func __TS_ConvertDaysOfWeek($dBitValue)
	Local $sMsg = ""
	If BitAND($dBitValue, 1) Then $sMsg = "Sunday"
	If BitAND($dBitValue, 2) Then $sMsg = $sMsg & ", Monday"
	If BitAND($dBitValue, 4) Then $sMsg = $sMsg & ", Tuesday"
	If BitAND($dBitValue, 8) Then $sMsg = $sMsg & ", Wednesday"
	If BitAND($dBitValue, 16) Then $sMsg = $sMsg & ", Thursday"
	If BitAND($dBitValue, 32) Then $sMsg = $sMsg & ", Friday"
	If BitAND($dBitValue, 64) Then $sMsg = $sMsg & ", Saturday"
	If StringLeft($sMsg, 2) = ", " Then $sMsg = StringMid($sMsg, 3)     ; Remove leading ", "
	Return $sMsg
EndFunc   ;==>__TS_ConvertDaysOfWeek

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_ConvertMonthsOfYear
; Description ...: Translates the MonthsOfYear values of a monthly or monthly-day-of-week Trigger to readable text
; Syntax.........: __TS_ConvertMonthsOfYear($dBitValue)
; Parameters ....: $dBitValue - Bitvalue to translate
; Return values .: Success - String holding all set bits as month-of-year (text) separated by ", "
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/de-de/windows/win32/taskschd/monthlytrigger-monthsofyear
; Example .......:
; ===============================================================================================================================
Func __TS_ConvertMonthsOfYear($dBitValue)
	Local $sMsg = ""
	If BitAND($dBitValue, 1) Then $sMsg = "January"
	If BitAND($dBitValue, 2) Then $sMsg = $sMsg & ", February"
	If BitAND($dBitValue, 4) Then $sMsg = $sMsg & ", March"
	If BitAND($dBitValue, 8) Then $sMsg = $sMsg & ", April"
	If BitAND($dBitValue, 16) Then $sMsg = $sMsg & ", May"
	If BitAND($dBitValue, 32) Then $sMsg = $sMsg & ", June"
	If BitAND($dBitValue, 64) Then $sMsg = $sMsg & ", July"
	If BitAND($dBitValue, 128) Then $sMsg = $sMsg & ", August"
	If BitAND($dBitValue, 256) Then $sMsg = $sMsg & ", September"
	If BitAND($dBitValue, 512) Then $sMsg = $sMsg & ", October"
	If BitAND($dBitValue, 1024) Then $sMsg = $sMsg & ", November"
	If BitAND($dBitValue, 2048) Then $sMsg = $sMsg & ", December"
	If StringLeft($sMsg, 2) = ", " Then $sMsg = StringMid($sMsg, 3)     ; Remove leading ", "
	Return $sMsg
EndFunc   ;==>__TS_ConvertMonthsOfYear

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __TS_ConvertWeeksOfMonth
; Description ...: Translates the WeeksOfMonth values of a monthly-day-of-week Trigger to readable text
; Syntax.........: __TS_ConvertWeeksOfMonth($dBitValue)
; Parameters ....: $dBitValue - Bitvalue to translate
; Return values .: Success - String holding all set bits as week-of-month (text) separated by ", "
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://docs.microsoft.com/de-de/windows/win32/taskschd/monthlydowtrigger-weeksofmonth
; Example .......:
; ===============================================================================================================================
Func __TS_ConvertWeeksOfMonth($dBitValue)
	Local $sMsg = ""
	If BitAND($dBitValue, 1) Then $sMsg = "First"
	If BitAND($dBitValue, 2) Then $sMsg = $sMsg & ", Second"
	If BitAND($dBitValue, 4) Then $sMsg = $sMsg & ", Third"
	If BitAND($dBitValue, 8) Then $sMsg = $sMsg & ", Fourth"
	If StringLeft($sMsg, 2) = ", " Then $sMsg = StringMid($sMsg, 3)     ; Remove leading ", "
	Return $sMsg
EndFunc   ;==>__TS_ConvertWeeksOfMonth
