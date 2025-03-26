#RequireAdmin
#include <GUIConstantsEx.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3> ; For _NowTime function
#include <ProcessConstants.au3> ; For $STDOUT_CHILD, $STDERR_CHILD
#include <StringConstants.au3> ; For $STR_STRIPLEADING
#include <FileConstants.au3>  ; For $FILE_ATTRIBUTE_READONLY

; =============================================================================
; Helper Functions for Process Exit Code (Replacement for ProcessGetExitCode)
; =============================================================================
Func _GetExitCode($iPID)
    ; Open the process with minimal access needed to query information
    Local $hProc = _OpenProcess($iPID)
    If $hProc = 0 Then
        Return -1 ; Could not open the process
    EndIf

    ; Call GetExitCodeProcess from kernel32.dll
    Local $aRet = DllCall("kernel32.dll", "bool", "GetExitCodeProcess", _
                          "handle", $hProc, _
                          "dword*", 0)
    ; Close the process handle
    DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProc)

    If @error Or Not $aRet[0] Then
        Return -1
    EndIf

    ; The exit code is in $aRet[2]
    Return $aRet[2]
EndFunc

Func _OpenProcess($iPID)
    ; PROCESS_QUERY_INFORMATION = 0x0400
    Local Const $PROCESS_QUERY_INFORMATION = 0x0400
    Local $aRet = DllCall("kernel32.dll", "handle", "OpenProcess", _
                          "dword", $PROCESS_QUERY_INFORMATION, _
                          "bool", False, _
                          "dword", $iPID)
    If @error Or $aRet[0] = 0 Then
        Return 0
    EndIf
    Return $aRet[0]
EndFunc

; =============================================================================
; Global Variables & Constants
; =============================================================================
Global Const $g_sToolsSubDir = "Tools"
Global Const $g_sWimlibSubDir = $g_sToolsSubDir & "\wimlib"
Global Const $g_sTempSubDir = "Temp_WimUpdater"

Global $g_sScriptDir = @ScriptDir
Global $g_sToolsFolder = $g_sScriptDir & "\" & $g_sToolsSubDir
Global $g_sWimlibFolder = $g_sScriptDir & "\" & $g_sWimlibSubDir
Global $g_sWimlibExe = $g_sWimlibFolder & "\wimlib-imagex.exe"
Global $g_sTempFolder = $g_sScriptDir & "\" & $g_sTempSubDir

Global $g_bLoggingEnabled = True
Global $g_sLogFile = $g_sScriptDir & "\WIM_Updater.log"

Global Const $UPDATE_TYPE_SETUP = 1
Global Const $UPDATE_TYPE_PE = 2
Global Const $WIMLIB_EXIT_IMAGE_NOT_FOUND = 18

; =============================================================================
; Initial Log Entry & Version Check
; =============================================================================
If Not FileExists($g_sLogFile) Then FileDelete($g_sLogFile)
_WriteLog("--- Script Started ---")
_WriteLog("AutoIt Version: " & @AutoItVersion)
_WriteLog("Script Directory: " & $g_sScriptDir)

; =============================================================================
; Embedded Files Installation & Initial Checks
; =============================================================================
_WriteLog("Checking Tools folder: " & $g_sToolsFolder)
If Not FileExists($g_sToolsFolder) Then
    _WriteLog("Tools folder not found. Creating and extracting files...")
    If Not DirCreate($g_sToolsFolder) Then _ExitWithError("Failed to create Tools directory: " & $g_sToolsFolder & @CRLF & "Check permissions. Error code: " & @error)
    If Not DirCreate($g_sWimlibFolder) Then _ExitWithError("Failed to create wimlib directory: " & $g_sWimlibFolder & @CRLF & "Check permissions. Error code: " & @error)

    If Not FileInstall("Tools\wimlib\wimlib-imagex.exe", $g_sWimlibExe, $FC_OVERWRITE) Then _ExitWithError("Failed to install wimlib-imagex.exe")
    If Not FileInstall("Tools\wimlib\libwim-15.dll", $g_sWimlibFolder & "\libwim-15.dll", $FC_OVERWRITE) Then _ExitWithError("Failed to install libwim-15.dll")
    If Not FileInstall("Tools\1.bmp", $g_sToolsFolder & "\1.bmp", $FC_OVERWRITE) Then _ExitWithError("Failed to install 1.bmp")
    If Not FileInstall("Tools\img0.jpg", $g_sToolsFolder & "\img0.jpg", $FC_OVERWRITE) Then _ExitWithError("Failed to install img0.jpg")
    If Not FileInstall("Tools\spwizimg.dll", $g_sToolsFolder & "\spwizimg.dll", $FC_OVERWRITE) Then _ExitWithError("Failed to install spwizimg.dll")
    _WriteLog("Tools extracted successfully.")
Else
    _WriteLog("Tools folder exists.")
EndIf

If Not FileExists($g_sWimlibExe) Then _ExitWithError("wimlib-imagex.exe is missing from " & $g_sWimlibFolder & ".")
_WriteLog("wimlib-imagex.exe found at: " & $g_sWimlibExe)

Local $hTestFile = FileOpen($g_sScriptDir & "\_testwrite.tmp", $FO_OVERWRITE)
If $hTestFile = -1 Then _ExitWithError("Cannot write to the script directory: " & $g_sScriptDir & @CRLF & "Check permissions.")
FileClose($hTestFile)
FileDelete($g_sScriptDir & "\_testwrite.tmp")
_WriteLog("Write access to script directory confirmed.")

; =============================================================================
; GUI Initialization
; =============================================================================
Global $hGUI = GUICreate("WIM Background Updater by UTZ v1.6", 450, 360)

Global $idLabelWIM = GUICtrlCreateLabel("Select WIM file:", 10, 15, 100, 20)
Global $idInputWIM = GUICtrlCreateInput("", 120, 10, 230, 20)
Global $idBrowseWIM = GUICtrlCreateButton("Browse...", 360, 10, 80, 20)

Global $idGroupPic = GUICtrlCreateGroup("Picture Selection", 10, 40, 430, 80)
Global $idRadioBuiltin = GUICtrlCreateRadio("Use built-in picture", 20, 60, 150, 20)
Global $idRadioCustom = GUICtrlCreateRadio("Use custom picture", 20, 85, 150, 20)
GUICtrlSetState($idRadioBuiltin, $GUI_CHECKED)
Global $idLabelCustomPic = GUICtrlCreateLabel("Custom Pic:", 180, 88, 60, 20)
Global $idInputCustomPic = GUICtrlCreateInput("", 240, 85, 110, 20)
Global $idBrowseCustomPic = GUICtrlCreateButton("Browse...", 360, 85, 80, 20)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $idLabelIndex = GUICtrlCreateLabel("Image Index:", 10, 135, 100, 20)
Global $idInputIndex = GUICtrlCreateInput("2", 120, 130, 50, 20)

Global $idGroupUpdateType = GUICtrlCreateGroup("Update Type", 10, 160, 430, 60)
Global $idRadioUpdateSetup = GUICtrlCreateRadio("Setup Background (.bmp)", 20, 180, 180, 20)
Global $idRadioUpdatePE = GUICtrlCreateRadio("PE Background (.jpg)", 220, 180, 180, 20)
GUICtrlSetState($idRadioUpdateSetup, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

Global $idCheckLogging = GUICtrlCreateCheckbox("Enable Logging (WIM_Updater.log)", 10, 230, 250, 20)
GUICtrlSetState($idCheckLogging, $GUI_CHECKED)

Global $idProcess = GUICtrlCreateButton("Process WIM", 170, 260, 110, 35)
Global $idStatus = GUICtrlCreateLabel("Status: Ready", 10, 305, 430, 45)
GUICtrlSetFont($idStatus, 9)

GUICtrlSetState($idLabelCustomPic, $GUI_DISABLE)
GUICtrlSetState($idInputCustomPic, $GUI_DISABLE)
GUICtrlSetState($idBrowseCustomPic, $GUI_DISABLE)

GUISetState(@SW_SHOW, $hGUI)
_WriteLog("GUI Initialized.")

; =============================================================================
; Main GUI Loop
; =============================================================================
While True
    Local $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            _WriteLog("GUI Closed. Exiting.")
            Exit

        Case $idBrowseWIM
            Local $sSelectedFile = FileOpenDialog("Select WIM File", $g_sScriptDir, "WIM files (*.wim)|All files (*.*)", 1 + 4, "", $hGUI)
            If @error Then ContinueLoop
            _WriteLog("WIM Browse selected: " & $sSelectedFile)
            GUICtrlSetData($idInputWIM, $sSelectedFile)
            GUICtrlSetData($idStatus, "Status: WIM file selected. Ready.")

        Case $idBrowseCustomPic
            Local $sFilter = "Image files (*.bmp;*.jpg;*.jpeg)|All files (*.*)"
            If GUICtrlRead($idRadioUpdateSetup) = $GUI_CHECKED Then
                $sFilter = "Bitmap files (*.bmp)|All files (*.*)"
            ElseIf GUICtrlRead($idRadioUpdatePE) = $GUI_CHECKED Then
                $sFilter = "JPEG files (*.jpg;*.jpeg)|All files (*.*)"
            EndIf
            Local $sSelectedPic = FileOpenDialog("Select Custom Picture", $g_sScriptDir, $sFilter, 1 + 4, "", $hGUI)
            If @error Then ContinueLoop
            _WriteLog("Custom Pic Browse selected: " & $sSelectedPic)
            GUICtrlSetData($idInputCustomPic, $sSelectedPic)
            GUICtrlSetData($idStatus, "Status: Custom picture selected. Ready.")

        Case $idRadioBuiltin
            _ControlCustomPicState(False)

        Case $idRadioCustom
            _ControlCustomPicState(True)

        Case $idCheckLogging
            $g_bLoggingEnabled = (GUICtrlRead($idCheckLogging) = $GUI_CHECKED)
            If $g_bLoggingEnabled Then _WriteLog("--- Logging Checkbox: Enabled ---")
            If Not $g_bLoggingEnabled Then _WriteLog("--- Logging Checkbox: Disabled ---")

        Case $idProcess
            _WriteLog("Process Button Clicked. Disabling UI.")
            _SetUIState(False)
            GUICtrlSetData($idStatus, "Status: Starting process...")
            Sleep(50) ; Allow UI to update

            Local $sWimFilePathOrig = GUICtrlRead($idInputWIM)
            Local $sIndexStr = GUICtrlRead($idInputIndex)
            Local $iUpdateType = (GUICtrlRead($idRadioUpdateSetup) = $GUI_CHECKED) ? $UPDATE_TYPE_SETUP : $UPDATE_TYPE_PE
            Local $bUseBuiltinPic = (GUICtrlRead($idRadioBuiltin) = $GUI_CHECKED)
            Local $sCustomPicPath = GUICtrlRead($idInputCustomPic)
            Local $sWimFilePathCurrent = ""
            Local $bCopiedWim = False

            ; --- Input Validation ---
            If $sWimFilePathOrig = "" Or Not FileExists($sWimFilePathOrig) Then
                _HandleError("Please select a valid WIM file that exists.", "Input Error")
                ContinueLoop
            EndIf
            _WriteLog("Original WIM File: " & $sWimFilePathOrig)

            Local $iIndex = Number($sIndexStr)
            If Not StringIsDigit($sIndexStr) Or $iIndex <= 0 Then
                _WriteLog("Invalid index '" & $sIndexStr & "' provided. Using default: 2 (will retry with 1 if needed).")
                $iIndex = 2
                GUICtrlSetData($idInputIndex, "2")
            EndIf
            _WriteLog("Selected Image Index: " & $iIndex)

            _WriteLog("Update Type: " & ($iUpdateType = $UPDATE_TYPE_SETUP ? "Setup" : "PE"))
            _WriteLog("Using Built-in Picture: " & $bUseBuiltinPic)

            Local $sPicSourcePath = ""
            If $bUseBuiltinPic Then
                $sPicSourcePath = $g_sToolsFolder & "\" & ($iUpdateType = $UPDATE_TYPE_SETUP ? "1.bmp" : "img0.jpg")
                _WriteLog("Using built-in picture: " & $sPicSourcePath)
                If Not FileExists($sPicSourcePath) Then
                     _HandleError("Built-in picture not found: " & $sPicSourcePath, "File Error")
                     ContinueLoop
                EndIf
            Else
                $sPicSourcePath = $sCustomPicPath
                _WriteLog("Using custom picture: " & $sPicSourcePath)
                If $sPicSourcePath = "" Or Not FileExists($sPicSourcePath) Then
                    _HandleError("Custom picture not selected or file does not exist: " & $sPicSourcePath, "Input Error")
                    ContinueLoop
                EndIf
                Local $sExt = StringLower(_PathGetExtension($sPicSourcePath))
                If $iUpdateType = $UPDATE_TYPE_SETUP And $sExt <> ".bmp" Then
                    _HandleError("Setup background requires a .bmp file. Selected: " & $sPicSourcePath, "Input Error")
                    ContinueLoop
                EndIf
                If $iUpdateType = $UPDATE_TYPE_PE And $sExt <> ".jpg" And $sExt <> ".jpeg" Then
                    _HandleError("PE background requires a .jpg or .jpeg file. Selected: " & $sPicSourcePath, "Input Error")
                    ContinueLoop
                EndIf
            EndIf

            ; --- Enhanced Read-Only Handling ---
            Local $bCopy = False
            Local $sWimFileAttrib = FileGetAttrib($sWimFilePathOrig)
            If StringInStr($sWimFileAttrib, "R") Then
                _WriteLog("WIM file has Read-Only attribute.")
                ; Attempt to remove read-only attribute from the original file.
                FileSetAttrib($sWimFilePathOrig, "-R")
                If @error Then
                    _WriteLog("Failed to remove read-only attribute from original file.")
                    $bCopy = True
                Else
                    _WriteLog("Successfully removed read-only attribute from original file.")
                    ; Check if the file's directory is writable.
                    Local $sWimFileDir = _PathGetDirectory($sWimFilePathOrig)
                    Local $hTest = FileOpen($sWimFileDir & "\_testwrite.tmp", $FO_OVERWRITE)
                    If $hTest = -1 Then
                        _WriteLog("Original file directory is not writable. Will copy file locally.")
                        $bCopy = True
                    Else
                        FileClose($hTest)
                        FileDelete($sWimFileDir & "\_testwrite.tmp")
                    EndIf
                EndIf
            EndIf

            If $bCopy Then
                Local $sWimFileName = _PathGetFileName($sWimFilePathOrig)
                Local $sNewWimPath = $g_sScriptDir & "\" & $sWimFileName
                _WriteLog("Attempting to copy WIM file to: " & $sNewWimPath)
                If FileExists($sNewWimPath) Then _WriteLog("Local copy exists. Overwriting.")
                If Not FileCopy($sWimFilePathOrig, $sNewWimPath, $FC_OVERWRITE) Then
                    _HandleError("Failed to copy WIM file." & @CRLF & "Source: " & $sWimFilePathOrig & @CRLF & "Dest: " & $sNewWimPath & @CRLF & "Check disk space/permissions.", "File Copy Error")
                    ContinueLoop
                EndIf
                _WriteLog("WIM file copied successfully to script directory. Removing read-only attribute from copy.")
                FileSetAttrib($sNewWimPath, "-R")
                If @error Then _WriteLog("WARN: Could not remove read-only attribute from the copied file. Processing might still fail.")
                $sWimFilePathCurrent = $sNewWimPath
                $bCopiedWim = True
                GUICtrlSetData($idInputWIM, $sWimFilePathCurrent)
                GUICtrlSetData($idStatus, "Status: WIM copied locally. Proceeding...")
                Sleep(100)
            Else
                _WriteLog("Using original WIM file path for processing.")
                $sWimFilePathCurrent = $sWimFilePathOrig
            EndIf

            _WriteLog("Processing WIM File: " & $sWimFilePathCurrent)
            GUICtrlSetData($idStatus, "Status: Processing WIM, please wait... This may take time.")
            GUISetState(@SW_DISABLE, $hGUI)

            Local $bResult = _ProcessWim($sWimFilePathCurrent, $iIndex, $iUpdateType, $sPicSourcePath)
            _WriteLog("Processing function returned: " & $bResult & ". Re-enabling GUI window.")
            GUISetState(@SW_ENABLE, $hGUI)

            If $bResult Then
                GUICtrlSetData($idStatus, "Status: Done! WIM processing completed successfully.")
                _WriteLog("--- WIM update process completed successfully. ---")
                MsgBox($MB_ICONINFORMATION + $MB_SYSTEMMODAL, "Success", "WIM file updated successfully!")
            Else
                GUICtrlSetData($idStatus, "Status: Error during processing! Check log file for details: " & @CRLF & $g_sLogFile)
                _WriteLog("--- WIM update process failed. ---")
            EndIf

            _WriteLog("Re-enabling UI controls.")
            Sleep(100)
            _SetUIState(True)
    EndSwitch
WEnd

; =============================================================================
; Function: _ProcessWim
; =============================================================================
Func _ProcessWim($sWimFile, $iIndex, $iUpdateType, $sSourcePicPath)
    Local $bSuccess = False
    _WriteLog("Starting _ProcessWim function.")
    GUICtrlSetData($idStatus, "Status: Preparing temporary files...")
    _WriteLog("Using Temp Folder: " & $g_sTempFolder)
    If FileExists($g_sTempFolder) Then
        _WriteLog("Removing existing temp folder: " & $g_sTempFolder)
        DirRemove($g_sTempFolder, $DIR_REMOVE)
        If @error Then Return _HandleProcError("Failed to remove existing temp directory: " & $g_sTempFolder & @CRLF & "Error: " & @error & ", Extended: " & @extended & @CRLF & "Close programs using it.", "Cleanup Error")
    EndIf
    _WriteLog("Creating temp folder: " & $g_sTempFolder)
    If Not DirCreate($g_sTempFolder) Then Return _HandleProcError("Failed to create temporary directory: " & $g_sTempFolder & @CRLF & "Error: " & @error & ", Extended: " & @extended, "Directory Error")

    Local $sTargetDllPath = $g_sToolsFolder & "\spwizimg.dll"
    GUICtrlSetData($idStatus, "Status: Copying files to temporary location...")
    If $iUpdateType = $UPDATE_TYPE_SETUP Then
        _WriteLog("Preparing files for Setup update.")
        Local $sTempWinSys32 = $g_sTempFolder & "\Windows\System32"
        Local $sTempSources = $g_sTempFolder & "\sources"
        If Not DirCreate($sTempWinSys32) Then Return _HandleProcError("Failed to create temp dir: " & $sTempWinSys32, "Directory Error")
        If Not DirCreate($sTempSources) Then Return _HandleProcError("Failed to create temp dir: " & $sTempSources, "Directory Error")
        If Not FileCopy($sSourcePicPath, $sTempSources & "\background.bmp", $FC_OVERWRITE) Then Return _HandleProcError("Failed to copy BMP to " & $sTempSources & "\background.bmp", "File Copy Error")
        If Not FileCopy($sSourcePicPath, $sTempWinSys32 & "\Setup.bmp", $FC_OVERWRITE) Then Return _HandleProcError("Failed to copy BMP to " & $sTempWinSys32 & "\Setup.bmp", "File Copy Error")
        If Not FileExists($sTargetDllPath) Then Return _HandleProcError("Required DLL not found: " & $sTargetDllPath, "File Error")
        If Not FileCopy($sTargetDllPath, $sTempSources & "\spwizimg.dll", $FC_OVERWRITE) Then Return _HandleProcError("Failed to copy DLL to " & $sTempSources & "\spwizimg.dll", "File Copy Error")
        _WriteLog("Files copied for Setup update.")
        Local $sRelSources = $g_sTempSubDir & '\sources'
        Local $sRelWindows = $g_sTempSubDir & '\Windows'
        Local $sArgsTemplateSources = 'update "%s" %%d --command="add %s \\sources"'
        Local $sArgsTemplateWindows = 'update "%s" %%d --command="add %s \\Windows"'
        GUICtrlSetData($idStatus, "Status: Updating WIM (Sources)... Please wait.")
        If Not _RunWimlibCommand($g_sWimlibExe, StringFormat($sArgsTemplateSources, $sWimFile, $sRelSources), $iIndex, "Sources") Then Return False
        GUICtrlSetData($idStatus, "Status: Updating WIM (Windows)... Please wait.")
        If Not _RunWimlibCommand($g_sWimlibExe, StringFormat($sArgsTemplateWindows, $sWimFile, $sRelWindows), $iIndex, "Windows") Then Return False
        $bSuccess = True
    ElseIf $iUpdateType = $UPDATE_TYPE_PE Then
        _WriteLog("Preparing files for PE update.")
        Local $sTempPEPath = $g_sTempFolder & "\Wallpaper\Windows"
        If Not DirCreate($sTempPEPath) Then Return _HandleProcError("Failed to create temp dir: " & $sTempPEPath, "Directory Error")
        If Not FileCopy($sSourcePicPath, $sTempPEPath & "\img0.jpg", $FC_OVERWRITE) Then Return _HandleProcError("Failed to copy JPG to " & $sTempPEPath & "\img0.jpg", "File Copy Error")
        _WriteLog("Files copied for PE update.")
        Local $sRelPEPath = $g_sTempSubDir & '\Wallpaper\Windows'
        Local $sArgsTemplatePE = 'update "%s" %%d --command="add %s \\Windows\\Web\\Wallpaper\\Windows"'
        GUICtrlSetData($idStatus, "Status: Updating WIM (PE Background)... Please wait.")
        If Not _RunWimlibCommand($g_sWimlibExe, StringFormat($sArgsTemplatePE, $sWimFile, $sRelPEPath), $iIndex, "PE Background") Then Return False
        $bSuccess = True
    Else
        Return _HandleProcError("Invalid Update Type specified: " & $iUpdateType, "Internal Script Error")
    EndIf

    _CleanupTempDir()
    Return $bSuccess
EndFunc

; =============================================================================
; Helper Function: _RunWimlibCommand (Handles Index Retry)
; =============================================================================
Func _RunWimlibCommand($sFullExePath, $sArgumentsTemplateWithPlaceholder, $iIndexAttempt, $sStepName)
    Local $iCurrentIndex = $iIndexAttempt
    Local $sArguments = StringFormat($sArgumentsTemplateWithPlaceholder, $iCurrentIndex)
    _WriteLog("Preparing to execute wimlib command for step [" & $sStepName & "] with Index " & $iCurrentIndex)
    _WriteLog("Exe: " & $sFullExePath)
    _WriteLog("Args: " & $sArguments)
    Local $iPID = Run('"' & $sFullExePath & '" ' & $sArguments, $g_sScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
    If @error Or $iPID = 0 Then Return _HandleProcError("Failed to launch wimlib-imagex (Index " & $iCurrentIndex & ") for step [" & $sStepName & "]. Error: " & @error, "Process Launch Error")
    _WriteLog("wimlib-imagex process launched (Index " & $iCurrentIndex & "). PID: " & $iPID)
    Local $sStdOut = ""
    Local $sStdErr = ""
    Local $sLineOut = ""
    Local $sLineErr = ""
    Local $iStartTime = TimerInit()
    Local $iExitCode = -1
    While ProcessExists($iPID)
        GUIGetMsg()
        While True
            $sLineErr = StderrRead($iPID, False, False)
            If @error Or $sLineErr = "" Then ExitLoop
            $sStdErr &= $sLineErr
            _WriteLog("WIMLIB_STDERR [" & $sStepName & "]: " & StringReplace($sLineErr, @CR, ""))
        WEnd
        While True
            $sLineOut = StdoutRead($iPID, False, False)
            If @error Or $sLineOut = "" Then ExitLoop
            $sStdOut &= $sLineOut
            _WriteLog("WIMLIB_STDOUT [" & $sStepName & "]: " & StringReplace($sLineOut, @CR, ""))
        WEnd
        If TimerDiff($iStartTime) > 500 Then
            Local $iElapsed = Round(TimerDiff($iStartTime) / 1000)
            GUICtrlSetData($idStatus, "Status: Running wimlib [" & $sStepName & "] (Index " & $iCurrentIndex & ")... (" & $iElapsed & "s)")
            $iStartTime = TimerInit()
        EndIf
        Sleep(100)
    WEnd
    $iExitCode = _GetExitCode($iPID)
    _WriteLog("wimlib-imagex process finished (Index " & $iCurrentIndex & "). PID: " & $iPID & ". Exit Code: " & $iExitCode)
    If $iExitCode = $WIMLIB_EXIT_IMAGE_NOT_FOUND And $iCurrentIndex <> 1 Then
        _WriteLog("WARN: wimlib failed with Exit Code 18 (Image Not Found) for Index " & $iCurrentIndex & ". Retrying with Index 1...")
        GUICtrlSetData($idStatus, "Status: Index " & $iCurrentIndex & " not found. Retrying with Index 1...")
        Sleep(100)
        $iCurrentIndex = 1
        $sArguments = StringFormat($sArgumentsTemplateWithPlaceholder, $iCurrentIndex)
        _WriteLog("Preparing retry execution for step [" & $sStepName & "] with Index " & $iCurrentIndex)
        _WriteLog("Args: " & $sArguments)
        $sStdOut = ""
        $sStdErr = ""
        $iPID = Run('"' & $sFullExePath & '" ' & $sArguments, $g_sScriptDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
        If @error Or $iPID = 0 Then Return _HandleProcError("Failed to launch wimlib-imagex (Retry Index 1) for step [" & $sStepName & "]. Error: " & @error, "Process Launch Error")
        _WriteLog("wimlib-imagex process launched (Retry Index 1). PID: " & $iPID)
        $iStartTime = TimerInit()
        While ProcessExists($iPID)
            GUIGetMsg()
            While True
                $sLineErr = StderrRead($iPID, False, False)
                If @error Or $sLineErr = "" Then ExitLoop
                $sStdErr &= $sLineErr
                _WriteLog("WIMLIB_STDERR [" & $sStepName & "/Retry]: " & StringReplace($sLineErr, @CR, ""))
            WEnd
            While True
                $sLineOut = StdoutRead($iPID, False, False)
                If @error Or $sLineOut = "" Then ExitLoop
                $sStdOut &= $sLineOut
                _WriteLog("WIMLIB_STDOUT [" & $sStepName & "/Retry]: " & StringReplace($sLineOut, @CR, ""))
            WEnd
            If TimerDiff($iStartTime) > 500 Then
                Local $iElapsed = Round(TimerDiff($iStartTime) / 1000)
                GUICtrlSetData($idStatus, "Status: Running wimlib [" & $sStepName & "] (Index " & $iCurrentIndex & ")... (" & $iElapsed & "s)")
                $iStartTime = TimerInit()
            EndIf
            Sleep(100)
        WEnd
        $iExitCode = _GetExitCode($iPID)
        _WriteLog("wimlib-imagex process finished (Retry Index 1). PID: " & $iPID & ". Exit Code: " & $iExitCode)
        If $iExitCode = 0 Then GUICtrlSetData($idInputIndex, "1")
    EndIf
    If $iExitCode <> 0 Then
        Local $sErrMsg = "wimlib-imagex failed during step [" & $sStepName & "]!" & @CRLF
        If $iCurrentIndex = 1 And $iIndexAttempt <> 1 Then $sErrMsg &= "(Tried Index 1 after initial failure)" & @CRLF
        $sErrMsg &= "Final Exit Code: " & $iExitCode & @CRLF & @CRLF
        $sErrMsg &= "Check log file (" & $g_sLogFile & ") for details." & @CRLF
        If StringStripWS($sStdErr, 8) <> "" Then
            $sErrMsg &= @CRLF & "wimlib Error Output Snippet (stderr):" & @CRLF & _
                         StringLeft(StringStripWS($sStdErr, $STR_STRIPLEADING), 500) & _
                         (StringLen($sStdErr) > 500 ? "..." : "")
        ElseIf StringStripWS($sStdOut, 8) <> "" Then
            $sErrMsg &= @CRLF & "wimlib Output Snippet (stdout):" & @CRLF & _
                         StringLeft(StringStripWS($sStdOut, $STR_STRIPLEADING), 500) & _
                         (StringLen($sStdOut) > 500 ? "..." : "")
        Else
            $sErrMsg &= "wimlib produced no output on stdout or stderr, despite failing with Exit Code " & $iExitCode & "."
        EndIf
        Return _HandleProcError($sErrMsg, "wimlib Execution Error")
    EndIf
    Return True ; Success
EndFunc

; =============================================================================
; Helper Function: _HandleError (General Errors)
; =============================================================================
Func _HandleError($sMessage, $sTitle = "Error")
    _WriteLog("ERROR: " & $sTitle & " - " & StringReplace($sMessage, @CRLF, " | "))
    MsgBox($MB_ICONERROR + $MB_SYSTEMMODAL, $sTitle, $sMessage)
    _WriteLog("Re-enabling UI after general error.")
    GUISetState(@SW_ENABLE, $hGUI)
    _SetUIState(True)
    GUICtrlSetData($idStatus, "Status: Error encountered. Ready.")
EndFunc

; =============================================================================
; Helper Function: _HandleProcError (Errors within Processing Steps)
; =============================================================================
Func _HandleProcError($sMessage, $sTitle = "Processing Error")
    _WriteLog("ERROR: " & $sTitle & " - " & StringReplace($sMessage, @CRLF, " | "))
    MsgBox($MB_ICONERROR + $MB_SYSTEMMODAL, $sTitle, $sMessage)
    _CleanupTempDir()
    Return False
EndFunc

; =============================================================================
; Helper Function: _CleanupTempDir
; =============================================================================
Func _CleanupTempDir()
    If FileExists($g_sTempFolder) Then
        _WriteLog("Cleaning up temporary directory: " & $g_sTempFolder)
        DirRemove($g_sTempFolder, $DIR_REMOVE)
        If @error Then
            _WriteLog("WARNING: Failed to remove temp dir: " & $g_sTempFolder & " Error: " & @error & ", Extended: " & @extended)
            MsgBox($MB_ICONWARNING, "Cleanup Warning", "Could not remove temp dir:" & @CRLF & $g_sTempFolder & @CRLF & "Remove manually.", 10)
        Else
            _WriteLog("Temporary directory removed successfully.")
        EndIf
    EndIf
EndFunc

; =============================================================================
; Helper Function: _SetUIState
; =============================================================================
Func _SetUIState($bEnable)
    _WriteLog("Setting UI State: " & ($bEnable ? "Enabled" : "Disabled"))
    Local $iState = ($bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idInputWIM, $iState)
    GUICtrlSetState($idBrowseWIM, $iState)
    GUICtrlSetState($idRadioBuiltin, $iState)
    GUICtrlSetState($idRadioCustom, $iState)
    GUICtrlSetState($idInputIndex, $iState)
    GUICtrlSetState($idRadioUpdateSetup, $iState)
    GUICtrlSetState($idRadioUpdatePE, $iState)
    GUICtrlSetState($idCheckLogging, $iState)
    GUICtrlSetState($idProcess, $iState)
    If $bEnable Then
        _ControlCustomPicState(GUICtrlRead($idRadioCustom) = $GUI_CHECKED)
    Else
        _ControlCustomPicState(False)
    EndIf
    _WriteLog("UI State Set.")
EndFunc

; =============================================================================
; Helper Function: _ControlCustomPicState
; =============================================================================
Func _ControlCustomPicState($bEnable)
    Local $iState = ($bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idLabelCustomPic, $iState)
    GUICtrlSetState($idInputCustomPic, $iState)
    GUICtrlSetState($idBrowseCustomPic, $iState)
EndFunc

; =============================================================================
; Function: _WriteLog
; =============================================================================
Func _WriteLog($sText)
    If $g_bLoggingEnabled Then
        Local $hFile = FileOpen($g_sLogFile, $FO_APPEND)
        If $hFile <> -1 Then
            FileWriteLine($hFile, _NowTime() & " | " & $sText)
            FileClose($hFile)
        Else
            Static $bLoggingErrorShown = False
            If Not $bLoggingErrorShown Then
                $bLoggingErrorShown = True
                MsgBox($MB_ICONWARNING + $MB_SYSTEMMODAL, "Logging Error", "Could not open log file for writing:" & @CRLF & $g_sLogFile)
                $bLoggingErrorShown = False
            EndIf
        EndIf
    EndIf
EndFunc

; =============================================================================
; Path Helper Functions
; =============================================================================
Func _PathGetExtension($sPath)
    Local $iPos = StringInStr($sPath, ".", 0, -1)
    If $iPos = 0 Then Return ""
    Return StringMid($sPath, $iPos)
EndFunc

Func _PathGetDirectory($sPath)
    Local $iLastSlash = StringInStr($sPath, "\", 0, -1)
    If $iLastSlash = 0 Then Return ""
    Return StringLeft($sPath, $iLastSlash)
EndFunc

Func _PathGetFileName($sPath)
    Return StringTrimLeft($sPath, StringInStr($sPath, "\", 0, -1))
EndFunc

; =============================================================================
; Function: _ExitWithError (For critical startup errors)
; =============================================================================
Func _ExitWithError($sMessage)
    _WriteLog("CRITICAL ERROR: " & StringReplace($sMessage, @CRLF, " | "))
    MsgBox($MB_ICONERROR + $MB_SYSTEMMODAL, "Critical Error", $sMessage)
    Exit(1)
EndFunc
