; #INDEX# =======================================================================================================================
; Title .........: OCC
; AutoIt Version : 3.3.14.5
; Language ......: French
; Description ...: Script .au3
; Author(s) .....: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================

; #ENVIRONMENT# =================================================================================================================
; AutoIt3Wrapper
#AutoIt3Wrapper_Res_ProductName=Outils Caisse Compagnon
#AutoIt3Wrapper_Res_Description=Permet de copier le NIR dans le presse papier à l'insertion d'un Carte Vitale
#AutoIt3Wrapper_Res_ProductVersion=1.0.2
#AutoIt3Wrapper_Res_FileVersion=1.0.2
#AutoIt3Wrapper_Res_CompanyName=CNAMTS/CPAM_ARTOIS/APPLINAT
#AutoIt3Wrapper_Res_LegalCopyright=yann.daniel@assurance-maladie.fr
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Compatibility=Win7
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Icon="static\icon.ico"
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Run_AU3Check=Y
#AutoIt3Wrapper_Run_Au3Stripper=N
#Au3Stripper_Parameters=/MO /RSLN
#AutoIt3Wrapper_AU3Check_Parameters=-q -d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w 7
#AutoIt3Wrapper_AU3Check_Stop_OnWarning=Y
; Includes Constants
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
#include <GUIConstants.au3>
; Includes
#include <Misc.au3>
#include <Array.au3>
#include <File.au3>
#include <Date.au3>
#include <String.au3>
#include <YDLogger.au3>
#include <YDTool.au3>
; Options
;#pragma compile(Icon, D:\Users\DANIEL-03598\Documents\Apps_portable\AutoIt3\_dev\nytrio_restants\ssh.ico)
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDLogger_Init()
Global $g_sAppName                = _YDTool_GetAppWrapperRes("ProductName")
Global $g_sAppDesc                = _YDTool_GetAppWrapperRes("Description")
Global $g_sAppVersion             = _YDTool_GetAppWrapperRes("FileVersion")
Global $g_sAppContact             = _YDTool_GetAppWrapperRes("LegalCopyright")
Global $g_sAppVersionV            = "v" & $g_sAppVersion
Global $g_sAppTitle               = $g_sAppName & " - " & $g_sAppVersionV
Global $g_sAppDataPath            = @ScriptDir & "\data"
Global $g_sAppStaticPath          = @ScriptDir & "\static"
Global $g_sAppLogsPath            = @ScriptDir & "\logs"
Global $g_sAppVendorPath          = @ScriptDir & "\vendor"
Global $g_sAppIconPath            = @ScriptDir & "\static\icon.ico"

Global $g_sOCTitle 				  = "Outils Caisse"
Global $g_sMsg                    = ""
Global $g_sOCExeName  			  = "OutilsCaisse.exe"
Global $g_sOCExeFile  			  = "C:\APPLINAT\Outils Caisse\" & $g_sOCExeName
Global $g_sLogFile                = "C:\ProgramData\santesocial\Outils Caisse\log\oc.log"
Global $g_sXmlFile 	              = "C:\ProgramData\santesocial\Outils Caisse\exports\Exp_Identification.xml"


; ===============================================================================================================================

; #VARIABLES DEBUG# =============================================================================================================
_YDTool_DebugGlobals()
_YDLogger_Var("$g_sAppName", $g_sAppName)
_YDLogger_Var("$g_sAppDesc", $g_sAppDesc)
_YDLogger_Var("$g_sAppVersion", $g_sAppVersion)
_YDLogger_Var("$g_sAppVersionV", $g_sAppVersionV)
_YDLogger_Var("$g_sAppTitle", $g_sAppTitle)
_YDLogger_Var("$g_sAppDataPath", $g_sAppDataPath)
_YDLogger_Var("$g_sAppStaticPath", $g_sAppStaticPath)
_YDLogger_Var("$g_sAppLogsPath", $g_sAppLogsPath)
_YDLogger_Var("$g_sAppVendorPath", $g_sAppVendorPath)
_YDLogger_Var("$g_sAppIconPath", $g_sAppIconPath)
If _Singleton($g_sAppName, 1) = 0 Then
    $g_sMsg = "L'application " & $g_sAppName & " est déjà en cours d'exécution !"
    _YDLogger_Log($g_sMsg)
    MsgBox($MB_SYSTEMMODAL, "Warning", $g_sMsg)
    Exit
EndIf
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
; On gere l'affichage de l'icone dans le tray
TraySetIcon($g_sAppIconPath)
TraySetToolTip($g_sAppTitle)
Global $idTrayOCLaunch = TrayCreateItem("Lancer Outils Caisse")
TrayCreateItem("")
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
Global $hWndActive, $iPIDActive
Global $iFileLastLine    = 1
Global $aNir 			 = ""
; #MAIN SCRIPT# =================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
	Global $iMsg = TrayGetMsg()
	Select
		Case $iMsg = $idTrayExit
			_ExitConfirm()
		Case $iMsg = $idTrayOCLaunch
			_OCLaunch()
		Case $iMsg = $idTrayAbout
			_About()
		Case Else
			_OCMain()
	EndSelect
	;------------------------------
    Sleep(10)
WEnd
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Description ...: Traitement principal si OC lance
; Syntax ........: _OCMain()
; Parameters ....:
; Return values .:
; Author ........: Yann DANIEL
; Last Modified .: 16/01/2019
; Notes .........:
;================================================================================================================================
Func _OCMain()
	Local $sFuncName = "_OCMain"
	$hWndActive = WinGetHandle("[ACTIVE]")
	$iPIDActive = WinGetProcess($hWndActive)
	If WinExists($g_sOCTitle) Then
		Local $aOCProcess = ProcessList($g_sOCExeName)
		If $aOCProcess[0][0] > 1 Then
			_YDLogger_Log("Fermeture du 2nd process : " & $aOCProcess[2][1], $sFuncName)
			ProcessClose($aOCProcess[2][1])
			_YDTool_SetMsgBoxError("Une instance d'Outils Caisse est déjà lancée !")
		EndIf
		Local $bCardInserted = _OCCheckCardInsertion()
		If $bCardInserted Then
			_SetNirInClipboard($aNir[1])
		Endif
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Lance Outils Caisse si pas deja lance
; Syntax ........: _OCCheckCardInsertion()
; Parameters ....:
; Return values .:
; Author ........: Yann DANIEL
; Last Modified .: 15/01/2019
; Notes .........:
;================================================================================================================================
Func _OCLaunch()
	Local $sFuncName = "_OCLaunch"
	_YDLogger_Log("", $sFuncName)
	If ProcessExists($g_sOCExeName) Then
		_YDTool_SetMsgBoxError("Outils Caisse est déjà en cours d'execution !")
	Else
		If FileExists($g_sOCExeFile) Then
			_YDLogger_Log("Lancement OC", $sFuncName)
			Run($g_sOCExeFile)
		EndIf
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Lance Outils Caisse si pas deja lance
; Syntax ........: _ExitConfirm()
; Parameters ....:
; Return values .:
; Author ........: Yann DANIEL
; Last Modified .: 15/01/2019
; Notes .........:
;================================================================================================================================
Func _ExitConfirm()
	Local $sFuncName = "_ExitConfirm"
	_YDLogger_Log("", $sFuncName)
 	Local $iResponse = MsgBox(4, $g_sAppName,"Etes-vous sûr de vouloir quitter " & $g_sAppName & " ?",30)
	If $iResponse = 6 Then
		_YDLogger_Log("Fermeture confirmee", $sFuncName)
		Exit
	Else
		_YDLogger_Log("Fermeture annulee", $sFuncName)
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Lance des actions lors de l'insertion de la carte vitale dans Outils Caisse
; Syntax ........: _OCCheckCardInsertion()
; Parameters ....:
; Return values .: True - Si la carte a été insérée et le NIR copié dans le presse papier
;                  False - Si un problème est survenu
; Author ........: Yann DANIEL
; Last Modified .: 15/01/2019
; Notes .........:
;================================================================================================================================
Func _OCCheckCardInsertion()
	Local $sFuncName = "_OCCheckCardInsertion"
	;Dim $g_sOCTitle, $iPIDActive, $hWndActive, $aNir
	Local $hWndOutilCaisse
	Local $iPIDOutilCaisse
	Local $sFileLine
	Local $hLogFile
	Local $sVitalePattern = "-->carte Vitale"
	Local $bCardInserted = False

	; On ne fait des recherches que si OC est lance
	If WinExists($g_sOCTitle) Then
		;------------------------------
		; On récupère le handle et le PID de OutilCaisse.exe
		$hWndOutilCaisse = WinGetHandle($g_sOCTitle)
		$iPIDOutilCaisse = WinGetProcess($hWndOutilCaisse)
		; On ne fait des recherches que lorsque le PID est OC et que la fenetre (chargement) change
		If $iPIDActive = $iPIDOutilCaisse And $hWndActive <> $hWndOutilCaisse Then
			;------------------------------
			; On boucle sur le fichier log pour detecter l insertion d une carte vitale
			WinWaitActive($hWndOutilCaisse)
			$hLogFile = FileOpen($g_sLogFile, 0)
			For $i = _FileCountLines($g_sLogFile) to $iFileLastLine Step -1
				$sFileLine = FileReadLine($hLogFile, $i)
				Local $iCardInserted = StringInStr($sFileLine, $sVitalePattern, 0, 1, 1)
				; Si pattern trouve on sort de la boucle
				If $iCardInserted > 0 Then
					;------------------------------
					_YDLogger_Log("Pattern trouve", $sFuncName)
					$bCardInserted = True
					$iFileLastLine = _FileCountLines($g_sLogFile) + 1
					ExitLoop
				EndIf
			Next
			FileClose($hLogFile)
			_YDLogger_Var("$iFileLastLine", $iFileLastLine, $sFuncName)
			_YDLogger_Var("$bCardInserted", $bCardInserted, $sFuncName)
			;_YDLogger_Var("$hWndOutilCaisse", $hWndOutilCaisse, $sFuncName)
			;_YDLogger_Var("$iPIDOutilCaisse", $iPIDOutilCaisse, $sFuncName)
			;Exit
			;------------------------------
			; On traite si la carte vient d etre inseree !
			If $bCardInserted = True Then
				;------------------------------
				_YDLogger_Log("Insertion carte vitale", $sFuncName)
				WinActivate($g_sOCTitle)
				While WinActive($g_sOCTitle) = 0
					Sleep(100)
				Wend
				_YDLogger_Log("Fenêtre OC activée", $sFuncName)
				_YDTool_SetTrayTip($g_sAppTitle, "Lecture de la carte vitale ...", 0, $TIP_ICONASTERISK)
				;------------------------------
				_YDLogger_Log("Sends pour export du XML", $sFuncName)
				Send("!i")
				Sleep(200)
				While WinGetHandle("[ACTIVE]") = $hWndOutilCaisse
					Sleep(100)
				Wend
				Send("{ENTER}")
				Sleep(200)
				;------------------------------
				_YDLogger_Log("Lecture du fichier : " & $g_sXmlFile, $sFuncName)
				_YDLogger_Log("FileGetTime : " & FileGetTime($g_sXmlFile, $FT_MODIFIED, $FT_STRING), $sFuncName)
				If Not FileExists($g_sXmlFile) Then
					_YDLogger_Error("Le fichier XML n'existe pas : " & $g_sXmlFile, $sFuncName)
					_YDTool_SetTrayTip($g_sAppTitle, "Erreur récupération infos Carte Vitale", 10000, $TIP_ICONHAND)
					Return False
				EndIf
				Local $hXmlFile = FileRead($g_sXmlFile)
				;------------------------------
				_YDLogger_Log("Récupération du NIR dans le fichier XML", $sFuncName)
				Local $aNirXml = _StringBetween($hXmlFile,"<nir>","</nir>")
				If Not IsArray($aNirXml) Then
					_YDLogger_Error("Erreur récupération valeur XML ($aNirXml = " & $aNirXml & ")", $sFuncName)
					_YDTool_SetTrayTip($g_sAppTitle, "Erreur récupération infos Carte Vitale", 10000, $TIP_ICONHAND)
					FileClose($g_sXmlFile)
					FileDelete($g_sXmlFile)
					Return False
				Endif
				FileClose($g_sXmlFile)
				FileDelete($g_sXmlFile)
				;------------------------------
				_YDLogger_Log("Split de $aNirXml", $sFuncName)
				$aNir = StringSplit($aNirXml[0], " ")
				If Not IsArray($aNir) Then
					_YDLogger_Error("Erreur tableau $aNir vide", $sFuncName)
					_YDTool_SetTrayTip($g_sAppTitle, "Erreur récupération infos Carte Vitale", 10000, $TIP_ICONHAND)
					Return False
				EndIf
				_YDLogger_Var("$aNir[1]", $aNir[1], $sFuncName)
				_YDLogger_Var("$aNir[2]", $aNir[2], $sFuncName)
				;------------------------------
				_YDLogger_Log("Copie du NIR dans le clipboard", $sFuncName)
				ClipPut($aNir[1])
				If @error <> 0 Then
					_YDLogger_Error("Erreur innatendue : " & @error, $sFuncName)
					Return False
				Else
					Return True
				EndIf
			EndIf
		EndIf
	Endif
	Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Copie le NIR dans le presse papier et affiche un TrayTip
; Syntax ........: _SetNirInClipboard($sNir)
; Parameters ....: $sNir - NIR sans la clé à copier
; Return values .: Success: Returns 0
;                  Failure:	Returns -1 and sets @error to 1
; Author ........: Yann DANIEL
; Last Modified .: 15/05/2018
; Notes .........:
;================================================================================================================================
Func _SetNirInClipboard($sNir)
	Local $sFuncName = "_SetNirInClipboard"
	_YDLogger_Log("Copie du NIR dans le clipboard", $sFuncName)
	ClipPut($sNir)
	_YDTool_SetTrayTip($g_sAppTitle, $sNir &  " : NIR copié dans le presse-papier", 5000, $TIP_ICONASTERISK)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _About
; Description ...: Fonction qui renvoi une GUI "A propos"
; Syntax.........: _About()
; Parameters ....:
; Return values .: True
; Author ........: yann.daniel@assurance-maladie.fr
; ===============================================================================================================================
Func _About()
    Local $sFuncName = "_About"
    Local $iAboutWidth = 550
    Local $iAboutHeight = 200
    Local $iAboutOkButtonWidth = 30
    Local $font = "Verdana"
	_YDLogger_Log("", $sFuncName)

    Local $hAboutGUI = GUICreate("A propos", $iAboutWidth, $iAboutHeight, -1, -1, BitOR($WS_POPUP,$WS_CAPTION))
    ; Titre
    GUISetFont(12, $iAboutWidth*2, 0, $font)
    GUICtrlCreateLabel($g_sAppName, 0, 0, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    ; Description + version
    GUISetFont(9, $iAboutWidth, 0, $font)
    GUICtrlCreateLabel($g_sAppDesc, 0, 40, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
    GUICtrlCreateLabel($g_sAppVersionV, 0, 80, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
	Local $idLinkContact = GUICtrlCreateLabel($g_sAppContact, 0, 120, $iAboutWidth, -1, BitOr($SS_CENTER,$BS_CENTER))
	;GUICtrlSetFont(-1, 24, 400, 4, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x0000FF)
	GUICtrlSetCursor(-1, 0)
    ; Bouton OK
    Local $idOkButton = GUICtrlCreateButton("OK", $iAboutWidth/2-$iAboutOkButtonWidth/2, 160, $iAboutOkButtonWidth, 25, BitOr($BS_MULTILINE,$BS_CENTER))
    ; Affichage GUI
    GUISetState(@SW_SHOW, $hAboutGUI)
    ; Loop GUI
    While 1
		Local $iMsg = GUIGetMsg()
		Select
			Case $iMsg = $idOkButton
				GUIDelete($hAboutGUI)
				ExitLoop
			Case $iMsg = $idLinkContact
		        ShellExecute("mailto:"&$g_sAppContact)
        EndSelect
        Sleep(50)
    WEnd
    Return True
EndFunc