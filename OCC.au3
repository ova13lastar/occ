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
#AutoIt3Wrapper_Res_Description=Permet de copier le NIR dans le presse papier a l'insertion d'un Carte Vitale
#AutoIt3Wrapper_Res_ProductVersion=1.2.1
#AutoIt3Wrapper_Res_FileVersion=1.2.1
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
; Includes YD
#include <YDGVars.au3>
#include <YDLogger.au3>
#include <YDTool.au3>
; Includes Constants
#include <StaticConstants.au3>
#Include <WindowsConstants.au3>
#include <TrayConstants.au3>
; Includes
#include <String.au3>
; Options
AutoItSetOption("MustDeclareVars", 1)
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("WinDetectHiddenText", 1)
AutoItSetOption("MouseCoordMode", 0)
AutoItSetOption("TrayMenuMode", 3)
OnAutoItExitRegister("_YDTool_ExitApp")
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
_YDGVars_Set("sAppName", _YDTool_GetAppWrapperRes("ProductName"))
_YDGVars_Set("sAppDesc", _YDTool_GetAppWrapperRes("Description"))
_YDGVars_Set("sAppVersion", _YDTool_GetAppWrapperRes("FileVersion"))
_YDGVars_Set("sAppContact", _YDTool_GetAppWrapperRes("LegalCopyright"))
_YDGVars_Set("sAppVersionV", "v" & _YDGVars_Get("sAppVersion"))
_YDGVars_Set("sAppTitle", _YDGVars_Get("sAppName") & " - " & _YDGVars_Get("sAppVersionV"))
_YDGVars_Set("sAppDirDataPath", @ScriptDir & "\data")
_YDGVars_Set("sAppDirStaticPath", @ScriptDir & "\static")
_YDGVars_Set("sAppDirLogsPath", @ScriptDir & "\logs")
_YDGVars_Set("sAppDirVendorPath", @ScriptDir & "\vendor")
_YDGVars_Set("sAppIconPath", @ScriptDir & "\static\icon.ico")
_YDGVars_Set("sAppConfFile", @ScriptDir & "\conf.ini")
_YDGVars_Set("iAppNbDaysToKeepLogFiles", 15)

_YDGVars_Set("sAppOCTitle", "Outils Caisse")
_YDGVars_Set("sAppOCExeName", "OutilsCaisse.exe")
_YDGVars_Set("sAppOCExeFile", "C:\APPLINAT\Outils Caisse\" & _YDGVars_Get("sAppOCExeName"))
_YDGVars_Set("sAppOCLogFile", "C:\ProgramData\santesocial\Outils Caisse\log\oc.log")
_YDGVars_Set("sAppOCXmlFile", "C:\ProgramData\santesocial\Outils Caisse\exports\Exp_Identification.xml")

_YDLogger_Init()
_YDLogger_LogAllGVars()
; ===============================================================================================================================

; #MAIN SCRIPT# =================================================================================================================
If Not _YDTool_IsSingleton() Then Exit
;------------------------------
; On supprime les anciens fichiers de log
_YDTool_DeleteOldFiles(_YDGVars_Get("sAppDirLogsPath"), _YDGVars_Get("iAppNbDaysToKeepLogFiles"))
;------------------------------
; On gere l'affichage de l'icone dans le tray
TraySetIcon(_YDGVars_Get("sAppIconPath"))
TraySetToolTip(_YDGVars_Get("sAppTitle"))
Global $idTrayOCLaunch = TrayCreateItem("Lancer Outils Caisse")
TrayCreateItem("")
Global $idTrayAbout = TrayCreateItem("A propos", -1, -1, -1)
Global $idTrayExit = TrayCreateItem("Quitter", -1, -1, -1)
TraySetState($TRAY_ICONSTATE_SHOW)
;------------------------------
Global $hWndActive, $iPIDActive, $hWndOutilCaisse, $iPIDOutilCaisse
Global $iFileLastLine    = 1
Global $aNir 			 = ""
; #MAIN SCRIPT# =================================================================================================================

; #MAIN LOOP# ====================================================================================================================
While 1
	Global $iMsg = TrayGetMsg()
	Select
		Case $iMsg = $idTrayExit
			_YDTool_ExitConfirm()
		Case $iMsg = $idTrayOCLaunch
			_OCLaunch()
		Case $iMsg = $idTrayAbout
			_YDTool_GUIShowAbout()
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
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 18/02/2019
; Notes .........:
;================================================================================================================================
Func _OCMain()
	Local $sFuncName = "_OCMain"
	; On recupere le Handle de la fenetre active et son PID
	$hWndActive = WinGetHandle("[ACTIVE]")
	$iPIDActive = WinGetProcess($hWndActive)
	; On ne travaille que si OC est lance
	If WinExists(_YDGVars_Get("sAppOCTitle")) Then
		; On liste les Process OC
		Local $aOCProcess = ProcessList(_YDGVars_Get("sAppOCExeName"))
		; Si un autre trouve, on le ferme
		If $aOCProcess[0][0] > 1 Then
			_YDLogger_Log("Fermeture du 2nd process : " & $aOCProcess[2][1], $sFuncName)
			ProcessClose($aOCProcess[2][1])
			_YDTool_SetMsgBoxError("Une instance d'Outils Caisse est déjà lancée !")
		EndIf
		; On verifie si une CV a ete inseree
		Local $bCardInserted = _OCCheckCardInsertion()
		; Si CV inseree on agit !!
		If $bCardInserted Then
			_YDLogger_Log("Carte vitale inseree !", $sFuncName)
			_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Lecture de la carte vitale ...", 0, $TIP_ICONASTERISK)
			; On tente de recuperer le NIR
			Local $bNirFound = _OCGetNir()
			; Si NIR trouve on le met dans le presse papier, sinon erreur
			If ($bNirFound And $aNir[1] <> "") Then
				_YDLogger_Log("NIR trouve !", $sFuncName)
				_SetNirInClipboard($aNir[1])
			Else
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Erreur récupération infos Carte Vitale", 10000, $TIP_ICONHAND)
			EndIf
		Endif
	EndIf
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de vérifier si une CV a été insérée
; Syntax ........: _OCCheckCardInsertion()
; Parameters ....:
; Return values .: True 	- La CV a été insérée
;                  False 	- Pas de CV insérée ou problème
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 19/02/2019
; Notes .........:
;================================================================================================================================
Func _OCCheckCardInsertion()
	Local $sFuncName = "_OCCheckCardInsertion"
	Local $sFileLine
	Local $hLogFile
	Local $sPattern = "- Consultation de la carte	Vitale"
	Local $bCardInserted = False
	Local $isFlash		 = False

	; On ne fait des recherches que si OC est lance
	If WinExists(_YDGVars_Get("sAppOCTitle")) Then
		;------------------------------
		; On gere la reactivation de la fenetre OC si flash (fenetre clignotante dans la barre des taches)
		If _YDTool_WinIsFlash(_YDGVars_Get("sAppOCTitle")) Then
			$isFlash = True
			WinActivate(_YDGVars_Get("sAppOCTitle"))
			Sleep(500)
		EndIf
		;------------------------------
		; On récupère le handle et le PID de OutilCaisse.exe
		$hWndOutilCaisse = WinGetHandle(_YDGVars_Get("sAppOCTitle"))
		$iPIDOutilCaisse = WinGetProcess($hWndOutilCaisse)
		; On ne fait des recherches que lorsque c est neccessaire
		If ($iPIDActive = $iPIDOutilCaisse And $hWndActive <> $hWndOutilCaisse And WinGetTitle($hWndActive) = "") Or ($isFlash = True) Then
			;------------------------------
			; On log les infos utiles
			_YDLogger_Var("$isFlash", $isFlash, $sFuncName, 2)
			_YDLogger_Var("$hWndActive", $hWndActive, $sFuncName, 2)
			_YDLogger_Var("$hWndOutilCaisse", $hWndOutilCaisse, $sFuncName, 2)
			_YDLogger_Var("$iPIDActive", $iPIDActive, $sFuncName, 2)
			_YDLogger_Var("$iPIDOutilCaisse", $iPIDOutilCaisse, $sFuncName, 2)
			_YDLogger_Var("WinGetTitle($hWndActive)", WinGetTitle($hWndActive), $sFuncName, 2)
			_YDLogger_Var("WinGetTitle($hWndOutilCaisse)", WinGetTitle($hWndOutilCaisse), $sFuncName, 2)
			;------------------------------
			; On attend que OC revienne au premier plan
			WinWaitActive($hWndOutilCaisse)
			$hLogFile = FileOpen(_YDGVars_Get("sAppOCLogFile"), 0)
			If $hLogFile = -1 Then
				_YDLogger_Error("Fichier impossible a ouvrir : " & $hLogFile, $sFuncName)
				_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), "Fichier OCLog inaccessible !", 0, $TIP_ICONASTERISK)
				Return False
			Endif
			; On recupere le nombre de lignes du fichier de log
			Local $iFileCountLine = _FileCountLines(_YDGVars_Get("sAppOCLogFile"))
			_YDLogger_Var("$iFileCountLine", $iFileCountLine, $sFuncName, 2)
			; Si le nb de ligne du fichier < compteur, on reinitialise le compteur a 1
			If $iFileCountLine < $iFileLastLine Then $iFileLastLine = 1
			_YDLogger_Var("$iFileLastLine", $iFileLastLine, $sFuncName, 2)
			; On boucle sur le fichier log pour detecter l insertion d une carte vitale
			For $i = $iFileCountLine to $iFileLastLine Step -1
				$sFileLine = FileReadLine($hLogFile, $i)
				Local $iCardInserted = StringInStr($sFileLine, $sPattern, 0, 1, 1)
				; Si pattern trouve on sort de la boucle
				If $iCardInserted > 0 Then
					;------------------------------
					_YDLogger_Log("Pattern trouve", $sFuncName)
					$bCardInserted = True
					$iFileLastLine = _FileCountLines(_YDGVars_Get("sAppOCLogFile")) + 1
					ExitLoop
				EndIf
			Next
			FileClose($hLogFile)
			;------------------------------
			; On log les infos utiles
			_YDLogger_Var("$bCardInserted", $bCardInserted, $sFuncName)
			_YDLogger_Var("$iFileLastLine", $iFileLastLine, $sFuncName)
			;------------------------------
			; On renvoi True si carte vitale rellement inseree
			If $bCardInserted = True Then Return True
		EndIf
	Endif
	;------------------------------
	; On renvoi False dans tous les autres cas
	Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Permet de récupérer le NIR via un export OC
; Syntax ........: _OCGetNir()
; Parameters ....:
; Return values .: True 	- Le NIR a été récupéré
;                  False 	- Problème lors de la récupération du NIR
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 19/02/2019
; Notes .........:
;================================================================================================================================
Func _OCGetNir()
	Local $sFuncName = "_OCGetNir"
	Local $iMaxLoops = 20
	Local $i

	; On tente d'activer la fenetre OC
	WinActivate(_YDGVars_Get("sAppOCTitle"))
	_YDLogger_Log("Attente activation Fenetre OC ...", $sFuncName, 2)
	$i = 1
	While WinActive(_YDGVars_Get("sAppOCTitle")) = 0 And WinGetState(_YDGVars_Get("sAppOCTitle")) <> 15
		_YDLogger_Log("Tentative " & $i, $sFuncName, 2)
		WinActivate(_YDGVars_Get("sAppOCTitle"))
		$i = $i + 1
		Sleep(100)
		If $i >= $iMaxLoops Then
			_YDLogger_Log("Impossible d'activer la Fenetre OC", $sFuncName, 2)
			Return False
		EndIf
	Wend
	_YDLogger_Log("Fenetre OC activee", $sFuncName, 2)
	;------------------------------
	_YDLogger_Log("Sends pour export du XML", $sFuncName)
	_YDLogger_Log("Send(!i)", $sFuncName, 2)
	Send("!i")
	_YDLogger_Log("Sleep(200)", $sFuncName, 2)
	Sleep(200)
	;------------------------------
	_YDLogger_Log("Attente activation Popup Export : " & WinGetHandle("[ACTIVE]"), $sFuncName, 2)
	$i = 1
	While WinActive("Information") = 0 And WinGetState("Information") <> 15
		_YDLogger_Log("Tentative " & $i, $sFuncName, 2)
		WinActivate("Information")
		$i = $i + 1
		Sleep(100)
		If $i >= $iMaxLoops Then
			_YDLogger_Log("Impossible d'activer la Popup Export", $sFuncName, 2)
			Return False
		EndIf
	Wend
	_YDLogger_Log("Popup Export activee : State = " & WinGetState("Information"), $sFuncName, 2)
	;------------------------------
	_YDLogger_Log("Send(ENTER)", $sFuncName, 2)
	Send("{ENTER}")
	_YDLogger_Log("Sleep(200)", $sFuncName, 2)
	Sleep(200)
	;------------------------------
	_YDLogger_Log("Attente reactivation Fenetre OC ...", $sFuncName, 2)
	$i = 1
	While WinActive(_YDGVars_Get("sAppOCTitle")) = 0 And WinGetState(_YDGVars_Get("sAppOCTitle")) <> 15
		_YDLogger_Log("Tentative " & $i, $sFuncName, 2)
		WinActivate(_YDGVars_Get("sAppOCTitle"))
		$i = $i + 1
		Sleep(100)
		If $i >= $iMaxLoops Then
			_YDLogger_Log("Impossible de reactiver la Fenetre OC", $sFuncName, 2)
			Return False
		Endif
	Wend
	_YDLogger_Log("Fenetre OC reactivee", $sFuncName, 2)
	;------------------------------
	Sleep(200)
	_YDLogger_Log("-------------------------", $sFuncName, 2)
	_YDLogger_Log("Lecture du fichier : " & _YDGVars_Get("sAppOCXmlFile"), $sFuncName)
	_YDLogger_Log("FileGetTime : " & FileGetTime(_YDGVars_Get("sAppOCXmlFile"), $FT_MODIFIED, $FT_STRING), $sFuncName, 2)
	;------------------------------
	If Not FileExists(_YDGVars_Get("sAppOCXmlFile")) Then
		_YDLogger_Log("Le fichier XML n'existe pas : " & _YDGVars_Get("sAppOCXmlFile"), $sFuncName)
		Return False
	EndIf
	Local $hXmlFile = FileRead(_YDGVars_Get("sAppOCXmlFile"))
	;------------------------------
	_YDLogger_Log("Récupération du NIR dans le fichier XML", $sFuncName)
	Local $aNirXml = _StringBetween($hXmlFile,"<nir>","</nir>")
	If Not IsArray($aNirXml) Then
		_YDLogger_Log("Erreur recuperation valeur XML ($aNirXml = " & $aNirXml & ")", $sFuncName)
		FileClose(_YDGVars_Get("sAppOCXmlFile"))
		FileDelete(_YDGVars_Get("sAppOCXmlFile"))
		Return False
	Endif
	FileClose(_YDGVars_Get("sAppOCXmlFile"))
	FileDelete(_YDGVars_Get("sAppOCXmlFile"))
	;------------------------------
	_YDLogger_Log("Split de $aNirXml", $sFuncName)
	$aNir = StringSplit($aNirXml[0], " ")
	If Not IsArray($aNir) Then
		_YDLogger_Log("Erreur tableau $aNir vide", $sFuncName)
		Return False
	EndIf
	_YDLogger_Var("$aNir[1]", $aNir[1], $sFuncName, 2)
	_YDLogger_Var("$aNir[2]", $aNir[2], $sFuncName, 2)
	;------------------------------
	_YDLogger_Log("Copie du NIR dans le clipboard", $sFuncName)
	ClipPut($aNir[1])
	;------------------------------
	If @error = 0 Then Return True
	;------------------------------
	_YDLogger_Error("Erreur innatendue : " & @error, $sFuncName)
	Return False
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Copie le NIR dans le presse papier et affiche un TrayTip
; Syntax ........: _SetNirInClipboard($_sNir)
; Parameters ....: $_sNir 		- NIR sans la clé à copier
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 18/02/2019
; Notes .........:
;================================================================================================================================
Func _SetNirInClipboard($_sNir)
	Local $sFuncName = "_SetNirInClipboard"
	_YDLogger_Var("$sNir", $_sNir, $sFuncName)
	ClipPut($_sNir)
	_YDTool_SetTrayTip(_YDGVars_Get("sAppTitle"), $_sNir &  " : NIR copié dans le presse-papier", 5000, $TIP_ICONASTERISK)
EndFunc

; #FUNCTION# ====================================================================================================================
; Description ...: Lance Outils Caisse si pas deja lance
; Syntax ........: _OCCheckCardInsertion()
; Parameters ....:
; Return values .:
; Author ........: yann.daniel@assurance-maladie.fr
; Last Modified .: 18/02/2019
; Notes .........:
;================================================================================================================================
Func _OCLaunch()
	Local $sFuncName = "_OCLaunch"
	If ProcessExists(_YDGVars_Get("sAppOCExeName")) Then
		_YDTool_SetMsgBoxError("Outils Caisse est déjà en cours d'execution !", $sFuncName)
	Else
		If FileExists(_YDGVars_Get("sAppOCExeFile")) Then
			_YDLogger_Log("Lancement OC", $sFuncName)
			Run(_YDGVars_Get("sAppOCExeFile"))
		EndIf
	EndIf
EndFunc
