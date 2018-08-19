; ----------------------------------------------------
; -------------------- Section I --------------------
; ----------------------------------------------------
; Version AutoIt :    3.2.8.1
; Langue     :        Francais
; Plateforme :        Windows
; Auteur    :        Paul Petit
;
; Fonction du script: Terminal d'inventaire pour LGPI
;
;
;
; Version 1.0 : 21/11/2007
;           - Première Version.
;

; ----------------------------------------------------
; -------------------- Section II --------------------
; ----------------------------------------------------
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=favicon.ico
#AutoIt3Wrapper_Res_Description=Logiciel de terminal d'inventaire
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_ProductName=Terminal d'inventaire
#AutoIt3Wrapper_Res_ProductVersion=1.0.0
#AutoIt3Wrapper_Res_CompanyName=Paul Petit
#AutoIt3Wrapper_Res_LegalCopyright=©2018 Paul Petit
#AutoIt3Wrapper_Res_Language=1036
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

; ----------------------------------------------------
; -------------------- Section III -------------------
; ----------------------------------------------------
#Region ### Includes ###

#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <GuiButton.au3>
#include <WindowsConstants.au3>
#include <Json.au3>
#include <Inet.au3>
#include <Array.au3>
#include <GuiEdit.au3>

#EndRegion ### Includes ###

; ----------------------------------------------------
; -------------------- Section IV --------------------
; ----------------------------------------------------
#Region ### Variables declaration ###

; Constantes pour les dimensions de la fenêtre et pour la centrer
Const $iGuiWidth = 800, $iGuiHeight = 555, $iGuiXPos = (@DesktopWidth / 2) - $iGuiWidth / 2, $iGuiYPos = (@DesktopHeight / 2) - $iGuiHeight / 2

; Variables contenat la liste des inventaires
Global $data

; Constante contenant le chemin du fichier de configuration
Global Const $sIniFile = @ScriptDir & "\config.ini"

#EndRegion ### Variables declaration ###

; ----------------------------------------------------
; -------------------- Section V ---------------------
; ----------------------------------------------------
#Region ### START Koda GUI section ### Form=c:\users\paul\desktop\autoit\form1.kxf

; Fenêtre principale
$Form1_1 = GUICreate("Terminal d'inventaire", $iGuiWidth, $iGuiHeight, $iGuiXPos, $iGuiYPos)

; Bouton pour la récupération des données
$Btn_retrive_data = GUICtrlCreateButton("Récupérer les données", 84, 392, 153, 41)

; Liste qui affiches toutes les commandes
$Label1 = GUICtrlCreateLabel("Dernières commandes", 56, 56, 109, 17)
$List1 = GUICtrlCreateList("", 48, 80, 225, 288, BitOR($LBS_NOTIFY, $WS_VSCROLL, $WS_BORDER), 0)

; Edit qui affiche le contenu des commandes
$Label_selected_inventory = GUICtrlCreateLabel("", 411, 15, 200, 17)
$Label_selected_inventory_date = GUICtrlCreateLabel("", 411, 32, 200, 17)
$Label2 = GUICtrlCreateLabel("Contenu de la commande", 411, 56, 125, 17)
$Edit1 = GUICtrlCreateEdit("", 388, 80, 337, 289, BitOR($GUI_SS_DEFAULT_EDIT, $ES_READONLY))

; Bouton d'envoi des données
$Btn_send_data = GUICtrlCreateButton("Transmettre les données à LGPI", 482, 480, 165, 41)
_GUICtrlButton_Enable($Btn_send_data, False)

; Slider pour choisir le delai de transmission
$Slider1 = GUICtrlCreateSlider(481, 432, 150, 45)
GUICtrlSetLimit($Slider1, 30, 5)
GUICtrlSetData($Slider1, IniRead($sIniFile, "data", "transmission_delay", 10))
GUICtrlSetTip($Slider1, "Délai de transmission")
$Label3 = GUICtrlCreateLabel("Délai de transmission (s)", 505, 412, 117, 17)
$Label_delay = GUICtrlCreateLabel(IniRead($sIniFile, "data", "transmission_delay", 10), 640, 436, 20, 17)

;Affichage de l'interface
GUISetState(@SW_SHOW)

#EndRegion ### END Koda GUI section ###

; ----------------------------------------------------
; -------------------- Section VI --------------------
; ----------------------------------------------------
#Region ### Boucle principale ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $Btn_retrive_data

			_GUICtrlButton_Enable($Btn_send_data, False)
			_GUICtrlButton_SetText($Btn_send_data, "Transmettre les données à LGPI")
			getData()

			_GUICtrlListBox_SetCurSel($List1, 0)
			_WinAPI_SetFocus(ControlGetHandle("Terminal d'inventaire", "", $List1))
			displayInventory(0)
			_GUICtrlButton_Enable($Btn_send_data)


		Case $List1

			displayInventory(_GUICtrlListBox_GetCurSel($List1))
			_GUICtrlButton_Enable($Btn_send_data)
			_GUICtrlButton_SetText($Btn_send_data, "Transmettre les données à LGPI")
		Case $Btn_send_data
			sendData()

		Case $Slider1
			IniWrite($sIniFile, "data", "transmission_delay", GUICtrlRead($Slider1))
			GUICtrlSetData($Label_delay, GUICtrlRead($Slider1))

	EndSwitch
WEnd

#EndRegion ### Boucle principale ###

; ----------------------------------------------------
; -------------------- Section VII -------------------
; ----------------------------------------------------
#Region ### Fonctions ###

Func _INetGetSourceEx($s_URL, $bString = True)
	Local $sString = InetRead($s_URL, 1)
	Local $nError = @error, $nExtended = @extended
	If $bString Then $sString = BinaryToString($sString, 4)
	Return SetError($nError, $nExtended, $sString)
EndFunc   ;==>_INetGetSourceEx

Func getData()
	;------------------------------------------------------
	; Cette fonction récupère les données et les affichent
	;------------------------------------------------------

	; On vide le contenu des listes ainsi que les labels
	GUICtrlSetData($List1, "")
	GUICtrlSetData($Edit1, "")
	GUICtrlSetData($Label_selected_inventory, "")
	GUICtrlSetData($Label_selected_inventory_date, "")

	; Le bouton de récupération de données va afficher 'Téléchargement..' et est désactivé
	_GUICtrlButton_SetText($Btn_retrive_data, 'Téléchargement ...')
	_GUICtrlButton_Enable($Btn_retrive_data, False)

	;---------------------------------------------------------------------------------;
	; Connexion a firebase et récupération des données

	; Récupération des données brutes
	Local $URL = 'https://terminal-inventaire.firebaseio.com/inventories.json?orderBy="date"'
	Local $webData = _INetGetSourceEx($URL)

	Local $dataDecoded = json_decode($webData)
	Local $keys = Json_ObjGetKeys($dataDecoded)

	;---------------------------------------------------------------------------------;
	; Il faut trier les données par date décroissante, on retroune le tableau
	; On créé une chaine de caractère en encodant et décodant
	Local $dataBonOrdre = "[" ;

	For $i = (UBound($keys) - 1) To 0 Step -1

		Local $value = Json_ObjGet($dataDecoded, $keys[$i])
		$dataBonOrdre &= Json_Encode($value)

		If $i <> 0 Then
			$dataBonOrdre &= ','
		EndIf

	Next

	$dataBonOrdre &= ']'

	$data = json_decode($dataBonOrdre)

	; On remplit la liste
	Local $list = ""

	For $i = 0 To (UBound($data) - 1) Step 1
		$list &= getFormattedDate(Round( Json_ObjGet($data[$i], "date") / 1000)) & ' - ' & Json_ObjGet($data[$i], "name") & "|"
	Next

	GUICtrlSetData($List1, $list)

	; On réinitialise le bouton de récupération de données
	_GUICtrlButton_SetText($Btn_retrive_data, 'Récupérer les données')
	_GUICtrlButton_Enable($Btn_retrive_data, True)

EndFunc   ;==>getData

Func displayInventory($dIndex)
	;------------------------------------------------
	; Affiche le contenu de la commande sélectionnée
	;------------------------------------------------

	;on affiche le nom de la commande au dessus du contenu

	GUICtrlSetData($Label_selected_inventory, Json_ObjGet($data[$dIndex], "name"))
	GUICtrlSetData($Label_selected_inventory_date, getFormattedDate(Round( Json_ObjGet($data[$dIndex], "date") / 1000)))

	;on recupère les bonnes données
	Local $inventoryToDisplay = Json_ObjGet($data[$dIndex], "entries")
	Local $stringToDisplay = ""

	For $i = 0 To (UBound($inventoryToDisplay) - 1) Step 1
		$stringToDisplay &= Json_ObjGet($inventoryToDisplay[$i], "qty")
		$stringToDisplay &= ' - '
		$stringToDisplay &= Json_ObjGet($inventoryToDisplay[$i], "code")
		$stringToDisplay &= "\r\n"
	Next

	_GUICtrlEdit_SetText($Edit1, StringFormat($stringToDisplay))

EndFunc   ;==>displayInventory

Func getFormattedDate($unixTimeStamp)
	;--------------------------------------------------
	; Revoie la date formatée à partir d'un time stamp
	;--------------------------------------------------
	Local $timezone = _Date_Time_GetTimeZoneInformation()
	$unixTimeStamp = $unixTimeStamp - ($timezone[1] + $timezone[7]) * 60
	Return (_DateTimeFormat(_DateAdd('s', $unixTimeStamp, "1970/01/01 00:00:00"), 0))
EndFunc   ;==>getFormattedDate

Func sendData()
	;---------------------------------------------------------------------------------------
	; Récupère les données de la commande sélectionnée, elles sont ensuite tapées dans LGPI
	;---------------------------------------------------------------------------------------

	;on récupère les données à envoyer
	Local $inventoryToDisplay = Json_ObjGet($data[_GUICtrlListBox_GetCurSel($List1)], "entries")
	Local $dataToSend = ""

	For $i = 0 To (UBound($inventoryToDisplay) - 1) Step 1
		$dataToSend &= Json_ObjGet($inventoryToDisplay[$i], "code")
		$dataToSend &= '{ENTER}'
		$dataToSend &= Json_ObjGet($inventoryToDisplay[$i], "qty")
		$dataToSend &= "{ENTER}"
	Next

	; On affiche une MsgBox de confirmation
	Local $msgres = MsgBox(64 + 1, "Transmission à LGPI", "Vous avez " & IniRead($sIniFile, "data", "transmission_delay", 10) & " secondes pour placer le curseur sur 'Scan Produit' dans LGPI")

	; Si 'ok' sélectionné, on fait la pause et son tape les données dans LGPI
	If $msgres == $IDOK Then
		Sleep(IniRead($sIniFile, "data", "transmission_delay", 10) * 1000)
		Send($dataToSend)
		MsgBox(64, "Transmission à LGPI", "Données transmises !")
		_GUICtrlButton_Enable($Btn_send_data, False)
		_GUICtrlButton_SetText($Btn_send_data, "Données transmines !")
	EndIf

EndFunc   ;==>sendData

#EndRegion ### Fonctions ###
