#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <File.au3>
#include <Array.au3>
#include "TestRail.au3"
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Crypt.au3>
#include <GuiComboBox.au3>

Global $run_ids
Global $html
Global $app_name = "Test Run Results Reporter"
Global $ini_filename = @ScriptDir & "\" & $app_name & ".ini"

Global $main_gui = GUICreate("TRRR - " & $app_name, 860, 600)

GUICtrlCreateGroup("TestRail Setup", 10, 10, 410, 110)
GUICtrlCreateLabel("TestRail Username", 20, 30, 100, 20)
Global $testrail_username_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "testrailusername", "sgriffin@janison.com"), 140, 30, 250, 20)
GUICtrlCreateLabel("TestRail Password", 20, 50, 100, 20)
Global $testrail_password_input = GUICtrlCreateInput("", 140, 50, 250, 20, $ES_PASSWORD)
;GUICtrlCreateLabel("TestRail Project", 20, 70, 100, 20)
;Global $testrail_project_combo = GUICtrlCreateCombo("", 140, 70, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
;GUICtrlCreateLabel("TestRail Plan", 20, 90, 100, 20)
;Global $testrail_plan_combo = GUICtrlCreateCombo("", 140, 90, 250, 20, BitOR($CBS_DROPDOWNLIST, $WS_VSCROLL))
GUICtrlCreateGroup("", -99, -99, 1, 1)

Local $testrail_encrypted_password = IniRead($ini_filename, "main", "testrailpassword", "")
Global $testrail_decrypted_password = ""

if stringlen($testrail_encrypted_password) > 0 Then

	$testrail_decrypted_password = _Crypt_DecryptData($testrail_encrypted_password, "applesauce", $CALG_AES_256)
	$testrail_decrypted_password = BinaryToString($testrail_decrypted_password)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_decrypted_password = ' & $testrail_decrypted_password & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	GUICtrlSetData($testrail_password_input, $testrail_decrypted_password)
Else

	$testrail_decrypted_password = ""
EndIf

GUICtrlCreateLabel("TestRail Runs", 20, 130, 100, 20)
Global $testrail_run_ids_input = GUICtrlCreateInput(IniRead($ini_filename, "main", "testrailruns", "903"), 140, 130, 680, 20)
Global $start_button = GUICtrlCreateButton("Start", 10, 160, 100, 20, -1, $BS_DEFPUSHBUTTON)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $display_report_button = GUICtrlCreateButton("Display Report", 120, 160, 100, 20)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $display_data_button = GUICtrlCreateButton("Display Data", 240, 160, 100, 20)
GUICtrlSetState(-1, $GUI_DISABLE)

Global $listview = GUICtrlCreateListView("Run ID|PID|Extract Status", 10, 200, 410, 300, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 200)
_GUICtrlListView_SetColumnWidth(-1, 1, 50)
_GUICtrlListView_SetColumnWidth(-1, 2, 200)
_GUICtrlListView_SetExtendedListViewStyle($listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))

Global $status_input = GUICtrlCreateInput("Enter the ""Run IDs"" and click ""Start""", 10, 600 - 25, 400, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $progress = GUICtrlCreateProgress(420, 600 - 25, 400, 20)

GUISetState(@SW_SHOW, $main_gui)

; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)

FileDelete(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Open(@ScriptDir & "\" & $app_name & ".sqlite")
_SQLite_Exec(-1, "CREATE TABLE report (RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails,TestCaseID,TestCaseOwner,TestCaseStatus,Issues);") ; CREATE a Table
;_SQLite_Exec(-1, "CREATE TABLE defects_in_tests (TestCaseID,BugID,UNIQUE(TestCaseID, BugID));") ; CREATE a Table
_SQLite_Exec(-1, "CREATE TABLE defect (BugID,BugSummary,Priority,TestCaseEpicStory,Impact,ActionRequired,FixDate,FixPhase);") ;,UNIQUE(TestCaseEpicStory, BugID));") ; CREATE a Table

; Startup TestRail

;GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
;_TestRailDomainSet("https://janison.testrail.com")
;_TestRailLogin(GUICtrlRead($testrail_username_input), GUICtrlRead($testrail_password_input))

;if StringLen(GUICtrlRead($testrail_password_input)) > 0 Then

	; Authentication

;	GUICtrlSetData($status_input, "Authenticating against TestRail ... ")
;	_TestRailAuth()

	GUICtrlSetState($start_button, $GUI_ENABLE)
;EndIf

GUICtrlSetData($status_input, "")
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")

; Loop until the user exits.
While 1

	; GUI msg loop...
	$msg = GUIGetMsg()

	Switch $msg

		Case $GUI_EVENT_CLOSE

			IniWrite($ini_filename, "main", "testrailusername", GUICtrlRead($testrail_username_input))
;			IniWrite($ini_filename, "main", "testrailproject", GUICtrlRead($testrail_project_combo))
;			IniWrite($ini_filename, "main", "testrailplan", GUICtrlRead($testrail_plan_combo))
			IniWrite($ini_filename, "main", "testrailruns", GUICtrlRead($testrail_run_ids_input))

			$testrail_encrypted_password = _Crypt_EncryptData(GUICtrlRead($testrail_password_input), "applesauce", $CALG_AES_256)
			IniWrite($ini_filename, "main", "testrailpassword", $testrail_encrypted_password)

			ExitLoop

		Case $start_button

			_SQLite_Exec(-1, "DELETE FROM report;") ; CREATE a Table
			_SQLite_Exec(-1, "DELETE FROM defect;") ; CREATE a Table

			GUICtrlSetData($progress, 0)
			GUICtrlSetState($testrail_run_ids_input, $GUI_DISABLE)
			GUICtrlSetState($start_button, $GUI_DISABLE)
			GUICtrlSetState($display_report_button, $GUI_DISABLE)
			GUICtrlSetState($display_data_button, $GUI_DISABLE)
			GUISetCursor(15, 1, $main_gui)
			_GUICtrlListView_DeleteAllItems($listview)

			; populate listview with testrail run ids

			Local $run_id = StringSplit(GUICtrlRead($testrail_run_ids_input), ",;|", 2)

			for $each in $run_id

				Local $pid = ShellExecute(@ScriptDir & "\data_extractor.exe", """" & GUICtrlRead($testrail_username_input) & """ """ & GUICtrlRead($testrail_password_input) & """ " & $each & " ""sgriffin@janison.com.au"" ""Gri04ffo..""", "", "", @SW_HIDE)
				GUICtrlCreateListViewItem($each & "|" & $pid & "|In Progress", $listview)
			Next

			While True

				Local $all_run_ids_done = True

				for $index = 0 to (_GUICtrlListView_GetItemCount($listview) - 1)

					Local $pid = _GUICtrlListView_GetItemText($listview, $index, 1)
					Local $status = _GUICtrlListView_GetItemText($listview, $index, 2)

					if StringCompare($status, "In Progress") = 0 Then

						$all_run_ids_done = False

						if ProcessExists($pid) = False Then

							_GUICtrlListView_SetItemText($listview, $index, "Done", 2)

						EndIf
					EndIf
				Next

				if $all_run_ids_done = True Then

					ExitLoop
				EndIf

				Sleep(1000)
			WEnd

			$html = 				"<!DOCTYPE html>" & @CRLF & _
									"<html>" & @CRLF & _
									"<head>" & @CRLF & _
									"<style>" & @CRLF & _
									"table, th, td {" & @CRLF & _
									"    border: 1px solid black;" & @CRLF & _
									"    border-collapse: collapse;" & @CRLF & _
									"    font-size: 12px;" & @CRLF & _
									"    font-family: Arial;" & @CRLF & _
									"}" & @CRLF & _
									".ds {width: 400px; text-align: left;}" & @CRLF & _
									".tes {width: 800px; text-align: left;}" & @CRLF & _
									".mti {width: 110px; text-align: center;}" & @CRLF & _
									".tt {width: 300px; text-align: left;}" & @CRLF & _
									".ati {width: 150px; text-align: center;}" & @CRLF & _
									".sd {width: 1000px; text-align: left;}" & @CRLF & _
									".tc {width: 300px; text-align: center;}" & @CRLF & _
									".tr {width: 110px; text-align: center;}" & @CRLF & _
									".trp {width: 110px; text-align: center; background-color: yellowgreen;}" & @CRLF & _
									".trf {width: 110px; text-align: center; background-color: lightcoral; color:white;}" & @CRLF & _
									".tru {width: 110px; text-align: center; background-color: lightgray;}" & @CRLF & _
									".trb {width: 110px; text-align: center; background-color: darkred; color: white;}" & @CRLF & _
									".pass {background-color: yellowgreen;}" & @CRLF & _
									".fail {background-color: lightcoral; color:white;}" & @CRLF & _
									".untested {background-color: lightgray;}" & @CRLF & _
									".mp {background-color: yellow;}" & @CRLF & _
									".rh {background-color: seagreen; color: white;}" & @CRLF & _
									".i {background-color: deepskyblue;}" & @CRLF & _
									"</style>" & @CRLF & _
									"</head>" & @CRLF & _
									"<body>" & @CRLF & _
									"<h2>Outstanding Defects Report</h2>" & @CRLF

			SQLite_to_HTML_table("SELECT BugID AS ""Defect ID"",BugSummary AS ""Summary"",Priority,'<table><tbody>' || GROUP_CONCAT(TestCaseEpicStory, """") || '</tbody></table>' AS ""Test Case - Epic - Story"",Impact,ActionRequired AS ""Action Required"",FixDate AS ""Fix Date"",FixPhase AS ""Fix Phase"" FROM defect GROUP BY BugID,BugSummary,Priority,Impact,ActionRequired,FixDate,FixPhase ORDER BY BugID;", "tr,ds,tr,tes,tr,tr,tr,tr", "", "")

			$html = $html &			"<h2>Test Run Results Report</h2>" & @CRLF

			for $each in $run_id

				SQLite_to_HTML_table("SELECT '<a href=""https://janison.testrail.com/index.php?/tests/view/' || ManualTestID || '"" target=""_blank"">' || ManualTestID || '</a><br>' || TestTitle AS ""Manual Test"",AutoTestID AS ""Auto Test"",TestResult AS ""Test Result"",StepDetails AS ""Step Details"",'<table><tbody><tr><td><a href=""https://janison.testrail.com/index.php?/cases/view/' || TestCaseID || '"" target=""_blank"">' || TestCaseID || '</a></td><td>' || TestCaseOwner || '</td><td>' || TestCaseStatus || '</td></tr></tbody></table>' AS ""Test Case - Owner - Status"",Issues AS ""Epic - Story - Bugs"" FROM report WHERE RunID = '" & $each & "' ORDER BY ManualTestID;", "tt,ati,tr,sd,tc,tc", "", $each)
			Next

			$html = $html &			"</body>" & @CRLF & _
									"</html> " & @CRLF

			FileDelete(@ScriptDir & "\html_report.html")
			FileWrite(@ScriptDir & "\html_report.html", $html)

			Local $aResult, $iRows, $iColumns, $iRval

			$iRval = _SQLite_GetTable2d(-1, "SELECT * FROM report;", $aResult, $iRows, $iColumns)

			If $iRval = $SQLITE_OK Then

				Local $data = _SQLite_Display2DResult($aResult, 0, True)
				FileDelete(@ScriptDir & "\data.txt")
				FileWrite(@ScriptDir & "\data.txt", $data)
			Else
				MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
			EndIf

			GUICtrlSetData($progress, 0)
			GUICtrlSetData($status_input, "")
			GUICtrlSetState($testrail_run_ids_input, $GUI_ENABLE)
			GUICtrlSetState($start_button, $GUI_ENABLE)
			GUICtrlSetState($display_report_button, $GUI_ENABLE)
			GUICtrlSetState($display_data_button, $GUI_ENABLE)
			GUISetCursor(2, 0, $main_gui)

		Case $display_report_button

			ShellExecute(@ScriptDir & "\html_report.html")

		case $display_data_button

			ShellExecute("notepad", "data.txt", @ScriptDir)


	EndSwitch

WEnd

GUIDelete($main_gui)
_SQLite_Close()
_SQLite_Shutdown()



Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg
    Local $hWndFrom, $iIDFrom, $iCode
    $hWndFrom = $lParam
    $iIDFrom = BitAND($wParam, 0xFFFF) ; Low Word
    $iCode = BitShift($wParam, 16) ; Hi Word

;	Switch $hWndFrom

;        Case GUICtrlGetHandle($testrail_project_combo)

;			Switch $iCode

 ;               Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

;					query_testrail_plans()
 ;           EndSwitch

  ;      Case GUICtrlGetHandle($testrail_plan_combo)

	;		Switch $iCode

     ;           Case $CBN_SELCHANGE ; Sent when the user changes the current selection in the list box of a combo box

	;				query_testrail_runs()

     ;       EndSwitch
 ;   EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_COMMAND

Func SQLite_to_HTML_table($query, $classes, $empty_message, $run_id)

	Local $class = StringSplit($classes, ",", 2)

	Local $aResult, $iRows, $iColumns, $iRval, $run_name = ""

;	$xx = "SELECT RunName AS ""Run Name"" FROM report WHERE RunID = '" & $run_id & "';"
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $xx = ' & $xx & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	if StringLen($run_id) > 0 Then

		$iRval = _SQLite_GetTable2d(-1, "SELECT RunName AS ""Run Name"" FROM report WHERE RunID = '" & $run_id & "';", $aResult, $iRows, $iColumns)

		If $iRval = $SQLITE_OK Then

			_SQLite_Display2DResult($aResult)

			$run_name = $aResult[1][0]
		EndIf

		$html = $html &	"<h3>Test Run " & $run_id & " - " & $run_name & "</h3>" & @CRLF
	EndIf

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

			$html = $html &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$html = $html &	"<table>" & @CRLF
			$html = $html & "<tr>"

			for $i = 0 to ($num_cols - 1)

				$html = $html & "<th class=""rh"">" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$html = $html & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$html = $html & "<tr>"

				for $j = 0 to ($num_cols - 1)

					if $j = 2 Then

	;					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $aResult[$i][$j] = ' & $aResult[$i][$j] & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

						Switch $aResult[$i][$j]

							case "Passed"

								$class[$j] = "trp"

							case "Failed"

								$class[$j] = "trf"

							case "Untested"

								$class[$j] = "tru"

							case "Blocked"

								$class[$j] = "trb"
						EndSwitch
					EndIf

					$html = $html & "<td class=""" & $class[$j] & """>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$html = $html & "</tr>" & @CRLF
			Next

			$html = $html &	"</table>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf
EndFunc

#cs
Func query_testrail_plans()

	Local $project_id_name = GUICtrlRead($testrail_project_combo)

	IniWrite($ini_filename, "main", "testrailproject", $project_id_name)

	Local $project_part = StringSplit($project_id_name, " - ", 3)

	GUICtrlSetData($status_input, "Querying TestRail Plans ... ")

	Local $plan_id_name = _TestRailGetPlansIDAndNameArray($project_part[0])
	Local $plan_id_str = ""

	for $i = 0 to (UBound($plan_id_name) - 1)

		if StringLen($plan_id_str) > 0 Then

			$plan_id_str = $plan_id_str & "|"
		EndIf

		$plan_id_str = $plan_id_str & $plan_id_name[$i][0] & " - " & $plan_id_name[$i][1]
	Next

	GUICtrlSetData($testrail_plan_combo, $plan_id_str)
	GUICtrlSetData($status_input, "")
	GUICtrlSetState($testrail_plan_combo, $GUI_ENABLE)

EndFunc

Func query_testrail_runs()

	Local $plan_id_name = GUICtrlRead($testrail_plan_combo)
	Local $plan_part = StringSplit($plan_id_name, " - ", 3)

	GUICtrlSetData($status_input, "Querying TestRail Runs ... ")
	Local $run_id = _TestRailGetPlanRunsID($plan_part[0])
	$run_ids = _ArrayToString($run_id)
	GUICtrlSetData($status_input, "")
	GUICtrlSetState($start_button, $GUI_ENABLE)

EndFunc
#ce