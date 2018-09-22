#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#RequireAdmin
;#AutoIt3Wrapper_usex64=n
#include <File.au3>
#include <Array.au3>
#include "Jira.au3"
#include "TestRail.au3"
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <SQLite.au3>
#include <SQLite.dll.au3>

Global $app_name = "Data Extractor"


;#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_run_id = $CmdLine[3]
;#ce

#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_run_id = $CmdLine[3]
#ce

Global $main_gui = GUICreate("TRRR - " & $app_name & " - Run ID " & $testrail_run_id, 840, 360)

Global $listview = GUICtrlCreateListView("Run ID|Test ID|Test Title|Automation Script Reference|Status|Comment", 10, 10, 820, 300, $LVS_SHOWSELALWAYS)
_GUICtrlListView_SetColumnWidth(-1, 0, 60)
_GUICtrlListView_SetColumnWidth(-1, 1, 60)
_GUICtrlListView_SetColumnWidth(-1, 2, 200)
_GUICtrlListView_SetColumnWidth(-1, 3, 200)
_GUICtrlListView_SetColumnWidth(-1, 4, 150)
_GUICtrlListView_SetColumnWidth(-1, 5, 2000)
_GUICtrlListView_SetExtendedListViewStyle($listview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))

Global $status_input = GUICtrlCreateInput("", 10, 360 - 25, 400, 20, $ES_READONLY, $WS_EX_STATICEDGE)
Global $progress = GUICtrlCreateProgress(420, 360 - 25, 400, 20)


GUISetState(@SW_SHOW, $main_gui)


; Startup SQLite

_SQLite_Startup()
ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)
_SQLite_Open(@ScriptDir & "\Test Run Results Reporter.sqlite")

; Startup TestRail

GUICtrlSetData($status_input, "Starting the TestRail connection ... ")
_TestRailDomainSet("https://janison.testrail.com")
_TestRailLogin($testrail_username, $testrail_password)

; Authentication

GUICtrlSetData($status_input, "Authenticating against TestRail ... ")
_TestRailAuth()


;GUICtrlSetData($status_input, "Querying TestRail Projects ... ")

;Local $project_id_name = _TestRailGetProjectsIDAndNameArray()
;Local $project_id_str = ""

;for $i = 0 to (UBound($project_id_name) - 1)

;	if StringLen($project_id_str) > 0 Then

;		$project_id_str = $project_id_str & "|"
;	EndIf

;	$project_id_str = $project_id_str & $project_id_name[$i][0] & " - " & $project_id_name[$i][1]
;Next

GUICtrlSetData($progress, 0)
GUISetCursor(15, 1, $main_gui)
_GUICtrlListView_DeleteAllItems($listview)
GUICtrlSetData($status_input, "")

Local $testrail_test_case_id = ""
Local $testrail_test_case_name = ""

_TestRailGetRun($testrail_run_id)
$run_detail = StringRegExp($testrail_json, '(?U)"name":"(.*)".*"config":"(.*)"', 3)
;FileWrite(@ScriptDir & "\fred.txt", $testrail_json)
Global $testrail_run_name = $run_detail[0] & " (" & $run_detail[1] & ")"

_TestRailGetTests($testrail_run_id)
$rr = StringRegExp($testrail_json, '(?U)"id":(\d+),.*"status_id":(\d+),.*"title":"(.*)".*"custom_auto_script_ref":"(.*)"', 3)
;FileWrite(@ScriptDir & "\fred.txt", $testrail_json)


for $i = 0 to (UBound($rr) - 1) step 4

	GUICtrlSetData($progress, ($i / UBound($rr)) * 100)

	Switch $rr[$i + 1]

		case "1"

			$rr[$i + 1] = "Passed"

		case "2"

			$rr[$i + 1] = "Blocked"

		case "4"

			$rr[$i + 1] = "Retest"

		case "5"

			$rr[$i + 1] = "Failed"

		case "6"

			$rr[$i + 1] = "Known Issue"

		case "7"

			$rr[$i + 1] = "Incomplete Automation"

		case "8"

			$rr[$i + 1] = "Removed from Run"

		case "9"

			$rr[$i + 1] = "In Progress"
	EndSwitch

	_TestRailGetResults($rr[$i])

	if StringLen($testrail_json) > 2 Then

		$ss = StringRegExp($testrail_json, '(?U)"comment":(.*),"version"', 3)
		Local $comment = $ss[0]

		if StringCompare($comment, "null") = 0 Then

			$comment = "-"
		Else

			$comment = StringTrimLeft($comment, 1)
			$comment = StringTrimRight($comment, 1)

			Local $comment_line = StringSplit($comment, "\n", 3)
			$comment = ""
			Local $comment_table = False
			Local $comment_class = ""

			for $j = 0 to (UBound($comment_line) - 1)

				if StringLen($comment_line[$j] > 2) Then

					Local $line_start = StringLeft($comment_line[$j], 2)

					if StringCompare($line_start, "||") = 0 Then

						if $comment_table = False Then

							$comment = $comment &	"<table><tbody>"
						EndIf

						if StringInStr($comment_line[$j], "|MANUAL PRE-CONDITIONS") > 0 Then

							$comment_class = " class=""mp"""
						EndIf

						if StringInStr($comment_line[$j], "|AUTOMATED PRE-CONDITIONS") > 0 Or StringInStr($comment_line[$j], "|TEST STEPS") > 0 Or StringInStr($comment_line[$j], "|POST-CONDITIONS") > 0 Then

							$comment_class = ""
						EndIf

						if StringInStr($comment_line[$j], "|EXTRA VIDEO PATH - ") > 0 Then

							Local $href_pos = StringInStr($comment_line[$j], "|EXTRA VIDEO PATH - ") + StringLen("|EXTRA VIDEO PATH - ")
							Local $href = StringMid($comment_line[$j], $href_pos)
							Local $href2 = StringReplace($href, "\\\\", "file://///")
							$href2 = StringReplace($href2, "\\", "/")
							$comment_line[$j] = StringReplace($comment_line[$j], $href, "<br><a href=""" & $href2 & """ target=""_blank"">" & $href & "</a>")
						EndIf

						$comment_table = True
						$comment_line[$j] = StringReplace($comment_line[$j], "||", "<tr><td class=""ati"">")
						$comment_line[$j] = StringReplace($comment_line[$j], "|DEPENDENCIES", "</td><td><b>DEPENDENCIES</b>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|PRE-CONDITIONS", "</td><td><b>PRE-CONDITIONS</b>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|MANUAL PRE-CONDITIONS", "</td><td" & $comment_class & "><b>MANUAL PRE-CONDITIONS</b>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|AUTOMATED PRE-CONDITIONS", "</td><td><b>AUTOMATED PRE-CONDITIONS</b>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|TEST STEPS", "</td><td><b>TEST STEPS</b>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|POST-CONDITIONS", "</td><td><b>POST-CONDITIONS</b>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|VIDEO PLAYBACK - ", "</td><td><b>VIDEO PLAYBACK</b><br>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|INFO|", "</td><td class=""i"">INFO</td><td" & $comment_class & ">")
						$comment_line[$j] = StringReplace($comment_line[$j], "|INFO", "</td><td class=""i"">INFO")
						$comment_line[$j] = StringReplace($comment_line[$j], "|PASS|", "</td><td class=""pass"">PASS</td><td>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|PASS", "</td><td class=""pass"">PASS")
						$comment_line[$j] = StringReplace($comment_line[$j], "|FAIL|", "</td><td class=""fail"">FAIL</td><td>")
						$comment_line[$j] = StringReplace($comment_line[$j], "|FAIL", "</td><td class=""fail"">FAIL")
						$comment_line[$j] = StringReplace($comment_line[$j], "Note - there are no manual activities required for this test.", "<i>Note - there are no manual activities required for this test.</i>")
						$comment_line[$j] = StringReplace($comment_line[$j], "\""", """")
						$comment_line[$j] = StringReplace($comment_line[$j], "\/", "/")
						$comment_line[$j] = StringReplace($comment_line[$j], "amp;", "&")
						$comment_line[$j] = $comment_line[$j] & "</td></tr>"
					Else

						if $comment_table = True Then

							$comment = $comment & "</tbody></table><br>"
						EndIf

						$comment_table = False

						; check if the comment line is a heading to a test case script

						$comment_line[$j] = StringReplace($comment_line[$j], $rr[$i + 3], "<b>" & $rr[$i + 3] & "</b>")

					EndIf
				Else

					if $comment_table = True Then

						$comment = $comment & "</tbody></table>"
					EndIf

					$comment_table = False
				EndIf

				if $comment_table = False and StringLen($comment_line[$j]) > 0 then ;and stringlen($comment) > 0 Then

					$comment_line[$j] = "" & $comment_line[$j] & "<br><br>"
				EndIf

				$comment = $comment & $comment_line[$j]
			Next

			if $comment_table = True Then

				$comment = $comment & "</tbody></table>"
			EndIf
		EndIf

		GUICtrlCreateListViewItem($testrail_run_id & "|" & $rr[$i] & "|" & $rr[$i + 2] & "|" & $rr[$i + 3] & "|" & $rr[$i + 1] & "|" & $comment, $listview)
		_SQLite_Exec(-1, "INSERT INTO report(RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails) VALUES ('" & $testrail_run_id & "','" & $testrail_run_name & "','" & $rr[$i] & "','" & $rr[$i + 2] & "','<b>" & $rr[$i + 3] & "</b>','" & $rr[$i + 1] & "','" & $comment & "');") ; INSERT Data
	Else

		GUICtrlCreateListViewItem($testrail_run_id & "|" & $rr[$i] & "|" & $rr[$i + 2] & "|" & $rr[$i + 3] & "|" & $rr[$i + 1] & "|-", $listview)
		_SQLite_Exec(-1, "INSERT INTO report(RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails) VALUES ('" & $testrail_run_id & "','" & $testrail_run_name & "','" & $rr[$i] & "','" & $rr[$i + 2] & "','" & $rr[$i + 3] & "','<b>" & $rr[$i + 1] & "</b>','-');") ; INSERT Data
	EndIf
Next

GUICtrlSetData($progress, 0)
GUICtrlSetData($status_input, "")
GUISetCursor(2, 0, $main_gui)
GUIDelete($main_gui)

_SQLite_Close()
_SQLite_Shutdown()
