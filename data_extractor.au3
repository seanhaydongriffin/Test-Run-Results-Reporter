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
Global $project_id
Global $suite_id
Global $run_name
Global $run_config
Global $test_id
Global $test_case_id
Global $test_status_id
Global $test_title
Global $test_refs
Global $test_auto_script_ref
Global $test_case_status
Global $test_case_owner


;#cs
Global $testrail_username = $CmdLine[1]
Global $testrail_password = $CmdLine[2]
Global $testrail_run_id = $CmdLine[3]
Global $jira_username = $CmdLine[4]
Global $jira_password = $CmdLine[5]
;#ce

#cs
Global $testrail_username = "sgriffin@janison.com"
Global $testrail_password = "Gri01ffo"
Global $testrail_run_id = 916
Global $jira_username = "sgriffin@janison.com.au"
Global $jira_password = "Gri04ffo.."
#ce

Global $main_gui = GUICreate("TRRR - " & $app_name & " - Run ID " & $testrail_run_id, 840, 360)

Global $listview = GUICtrlCreateListView("Run ID|Test ID|Test Title|Automation Script Reference|Status|Comment|Case Owner|Case Status|Issues", 10, 10, 820, 300, $LVS_SHOWSELALWAYS)
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
_SQLite_SetTimeout(-1, 10000)

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



; Startup Jira

GUICtrlSetData($status_input, "Starting the Jira connection ... ")
_JiraSetup()
_JiraDomainSet("https://janisoncls.atlassian.net")
_JiraLogin($jira_username, $jira_password)



GUICtrlSetData($progress, 0)
GUISetCursor(15, 1, $main_gui)
_GUICtrlListView_DeleteAllItems($listview)
GUICtrlSetData($status_input, "")

Local $testrail_test_case_id = ""
Local $testrail_test_case_name = ""
Local $user_dict = ObjCreate("Scripting.Dictionary")
Local $story_dict = ObjCreate("Scripting.Dictionary")
Local $epic_dict = ObjCreate("Scripting.Dictionary")
Local $story_epic_dict = ObjCreate("Scripting.Dictionary")
Local $outstanding_defect_description_dict = ObjCreate("Scripting.Dictionary")
Local $outstanding_defect_test_cases_impacted_dict = ObjCreate("Scripting.Dictionary")


GUICtrlSetData($status_input, "Getting TestRail Users List ... ")
_TestRailGetUsers()
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit
Local $user = StringRegExp($testrail_json, '(?U)"id":(\d+),.*"name":"(.*)"', 3)

for $i = 0 to (UBound($user) - 1) step 2

	$user_dict.Add($user[$i], $user[$i + 1])
Next

GUICtrlSetData($status_input, "Getting details for TestRail Run ID " & $testrail_run_id & " ... ")
_TestRailGetRun($testrail_run_id)
$run_detail = StringRegExp($testrail_json, '(?U)"suite_id":(\d+),.*"name":"(.*)".*"config":"(.*)".*"project_id":(\d+),', 3)
;FileWrite(@ScriptDir & "\fred.txt", $testrail_json)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit
$suite_id = $run_detail[0]
$run_name = $run_detail[1]
$run_config = $run_detail[2]
$project_id = $run_detail[3]


Global $testrail_run_name = $run_name & " (" & $run_config & ")"
$testrail_run_name = StringReplace($testrail_run_name, "'", "''")

GUICtrlSetData($status_input, "Getting Tests for TestRail Run ID " & $testrail_run_id & " ... ")
_TestRailGetTests($testrail_run_id)
Local $rr = StringRegExp($testrail_json, '(?U)"id":(\d+),.*"case_id":(\d+),.*"status_id":(\d+),.*"title":"(.*)".*"refs":"(.*)".*"custom_auto_script_ref":"(.*)".*"custom_test_status":(\d+),.*"custom_owner":(.*),', 3)
;FileWrite(@ScriptDir & "\fred.txt", $testrail_json)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit
GUICtrlSetData($status_input, "")

Local $sqlite_insert[UBound($rr)]

for $i = 0 to (UBound($rr) - 1) step 8

	$test_id = $rr[$i]
	$test_case_id = $rr[$i + 1]
	$test_status_id = $rr[$i + 2]
	$test_title = $rr[$i + 3]
	$test_refs = $rr[$i + 4]
	$test_auto_script_ref = $rr[$i + 5]
	$test_case_status = $rr[$i + 6]
	$test_case_owner = $rr[$i + 7]


	$test_title = StringReplace($test_title, "'", "''")

	GUICtrlSetData($progress, ($i / UBound($rr)) * 100)

;	_TestRailGetCase($test_case_id)
;ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
;Exit

	; get the Jira stories and epics

	Local $story_epic_str = ""

	if $story_epic_dict.Exists($test_refs) = True Then

;		$story_epic_str = $test_refs & " [" & $story_epic_dict.Item($test_refs) & "]"
		$story_epic_str = "<tr><td><a href=""https://janisoncls.atlassian.net/browse/" & $story_epic_dict.Item($test_refs) & """ target=""_blank"">" & $story_epic_dict.Item($test_refs) & "</a></td><td><a href=""https://janisoncls.atlassian.net/browse/" & $test_refs & """ title=""" & $test_title & """ target=""_blank"">" & $test_refs & "</a></td>"
	Else

		_JiraSearchIssues("summary,issuetype,customfield_10008", "key='" & $test_refs & "'")

		if StringInStr($jira_json, """errorMessages"":") = 0 Then

			$jira = StringRegExp($jira_json, '(?U)"summary":"(.*)".*"customfield_10008":"(.*)"', 3)
	;		ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
			$story_dict.Add($test_refs, $jira[0])
			$story_epic_dict.Add($test_refs, $jira[1])
;			$story_epic_str = $test_refs & " [" & $jira[1] & "]"
			$story_epic_str = "<tr><td><a href=""https://janisoncls.atlassian.net/browse/" & $jira[1] & """ target=""_blank"">" & $jira[1] & "</a></td><td><a href=""https://janisoncls.atlassian.net/browse/" & $test_refs & """ title=""" & $test_title & """ target=""_blank"">" & $test_refs & "</a></td>"
		EndIf
	EndIf

	; get the bugs linked to the Jira story

	Local $story_epic_bug_str = $story_epic_str

;	_JiraSearchIssues("summary,issuetype", "issuetype=Bug AND status != Done AND issue in linkedIssues(""" & $test_refs & """)")
	_JiraSearchIssues("summary,issuetype,priority", "issuetype=Bug AND status != Done AND issue in linkedIssues(""" & $test_refs & """)")
;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	if StringInStr($jira_json, ",""total"":0,") = 0 Then

		$jira = StringRegExp($jira_json, '(?U)"key":"(.*)".*"summary":"(.*)","issuetype".*"priority":{.*"name":"(.*)"', 3)

		for $j = 0 to (UBound($jira) - 1) Step 3

			if $j > 0 Then

				$story_epic_bug_str = $story_epic_bug_str & $story_epic_str
			EndIf

			$jira[$j + 1] = StringReplace($jira[$j + 1], "\""", "&#34;")
			$jira[$j + 1] = StringReplace($jira[$j + 1], "'", "&#39;")

			$story_epic_bug_str = $story_epic_bug_str & "<td style=""background-color:red""><a href=""https://janisoncls.atlassian.net/browse/" & $jira[$j] & """ title=""" & $jira[$j + 1] & """ style=""color:white"" target=""_blank"">" & $jira[$j] & "</a></td></tr>"

			_SQLite_Exec(-1, "BEGIN TRANSACTION;")
	;		_SQLite_Exec(-1, "INSERT INTO defects_in_tests(TestCaseID,BugID) VALUES ('" & $test_case_id & "','" & $jira[0] & "');")
			_SQLite_Exec(-1, "INSERT INTO defect(BugID,BugSummary,Priority,TestCaseEpicStory,Impact,ActionRequired,FixDate,FixPhase) VALUES ('<a href=""https://janisoncls.atlassian.net/browse/" & $jira[$j] & """ target=""_blank"">" & $jira[$j] & "</a>','" & $jira[$j + 1] & "','" & $jira[$j + 2] & "','<tr><td><a href=""https://janison.testrail.com/index.php?/cases/view/" & $test_case_id & """ target=""_blank"">" & $test_case_id & "</a></td><td style=""width:75px""><a href=""https://janisoncls.atlassian.net/browse/" & $story_epic_dict.Item($test_refs) & """ target=""_blank"">" & $story_epic_dict.Item($test_refs) & "</a></td><td><a href=""https://janisoncls.atlassian.net/browse/" & $test_refs & """ target=""_blank"">" & $test_refs & "</a> - " & $test_title & "</td></tr>','','','','');")
			_SQLite_Exec(-1, "COMMIT TRANSACTION;")
		Next

	Else

		$story_epic_bug_str = $story_epic_bug_str & "<td></td></tr>"
	EndIf

	if StringLen($story_epic_bug_str) > 0 Then

		$story_epic_bug_str = "<table><tbody>" & $story_epic_bug_str & "</tbody></table>"
	EndIf

;	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $story_epic_bug_str = ' & $story_epic_bug_str & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console


	Switch $test_status_id

		case "1"

			$test_status_id = "Passed"

		case "2"

			$test_status_id = "Blocked"

		case "3"

			$test_status_id = "Untested"

		case "4"

			$test_status_id = "Retest"

		case "5"

			$test_status_id = "Failed"

		case "6"

			$test_status_id = "Known Issue"

		case "7"

			$test_status_id = "Incomplete Automation"

		case "8"

			$test_status_id = "Removed from Run"

		case "9"

			$test_status_id = "In Progress"
	EndSwitch


	Switch $test_case_status

		case "0"

			$test_case_status = ""

		case "1"

			$test_case_status = "Not Started"

		case "2"

			$test_case_status = "In Progress"

		case "3"

			$test_case_status = "In Review"

		case "4"

			$test_case_status = "Complete"

		case "5"

			$test_case_status = "Blocked"
	EndSwitch

	if StringCompare($test_case_owner, "null") = 0 Then

		$test_case_owner = ""
	Else

		$test_case_owner = $user_dict.Item($test_case_owner)
	EndIf

	_TestRailGetResults($test_id)

	if StringLen($testrail_json) > 2 Then

		$ss = StringRegExp($testrail_json, '(?U)"comment":(.*),"version"', 3)
		Local $comment = $ss[0]

		if StringCompare($comment, "null") = 0 Then

			$comment = "-"
		Else

			$comment = StringTrimLeft($comment, 1)
			$comment = StringTrimRight($comment, 1)
			$comment = StringReplace($comment, "'", "''")

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

						$comment_line[$j] = StringReplace($comment_line[$j], $test_auto_script_ref, "<b>" & $test_auto_script_ref & "</b>")

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

;		FileWrite(@ScriptDir & "\" & $test_id & ".txt", "INSERT INTO report(RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails) VALUES ('" & $testrail_run_id & "','" & $testrail_run_name & "','" & $test_id & "','" & $test_status_id & "','<b>" & $test_title & "</b>','" & $test_case_id & "','" & $comment & "');")
		GUICtrlCreateListViewItem($testrail_run_id & "|" & $test_id & "|" & $test_title & "|" & $test_auto_script_ref & "|" & $test_status_id & "|" & $comment & "|" & $test_case_owner & "|" & $test_case_status & "|" & $story_epic_bug_str, $listview)
		$sqlite_insert[$i] = "INSERT INTO report(RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails,TestCaseID,TestCaseOwner,TestCaseStatus,Issues) VALUES ('" & $testrail_run_id & "','" & $testrail_run_name & "','" & $test_id & "','" & $test_title & "','<b>" & $test_auto_script_ref & "</b>','" & $test_status_id & "','" & $comment & "','" & $test_case_id & "','" & $test_case_owner & "','" & $test_case_status & "','" & $story_epic_bug_str & "');"
	Else

;		FileWrite(@ScriptDir & "\" & $test_id & ".txt", "INSERT INTO report(RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails) VALUES ('" & $testrail_run_id & "','" & $testrail_run_name & "','" & $test_id & "','" & $test_status_id & "','" & $test_title & "','<b>" & $test_case_id & "</b>','-');")
		GUICtrlCreateListViewItem($testrail_run_id & "|" & $test_id & "|" & $test_title & "|" & $test_auto_script_ref & "|" & $test_status_id & "|-|" & $test_case_owner & "|" & $test_case_status & "|" & $story_epic_bug_str, $listview)
		$sqlite_insert[$i] = "INSERT INTO report(RunID,RunName,ManualTestID,TestTitle,AutoTestID,TestResult,StepDetails,TestCaseID,TestCaseOwner,TestCaseStatus,Issues) VALUES ('" & $testrail_run_id & "','" & $testrail_run_name & "','" & $test_id & "','" & $test_title & "','" & $test_auto_script_ref & "','" & $test_status_id & "','-','" & $test_case_id & "','" & $test_case_owner & "','" & $test_case_status & "','" & $story_epic_bug_str & "');"
	EndIf
Next

_SQLite_Exec(-1, "BEGIN TRANSACTION;")

for $i = 0 to (UBound($rr) - 1) step 8

	_SQLite_Exec(-1, $sqlite_insert[$i])
Next

_SQLite_Exec(-1, "COMMIT TRANSACTION;")


GUICtrlSetData($status_input, "Getting Test Cases for TestRail Suite ID " & $suite_id & " ... ")
_TestRailGetCases($project_id, $suite_id)
ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $testrail_json = ' & $testrail_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
$case_detail = StringRegExp($testrail_json, '(?U)"id":(\d+),.*"title":"(.*)","section_id"', 3)
;Exit

Local $sqlite_run_suite_case_insert[UBound($case_detail)]

for $i = 0 to (UBound($case_detail) - 1) step 2

	$test_case_id = $case_detail[$i]
	$test_case_title = $case_detail[$i + 1]
	$sqlite_run_suite_case_insert[$i] = "INSERT INTO run_suite_case(RunID,RunName,SuiteID,CaseID,CaseTitle) VALUES ('" & $testrail_run_id & "','" & $run_name & "','" & $suite_id & "','" & $test_case_id & "','" & $test_case_title & "');"
Next

Local $story_epic_str = "<table><tbody>"

For $vKey In $story_epic_dict

	$story_epic_str = $story_epic_str & "<tr><td style=""width:75px""><a href=""https://janisoncls.atlassian.net/browse/" & $story_epic_dict.Item($vKey) & """ target=""_blank"">" & $story_epic_dict.Item($vKey) & "</a></td><td><a href=""https://janisoncls.atlassian.net/browse/" & $vKey & """ target=""_blank"">" & $vKey & "</a> - " & $story_dict.Item($vKey) & "</td></tr>"
Next

$story_epic_str = $story_epic_str & "</tbody></table>"

_SQLite_Exec(-1, "BEGIN TRANSACTION;")

for $i = 0 to (UBound($case_detail) - 1) step 2

	_SQLite_Exec(-1, $sqlite_run_suite_case_insert[$i])
Next

_SQLite_Exec(-1, "INSERT INTO run_epic_story(RunID,EpicStory) VALUES ('" & $testrail_run_id & "','" & $story_epic_str & "');")

_SQLite_Exec(-1, "COMMIT TRANSACTION;")


GUICtrlSetData($progress, 0)
GUICtrlSetData($status_input, "")
GUISetCursor(2, 0, $main_gui)
GUIDelete($main_gui)

_SQLite_Close()
_SQLite_Shutdown()

GUICtrlSetData($status_input, "Closing Jira ... ")
_JiraShutdown()
