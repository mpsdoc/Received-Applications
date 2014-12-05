'Received-Applications
'=====================
'
'Script for Received Applications
'
>>>>>> SCRIPT BEGINS HERE AND BELOW
'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - Received Applications"
start_time = timer
'
'FUNCTIONS----------------------------------------------------------------------------------------------------
'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("F:\BlueZone\BlueZone Scripts\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'The Dialog
BeginDialog received_applications, 0, 0, 386, 300, "Received Applications"
  DropListBox 35, 5, 100, 15, "CHOOSE ONE"+chr(9)+"Application"+chr(9)+"Recertification"+chr(9)+"Reapplication", contact_type_list
  DropListBox 185, 5, 70, 15, "CHOOSE ONE"+chr(9)+"in Mail"+chr(9)+"in Office"+chr(9)+"via ApplyMN"+chr(9)+"by Fax", rcvd_type_list
  DropListBox 285, 5, 70, 15, "CHOOSE ONE"+chr(9)+"client"+chr(9)+"AREP"+chr(9)+"SWKR", who_contacted_list
  EditBox 70, 35, 60, 15, phone_number
  EditBox 60, 60, 85, 15, case_number
  CheckBox 5, 80, 30, 10, "Cash", cash_check
  CheckBox 45, 80, 30, 10, "HC", HC_check
  CheckBox 85, 80, 35, 10, "SNAP", SNAP_check
  CheckBox 125, 80, 35, 10, "EMER", EMER_check
  EditBox 115, 105, 60, 15, appt_date
  EditBox 230, 105, 50, 15, appt_time
  CheckBox 10, 130, 155, 10, "Check here if you screened for expedited FS.", expedite_fs
  CheckBox 10, 150, 70, 10, "Check here if you", appt_lttr
  DropListBox 85, 145, 65, 15, "CHOOSE ONE"+chr(9)+"mailed"+chr(9)+"handed", distr_type_list
  CheckBox 10, 170, 200, 10, "Check here if expedited client refused same day interview.", same_day
  EditBox 85, 190, 250, 15, verif_rcvd
  EditBox 85, 215, 250, 15, add_comm
  EditBox 175, 240, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 270, 280, 50, 15
    CancelButton 330, 280, 50, 15
  Text 260, 10, 20, 10, "from"
  Text 160, 150, 90, 10, "appointment letter to client."
  Text 5, 65, 50, 10, "Case number: "
  Text 5, 195, 75, 10, "Verifications Received: "
  GroupBox 5, 25, 370, 30, "Optional info:"
  Text 5, 220, 75, 10, "Additional Comments: "
  Text 200, 110, 25, 10, "Time:"
  Text 95, 245, 70, 10, "Sign your case note: "
  Text 5, 110, 85, 10, "Appointment Information: "
  Text 95, 110, 20, 10, "Date:"
  Text 15, 40, 50, 10, "Phone number: "
  Text 145, 10, 35, 10, "Received"
  Text 10, 10, 25, 10, "Type"
EndDialog


'Connect to BlueZone
EMConnect ""

'Loops a PF3 back to SELF
Do
	EMSendKey "<PF3>"
	EMWaitReady 0, 0
	EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"


'Find the Maxis case number, if there is one
call MAXIS_case_number_finder(case_number)


'Runs the dialog
Dialog received_applications
If ButtonPressed = 0 then StopScript

'This loop segment requires worker select type of application, how received and from who
Do
	Do
		Do
			Do
				Dialog received_applications	
				If contact_type_list = "CHOOSE ONE" then MsgBox "You must choose a type of application."
			Loop until contact_type_list <> "CHOOSE ONE"
			If rcvd_type_list = "CHOOSE ONE" then MsgBox "You must choose how application was received"
		Loop until rcvd_type_list <> "CHOOSE ONE"
		If who_contacted_list = "CHOOSE ONE" then MsgBox "You must choose who submitted application"
	Loop until who_contacted_list <> "CHOOSE ONE"
	If appt_lttr = 1 and distr_type_list = "CHOOSE ONE" then MsgBox "You must choose how you gave appointment letter to client."
Loop until appt_lttr = 1 and distr_type_list <> "CHOOSE ONE"


'Check to make sure not locked out of MAXIS
EMSendKey "<enter>"
EMWaitReady 0,0
EMReadScreen MAXIS_check, 5,1,39
If MAXIS_check <> "MAXIS" then
	MsgBox "You are not active in MAXIS. You may be passworded out."
	StopScript
End if


'Navigates to CASE/NOTE for provided Case Number
call navigate_to_screen("CASE", "NOTE")

'PF9 to open new CASE/NOTE
PF9

'Converts Program Checkbox to text for CASE/NOTE
If cash_check = 1 then programs_applied_for = programs_applied_for & "cash, "
If HC_check = 1 then programs_applied_for = programs_applied_for & "HC, "
If SNAP_check = 1 then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_check = 1 then programs_applied_for = programs_applied_for & "emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)


'Writes to the CASE/NOTE
EMWriteScreen ">>>Received Application<<<", 4, 3
	EMWriteScreen "* " & contact_type_list & " received " & rcvd_type_list & " from " & who_contacted_list, 5, 3
EMSendKey "<newline>"
EMSendKey "<newline>"
  	If phone_number <> "" then Call write_editbox_in_case_note("Phone number", phone_number, 6)
  	If programs_applied_for <> "" then call write_editbox_in_case_note("Programs applied for", programs_applied_for, 6)
  	If verif_rcvd <> "" then call write_editbox_in_case_note("Verifications Received", verif_rcvd, 6)
  	If expedite_fs = 1 then call write_new_line_in_case_note("* Screened for Expedited SNAP. See Previous CASE/NOTE")
  	If same_day = 1 then call write_new_line_in_case_note("* Client Refused Same Day Appointment.")
  	If appt_lttr = 1 then 
		If distr_type = "CHOOSE ONE" then MsgBox "You must choose how client was notified of appointment date/time."
    		If distr_type = "mailed" then 
      		call write_new_line_in_case_note("* Appointment Letter Mailed to Client.")
    		Else 
      		call write_new_line_in_case_note ("* Appointment Letter Handed to Client in Office.")
    		End If
  	End If
	If appt_date <> "" then call write_editbox_in_case_note("Appointment Date", appt_date, 6)
  	If appt_time <> "" then call write_editbox_in_case_note("Appointment Time", appt_time, 6)
	If add_comm <> "" then call write_editbox_in_case_note("Additional Comments", add_comm, 6)
  	Call write_new_line_in_case_note("---")
  	Call write_new_line_in_case_note(worker_signature)


'Will Automatically run Expedited Screening Script if Initial Application for SNAP
If SNAP_check = 1 then
  If contact_type_list = "Application" then
    Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
    Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\NOTE - expedited screening.vbs")
    text_from_the_other_script = fso_command.ReadAll
    fso_command.Close
    Execute text_from_the_other_script
  End If
End If


'Script Ends
script_end_procedure("")

