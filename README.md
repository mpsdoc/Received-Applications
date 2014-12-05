'Received-Applications
'=====================
'
'Script for Received Applications
'
>>>>>> SCRIPT BEGINS HERE AND BELOW
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
