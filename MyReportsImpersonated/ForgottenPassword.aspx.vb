Imports System.Data

Partial Class _ForgottenPassword
	Inherits System.Web.UI.Page

	Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
		'     <add name="BIWindowsConnectionString" connectionString="Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;application name=WHUA"/>

		'Initial catalog=BI;data source=10.11.24.21;Integrated Security=SSPI;persist security info=True;
		'Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;Persist security info=True;application name=WHUA


		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' IF THE WEBSITE KEEP PROMPTING YOU FOR WINDOWS CREDENTIALS, THEN:
		'Authentication > Enable Windows Authentication, then Right-Click to set the Providers.
		'----->>>>NTLM needs to be FIRST!
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------




		lblStatus.Text = ""

		'Dim stest As String = "123456789"
		'lblStatus.Text = stest.Substring(0, 4) & "<br />"
		'lblStatus.Text &= stest.Substring(4)

		lblWindowsUsername.Text = HttpContext.Current.User.Identity.Name


		If Me.IsPostBack = False Then

			txtVerificationCode.Text = ""

			lnkcmdVerifyMobile.Visible = False
			lnkcmdVerifyEmail.Visible = False
			lnkcmdReSendVerifyCodeMobile.Visible = False
			lnkcmdReSendVerifyCodeEmail.Visible = False

			lnkcmdVerifyMobile.Text = "Verify"
			lnkcmdVerifyEmail.Text = "Verify"

			lnkcmdSendPasswordToMobile.Visible = False
			lnkcmdSendPasswordToEmail.Visible = False

			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = False

			lnkcmdCancelEdit.Visible = False
			ucContactDetails1.Visible = False

		End If

		Try

			Dim bVerifyRequired As Boolean = False

			' get the phone/email and then display HR or haveDetails panels accodingly
			Dim sGetError As String = ""
			Dim sDBPhone As String = ""
			Dim sDBEmail As String = ""

			' 1. check if the phone/email has been entered
			' 2. shiw verify link
			sGetError = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)




		Catch ex As Exception
		End Try




	End Sub


	Public Function GetPhoneAndEmail(ByRef sDBPhone As String, ByRef sDBEmail As String, ByRef bVerifyRequired As Boolean) As String

		lnkcmdVerifyMobile.Visible = False
		lnkcmdVerifyEmail.Visible = False
		lnkcmdReSendVerifyCodeMobile.Visible = False
		lnkcmdReSendVerifyCodeEmail.Visible = False

		lnkcmdSendPasswordToMobile.Visible = False
		lnkcmdSendPasswordToEmail.Visible = False


		SaveToLog("GetPhoneAndEmail; lblWindowsUsername.Text=" & lblWindowsUsername.Text)

		Dim sGetError As String = ""
		Using lib1 As New clsDB(BIWindowsConnectionString)
			' 1. get Stored info for comparison
			sGetError = lib1.GetPhoneOrEmail(lblWindowsUsername.Text, sDBPhone, sDBEmail)
		End Using

		lblMobile.Text = sDBPhone
		lblEmail.Text = sDBEmail

		ucContactDetails1.ExistingMobileNumber = lblMobile.Text
		ucContactDetails1.ExistingEmailAddress = lblEmail.Text
		ucContactDetails1.WindowsUsername = lblWindowsUsername.Text

		If lblMobile.Text.Trim = "" And lblEmail.Text.Trim = "" Then
			litContactIntro.Text = "There are no contact details entered: Please contact HR and provide them with either a mobile number or personal email address. Then refresh this screen"
			cmdEditDetails.Visible = False
			Session("action") = "NewPasswordStep1"
		Else
			litContactIntro.Text = "If you <b>forget your password</b> or <b>request a new one</b>, you will need to provide a mobile number or personal email address, where we can send a temporary password to. You will need to verify these first by clicking the &quot;Verify&quot; link below. <br /><br />A <b>Verify</b> link will be shown if you need to verify the phone or email. When you click the link, it will send a code to your mobile or email which you will need to enter in the next step to complete verification. If you see a <b>ReSend Code</b> link, then you will be able to resend the Verification Code. Once verified, a &quot;Send Password&quot; link will appear, which you can use to send a temporary Password to your phone or email.<br /><br />The <b>Send Password</b> link will send a temporary Password to either your phone or email address. You will be asked to change it when you first login. <br /><br /><b>NOTE</b>: It can take several minutes for an SMS message to be sent to your phone. We also recommend that you enter both a mobile and personal email address, so you have a choice of where to send a temporary password."
			'Session("action") = "EditContactDetails"
			''pnlNewPasswordStep2.Visible = False
			'ucContactDetails1.Visible = True
			lnkcmdCancelEdit.Visible = True
		End If


		Dim sSession_EmailFinalVerificationCode As String = ""
		Dim sSession_PhoneFinalVerificationCode As String = ""
		Try
			If HttpContext.Current.Session("EmailFinalVerificationCode") IsNot Nothing Then
				sSession_EmailFinalVerificationCode = HttpContext.Current.Session("EmailFinalVerificationCode")
			End If
		Catch ex As Exception
		End Try
		Try
			If HttpContext.Current.Session("PhoneFinalVerificationCode") IsNot Nothing Then
				sSession_PhoneFinalVerificationCode = HttpContext.Current.Session("PhoneFinalVerificationCode")
			End If
		Catch ex As Exception
		End Try

		If sGetError.Trim <> "" Then
			lblStatus.Text = sGetError
		Else

			If sDBPhone.Trim <> "" Or sDBEmail.Trim <> "" Then

				Dim outPhoneCode As String = ""
				Dim outEmailCode As String = ""
				Using lib1 As New clsDB(BIWindowsConnectionString)
					' 1. get Stored info for comparison
					sGetError = lib1.GetVerificationCodes(lblWindowsUsername.Text, outPhoneCode, outEmailCode)
				End Using

				If sDBPhone.Trim <> "" Then
					' there is a saved phone
					bVerifyRequired = True ' show contact details because we have a phone OR email
					If outPhoneCode.Trim <> "" Then
						lnkcmdVerifyMobile.Visible = True

						If sSession_PhoneFinalVerificationCode.Trim <> "" Then
							' there is a code in the session: 
							' 1. show the resend button
							' 2. change the text for the original Verify link to Enter Code instead
							lnkcmdReSendVerifyCodeMobile.Visible = True
							lnkcmdVerifyMobile.Text = "Enter Code"
						End If

					Else
						lnkcmdSendPasswordToMobile.Visible = True
					End If
				End If

				If sDBEmail.Trim <> "" Then
					' there is a saved email
					bVerifyRequired = True ' show contact details because we have a phone OR email
					If outEmailCode.Trim <> "" Then
						lnkcmdVerifyEmail.Visible = True

						If sSession_EmailFinalVerificationCode.Trim <> "" Then
							' there is a code in the session: 
							' 1. show the resend button
							' 2. change the text for the original Verify link to Enter Code instead
							lnkcmdReSendVerifyCodeEmail.Visible = True
							lnkcmdVerifyEmail.Text = "Enter Code"
						End If

					Else
						lnkcmdSendPasswordToEmail.Visible = True
					End If
				End If

			Else
			End If

		End If

		' if we have a temporary passwrod in sessions, change teh Send Password text to be "Re-Send Password"
		Dim sSessionSendNewPassword As String = ""
		Try
			If Session("SendNewPassword") IsNot Nothing Then
				sSessionSendNewPassword = Session("SendNewPassword")
			End If
		Catch ex As Exception
		End Try

		If sSessionSendNewPassword.Trim <> "" Then
			lnkcmdSendPasswordToEmail.Text = "ReSend password"
			lnkcmdSendPasswordToMobile.Text = "ReSend password"
		Else
			lnkcmdSendPasswordToEmail.Text = "Send password"
			lnkcmdSendPasswordToMobile.Text = "Send password"
		End If


		Try
			If Session("action") = "NewPasswordStep3updated" Then
				' this is coming from the EDIT CONTACTS details screen and the updates need to be verified
				Session("action") = "NewPasswordStep3"
				' don't go to verify screen - need to go to check verificaton code screen

			Else
				' this can happen regardless of whether it is a new password, forgotten password or account info
				If Session("action") = "EditContactDetails" Then
					' don;t reset the session if we are editing the contact
				Else
					If bVerifyRequired = True Then
						Session("action") = "NewPasswordStep2"
					End If
				End If

			End If
		Catch ex As Exception

		End Try


		If Session("action") = "EditContactDetails" Then

			lblHeading.Text = "Edit contact details"
			pnlNewPasswordStep2.Visible = False
			ucContactDetails1.Visible = True
			lnkcmdCancelEdit.Visible = True

		ElseIf Session("action") = "NewPasswordStep1" Then
			' call HR and refresh
			lblHeading.Text = "Request New Password"

			pnlNewPasswordStep1.Visible = True
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = False


		ElseIf Session("action") = "NewPasswordStep2" Then
			' show verify and send code
			lblHeading.Text = "Contact details"

			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = True
			pnlNewPasswordStep3.Visible = False

		ElseIf Session("action") = "NewPasswordStep3" Then
			' enter the code
			lblHeading.Text = "Enter verification code"

			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = True

		ElseIf Session("action") = "NewPasswordStep4" Then
			' force them to enter new password
			lblHeading.Text = "Create New password"

			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = False


		Else
			' show verify and send code
			lblHeading.Text = "Contact details"

			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = True
			pnlNewPasswordStep3.Visible = False

		End If



		Return sGetError
	End Function

	Protected Function UpdatePhoneOrEmail(
		ByVal PhoneEntered As String,
		ByVal EmailEntered As String
		) As String

		Dim sReturn As String = ""

		' save the details and if they have chnaged, then go to the verify page
		Dim sGetError As String = ""

		Dim sDBPhone As String = ""
		Dim sDBEmail As String = ""

		' 1. check if the phone/email has changed
		' 2. update the email/phone
		' 3. verify is either has changed
		Using lib1 As New clsDB(BIWindowsConnectionString)
			' 1. get Stored info for comparison
			sGetError = lib1.GetPhoneOrEmail(lblWindowsUsername.Text, sDBPhone, sDBEmail)
		End Using

		If sGetError.Trim <> "" Then
			sReturn = sGetError
		Else

			' now compare mobile and email
			If PhoneEntered.Trim <> "" Then
				If sDBPhone.Trim <> "" Then
					If sDBPhone.Trim.Replace(" ", "") <> PhoneEntered.Trim.Replace(" ", "") Then
						' entered phone is different from stored phone
						' save it and then redirect to verify page

						' 1. SAVE it without spaces and with a +
						' 2. get verification code
						' 3. Send SMS and 
						' 4. redirect to code entry page

						' 1. SAVE
						Using lib1 As New clsDB(BIWindowsConnectionString)
							' 1. get Stored info for comparison
							sGetError = lib1.SaveUserPhone(lblWindowsUsername.Text, PhoneEntered.Trim.Replace(" ", ""))
						End Using

						If sGetError.Trim <> "" Then
							sReturn = sGetError
						Else

							' 2. get verification code and send
							Session("PhoneFinalVerificationCode") = ""
							SendVerificationSMS(PhoneEntered.Trim.Replace(" ", ""), lblWindowsUsername.Text, True)

						End If


					Else
						' not changed
						sReturn = "Mobile number already saved"
					End If

				End If
			End If

			If EmailEntered.Trim <> "" Then
				If sDBEmail.Trim <> "" Then
					If sDBEmail.Trim.Replace(" ", "") <> EmailEntered.Trim Then
						' entered phone is different from stored phone
						' save it and then redirect to verify page

						' 1. SAVE it without spaces and with a +
						' 2. get verification code
						' 3. Send emil and 
						' 4. redirect to code entry page

						' 1. SAVE
						Using lib1 As New clsDB(BIWindowsConnectionString)
							' 1. get Stored info for comparison
							sGetError = lib1.SaveUserEmail(lblWindowsUsername.Text, EmailEntered.Trim)
						End Using

						If sGetError.Trim <> "" Then
							sReturn = sGetError
						Else

							' 2. get verification code and send
							Session("EmailFinalVerificationCode") = ""
							SendVerificationEmail(EmailEntered.Trim, lblWindowsUsername.Text, True)

						End If

					Else
						' not changed
						sReturn = "Mobile saved successfully"
					End If

				End If
			End If


		End If

		Return sReturn
	End Function







	Protected Sub lnkcmdVerifyMobile_Click(sender As Object, e As EventArgs) Handles lnkcmdVerifyMobile.Click
		' go to the relevant verify page
		txtVerificationCode.Text = ""
		Dim sSMSError As String = SendVerificationSMS(lblMobile.Text.Trim, lblWindowsUsername.Text, False)
		' 4. redirect to code entry page
		If sSMSError.Trim <> "" Then

			If sSMSError.Trim = "ENTER_VERIFICATION_CODE" Then

				lblHeading.Text = "Verify code"
				Session("action") = "NewPasswordStep3"
				pnlNewPasswordStep1.Visible = False
				pnlNewPasswordStep2.Visible = False
				pnlNewPasswordStep3.Visible = True

			Else
				lblStatus.Text = sSMSError.Trim
			End If


		Else
			lblHeading.Text = "Verify code"
			Session("action") = "NewPasswordStep3"
			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = True

		End If

	End Sub


	Protected Sub lnkcmdVerifyEmail_Click(sender As Object, e As EventArgs) Handles lnkcmdVerifyEmail.Click
		' go to the relevant verify page
		txtVerificationCode.Text = ""
		Dim sSMSError As String = SendVerificationEmail(lblEmail.Text.Trim, lblWindowsUsername.Text, False)
		' 4. redirect to code entry page
		If sSMSError.Trim <> "" Then
			If sSMSError.Trim = "ENTER_VERIFICATION_CODE" Then

				lblHeading.Text = "Verify code"
				Session("action") = "NewPasswordStep3"
				pnlNewPasswordStep1.Visible = False
				pnlNewPasswordStep2.Visible = False
				pnlNewPasswordStep3.Visible = True

			Else
				lblStatus.Text = sSMSError.Trim
			End If
		Else

			lblHeading.Text = "Verify code"
			Session("action") = "NewPasswordStep3"
			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = True

		End If
	End Sub


	Protected Sub lnkcmdReSendVerifyCodeMobile_Click(sender As Object, e As EventArgs) Handles lnkcmdReSendVerifyCodeMobile.Click
		' resend the verify code and go to enter code screen
		txtVerificationCode.Text = ""
		Dim sSMSError As String = SendVerificationSMS(lblMobile.Text.Trim, lblWindowsUsername.Text, True)
		' 4. redirect to code entry page
		If sSMSError.Trim <> "" Then

			If sSMSError.Trim = "ENTER_VERIFICATION_CODE" Then

				lblHeading.Text = "Verify code"
				Session("action") = "NewPasswordStep3"
				pnlNewPasswordStep1.Visible = False
				pnlNewPasswordStep2.Visible = False
				pnlNewPasswordStep3.Visible = True

			Else
				lblStatus.Text = sSMSError.Trim
			End If


		Else
			lblHeading.Text = "Verify code"
			Session("action") = "NewPasswordStep3"
			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = True

		End If
	End Sub
	Protected Sub lnkcmdReSendVerifyCodeEmail_Click(sender As Object, e As EventArgs) Handles lnkcmdReSendVerifyCodeEmail.Click
		' resend the verify code and go to enter code screen
		txtVerificationCode.Text = ""
		Dim sSMSError As String = SendVerificationEmail(lblEmail.Text.Trim, lblWindowsUsername.Text, True)
		' 4. redirect to code entry page
		If sSMSError.Trim <> "" Then
			If sSMSError.Trim = "ENTER_VERIFICATION_CODE" Then

				lblHeading.Text = "Verify code"
				Session("action") = "NewPasswordStep3"
				pnlNewPasswordStep1.Visible = False
				pnlNewPasswordStep2.Visible = False
				pnlNewPasswordStep3.Visible = True

			Else
				lblStatus.Text = sSMSError.Trim
			End If
		Else

			lblHeading.Text = "Verify code"
			Session("action") = "NewPasswordStep3"
			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = True

		End If
	End Sub


	'Protected Sub lnkGetNewPassword_Click(sender As Object, e As EventArgs) Handles lnkGetNewPassword.Click

	'    ' ----------------------- test ------------------------
	'    'Dim sError1 As String = SendSMS("07910030684@evotext.net", "Hi Stu")
	'    ''Dim sError As String = SendSMS("07910030684@evotext.net", "test")
	'    'Dim sError1 As String = SendSMS2("07507342793@evotext.net", "testrr")

	'    'If sError1.Trim <> "" Then
	'    '    lblStatus.Text &= "With Broadlook: " & sError1
	'    'Else
	'    '    lblStatus.Text &= "With Broadlook: " & "SMS sent with code to login - must be changed on first login" & "<br /><br />"
	'    'End If

	'    'Dim sError2 As String = SendSMS("07507342793@evotext.net", "Hi Me")

	'    'If sError2.Trim <> "" Then
	'    '    lblStatus.Text &= "Without Broadlook: " & sError2
	'    'Else
	'    '    lblStatus.Text &= "Without Broadlook: " & "SMS sent with code to login - must be changed on first login" & "<br /><br />"
	'    'End If

	'    'Dim sError As String = SendSMS2("sanjay.patel@evolutionjobs.co.uk", "test")
	'    ' ----------------------- test ------------------------




	'    Dim dtCons As New System.Data.DataTable
	'    Using lib1 As New clsMyReports(BIWindowsConnectionString)
	'        dtCons = lib1.GetNewPasswordForAUser("Evolution2", 517)
	'    End Using

	'    Dim sCode As String = ""
	'    If dtCons IsNot Nothing Then

	'        Dim iCount As Integer = 0
	'        Dim row As DataRow
	'        For Each row In dtCons.Rows

	'            If Not IsDBNull(row(0)) Then
	'                sCode = row(0)
	'                Exit For
	'            End If
	'            iCount += 1
	'        Next

	'        ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'        ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'        ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'        ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'        ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'        Dim sError As String = SendSMS2("07507342793@evotext.net", "test")
	'        'Dim sError As String = SendSMS("sanjay.patel@evolutionjobs.co.uk", "test")
	'        'Dim sError As String = SendSMS("07507342793@evotext.net", sCode)
	'        'SaveTextToFile("", "DEVELOPMENT - s_SendSMS SENT TO DEV TEL")

	'        If sError.Trim <> "" Then
	'            lblStatus.Text = sError
	'        Else
	'            lblStatus.Text = "SMS sent with code to login - must be changed on first login"
	'        End If

	'    End If


	'End Sub






	Protected Sub cmdVerifyCode_Click(sender As Object, e As EventArgs) Handles cmdVerifyCode.Click

		' SPLIT OUT THE VERIFICATION CODE AND PASSWORD; FIRST 4 DIGITS=VERIFICATIONCODE AND NEXT 4 DIGITS = NEW PASSWORD
		' E.G. 79AB9898: 79AB = VERIFICATION CODE AND 8989 = NEW PASSWORD
		Dim sVerificationCode As String = ""
		Dim sNewPassword As String = ""
		If txtVerificationCode.Text.Trim <> "" Then

			' check to see if the verification code is 4 chars (without pass) or 8 chars (with password) long
			If txtVerificationCode.Text.Trim.Length = 4 Then
				sVerificationCode = txtVerificationCode.Text

			ElseIf txtVerificationCode.Text.Trim.Length = 8 Then
				' we are saving the 8 character code
				sVerificationCode = txtVerificationCode.Text.Trim '.Substring(0, 4)
				lblTemp.Text = txtVerificationCode.Text.Trim.Substring(4) ' new password any letters from index = 4 to end



			Else
				lblStatus.Text = "ERROR: Please enter a valid verification code"
				sVerificationCode = "" ' clear it out so it exits
			End If



			If sVerificationCode.Trim <> "" Then

				Dim sGetError As String = ""
				Dim outPhoneCode As String = ""
				Dim outEmailCode As String = ""
				Using lib1 As New clsDB(BIWindowsConnectionString)
					' 1. get Stored info for comparison
					sGetError = lib1.GetVerificationCodes(lblWindowsUsername.Text, outPhoneCode, outEmailCode)
				End Using

				If sGetError.Trim <> "" Then
					lblStatus.Text = sGetError
				Else

					If sVerificationCode.Trim = outPhoneCode.Trim Then
						' OK = Email Verified or Phone Verified or Email Verification Failed or Phone Verification Failed
						Using lib1 As New clsDB(BIWindowsConnectionString)
							' 1. get Stored info for comparison
							sGetError = lib1.CheckVerificationCode(lblWindowsUsername.Text, sVerificationCode.Trim, "")
						End Using

						If sGetError.Trim.ToLower.Contains(("Failed").ToLower) Then
							' failed - return error
							lblStatus.Text = "ERROR: " & sGetError
						Else
							' OK: check session to see if they clicked the forgotten password and if they did, send SMS with new password
							' if they didn't, then redirect them to the login page


							lblStatus.Text = "Phone verified successfully"

							HttpContext.Current.Session("PhoneFinalVerificationCode") = ""


							' refresh the details
							' get the phone/email and then display HR or haveDetails panels accodingly
							Dim bVerifyRequired As Boolean = False
							Dim sGetErrorx As String = ""
							Dim sDBPhone As String = ""
							Dim sDBEmail As String = ""

							' 1. check if the phone/email has been entered
							' 2. shiw verify link
							sGetErrorx = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)


							If txtVerificationCode.Text.Trim.Length = 8 Then
								Session("NewPassword") = lblTemp.Text
								Session("action") = "TemporaryPassword"
								' redirect to login page and force them to change their password
								Dim sFromManageMyAccount As String = ""
								Dim sRedirectURL As String = "login.aspx?cp=1"
								Try
									If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
										sRedirectURL = Session("AutoLoginLinkChangePassword")
									End If
									If Session("LoginPage") IsNot Nothing Then
										sFromManageMyAccount = Session("LoginPage")
									End If
								Catch ex As Exception
									sRedirectURL = "login.aspx?cp=1"
								End Try

								If sFromManageMyAccount = "From_Manage_My_Account" Then
									' we came from the manage my account link, which means we don;t need to change password, so
									' clear the sessions for changing passord and just go back to forgot my page list
									Session("NewPassword") = ""
									Session("action") = "NewPasswordStep2"

									' show verify and send code
									lblHeading.Text = "Contact details"

									pnlNewPasswordStep1.Visible = False
									pnlNewPasswordStep2.Visible = True
									pnlNewPasswordStep3.Visible = False

								Else
									Session("ChangePasswordMessage") = ChangePasswordMessage
									Response.Redirect(sRedirectURL)
								End If

							Else
								Session("action") = "NewPasswordStep2"
								' show verify and send code
								lblHeading.Text = "Contact details"

								pnlNewPasswordStep1.Visible = False
								pnlNewPasswordStep2.Visible = True
								pnlNewPasswordStep3.Visible = False
							End If



						End If

					ElseIf sVerificationCode.Trim = outEmailCode.Trim Then
						' OK = Email Verified or Phone Verified or Email Verification Failed or Phone Verification Failed
						Using lib1 As New clsDB(BIWindowsConnectionString)
							' 1. get Stored info for comparison
							sGetError = lib1.CheckVerificationCode(lblWindowsUsername.Text, "", sVerificationCode.Trim)
						End Using

						If sGetError.Trim.ToLower.Contains(("Failed").ToLower) Then
							' failed - return error
							lblStatus.Text = "ERROR: " & sGetError
						Else
							' OK: check session to see if they clicked the forgotten password and if they did, send SMS with new password
							' if they didn't, then redirect them to the login page

							lblStatus.Text = "Email verified successfully"

							HttpContext.Current.Session("EmailFinalVerificationCode") = ""

							' refresh the details
							' get the phone/email and then display HR or haveDetails panels accodingly
							Dim bVerifyRequired As Boolean = False
							Dim sGetErrorx As String = ""
							Dim sDBPhone As String = ""
							Dim sDBEmail As String = ""

							' 1. check if the phone/email has been entered
							' 2. shiw verify link
							sGetErrorx = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)

							If txtVerificationCode.Text.Trim.Length = 8 Then
								Session("NewPassword") = lblTemp.Text
								Session("action") = "TemporaryPassword"
								' redirect to login page and force them to change their password
								Dim sFromManageMyAccount As String = ""
								Dim sRedirectURL As String = "login.aspx?cp=1"
								Try
									If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
										sRedirectURL = Session("AutoLoginLinkChangePassword")
									End If
									If Session("LoginPage") IsNot Nothing Then
										sFromManageMyAccount = Session("LoginPage")
									End If
								Catch ex As Exception
									sRedirectURL = "login.aspx?cp=1"
								End Try


								If sFromManageMyAccount = "From_Manage_My_Account" Then
									' we came from the manage my account link, which means we don;t need to change password, so
									' clear the sessions for changing passord and just go back to forgot my page list
									Session("NewPassword") = ""
									Session("action") = "NewPasswordStep2"

									' show verify and send code
									lblHeading.Text = "Contact details"

									pnlNewPasswordStep1.Visible = False
									pnlNewPasswordStep2.Visible = True
									pnlNewPasswordStep3.Visible = False

								Else
									Session("ChangePasswordMessage") = ChangePasswordMessage
									Response.Redirect(sRedirectURL)
								End If


							Else
								Session("action") = "NewPasswordStep2"
								' show verify and send code
								lblHeading.Text = "Contact details"

								pnlNewPasswordStep1.Visible = False
								pnlNewPasswordStep2.Visible = True
								pnlNewPasswordStep3.Visible = False
							End If






						End If

					Else

						' error - 
						lblStatus.Text = "ERROR: Wrong code entered"
					End If


				End If




			Else
				lblStatus.Text = "ERROR: Please enter a verification code"


			End If




		Else
			lblStatus.Text = "ERROR: Please enter a verification code"
		End If

		If lblStatus.Text.Contains("ERROR") Then
			' allow them to retry entering another code

			' enter the code
			lblHeading.Text = "Enter verification code"

			pnlNewPasswordStep1.Visible = False
			pnlNewPasswordStep2.Visible = False
			pnlNewPasswordStep3.Visible = True
		End If


	End Sub



	'Protected Sub cmdVerifyCode_Click(sender As Object, e As EventArgs) Handles cmdVerifyCode.Click

	'	' SPLIT OUT THE VERIFICATION CODE AND PASSWORD; FIRST 4 DIGITS=VERIFICATIONCODE AND NEXT 4 DIGITS = NEW PASSWORD
	'	' E.G. 79AB9898: 79AB = VERIFICATION CODE AND 8989 = NEW PASSWORD
	'	Dim sVerificationCode As String = ""
	'	Dim sNewPassword As String = ""
	'	If txtVerificationCode.Text.Trim <> "" Then

	'		' check to see if the verification code is 4 chars (without pass) or 8 chars (with password) long
	'		If txtVerificationCode.Text.Trim.Length = 4 Then
	'			sVerificationCode = txtVerificationCode.Text

	'		ElseIf txtVerificationCode.Text.Trim.Length = 8 Then
	'			sVerificationCode = txtVerificationCode.Text.Trim.Substring(0, 4)
	'			lblTemp.Text = txtVerificationCode.Text.Trim.Substring(4) ' new password any letters from index = 4 to end



	'		Else
	'			lblStatus.Text = "ERROR: Please enter a valid verification code"
	'			Exit Sub
	'		End If



	'		If sVerificationCode.Trim <> "" Then

	'			Dim sGetError As String = ""
	'			Dim outPhoneCode As String = ""
	'			Dim outEmailCode As String = ""
	'			Using lib1 As New clsDB(BIWindowsConnectionString)
	'				' 1. get Stored info for comparison
	'				sGetError = lib1.GetVerificationCodes(lblWindowsUsername.Text, outPhoneCode, outEmailCode)
	'			End Using

	'			If sGetError.Trim <> "" Then
	'				lblStatus.Text = sGetError
	'			Else

	'				If sVerificationCode.Trim = outPhoneCode.Trim Then
	'					' OK = Email Verified or Phone Verified or Email Verification Failed or Phone Verification Failed
	'					Using lib1 As New clsDB(BIWindowsConnectionString)
	'						' 1. get Stored info for comparison
	'						sGetError = lib1.CheckVerificationCode(lblWindowsUsername.Text, sVerificationCode.Trim, "")
	'					End Using

	'					If sGetError.Trim.ToLower.Contains(("Failed").ToLower) Then
	'						' failed - return error
	'						lblStatus.Text = "ERROR: " & sGetError
	'					Else
	'						' OK: check session to see if they clicked the forgotten password and if they did, send SMS with new password
	'						' if they didn't, then redirect them to the login page


	'						lblStatus.Text = "Phone verified successfully"

	'						' refresh the details
	'						' get the phone/email and then display HR or haveDetails panels accodingly
	'						Dim bVerifyRequired As Boolean = False
	'						Dim sGetErrorx As String = ""
	'						Dim sDBPhone As String = ""
	'						Dim sDBEmail As String = ""

	'						' 1. check if the phone/email has been entered
	'						' 2. shiw verify link
	'						sGetErrorx = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)


	'						If txtVerificationCode.Text.Trim.Length = 8 Then
	'							Session("NewPassword") = lblTemp.Text
	'							Session("action") = "TemporaryPassword"
	'							' redirect to login page and force them to change their password
	'							Dim sFromManageMyAccount As String = ""
	'							Dim sRedirectURL As String = "login.aspx?cp=1"
	'							Try
	'								If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
	'									sRedirectURL = Session("AutoLoginLinkChangePassword")
	'								End If
	'								If Session("LoginPage") IsNot Nothing Then
	'									sFromManageMyAccount = Session("LoginPage")
	'								End If
	'							Catch ex As Exception
	'								sRedirectURL = "login.aspx?cp=1"
	'							End Try

	'							If sFromManageMyAccount = "From_Manage_My_Account" Then
	'								' we came from the manage my account link, which means we don;t need to change password, so
	'								' clear the sessions for changing passord and just go back to forgot my page list
	'								Session("NewPassword") = ""
	'								Session("action") = "NewPasswordStep2"

	'								' show verify and send code
	'								lblHeading.Text = "Contact details"

	'								pnlNewPasswordStep1.Visible = False
	'								pnlNewPasswordStep2.Visible = True
	'								pnlNewPasswordStep3.Visible = False

	'							Else
	'								Session("ChangePasswordMessage") = ChangePasswordMessage
	'								Response.Redirect(sRedirectURL)
	'							End If

	'						Else
	'							Session("action") = "NewPasswordStep2"
	'							' show verify and send code
	'							lblHeading.Text = "Contact details"

	'							pnlNewPasswordStep1.Visible = False
	'							pnlNewPasswordStep2.Visible = True
	'							pnlNewPasswordStep3.Visible = False
	'						End If



	'					End If

	'				ElseIf sVerificationCode.Trim = outEmailCode.Trim Then
	'					' OK = Email Verified or Phone Verified or Email Verification Failed or Phone Verification Failed
	'					Using lib1 As New clsDB(BIWindowsConnectionString)
	'						' 1. get Stored info for comparison
	'						sGetError = lib1.CheckVerificationCode(lblWindowsUsername.Text, "", sVerificationCode.Trim)
	'					End Using

	'					If sGetError.Trim.ToLower.Contains(("Failed").ToLower) Then
	'						' failed - return error
	'						lblStatus.Text = "ERROR: " & sGetError
	'					Else
	'						' OK: check session to see if they clicked the forgotten password and if they did, send SMS with new password
	'						' if they didn't, then redirect them to the login page

	'						lblStatus.Text = "Email verified successfully"

	'						' refresh the details
	'						' get the phone/email and then display HR or haveDetails panels accodingly
	'						Dim bVerifyRequired As Boolean = False
	'						Dim sGetErrorx As String = ""
	'						Dim sDBPhone As String = ""
	'						Dim sDBEmail As String = ""

	'						' 1. check if the phone/email has been entered
	'						' 2. shiw verify link
	'						sGetErrorx = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)

	'						If txtVerificationCode.Text.Trim.Length = 8 Then
	'							Session("NewPassword") = lblTemp.Text
	'							Session("action") = "TemporaryPassword"
	'							' redirect to login page and force them to change their password
	'							Dim sFromManageMyAccount As String = ""
	'							Dim sRedirectURL As String = "login.aspx?cp=1"
	'							Try
	'								If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
	'									sRedirectURL = Session("AutoLoginLinkChangePassword")
	'								End If
	'								If Session("LoginPage") IsNot Nothing Then
	'									sFromManageMyAccount = Session("LoginPage")
	'								End If
	'							Catch ex As Exception
	'								sRedirectURL = "login.aspx?cp=1"
	'							End Try


	'							If sFromManageMyAccount = "From_Manage_My_Account" Then
	'								' we came from the manage my account link, which means we don;t need to change password, so
	'								' clear the sessions for changing passord and just go back to forgot my page list
	'								Session("NewPassword") = ""
	'								Session("action") = "NewPasswordStep2"

	'								' show verify and send code
	'								lblHeading.Text = "Contact details"

	'								pnlNewPasswordStep1.Visible = False
	'								pnlNewPasswordStep2.Visible = True
	'								pnlNewPasswordStep3.Visible = False

	'							Else
	'								Session("ChangePasswordMessage") = ChangePasswordMessage
	'								Response.Redirect(sRedirectURL)
	'							End If


	'						Else
	'							Session("action") = "NewPasswordStep2"
	'							' show verify and send code
	'							lblHeading.Text = "Contact details"

	'							pnlNewPasswordStep1.Visible = False
	'							pnlNewPasswordStep2.Visible = True
	'							pnlNewPasswordStep3.Visible = False
	'						End If






	'					End If

	'				Else

	'					' error - 
	'					lblStatus.Text = "ERROR: Wrong code entered"
	'				End If


	'			End If




	'		Else
	'			lblStatus.Text = "Please enter a verification code"


	'		End If




	'	Else
	'		lblStatus.Text = "Please enter a verification code"
	'	End If




	'End Sub




	Protected Sub lnkcmdSendPasswordToMobile_Click(sender As Object, e As EventArgs) Handles lnkcmdSendPasswordToMobile.Click
		Dim sError As String = SendPasswordToSMS(lblMobile.Text)
		If sError.Trim <> "" Then
			lblStatus.Text = sError
		End If

	End Sub




	Protected Sub lnkcmdSendPasswordToEmail_Click(sender As Object, e As EventArgs) Handles lnkcmdSendPasswordToEmail.Click
		Dim sError As String = SendPasswordToEmail(lblEmail.Text)
		If sError.Trim <> "" Then
			lblStatus.Text = sError
		End If
	End Sub



	Protected Function SendPasswordToSMS(ByVal PhoneEntered As String) As String
		Dim sReturn As String = ""
		Dim sGetError As String = ""
		Dim outPassword As String = ""

		Dim sSessionSendNewPassword As String = ""
		Try
			If Session("SendNewPassword") IsNot Nothing Then
				sSessionSendNewPassword = Session("SendNewPassword")
			End If
		Catch ex As Exception
		End Try

		SaveToLog("SendPasswordToSMS; lblWindowsUsername.Text=" & lblWindowsUsername.Text & "; sSessionSendNewPassword=" & sSessionSendNewPassword)


		If sSessionSendNewPassword <> "" Then
			outPassword = sSessionSendNewPassword
		Else
			Using lib1 As New clsDB(BIWindowsConnectionString)
				' 1. get Stored info for comparison
				sGetError = lib1.GetNewPassword(lblWindowsUsername.Text, outPassword)
			End Using
		End If

		SaveToLog("SendPasswordToSMS; lblWindowsUsername.Text=" & lblWindowsUsername.Text & "; sGetError=" & sGetError & "; outPassword=" & outPassword)


		If sGetError.Trim <> "" Then
			sReturn = sGetError.Trim
		Else


			' 3. Send SMS 
			' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
			' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
			' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
			' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
			' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
			Dim sSMSMessage As String = String.Format(sSMSNewPasswordMessage, outPassword)
			Dim sSMSError As String = SendSMS(PhoneEntered.Trim.Replace(" ", "").Replace("+", "00") & SMSDomain, sSMSMessage)

			SaveToLog("SendPasswordToSMS; lblWindowsUsername.Text=" & lblWindowsUsername.Text & "; sSMSError=" & sSMSError)

			' 4. redirect to code entry page
			If sSMSError.Trim <> "" Then
				sReturn = sSMSError.Trim
			Else
				Session("action") = "TemporaryPassword"
				' redirect to login page and force them to change their password
				Dim sFromManageMyAccount As String = ""
				Dim sRedirectURL As String = "login.aspx?cp=1"
				Try
					If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
						sRedirectURL = Session("AutoLoginLinkChangePassword")
					End If
					If Session("LoginPage") IsNot Nothing Then
						sFromManageMyAccount = Session("LoginPage")
					End If
				Catch ex As Exception
					sRedirectURL = "login.aspx?cp=1"
				End Try

				Session("SendNewPassword") = outPassword

				If sFromManageMyAccount = "From_Manage_My_Account" Then
					' we came from manage my account link
					' we want to go straight to change password
					sRedirectURL = "login.aspx?cp=1"

					Session("ChangePasswordMessage") = SendPasswordMessage

					Response.Redirect(sRedirectURL)

				Else
					' normal
					Session("ChangePasswordMessage") = ChangePasswordMessage

					Response.Redirect(sRedirectURL)

				End If


			End If

		End If

		Return sReturn
	End Function


	Protected Function SendPasswordToEmail(ByVal EmailEntered As String) As String
		Dim sReturn As String = ""
		Dim sGetError As String = ""
		Dim outPassword As String = ""

		Dim sSessionSendNewPassword As String = ""
		Try
			If Session("SendNewPassword") IsNot Nothing Then
				sSessionSendNewPassword = Session("SendNewPassword")
			End If
		Catch ex As Exception
		End Try

		SaveToLog("SendPasswordToEmail; lblWindowsUsername.Text=" & lblWindowsUsername.Text & "; sSessionSendNewPassword=" & sSessionSendNewPassword)

		If sSessionSendNewPassword <> "" Then
			outPassword = sSessionSendNewPassword
		Else
			Using lib1 As New clsDB(BIWindowsConnectionString)
				' 1. get Stored info for comparison
				sGetError = lib1.GetNewPassword(lblWindowsUsername.Text, outPassword)
			End Using
		End If

		SaveToLog("SendPasswordToEmail; lblWindowsUsername.Text=" & lblWindowsUsername.Text & "; sGetError=" & sGetError & "; outPassword=" & outPassword)


		If sGetError.Trim <> "" Then
			sReturn = sGetError.Trim
		Else

			If outPassword.Trim <> "" Then

				' 3. Send email 
				Dim sSMSMessage As String = String.Format(sEmailNewPasswordMessage, outPassword)
				' sendng SMS is same as sending email!!!
				Dim sSMSError As String = SendSMS(EmailEntered.Trim, sSMSMessage)

				SaveToLog("SendPasswordToEmail; lblWindowsUsername.Text=" & lblWindowsUsername.Text & "; sSMSError=" & sSMSError)

				' 4. redirect to code entry page
				If sSMSError.Trim <> "" Then
					sReturn = sSMSError.Trim
				Else
					Session("action") = "TemporaryPassword"
					' redirect to login page and force them to change their password
					Dim sFromManageMyAccount As String = ""
					Dim sRedirectURL As String = "login.aspx?cp=1"
					Try
						If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
							sRedirectURL = Session("AutoLoginLinkChangePassword")
						End If
						If Session("LoginPage") IsNot Nothing Then
							sFromManageMyAccount = Session("LoginPage")
						End If

					Catch ex As Exception
						sRedirectURL = "login.aspx?cp=1"
					End Try


					Session("SendNewPassword") = outPassword

					If sFromManageMyAccount = "From_Manage_My_Account" Then
						' we came from manage my account link
						' we want to go straight to change password
						sRedirectURL = "login.aspx?cp=1"

						Session("ChangePasswordMessage") = SendPasswordMessage

						Response.Redirect(sRedirectURL)

					Else
						' normal
						Session("ChangePasswordMessage") = ChangePasswordMessage

						Response.Redirect(sRedirectURL)

					End If



				End If

			Else
				sReturn = "ERROR: Getting temporary password"
			End If


		End If

		Return sReturn

	End Function



	Protected Sub lnkcmdCancelVerify_Click(sender As Object, e As EventArgs) Handles lnkcmdCancelVerify.Click
		'Dim sRedirectURL As String = ""
		'Try
		'	If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
		'		sRedirectURL = Session("AutoLoginLinkChangePassword")
		'	End If
		'Catch ex As Exception
		'	sRedirectURL = "Report.aspx"
		'End Try


		Session("action") = "NewPasswordStep2"
		' show verify and send code
		lblHeading.Text = "Contact details"

		pnlNewPasswordStep1.Visible = False
		pnlNewPasswordStep2.Visible = True
		pnlNewPasswordStep3.Visible = False


	End Sub


	Protected Sub cmdEditDetails_Click(sender As Object, e As EventArgs) Handles cmdEditDetails.Click
		Session("LoginPage") = "From_Manage_My_Account"
		Session("action") = "EditContactDetails"
		lblHeading.Text = "Edit contact details"
		pnlNewPasswordStep2.Visible = False
		ucContactDetails1.Visible = True
		lnkcmdCancelEdit.Visible = True
	End Sub


	Protected Sub lnkcmdMyReports_Click(sender As Object, e As EventArgs) Handles lnkcmdMyReports.Click
		Response.Redirect("Report.aspx")
	End Sub
	Protected Sub lnkcmdLogin_Click(sender As Object, e As EventArgs) Handles lnkcmdLogin.Click

		Dim sReplaceCPString1 As String = "&cp=1"
		Dim sReplaceCPString2 As String = "?cp=1"
		Try
			If Session("SendNewPassword") IsNot Nothing Then
				If Session("SendNewPassword").trim <> "" Then
					' musy change password, so it needs the &cp=1 parameter, hence don;t remove it!!!!
					sReplaceCPString1 = ""
					sReplaceCPString2 = ""
				End If
			End If
		Catch ex As Exception
		End Try


		Dim sURL As String = ""
		Try
			If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
				sURL = Session("AutoLoginLinkChangePassword")
			End If
		Catch ex As Exception
		End Try

		If sURL.Trim = "" Then
			Dim sRedirect As String = ""
			Try
				If Session("LoginPage") <> "" Then
					sRedirect = Session("LoginPage").ToString.Replace("From_Manage_My_Account", "")
				End If
			Catch ex As Exception
				sRedirect = ""
			End Try

			If sRedirect.Trim <> "" Then
				sURL = sRedirect
			Else
				sURL = "Login.aspx"
			End If

		End If

		If sURL.Trim <> "" Then

			If sReplaceCPString1.Trim <> "" Then
				sURL = sURL.Replace(sReplaceCPString1, "")
			End If

			If sReplaceCPString2.Trim <> "" Then
				sURL = sURL.Replace(sReplaceCPString2, "")
			End If

		Else
			sURL = "Login.aspx"
		End If
		Response.Redirect(sURL)
	End Sub
	Protected Sub lnkcmdCancelEdit_Click(sender As Object, e As EventArgs) Handles lnkcmdCancelEdit.Click
		Session("action") = "NewPasswordStep2"
		lblHeading.Text = "Contact details"
		pnlNewPasswordStep2.Visible = True
		ucContactDetails1.Visible = False
		lnkcmdCancelEdit.Visible = False
	End Sub


	Protected Sub lnkcmdRefreshFromStep1_Click(sender As Object, e As EventArgs) Handles lnkcmdRefreshFromStep1.Click
		' refresh the details
		' get the phone/email and then display HR or haveDetails panels accodingly
		Dim bVerifyRequired As Boolean = False
		Dim sGetErrorx As String = ""
		Dim sDBPhone As String = ""
		Dim sDBEmail As String = ""

		' 1. check if the phone/email has been entered
		' 2. shiw verify link
		sGetErrorx = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)
	End Sub
	Protected Sub lnkcmdRefreshFromStep2_Click(sender As Object, e As EventArgs) Handles lnkcmdRefreshFromStep2.Click
		' refresh the details
		' get the phone/email and then display HR or haveDetails panels accodingly
		Dim bVerifyRequired As Boolean = False
		Dim sGetErrorx As String = ""
		Dim sDBPhone As String = ""
		Dim sDBEmail As String = ""

		' 1. check if the phone/email has been entered
		' 2. shiw verify link
		sGetErrorx = GetPhoneAndEmail(sDBPhone, sDBEmail, bVerifyRequired)
	End Sub



End Class
