Imports Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.IO
Public Module functions

	' the impersonating account - will run "execute as evolutionjob\patels" for example
	Public ImpersonatingAccount As String = "MyReportAccess" '"ed"
	Public WindowsDomain As String = "evolutionjobs"

	'Public DevelopmentEmails As String = "sanjay.patel@evolutionjobs.co.uk"
	Public DevelopmentEmails As String = "sanjay.patel@evolutionjobs.co.uk"

	Public sSMSVerificationMessage As String = "My Reports Verification Code: {0}"
	Public sEmailVerificationMessage As String = "My Reports Verification Code: {0}. Please enter this in the My Reports verification page."

	Public sSMSNewPasswordMessage As String = "My Reports temporary Password: {0}"
	Public sEmailNewPasswordMessage As String = "My Reports temporary Password: {0}. Please enter this in the My Reports login page and change it when prompted."

	Public ChangePasswordMessage As String = "Your email or phone has been verified and a temporary password has been sent. <br />Please enter a new password before proceeding."
	Public SendPasswordMessage As String = "A temporary Password has been sent to your phone or email. <br />NOTE: It can take several minutes for SMS messages to be sent to your phone.<br />Please enter it below and choose a new password before proceeding."

	Public TVPingHistoryActionFK As Long = 1797

	Public SMSDomain As String = "@sms.textapp.net" ' "@evotext.net"




	Public Property WindowsUsername() As String
		Get
			Dim sWinUsername As String = ""
			If HttpContext.Current.Session("WinUsername") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("WinUsername")) <> "" Then
					sWinUsername = HttpContext.Current.Session("WinUsername")
				End If
			End If
			Return sWinUsername
		End Get
		Set(ByVal value As String)
			HttpContext.Current.Session("WinUsername") = value
		End Set
	End Property


	Public Property WindowsPassword() As String
		Get
			Dim sWinPassword As String = ""
			If HttpContext.Current.Session("WinPassword") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("WinPassword")) <> "" Then
					sWinPassword = HttpContext.Current.Session("WinPassword")
				End If
			End If
			Return sWinPassword
		End Get
		Set(ByVal value As String)
			HttpContext.Current.Session("WinPassword") = value
		End Set
	End Property


	Public Property BIWindowsConnectionString() As String
		Get
			Dim sBIWindowsConnectionString As String = ""
			If HttpContext.Current.Session("ConnectionString") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("ConnectionString")) <> "" Then
					sBIWindowsConnectionString = HttpContext.Current.Session("ConnectionString")
				End If
			End If
			Return sBIWindowsConnectionString
		End Get
		Set(ByVal value As String)
			HttpContext.Current.Session("ConnectionString") = value
		End Set
	End Property

	Public Property EdConnectionString() As String
		Get
			Dim sEdConnectionString As String = ""
			If HttpContext.Current.Session("EdConnectionString") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("EdConnectionString")) <> "" Then
					sEdConnectionString = HttpContext.Current.Session("EdConnectionString")
				End If
			End If

			Dim sEDDecrypted As String = libEdContracts.Utils.CryptoAESDerypter("aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=", libEdContracts.Utils.CryptoGetSalt(""))
			sEdConnectionString = sEdConnectionString.Replace("aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=", sEDDecrypted)

			Return sEdConnectionString
		End Get
		Set(ByVal value As String)
			HttpContext.Current.Session("EdConnectionString") = value
		End Set
	End Property

	Public Property PasswordAdminEdConnectionString() As String
		Get
			Dim sPasswordAdminEdConnectionStringPA As String = ""
			If HttpContext.Current.Session("PasswordAdminEdConnectionString") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("PasswordAdminEdConnectionString")) <> "" Then
					sPasswordAdminEdConnectionStringPA = HttpContext.Current.Session("PasswordAdminEdConnectionString")
				End If
			End If
			Return sPasswordAdminEdConnectionStringPA
		End Get
		Set(ByVal value As String)
			HttpContext.Current.Session("PasswordAdminEdConnectionString") = value
		End Set
	End Property

	Public Property GrantImpersonateConnectionString() As String
		Get
			Dim sGrantImpersonateConnectionString As String = ""
			If HttpContext.Current.Session("GrantImpersonateConnectionString") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("GrantImpersonateConnectionString")) <> "" Then
					sGrantImpersonateConnectionString = HttpContext.Current.Session("GrantImpersonateConnectionString")
				End If
			End If

			Dim sIMPDecrypted As String = libEdContracts.Utils.CryptoAESDerypter("OTDLl5F+O6tnadA/xr/AyJd1hyjwQ1WZxO0hx/OVqQNY1qKiIqEDmdwmZN2ZpMVj", libEdContracts.Utils.CryptoGetSalt(""))
			sGrantImpersonateConnectionString = sGrantImpersonateConnectionString.Replace("OTDLl5F+O6tnadA/xr/AyJd1hyjwQ1WZxO0hx/OVqQNY1qKiIqEDmdwmZN2ZpMVj", sIMPDecrypted)

			Return sGrantImpersonateConnectionString
		End Get
		Set(ByVal value As String)
			HttpContext.Current.Session("GrantImpersonateConnectionString") = value
		End Set
	End Property


	Public Property BIApplicationFK() As Integer
		Get
			Dim iBIApplicationFK As Integer = 0
			If HttpContext.Current.Session("BIApplicationFK") IsNot Nothing Then
				If CStr(HttpContext.Current.Session("BIApplicationFK")) <> "" Then
					iBIApplicationFK = CInt(HttpContext.Current.Session("BIApplicationFK"))
				End If
			End If
			Return iBIApplicationFK
		End Get
		Set(ByVal value As Integer)
			HttpContext.Current.Session("BIApplicationFK") = value
		End Set
	End Property


	Public Function IsBackup() As Boolean
		Try
			If BIWindowsConnectionString.ToUpper.Contains("EVOSVR21") Or BIWindowsConnectionString.ToUpper.Contains("10.11.24.21;") Then
				Return True
			Else
				Return False
			End If
		Catch ex As Exception
			Return False
		End Try
	End Function

	Public Function GetRandomDigits(ByVal iNumberOfDigits As Integer) As String
		Dim sDigits As String = ""
		If iNumberOfDigits = 0 Then
			iNumberOfDigits = 4
		End If

		Dim RandGen As New Random(Guid.NewGuid().GetHashCode())
		Dim rnx As Long = RandGen.Next(1000, 10000)
		If rnx.ToString.Length = iNumberOfDigits Then
			sDigits = rnx.ToString
		ElseIf rnx.ToString.Length > iNumberOfDigits Then
			sDigits = rnx.ToString.Substring(0, iNumberOfDigits)
		Else
			' it's less than 4 digitst

		End If
		Return sDigits
	End Function


	Public Function SaveErrorsToFile(ByVal strData As String) As Boolean
		Dim bAns As Boolean = False

		Try
			Dim sLogFolder As String = "C:\Temp"
			If IO.Directory.Exists(sLogFolder) = False Then
				IO.Directory.CreateDirectory(sLogFolder)
			End If
			Dim sFilePath As String = IO.Path.Combine(sLogFolder, "EvoWeb3_Errors" & Date.Now.ToString("yyyyMMdd") & ".txt")

			Dim Contents As String = ""

			Dim objReader As IO.StreamWriter

			objReader = New IO.StreamWriter(sFilePath, True)
			objReader.WriteLine(Date.Now.ToString & ": " & strData)
			objReader.Close()
			bAns = True
		Catch Ex As Exception
			'Throw Ex
		End Try
		Return bAns
	End Function


	Public Sub logError(ByVal strPage As String, ByVal strMessage As String, ByVal strStackTrace As String)
		If InStr(strMessage, "Thread was being aborted") Then Exit Sub

		SaveErrorsToFile(strPage & Environment.NewLine & strMessage & Environment.NewLine & strStackTrace)

		If strPage.ToUpper <> "DO_NOT_SENDMAIL" Then
			SendEmail(
				"sanjay.patel@evolutionjobs.co.uk",
				"do-not-reply@evolutionjobs.co.uk",
				"MY-REPORTS(EVOWEB3) - ERROR",
				strPage & "<br/><br/>" & strMessage & "<br/><br/>" & strStackTrace & "<br/><br/>",
				True,
				"",
				"")
		End If

	End Sub

	Public Sub SendEmail(
	ByVal strToAddress As String,
	ByVal strFromAddress As String,
	ByVal strSubject As String,
	ByVal strBody As String,
	ByVal isHTML As Boolean,
	ByVal strCC As String,
	ByVal strBCC As String)


		'SaveToErrorLog("SendEMail-START: [" & strSubject & "] TO:" & strToAddress & "; From:" & strFromAddress)


		Try
			Dim objMail As New MailMessage()
			Dim fromAddress As New Net.Mail.MailAddress(strFromAddress)
			Dim msg As New MailMessage()
			msg.From = New MailAddress(strFromAddress)

			'' ----------------------------------------
			'Dim sSourceDSN As String = "Data Source=10.11.24.5;Password=aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=;Persist Security Info=True;User ID=Ed;Initial Catalog=Ed;Connect Timeout=30"
			'Dim sDecryptedPass As String = clsDTEmailAccount.clsAccountDetails.CryptoDecrypt("aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=", "support@evolutionjobs.co.uk")
			Dim DBConnectionString As String = EdConnectionString 'sSourceDSN.Replace("aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=", sDecryptedPass)
			Dim EmailServer As String = ""
			Dim EmailAccount As String = ""
			Dim EmailPasswordDecrypted As String = ""
			Using eml1 As New clsDTEmailAccount.clsAccountDetails(DBConnectionString)

				eml1.GetEmailAccountDetails(
				 EmailServer,
				 EmailAccount,
				 EmailPasswordDecrypted
				)
			End Using
			'' ----------------------------------------

			strToAddress = strToAddress.Replace(",", ";")
			If strToAddress <> "" Then
				If strToAddress.Trim.Contains(";") Then
					Dim sToSplit() As String = strToAddress.Split(CChar(";"))
					If sToSplit IsNot Nothing Then
						For iCC As Integer = 0 To sToSplit.Count - 1
							If sToSplit(iCC) IsNot Nothing Then
								If sToSplit(iCC) <> "" Then
									msg.To.Add(New MailAddress(sToSplit(iCC).Trim))
								End If
							End If
						Next
					End If
				Else
					msg.To.Add(New MailAddress(strToAddress))
				End If

			End If
			'msg.To.Add(New Net.Mail.MailAddress(strToAddress))

			strCC = strCC.Replace(",", ";")
			If strCC <> "" Then
				If strCC.Trim.Contains(";") Then
					Dim sCCSplit() As String = strCC.Split(CChar(";"))
					If sCCSplit IsNot Nothing Then
						For iCC As Integer = 0 To sCCSplit.Count - 1
							If sCCSplit(iCC) IsNot Nothing Then
								If sCCSplit(iCC) <> "" Then
									msg.CC.Add(New MailAddress(sCCSplit(iCC).Trim))
								End If
							End If
						Next
					End If
				Else
					msg.CC.Add(New MailAddress(strCC))
				End If

			End If

			strBCC = strBCC.Replace(",", ";")
			If strBCC <> "" Then
				If strBCC.Trim.Contains(";") Then
					Dim sBCCSplit() As String = strBCC.Split(CChar(";"))
					If sBCCSplit IsNot Nothing Then
						For iBCC As Integer = 0 To sBCCSplit.Count - 1
							If sBCCSplit(iBCC) IsNot Nothing Then
								If sBCCSplit(iBCC) <> "" Then
									msg.Bcc.Add(New MailAddress(sBCCSplit(iBCC).Trim))
								End If
							End If
						Next
					End If
				Else
					msg.Bcc.Add(New MailAddress(strBCC))
				End If

			End If

			msg.Body = strBody '& Environment.NewLine & Environment.NewLine & My.Computer.Name.ToString.ToUpper
			msg.Subject = strSubject
			msg.IsBodyHtml = isHTML
			Dim mailSender As New System.Net.Mail.SmtpClient()
			mailSender.Host = "evoexch1" ' My.Settings.EmailHost ' ConfigurationManager.AppSettings("Emailhost")

			If EmailServer.Trim <> "" Then
				mailSender.Host = EmailServer
			End If

			Dim oCredential As Net.NetworkCredential = New Net.NetworkCredential(EmailAccount, EmailPasswordDecrypted)
			mailSender.UseDefaultCredentials = False
			mailSender.Credentials = oCredential

			mailSender.Send(msg)
			objMail = Nothing
			fromAddress = Nothing
			msg = Nothing
			mailSender = Nothing
		Catch ex As Exception
			SaveToLog("EXCEPTION SendEMail: " & ex.Message & ex.StackTrace)
			logError("EvoWeb3-SendEmail", "EXCEPTION SendEMail: " & ex.Message, ex.StackTrace)
			'SaveToDailyLog("$$$$$$ EXCEPTION $$$$$$$$: SendEMail: " & ex.Message & ex.StackTrace)
		End Try

	End Sub



	Public Function SendSMS(ByVal strTo As String, ByVal strMessage As String) As String

		' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
		' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
		' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
		' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
		' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
		' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN


		Dim sError As String = ""
		'Dim DBConnectionString As String = ConfigurationManager.ConnectionStrings("EdConnectionString").ConnectionString

		'' ----------------------------------------
		'Dim sDecryptedPass As String = clsDTEmailAccount.clsAccountDetails.CryptoDecrypt("aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=", "support@evolutionjobs.co.uk")
		Dim DBConnectionString As String = EdConnectionString 'sSourceDSN.Replace("aLX6BSwDgW4f3aBQDvXY3fCzioqag3+4bPaPjKmiqpk=", sDecryptedPass)
		Dim EmailServer As String = ""
		Dim EmailAccount As String = ""
		Dim EmailPasswordDecrypted As String = ""
		Using eml1 As New clsDTEmailAccount.clsAccountDetails(DBConnectionString)

			eml1.GetEmailAccountDetails(
				 EmailServer,
				 EmailAccount,
				 EmailPasswordDecrypted
				)
		End Using
		'' ----------------------------------------



		Dim ObjMM As New Net.Mail.MailMessage
		Dim SMTP As New Net.Mail.SmtpClient
		SMTP.Host = EmailServer

		ObjMM.From = New Net.Mail.MailAddress("ITTeam@evolutionjobs.co.uk", "")
		ObjMM.To.Add(New Net.Mail.MailAddress(strTo))
		' ObjMM.Subject = ""

		Dim sBody As String = strMessage
		ObjMM.Body = sBody

		ObjMM.IsBodyHtml = False

		Try
			Dim oCredential As Net.NetworkCredential = New Net.NetworkCredential(EmailAccount, EmailPasswordDecrypted)
			SMTP.UseDefaultCredentials = False
			SMTP.Credentials = oCredential

			SMTP.Send(ObjMM)
		Catch ex As Exception
			sError = "ERROR:" & ex.Message
		Finally
			ObjMM.Dispose()
		End Try

		Return sError
	End Function


	Function IsValidEmail(ByVal Email As String, Optional ByVal blnAllowMultipleEmails As Boolean = False) As Boolean

		'Email = NulltoString(Email)
		Dim strRegex As String = "^([a-zA-Z0-9~_\-\.]+)@((\[[0-9]{1,3}" &
			  "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" &
			  ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"

		Dim re As Regex = New Regex(strRegex)

		' Allow apostrophe's in emails addresses - since the validator wil fail apostrophes,
		' remove them from the check
		Email = Email.Replace("'", "")

		If Email.Contains(";") = True And blnAllowMultipleEmails = True Then
			Dim strEmailItem() As String = Email.Split(";")
			For Each strItem As String In strEmailItem
				If re.IsMatch(strItem.Trim) = False Then Return False
			Next
			Return True
		Else
			If re.IsMatch(Email) Then
				Return True
			Else
				Return False
			End If
		End If

	End Function


	Public Function SaveToLog(ByVal strData As String) As Boolean
		Dim bAns As Boolean = False
		Dim objStreamWriter As IO.StreamWriter
		Try
			If IO.Directory.Exists("C:\Temp") = False Then
				IO.Directory.CreateDirectory("C:\Temp")
			End If
			objStreamWriter = New IO.StreamWriter("C:\Temp\MyReports_" & Date.Now.ToString("yyyyMMdd") & ".txt", True)
			objStreamWriter.WriteLine(Date.Now.ToString & ": [" & WindowsUsername & "] " & strData)
			objStreamWriter.Close()
			bAns = True
		Catch Ex As Exception
			'Throw Ex
		Finally
			If objStreamWriter IsNot Nothing Then
				objStreamWriter.Dispose()
			End If
		End Try
		Return bAns
	End Function


	Public Function SendVerificationSMS(ByVal PhoneEntered As String, ByVal WindowsUsername As String, ByVal ReSendCode As Boolean) As String
		Dim sReturn As String = ""
		Dim sGetError As String = ""
		Dim outPhoneCode As String = ""
		Dim outEmailCode As String = ""

		SaveToLog(WindowsUsername & "SendVerificationSMS-START")

		Dim sSessionFinalVerificationCode As String = ""
		Try
			If HttpContext.Current.Session("PhoneFinalVerificationCode") IsNot Nothing Then
				sSessionFinalVerificationCode = HttpContext.Current.Session("PhoneFinalVerificationCode")
			End If
		Catch ex As Exception
		End Try

		If sSessionFinalVerificationCode.Trim <> "" Then

			' the first time the VERIFY button is clicked, we send out the code
			' the second time it is clicked wwe just go to the enter code screen IF the session has a code stored
			' BUT if ReSendCode is set, then we re-send the code

			If ReSendCode Then
				' send out the code again
				outPhoneCode = sSessionFinalVerificationCode
			Else
				' dont send out the code - just go to enter code screen
				Return "ENTER_VERIFICATION_CODE"
			End If



		Else

			SaveToLog(WindowsUsername & "SendVerificationSMS-GetVerificationCodes-A1")

			Using lib1 As New clsDB(BIWindowsConnectionString)
				' 1. get Stored info for comparison
				sGetError = lib1.GetVerificationCodes(WindowsUsername, outPhoneCode, outEmailCode)
			End Using

			SaveToLog(WindowsUsername & "SendVerificationSMS-GetVerificationCodes-A2; outPhoneCode=" & outPhoneCode & "; outEmailCode=" & outEmailCode)

		End If


		If sGetError.Trim <> "" Then
			sReturn = sGetError.Trim

			SaveToLog(WindowsUsername & "SendVerificationSMS-GetVerificationCodes-B; sGetError=" & sGetError)
		Else

			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------
			Dim sManageMyAcc As String = ""
			Try
				If HttpContext.Current.Session("LoginPage") IsNot Nothing Then
					sManageMyAcc = HttpContext.Current.Session("LoginPage")
				End If
			Catch ex As Exception
			End Try
			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------

			Dim sPWError As String = ""
			Dim sTempPW As String = ""
			If sManageMyAcc = "From_Manage_My_Account" Then
				' this is from manage my account, so does not require a password check, because the user is not verifying and requesting a new password

			Else

				SaveToLog(WindowsUsername & "SendVerificationSMS-GetNewPassword-C")

				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' this is to avoid somone having to click verify, then get the code, enter it and verify, then click the Send Password link, get it from phone/email...
				' this way, they get ONE code sent to their mobile which they use to verify and login (change pass first)
				Using lib1 As New clsDB(BIWindowsConnectionString)
					' 1. get Stored info for comparison
					sPWError = lib1.GetNewPassword(WindowsUsername, sTempPW)
				End Using
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
				' ======================== append the new password to the code sent only if forgot pass or new pass ================================
			End If


			SaveToLog(WindowsUsername & "SendVerificationSMS-GetNewPassword-D; outPhoneCode=" & outPhoneCode & "; sPWError=" & sPWError)

			If sPWError.Trim <> "" Then
				sReturn = sPWError.Trim
				logError("SendVerificationSMS1", "ERROR: getting phone and email verification SMS codes; WindowsUsername=" & WindowsUsername & ": " & sPWError, Nothing)

			Else

				If outPhoneCode.Trim <> "" Then

					' ++++++++++++++++++++++++++++ IMPORTANT +++++++++++++++++++++++++++++
					' NOTE: we need a way to stop sending out verification codes everytime one is requested:
					' if ther is code in the database, then re-send that one. The only problem is we need to check
					' if it requires a password concatenated and if it hasnt been concatenated, then we need to do so and save it
					' ++++++++++++++++++++++++++++ IMPORTANT +++++++++++++++++++++++++++++

					SaveToLog(WindowsUsername & "SendVerificationSMS-A; outPhoneCode=" & outPhoneCode)

					Dim sRetSaveCodeError As String = ""
					Dim sFinalCode As String = outPhoneCode.Trim

					If sManageMyAcc = "From_Manage_My_Account" Then
						' no password required
						' verification code should be 4 chars long

						' ***** just send that again unchanged = default *****
					Else
						' password is required, so the verification code should be 8 chars long
						' if it isn;t then append the pass and save it

						' if it is , then just resent that
						If outPhoneCode.Trim.Length = 8 Then
							' OK it is 8 chars - just send that out again

							' ***** just send that again unchanged = default *****
						Else
							' not OK
							If outPhoneCode.Trim.Length = 4 Then
								' only 4 chars - need to append password

								If sTempPW.Trim <> "" Then

									'sFinalCode &= sTempPW.Trim

									' generate only numbers: 4 digit
									Dim s4Digits As String = ""

									s4Digits = GetRandomDigits(4)

									SaveToLog(WindowsUsername & "SendVerificationSMS-B")

									If s4Digits.Trim = "" Then
										s4Digits = outPhoneCode.Trim
									End If

									'Dim RandGen As New Random(Guid.NewGuid().GetHashCode())
									'Dim rnx As Long = RandGen.Next(1000, 10000)
									'If rnx.ToString.Length = 4 Then
									'	s4Digits = rnx.ToString
									'ElseIf rnx.ToString.Length > 4 Then
									'	s4Digits = rnx.ToString.Substring(0, 4)
									'Else
									'	' it's less than 4 digitst
									'	s4Digits = outPhoneCode.Trim
									'End If

									sFinalCode = s4Digits.ToString & sTempPW.Trim


									Using lib1 As New clsDB(BIWindowsConnectionString)
										sRetSaveCodeError = lib1.UpdateVerificationCode(WindowsUsername, sFinalCode, "")
									End Using

									SaveToLog(WindowsUsername & "SendVerificationSMS-C")

								End If

							Else
								' ERROR - should be 4 or 8 chars only
								sRetSaveCodeError = "ERROR: Verification code is incorrect"

							End If


						End If

					End If


					If sRetSaveCodeError.Trim <> "" Then

						SaveToLog(WindowsUsername & "SendVerificationSMS-D")

						sReturn = sRetSaveCodeError

					Else


						SaveToLog(WindowsUsername & "SendVerificationSMS-E-Pre_SendSMS")


						' 3. Send SMS 
						' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
						' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
						' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
						' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
						' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
						Dim sSMSMessage As String = String.Format(sSMSVerificationMessage, sFinalCode)
						Dim sSMSError As String = ""
						sSMSError = SendSMS(PhoneEntered.Trim.Replace(" ", "").Replace("+", "00") & SMSDomain, sSMSMessage)

						'If PhoneEntered.Contains("7507342793") Then
						'	sSMSError = SendSMS(PhoneEntered.Trim.Replace(" ", "").Replace("+", "00") & "@sms.textapp.net", sSMSMessage & "(NEW)")

						'Else
						'	sSMSError = SendSMS(PhoneEntered.Trim.Replace(" ", "").Replace("+"SMSDomain, "00") & SMSDomain, sSMSMessage)

						'End If

						SaveToLog(WindowsUsername & "SendVerificationSMS-F-POST_SendSMS")

						' 4. redirect to code entry page
						If sSMSError.Trim <> "" Then
							sReturn = sSMSError.Trim
						Else
							' store it in sessions so we can keep resending it, rather than getting a new one every time, which will be a pain is
							' the user presses send password/verify multiple times and when the SMS takes so long to arrive
							HttpContext.Current.Session("PhoneFinalVerificationCode") = sFinalCode

							'Session("action") = "NewPasswordStep3"
							'pnlNewPasswordStep1.Visible = False
							'pnlNewPasswordStep2.Visible = False
							'pnlNewPasswordStep3.Visible = True
							'pnlNewPasswordStep4.Visible = False
						End If

					End If


				Else
					sReturn = "ERROR: Retrieving verification code"
					logError("SendVerificationSMS2-a", "ERROR-a: getting phone and email verification SMS codes; WindowsUsername=" & WindowsUsername & ": " & sGetError, Nothing)
				End If


			End If


		End If

		Return sReturn
	End Function


	Public Function SendVerificationEmail(ByVal EmailEntered As String, ByVal WindowsUsername As String, ByVal ReSendCode As Boolean) As String
		Dim sReturn As String = ""
		Dim sGetError As String = ""
		Dim outPhoneCode As String = ""
		Dim outEmailCode As String = ""

		If IsValidEmail(EmailEntered) = False Then
			sGetError = "ERROR: Please enter a valid email address"
			Return sGetError
		End If

		Dim sSessionFinalVerificationCode As String = ""
		Try
			If HttpContext.Current.Session("EmailFinalVerificationCode") IsNot Nothing Then
				sSessionFinalVerificationCode = HttpContext.Current.Session("EmailFinalVerificationCode")
			End If
		Catch ex As Exception
		End Try

		If sSessionFinalVerificationCode.Trim <> "" Then
			' the first time the VERIFY button is clicked, we send out the code
			' the second time it is clicked wwe just go to the enter code screen IF the session has a code stored
			' BUT if ReSendCode is set, then we re-send the code

			If ReSendCode Then
				' send out the code again
				outEmailCode = sSessionFinalVerificationCode
			Else
				' dont send out the code - just go to enter code screen
				Return "ENTER_VERIFICATION_CODE"
			End If


		Else
			Using lib1 As New clsDB(BIWindowsConnectionString)
				' 1. get Stored info for comparison
				sGetError = lib1.GetVerificationCodes(WindowsUsername, outPhoneCode, outEmailCode)
			End Using
		End If


		If sGetError.Trim <> "" Then
			sReturn = sGetError.Trim
			logError("SendVerificationEmail", "ERROR: getting phone and email verification codes; WindowsUsername=" & WindowsUsername & ": " & sGetError, Nothing)

		Else

			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------
			Dim sManageMyAcc As String = ""
			Try
				If HttpContext.Current.Session("LoginPage") IsNot Nothing Then
					sManageMyAcc = HttpContext.Current.Session("LoginPage")
				End If
			Catch ex As Exception
			End Try
			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------
			' ------------- check if password check is required -------------------

			Dim sPWError As String = ""
			Dim sTempPW As String = ""
			If sManageMyAcc = "From_Manage_My_Account" Then
				' this is from manage my account, so does not require a password check, because the user is not verifying and requesting a new password

			Else

				' ======================== append the new password to the code sent ================================
				' ======================== append the new password to the code sent ================================
				' ======================== append the new password to the code sent ================================
				' ======================== append the new password to the code sent ================================
				' this is to avoid somone having to click verify, then get the code, enter it and verify, then click the Send Password link, get it from phone/email...
				' this way, they get ONE code sent to their mobile which they use to verify and login (change pass first)
				Using lib1 As New clsDB(BIWindowsConnectionString)
					' 1. get Stored info for comparison
					sPWError = lib1.GetNewPassword(WindowsUsername, sTempPW)
				End Using
				' ======================== append the new password to the code sent ================================
				' ======================== append the new password to the code sent ================================
				' ======================== append the new password to the code sent ================================
				' ======================== append the new password to the code sent ================================

			End If



			If sPWError.Trim <> "" Then
				sReturn = sPWError.Trim
				logError("SendVerificationEmail1", "ERROR: getting phone and email verification codes; WindowsUsername=" & WindowsUsername & ": " & sPWError, Nothing)

			Else


				If outEmailCode.Trim <> "" Then

					' ++++++++++++++++++++++++++++ IMPORTANT +++++++++++++++++++++++++++++
					' NOTE: we need a way to stop sending out verification codes everytime one is requested:
					' if ther is code in the database, then re-send that one. The only problem is we need to check
					' if it requires a password concatenated and if it hasnt been concatenated, then we need to do so and save it
					' ++++++++++++++++++++++++++++ IMPORTANT +++++++++++++++++++++++++++++

					Dim sRetSaveCodeError As String = ""
					Dim sFinalCode As String = outEmailCode.Trim

					If sManageMyAcc = "From_Manage_My_Account" Then
						' no password required
						' verification code should be 4 chars long

						' ***** just send that again unchanged = default *****


					Else
						' password is required, so the verification code should be 8 chars long
						' if it isn;t then append the pass and save it

						' if it is , then just resent that

						If outEmailCode.Trim.Length = 8 Then
							' OK it is 8 chars - just send that out again

							' ***** just send that again unchanged = default *****

						Else
							' not OK

							If outEmailCode.Trim.Length = 4 Then
								' only 4 chars - need to append password

								If sTempPW.Trim <> "" Then

									'sFinalCode &= sTempPW.Trim

									' generate only numbers: 4 digit

									Dim s4Digits As String = ""

									s4Digits = GetRandomDigits(4)

									If s4Digits.Trim = "" Then
										s4Digits = outPhoneCode.Trim
									End If


									'Dim RandGen As New Random(Guid.NewGuid().GetHashCode())
									'Dim rnx As Long = RandGen.Next(1000, 10000)
									'If rnx.ToString.Length = 4 Then
									'	s4Digits = rnx.ToString
									'ElseIf rnx.ToString.Length > 4 Then
									'	s4Digits = rnx.ToString.Substring(0, 4)
									'Else
									'	' it's less than 4 digitst
									'	s4Digits = outEmailCode.Trim
									'End If

									sFinalCode = s4Digits.ToString & sTempPW.Trim


									Using lib1 As New clsDB(BIWindowsConnectionString)
										sRetSaveCodeError = lib1.UpdateVerificationCode(WindowsUsername, "", sFinalCode)
									End Using
								End If
							Else
								' ERROR - should be 4 or 8 chars only
								sRetSaveCodeError = "ERROR: Verification code is incorrect"
							End If

						End If

					End If


					If sRetSaveCodeError.Trim <> "" Then
						sReturn = sRetSaveCodeError
					Else



						' 3. Send email 
						Dim sSMSMessage As String = String.Format(sEmailVerificationMessage, sFinalCode)
						' sendng SMS is same as sending email!!!
						Dim sSMSError As String = SendSMS(EmailEntered.Trim, sSMSMessage)


						' 4. redirect to code entry page
						If sSMSError.Trim <> "" Then
							sReturn = sSMSError.Trim
						Else
							' store it in sessions so we can keep resending it, rather than getting a new one every time, which will be a pain is
							' the user presses send password/verify multiple times and when the SMS takes so long to arrive

							HttpContext.Current.Session("EmailFinalVerificationCode") = sFinalCode

							'Session("action") = "NewPasswordStep3"
							'pnlNewPasswordStep1.Visible = False
							'pnlNewPasswordStep2.Visible = False
							'pnlNewPasswordStep3.Visible = True
							'pnlNewPasswordStep4.Visible = False

						End If


					End If


				Else
					sReturn = "ERROR: Retrieving verification code"
					logError("SendVerificationEmail2", "ERROR: getting phone and email verification codes; WindowsUsername=" & WindowsUsername & ": " & sGetError, Nothing)

				End If


			End If


		End If

		Return sReturn

	End Function




	'Public Function SendVerificationSMS(ByVal PhoneEntered As String, ByVal WindowsUsername As String) As String
	'	Dim sReturn As String = ""
	'	Dim sGetError As String = ""
	'	Dim outPhoneCode As String = ""
	'	Dim outEmailCode As String = ""
	'	Using lib1 As New clsDB(BIWindowsConnectionString)
	'		' 1. get Stored info for comparison
	'		sGetError = lib1.GetVerificationCodes(WindowsUsername, outPhoneCode, outEmailCode)
	'	End Using

	'	If sGetError.Trim <> "" Then
	'		sReturn = sGetError.Trim
	'	Else

	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------
	'		Dim sManageMyAcc As String = ""
	'		Try
	'			If HttpContext.Current.Session("LoginPage") IsNot Nothing Then
	'				sManageMyAcc = HttpContext.Current.Session("LoginPage")
	'			End If
	'		Catch ex As Exception
	'		End Try
	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------

	'		Dim sPWError As String = ""
	'		Dim sTempPW As String = ""
	'		If sManageMyAcc = "From_Manage_My_Account" Then
	'			' this is from manage my account, so does not require a password check, because the user is not verifying and requesting a new password

	'		Else
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' this is to avoid somone having to click verify, then get the code, enter it and verify, then click the Send Password link, get it from phone/email...
	'			' this way, they get ONE code sent to their mobile which they use to verify and login (change pass first)
	'			Using lib1 As New clsDB(BIWindowsConnectionString)
	'				' 1. get Stored info for comparison
	'				sPWError = lib1.GetNewPassword(WindowsUsername, sTempPW)
	'			End Using
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'			' ======================== append the new password to the code sent only if forgot pass or new pass ================================
	'		End If




	'		If sPWError.Trim <> "" Then
	'			sReturn = sPWError.Trim
	'		Else


	'			If outPhoneCode.Trim <> "" Then

	'				Dim sRetSaveCodeError As String = ""
	'				Dim sFinalCode As String = outPhoneCode.Trim
	'				If sTempPW.Trim <> "" Then

	'					'sFinalCode &= sTempPW.Trim

	'					' generate only numbers: 4 digit
	'					Dim s4Digits As Integer = 0
	'					Dim RandGen As New Random(Guid.NewGuid().GetHashCode())
	'					Dim rnx As Long = RandGen.Next(1000, 10000)
	'					If rnx.ToString.Length = 4 Then
	'						s4Digits = rnx.ToString
	'					ElseIf rnx.ToString.Length > 4 Then
	'						s4Digits = rnx.ToString.Substring(0, 4)
	'					Else
	'						' it's less than 4 digitst
	'						s4Digits = outEmailCode.Trim
	'					End If

	'					sFinalCode = s4Digits.ToString & sTempPW.Trim


	'					Using lib1 As New clsDB(BIWindowsConnectionString)
	'						sRetSaveCodeError = lib1.UpdateVerificationCode(WindowsUsername, sFinalCode, "")
	'					End Using
	'				End If


	'				If sRetSaveCodeError.Trim <> "" Then

	'					sReturn = sRetSaveCodeError

	'				Else
	'					' 3. Send SMS 
	'					' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'					' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'					' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'					' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'					' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
	'					Dim sSMSMessage As String = String.Format(sSMSVerificationMessage, sFinalCode)
	'					Dim sSMSError As String = SendSMS(PhoneEntered.Trim.Replace(" ", "").Replace("+", "00") & "@evotext.net", sSMSMessage)


	'					' 4. redirect to code entry page
	'					If sSMSError.Trim <> "" Then
	'						sReturn = sSMSError.Trim
	'					Else
	'						'Session("action") = "NewPasswordStep3"
	'						'pnlNewPasswordStep1.Visible = False
	'						'pnlNewPasswordStep2.Visible = False
	'						'pnlNewPasswordStep3.Visible = True
	'						'pnlNewPasswordStep4.Visible = False
	'					End If

	'				End If


	'			Else
	'				sReturn = "ERROR: Retrieving verification code"
	'			End If


	'		End If







	'	End If

	'	Return sReturn
	'End Function


	'Public Function SendVerificationEmail(ByVal EmailEntered As String, ByVal WindowsUsername As String) As String
	'	Dim sReturn As String = ""
	'	Dim sGetError As String = ""
	'	Dim outPhoneCode As String = ""
	'	Dim outEmailCode As String = ""

	'	If IsValidEmail(EmailEntered) = False Then
	'		sGetError = "Please enter a valid email address"
	'		Return sGetError
	'	End If

	'	Using lib1 As New clsDB(BIWindowsConnectionString)
	'		' 1. get Stored info for comparison
	'		sGetError = lib1.GetVerificationCodes(WindowsUsername, outPhoneCode, outEmailCode)
	'	End Using

	'	If sGetError.Trim <> "" Then
	'		sReturn = sGetError.Trim
	'	Else

	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------
	'		Dim sManageMyAcc As String = ""
	'		Try
	'			If HttpContext.Current.Session("LoginPage") IsNot Nothing Then
	'				sManageMyAcc = HttpContext.Current.Session("LoginPage")
	'			End If
	'		Catch ex As Exception
	'		End Try
	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------
	'		' ------------- check if password check is required -------------------

	'		Dim sPWError As String = ""
	'		Dim sTempPW As String = ""
	'		If sManageMyAcc = "From_Manage_My_Account" Then
	'			' this is from manage my account, so does not require a password check, because the user is not verifying and requesting a new password

	'		Else

	'			' ======================== append the new password to the code sent ================================
	'			' ======================== append the new password to the code sent ================================
	'			' ======================== append the new password to the code sent ================================
	'			' ======================== append the new password to the code sent ================================
	'			' this is to avoid somone having to click verify, then get the code, enter it and verify, then click the Send Password link, get it from phone/email...
	'			' this way, they get ONE code sent to their mobile which they use to verify and login (change pass first)
	'			Using lib1 As New clsDB(BIWindowsConnectionString)
	'				' 1. get Stored info for comparison
	'				sPWError = lib1.GetNewPassword(WindowsUsername, sTempPW)
	'			End Using
	'			' ======================== append the new password to the code sent ================================
	'			' ======================== append the new password to the code sent ================================
	'			' ======================== append the new password to the code sent ================================
	'			' ======================== append the new password to the code sent ================================

	'		End If



	'		If sPWError.Trim <> "" Then
	'			sReturn = sPWError.Trim
	'		Else


	'			If outEmailCode.Trim <> "" Then

	'				Dim sRetSaveCodeError As String = ""
	'				Dim sFinalCode As String = outEmailCode.Trim
	'				If sTempPW.Trim <> "" Then

	'					'sFinalCode &= sTempPW.Trim

	'					' generate only numbers: 4 digit
	'					Dim s4Digits As Integer = 0
	'					Dim RandGen As New Random(Guid.NewGuid().GetHashCode())
	'					Dim rnx As Long = RandGen.Next(1000, 10000)
	'					If rnx.ToString.Length = 4 Then
	'						s4Digits = rnx.ToString
	'					ElseIf rnx.ToString.Length > 4 Then
	'						s4Digits = rnx.ToString.Substring(0, 4)
	'					Else
	'						' it's less than 4 digitst
	'						s4Digits = outEmailCode.Trim
	'					End If

	'					sFinalCode = s4Digits.ToString & sTempPW.Trim


	'					Using lib1 As New clsDB(BIWindowsConnectionString)
	'						sRetSaveCodeError = lib1.UpdateVerificationCode(WindowsUsername, "", sFinalCode)
	'					End Using
	'				End If


	'				If sRetSaveCodeError.Trim <> "" Then
	'					sReturn = sRetSaveCodeError
	'				Else

	'					' 3. Send email 
	'					Dim sSMSMessage As String = String.Format(sEmailVerificationMessage, sFinalCode)
	'					' sendng SMS is same as sending email!!!
	'					Dim sSMSError As String = SendSMS(EmailEntered.Trim, sSMSMessage)


	'					' 4. redirect to code entry page
	'					If sSMSError.Trim <> "" Then
	'						sReturn = sSMSError.Trim
	'					Else
	'						'Session("action") = "NewPasswordStep3"
	'						'pnlNewPasswordStep1.Visible = False
	'						'pnlNewPasswordStep2.Visible = False
	'						'pnlNewPasswordStep3.Visible = True
	'						'pnlNewPasswordStep4.Visible = False

	'					End If


	'				End If


	'			Else
	'				sReturn = "ERROR: Retrieving verification code"
	'			End If


	'		End If


	'	End If

	'	Return sReturn

	'End Function










	Public Function UpdateUserPhone(
	 ByVal windowslogin As String,
	 ByVal inPhone As String
	 ) As String
		Dim sReturn As String = ""
		Dim sGetError As String = ""
		Dim outPhoneCode As String = ""
		Dim outEmailCode As String = ""
		Using lib1 As New clsDB(BIWindowsConnectionString)
			' 1. get Stored info for comparison
			sGetError = lib1.SaveUserPhone(windowslogin, inPhone)
		End Using

		If sGetError.Trim <> "" Then
			sReturn = sGetError.Trim
		End If

		Return sReturn
	End Function



	Public Function UpdateUserEmail(
	 ByVal windowslogin As String,
	 ByVal inEmail As String
	 ) As String
		Dim sReturn As String = ""
		Dim sGetError As String = ""
		Dim outPhoneCode As String = ""
		Dim outEmailCode As String = ""
		Using lib1 As New clsDB(BIWindowsConnectionString)
			' 1. get Stored info for comparison
			sGetError = lib1.SaveUserEmail(windowslogin, inEmail)
		End Using

		If sGetError.Trim <> "" Then
			sReturn = sGetError.Trim
		End If

		Return sReturn
	End Function



	Public Sub ClearSessions()
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		Try
			' we must keep the Windows details
			HttpContext.Current.Session("AutoLoginLinkChangePassword") = ""
			HttpContext.Current.Session("action") = ""
			HttpContext.Current.Session("NewPassword") = ""
			HttpContext.Current.Session("LoginPage") = ""
			HttpContext.Current.Session("ChangePasswordMessage") = ""
			HttpContext.Current.Session("EmailFinalVerificationCode") = ""
			HttpContext.Current.Session("PhoneFinalVerificationCode") = ""
			HttpContext.Current.Session("SendNewPassword") = ""
		Catch ex As Exception
		End Try
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password
		' clear the session holding the login url for New Password

	End Sub


	'Public Function CheckWindowsUsername(ByVal WindowsUserName As String) As Long

	'    Dim oRet As New Object
	'    Dim lRet As Long = -1
	'    Try
	'        Using conn As New SqlClient.SqlConnection(EdConnectionString)
	'            conn.Open()
	'            Using oCmd As New SqlClient.SqlCommand("SELECT [dbo].[udf_SP_WIN_CheckWindowsUsername](@WindowsUserName) ", conn)
	'                oCmd.CommandType = CommandType.Text
	'                oCmd.CommandTimeout = 3600
	'                oCmd.Parameters.Add("@WindowsUserName", SqlDbType.VarChar, 100).Value = WindowsUserName
	'                oRet = oCmd.ExecuteScalar()

	'                If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'                    lRet = CLng(oRet)
	'                End If

	'            End Using
	'            conn.Close()
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    End Try
	'    Return lRet
	'End Function


End Module
