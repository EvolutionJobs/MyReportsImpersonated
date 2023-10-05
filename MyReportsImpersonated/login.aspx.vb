Imports System.Data
Imports System.Web

Partial Class WagesLogin
    Inherits System.Web.UI.Page




	Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

		' live site
		'http://EvoWeb3/Login.aspx?appfk=22


		' Test site
		'http://EvoWeb3:91/Login.aspx?appfk=22


		'     <add name="BIWindowsConnectionString" connectionString="Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;application name=WHUA"/>

		'Initial catalog=BI;data source=10.11.24.21;Integrated Security=SSPI;persist security info=True;
		'Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;Persist security info=True;application name=WHUA

		'If Not HttpContext.Current.Request.ServerVariables("HTTP_REFERER") = Nothing Then
		'    lblTest.Text = "HTTP_REFERER=" & HttpContext.Current.Request.ServerVariables("HTTP_REFERER")
		'End If

		' PROCESS:
		' Landing page will be the MyReports page
		' Each Link on the MyReports page will point to the login page but with an appfk query string param


		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' IF THE WEBSITE KEEP PROMPTING YOU FOR WINDOWS CREDENTIALS, THEN:
		'Authentication > Enable Windows Authentication, then Right-Click to set the Providers.
		' ----->>>>NEGOTIATE needs to be FIRST!
		' ------SET: <identity impersonate="true" />

		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ASLO must set :
		' In the dropdown menu select system.webServer > security > authentication > windowsAuthentication > 
		' Change useAppPoolCredentials to True as in:
		' http://woshub.com/configuring-kerberos-authentication-on-iis-website/
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################
		' ############### useAppPoolCredentials ########################

		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------
		' ------------- IMPORTANT ----------------------









		' test site: \\EvoWeb3\intranet\MyReports2017TestV15Test
		'http://EvoWeb3:91/report.aspx?appfk=22
		'http://EvoWeb3:91/login.aspx?appfk=22

		'http://EvoWeb3
		'http://EvoWeb3/Report.aspx
		'http://EvoWeb3/login.aspx?appfk=22
		'http://EvoWeb3/login.aspx?appfk=22

		' http://localhost:52354/SSRSLayer/Default.aspx?appfk=22

		Dim sWinUsername As String = HttpContext.Current.User.Identity.Name

		lblStatus.Text = ""

		lblCurrentURL.Text = Request.Url.ToString
		lblURL.Text = ""
		lblSpecialMessage.Text = ""
		lblOldPassword.Text = "Old (or Temporary) Password:"
		pnlDefault.Visible = False



		'Dim RandGen As New Random(Guid.NewGuid().GetHashCode())
		'Dim rnx As Long = RandGen.Next(1000, 10000)

		'lblStatus.Text = "rnx=" & rnx.ToString


		Dim iChangePassword As Integer = 0
		Dim inApplicationFK As Integer = 0
		Dim PasswordProtect As Boolean = False
		If Request("appfk") <> "" Then
			inApplicationFK = CInt(Request("appfk"))
		End If
		If Request("cp") <> "" Then
			iChangePassword = CInt(Request("cp"))
		End If

		Try

			SaveToLog(sWinUsername & " Login_LOAD_____Pre_ImpersonateValidUser_____ [HttpContext.Current.User.Identity.Name=" & sWinUsername & "] HttpContext.Current.Session.WindowsUsername=" & WindowsUsername & "; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & ": inApplicationFK=" & inApplicationFK & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ")


			' get the user's details to see which link to show/hide
			Dim bVerifyRequired As Boolean = False
			Dim sGetError As String = ""
			Dim sDBPhone As String = ""
			Dim sDBEmail As String = ""
			'Using lib1 As New clsDB(BIWindowsConnectionString)
			'	' 1. get Stored info for comparison
			'	sGetError = lib1.GetPhoneOrEmail(lblWindowsUsername.Text, sDBPhone, sDBEmail)
			'End Using

			If WindowsPassword.Trim.Length > 0 Then

				' --- with manual impersonation ------
				' --- with manual impersonation ------
				' --- with manual impersonation ------
				Try
					' IF YOU USE IMPERSONATION HERE, THEN YOU must USE INLINE sql TO "Execute As Login = 'evolutionjobs\patels';"
					Dim objImpersonation As New UserImpersonation()
					If (objImpersonation.ImpersonateValidUser(WindowsUsername, WindowsPassword, WindowsDomain)) Then

						' ---- impoersonate user -------
						' ---- impoersonate user -------
						' ---- impoersonate user -------
						Using lib1 As New clsDB(PasswordAdminEdConnectionString)
							' 1. get Stored info for comparison
							sGetError = lib1.GetPhoneOrEmail(lblWindowsUsername.Text, sDBPhone, sDBEmail)
						End Using
						' ---- impoersonate user -------
						' ---- impoersonate user -------
						' ---- impoersonate user -------



						objImpersonation.UndoImpersonation()
					Else
						objImpersonation.UndoImpersonation()
						Response.Redirect("Error.aspx?e=User+does+not+have+enough+permissions+to+perform+the+task")

						'Throw New Exception("User does not have enough permissions to perform the task")
					End If
				Catch ex As Exception
					' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
					' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
					' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
					'Response.Redirect("Error.aspx?e=" & ex.Message.Replace(" ", "+"))
					'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				End Try
				' --- with manual impersonation ------
				' --- with manual impersonation ------
				' --- with manual impersonation ------


			Else


				Using lib1 As New clsDB(BIWindowsConnectionString)
					' 1. get Stored info for comparison
					sGetError = lib1.GetPhoneOrEmail(lblWindowsUsername.Text, sDBPhone, sDBEmail)
				End Using

			End If


			lnkcmdManageMyAccount.Visible = True
			lnkcmdChangePassword.Visible = True
			litPipe1.Visible = True
			lnkcmdFogottenPassword.Visible = True
			litPipe2.Visible = True
			lnkcmdRequestNewPassword.Visible = True


			If sGetError.Trim <> "" Then
				' ignore - do nothing: keep all links
			Else
				If sDBPhone.Trim = "" And sDBEmail.Trim = "" Then
					' no details, hide the forgot pass
					lnkcmdChangePassword.Visible = False
					litPipe1.Visible = False
					lnkcmdFogottenPassword.Visible = False
					litPipe2.Visible = False
					lnkcmdRequestNewPassword.Visible = True
					lnkcmdManageMyAccount.Visible = True
				Else
					' there are some details - hid eht erequest now password
					lnkcmdChangePassword.Visible = True
					litPipe1.Visible = True
					lnkcmdFogottenPassword.Visible = True
					litPipe2.Visible = False
					lnkcmdRequestNewPassword.Visible = False
					lnkcmdManageMyAccount.Visible = True

				End If
			End If
		Catch ex As Exception

		End Try


		Try
			' force user to change their password
			If Session("action") = "TemporaryPassword" Then
				iChangePassword = 1
			End If
		Catch ex As Exception
		End Try


		Dim ReportName As String = ""
		Dim ReportPath As String = ""




		If WindowsPassword.Trim.Length > 0 Then


			SaveToLog(sWinUsername & " Login_____GetReportDetails_Impersonation_____ [" & sWinUsername & "] WindowsUsername=" & WindowsUsername & "; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & ": inApplicationFK=" & inApplicationFK & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)

			' --- with manual impersonation ------
			' --- with manual impersonation ------
			' --- with manual impersonation ------
			Try
				' IF YOU USE IMPERSONATION HERE, THEN YOU must USE INLINE sql TO "Execute As Login = 'evolutionjobs\patels';"
				Dim objImpersonation As New UserImpersonation()
				If (objImpersonation.ImpersonateValidUser(WindowsUsername, WindowsPassword, WindowsDomain)) Then

					' ---- impoersonate user -------
					' ---- impoersonate user -------
					' ---- impoersonate userwas PasswordAdminEdConnectionString -------
					Using lib1 As New clsDB(PasswordAdminEdConnectionString)
						PasswordProtect = lib1.GetReportDetails_Impersonation(inApplicationFK, ReportPath, ReportName)
					End Using
					SaveToLog(sWinUsername & " Login_____PASSED_IMPERSONATION_____ [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; PasswordProtect=" & PasswordProtect & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath & "; Request.Url.ToString=" & Request.Url.ToString)

					' ---- impoersonate user -------
					' ---- impoersonate user -------
					' ---- impoersonate user -------




					objImpersonation.UndoImpersonation()
				Else
					objImpersonation.UndoImpersonation()
					Response.Redirect("Error.aspx?e=User+does+not+have+enough+permissions+to+perform+the+task")

					'Throw New Exception("User does not have enough permissions to perform the task")
				End If
			Catch ex As Exception
				' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'Response.Redirect("Error.aspx?e=" & ex.Message.Replace(" ", "+"))
				'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			End Try
			' --- with manual impersonation ------
			' --- with manual impersonation ------
			' --- with manual impersonation ------


		Else
			SaveToLog(sWinUsername & " Login_____GetReportDetails (May fail)_____ [" & sWinUsername & "]; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " : inApplicationFK=" & inApplicationFK & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)

			Try
				Using lib1 As New clsDB(BIWindowsConnectionString)
					PasswordProtect = lib1.GetReportDetails(inApplicationFK, ReportPath, ReportName)
				End Using
			Catch ex As Exception
				' this will error with NTAuthority/anonymous, so use this to set the session vars for impersonation
				Response.Redirect("LoginWindows.aspx?appfk=" & inApplicationFK)
			End Try

		End If






		'Try

		'	Using lib1 As New clsDB(BIWindowsConnectionString)
		'		PasswordProtect = lib1.GetReportDetails(inApplicationFK, ReportPath, ReportName)
		'	End Using

		'Catch ex As Exception
		'End Try


		SaveToLog("LOGIN_LOAD [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; iChangePassword=" & iChangePassword & "; PasswordProtect=" & PasswordProtect & "; SMSDomain=" & SMSDomain & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)

		If PasswordProtect = False Then
			' redirect to report page and show the report
			If inApplicationFK > 0 Then
				' only redirect if there is an appfk set

				SaveToLog("LOGIN_LOAD___REDIRECT____ [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; REDIRECT_TO: report.aspx?appfk=" & inApplicationFK)

				Response.Redirect("report.aspx?appfk=" & inApplicationFK)
				Exit Sub
			End If
		End If

		lblURL.Text = "" '"<a href='Report.aspx'>Proceed to My Reports</a>"
		'lblURL.Text = "<a href='http://EvoWeb3'>MY Reports </a>"
		'lblWindowsUsername.Text = HttpContext.Current.User.Identity.Name

		If Me.IsPostBack = False Then

			If inApplicationFK > 0 Then

				If iChangePassword > 0 Then
					pnlLogin.Visible = False
					pnlChangePassword.Visible = True
				Else

					pnlLogin.Visible = True
					pnlChangePassword.Visible = False

					Try
						' autologin from the New Password page
						Try
							If Session("NewPassword") <> "" Then
								txtPassword.Text = Session("NewPassword")
							End If
						Catch ex As Exception

						End Try

						DoLogin(False)
					Catch ex As Exception
					End Try

				End If

			Else
				If iChangePassword > 0 Then
					pnlLogin.Visible = False
					pnlChangePassword.Visible = True
				Else
					pnlLogin.Visible = False
					pnlChangePassword.Visible = False

					' autologin from the New Password page
					Try
						If Session("NewPassword") <> "" Then
							txtPassword.Text = Session("NewPassword")
						End If
					Catch ex As Exception

					End Try

					'lblStatus.Text = "ERROR: Application value missing"
				End If


			End If

		End If

		If pnlChangePassword.Visible Then
			Try
				If Session("ChangePasswordMessage") <> "" Then
					lblSpecialMessage.Text = Session("ChangePasswordMessage")
					lblOldPassword.Text = "Temporary Password:"
				End If
			Catch ex As Exception
			End Try

			' preset the Old Password
			lblOldPassword.Visible = True
			txtOldPassword.Visible = True
			Try
				If Session("NewPassword") <> "" Then
					txtOldPassword.Text = Session("NewPassword")
					lblTemp.Text = Session("NewPassword")
					lblOldPassword.Visible = False
					txtOldPassword.Visible = False
				End If
			Catch ex As Exception
			End Try

		End If


		If pnlLogin.Visible = False And pnlChangePassword.Visible = False Then
			pnlDefault.Visible = True
		End If


		' clear sessions ONY if all is OK: can't do this here because if they enter data worngly, the  it resets everything because the sessions are lost
		'ClearSessions()

	End Sub

	Protected Sub cmdSubmit_Click(sender As Object, e As System.EventArgs) Handles cmdSubmit.Click
        DoLogin(False)
    End Sub

	Public Sub DoLogin(ByVal inFromPassword As Boolean)

		If txtPassword.Text.Trim = "" Then
			lblStatus.Text = "ERROR: Please enter a password"
			Exit Sub
		End If
		Dim iChangePassword As Integer = 0
		Dim inApplicationFK As Integer = 0
		Try
			If Request("appfk") <> "" Then
				inApplicationFK = CInt(Request("appfk"))
			End If
			If Request("cp") <> "" Then
				iChangePassword = CInt(Request("cp"))
			End If
		Catch ex As Exception
		End Try

		Dim sWinUsername As String = HttpContext.Current.User.Identity.Name
		SaveToLog("Login_DoLogin [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK)

		If inApplicationFK <= 0 Then

			If iChangePassword > 0 Then
				' allow through to login and show links
			Else
				lblStatus.Text = "ERROR: no application specified"
				Exit Sub
			End If

		End If


		Dim PasswordStatus As String = ""
		Dim sResultFK As String = ""
		Dim sReportURL As String = ""
		Dim sSeed As String = ""

		Try

			Dim impersonationContext As System.Security.Principal.WindowsImpersonationContext
			Dim currentWindowsIdentity As System.Security.Principal.WindowsIdentity

			currentWindowsIdentity = CType(User.Identity, System.Security.Principal.WindowsIdentity)
			impersonationContext = currentWindowsIdentity.Impersonate()

			Dim dtCons As New System.Data.DataTable
			'Using lib1 As New clsDB(BIWindowsConnectionString)
			'	dtCons = lib1.CheckPassword(txtPassword.Text, inApplicationFK)
			'End Using



			If WindowsPassword.Trim.Length > 0 Then

				SaveToLog(sWinUsername & " Login-WindowsPassword_____WindowsPassword=" & WindowsPassword)

				' --- with manual impersonation ------
				' --- with manual impersonation ------
				' --- with manual impersonation ------
				Try
					' IF YOU USE IMPERSONATION HERE, THEN YOU must USE INLINE sql TO "Execute As Login = 'evolutionjobs\patels';"
					Dim objImpersonation As New UserImpersonation()
					If (objImpersonation.ImpersonateValidUser(WindowsUsername, WindowsPassword, WindowsDomain)) Then

						' ---- impoersonate user -------
						' ---- impoersonate user -------
						' ---- impoersonate user was PasswordAdminEdConnectionString-------
						SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____PRE_ImpersonateValidUser")
						Using lib1 As New clsDB(PasswordAdminEdConnectionString)
							dtCons = lib1.CheckPassword_Impersonation(txtPassword.Text, inApplicationFK)
						End Using

						SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____POST_ImpersonateValidUser")
						'SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____POST_ImpersonateValidUser_____ [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; Request.Url.ToString=" & Request.Url.ToString)

						' ---- impoersonate user -------
						' ---- impoersonate user -------
						' ---- impoersonate user -------




						'objImpersonation.UndoImpersonation()
					Else
						'objImpersonation.UndoImpersonation()
						Response.Redirect("Error.aspx?e=User+does+not+have+enough+permissions+to+perform+the+task")

						'Throw New Exception("User does not have enough permissions to perform the task")
					End If
				Catch ex As Exception
					' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
					' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
					' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
					'Response.Redirect("Error.aspx?e=" & ex.Message.Replace(" ", "+"))
					'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				End Try
				' --- with manual impersonation ------
				' --- with manual impersonation ------
				' --- with manual impersonation ------

			Else

				Using lib1 As New clsDB(BIWindowsConnectionString)
					dtCons = lib1.CheckPassword(txtPassword.Text, inApplicationFK)
				End Using


			End If


			SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____dtCons_____ [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; Request.Url.ToString=" & Request.Url.ToString)


			If dtCons IsNot Nothing Then

				If dtCons.Rows.Count > 0 Then
					SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____dtCons.Rows.Count=" & dtCons.Rows.Count)

					Dim iCount As Integer = 0
					Dim row As DataRow
					For Each row In dtCons.Rows

						If Not IsDBNull(row("PasswordStatus")) Then
							PasswordStatus = row("PasswordStatus")
						End If
						If Not IsDBNull(row("ResultFK")) Then
							sResultFK = row("ResultFK")
						End If
						If Not IsDBNull(row("ReportURL")) Then
							sReportURL = row("ReportURL")
						End If
						If Not IsDBNull(row("Seed")) Then
							sSeed = row("Seed")
						End If
						' & "&RandNum=" & First(Fields!Seed.Value, "GiveItAGo")
						iCount += 1
					Next
				End If

			End If

		Catch ex As Exception
			SaveToLog("Login_DoLogin ****EXCEPTION**** ERROR: checking password [" & sWinUsername & "]: PasswordStatus=" & PasswordStatus & "; sResultFK=" & sResultFK & "; sSeed=" & sSeed & "; sReportURL=" & sReportURL & "; EX=" & ex.Message & ex.StackTrace)

			lblStatus.Text = ("ERROR: checking password: " & ex.Message)

		Finally
		End Try


		SaveToLog("Login_DoLogin [" & sWinUsername & "]: PasswordStatus=" & PasswordStatus & "; sResultFK=" & sResultFK & "; sSeed=" & sSeed & "; sReportURL=" & sReportURL)


		If sResultFK = "5" Then
			' sResultFK = "5"
			Dim sURL As String = sReportURL & "&RandNum=" & sSeed
			If sReportURL.Trim <> "" Then
				sURL = sReportURL & "&RandNum=" & sSeed
			End If

			' thi falls foul of popup blockers
			'ClientScript.RegisterStartupScript(Me.Page.GetType(), "", "window.open('" & sURL & "','LoginRedirect','height=600,width=800');", True)

			' http://EvoWeb3/Wageslogin.aspx?appfk=22
			' http://EvoWeb3/Wageslogin.aspx?appfk=22
			'http://localhost:54391/WagesLogin/WagesPage.aspx

			' if this is coming from the old MyBI reports page using :8099, then redirect using the reportsURL

			If inFromPassword Then
				' do not redirect =- show links 

				Dim sRedURL As String = ""
				If Request.Url.ToString.Contains(":8099") Then
					sRedURL = sURL
				Else
					sRedURL = "Report.aspx?RandNum=" & sSeed & "&appfk=" & inApplicationFK
				End If


				If inApplicationFK > 0 Then


					lblURL.Text = "<a href='" & sRedURL & "'>Proceed to Report &gt;&gt;</a>"

					If lblStatus.Text = "Password changed successfully" Then

						' if we are coming from the change password panel, then the succes message will still be three
						' hence, if it was changed succesfuly, then redirect them to the report????

						pnlChangePassword.Visible = False

						' clear sessions ONY if all is OK
						ClearSessions()

						'Response.Redirect(sRedURL)

					End If

				Else

					lblURL.Text = "" '"<a href='Report.aspx'>Proceed to Report &gt;&gt;</a>"
					'lblURL.Text = "<a href='http://EvoWeb3'>MY Reports </a>"


					pnlChangePassword.Visible = False

					ClearSessions()


				End If

			Else

				Dim sRedURL As String = ""
				If Request.Url.ToString.Contains(":8099") Then
					sRedURL = sURL
				Else
					sRedURL = "Report.aspx?RandNum=" & sSeed & "&appfk=" & inApplicationFK
				End If
				'lblStatus.Text = "sRedURL = " & sRedURL

				SaveToLog("Login_DoLogin [" & sWinUsername & "]: PasswordStatus=" & PasswordStatus & "; sResultFK=" & sResultFK & "; sSeed=" & sSeed & "; Request.Url.ToString=" & Request.Url.ToString & "; sRedURL=" & sRedURL)





				ClearSessions()






				Response.Redirect(sRedURL)
			End If



			'If sResultFK = "5" Then
			'    ' http://evosvr05/ReportServer/Pages/ReportViewer.aspx?%2fWagesPages%2fOpeningPage&rs:Command=Render&RandNum=517251
			'    '"http://evosvr05/ReportServer/Pages/ReportViewer.aspx?%2fWagesPages%2fOpeningPage&rs:Command=Render&RandNum="
			'Else
			'    lblStatus.Text = PasswordStatus
			'End If


		ElseIf sResultFK = "3" Then
			' Please change your temporary reports password before trying to do anything else. 

			If PasswordStatus.Trim <> "change your temporary" Then
				If PasswordStatus.Trim.Contains("change your temporary") Then
					pnlLogin.Visible = False
					pnlChangePassword.Visible = True
					txtOldPassword.Text = txtPassword.Text
				End If
				lblStatus.Text = PasswordStatus
			Else
				lblStatus.Text = "ERROR: no Report URL returned"
			End If

		Else

			If PasswordStatus.Trim <> "" Then
				lblStatus.Text = PasswordStatus
			Else
				lblStatus.Text = "ERROR: no Report URL returned"
			End If


		End If



		'Dim sWinUsername As String = HttpContext.Current.User.Identity.Name

		'Dim lWinUsernameBIUserFK As Long = 0
		'Dim sWinObj As New Object

		'Using lib1 As New clsDB(ConnectionString)
		'    sWinObj = lib1.RunSQLScalar("SELECT Ed.[dbo].[udf_SP_WIN_CheckWindowsUsername]('" & sWinUsername & "')")
		'End Using

		'If sWinObj IsNot Nothing AndAlso Not IsDBNull(sWinObj) Then
		'    Long.TryParse(sWinObj, lWinUsernameBIUserFK)
		'End If

		'If lWinUsernameBIUserFK > 0 Then
		'    lblStatus.Text = "UserID=" & lWinUsernameBIUserFK
		'End If
	End Sub

	Public Function SwapName(ByVal strName As String) As String
        Dim int2ndStart As Integer
        Dim int3rdStart As Integer
        Dim strFirstName As String
        Dim strLastName As String
        Dim strThirdName As String

        int2ndStart = InStr(1, strName, " ", vbTextCompare)
        int3rdStart = InStr(int2ndStart + 1, strName, " ", vbTextCompare)

        If int3rdStart > 0 Then
            strFirstName = Left(strName, int2ndStart)
            strLastName = Mid(strName, int2ndStart + 1, (int3rdStart - int2ndStart) - 1)
            strThirdName = Right(strName, Len(strName) - int3rdStart)
            strFirstName = Trim(strFirstName)
            strLastName = Trim(strLastName)

            SwapName = strLastName & " " & strThirdName & " " & strFirstName
            'SwapName = strThirdName & " " & strLastName & " " & strFirstName
        Else
            int2ndStart = InStr(1, strName, " ", vbTextCompare)
            strFirstName = Left(strName, int2ndStart)
            strLastName = Right(strName, Len(strName) - Len(strFirstName))
            strFirstName = Trim(strFirstName)
            strLastName = Trim(strLastName)

            SwapName = strLastName & " " & strFirstName
        End If

        Return SwapName

    End Function


    Protected Sub cmdTest_Click(sender As Object, e As System.EventArgs) Handles cmdTest.Click

        If txtTest.Text.Trim <> "" Then
            lblTest.Text = SwapName(txtTest.Text.Trim)
        End If

    End Sub

	Protected Sub lnkcmdChangePassword_Click(sender As Object, e As System.EventArgs) Handles lnkcmdChangePassword.Click
		ClearSessions()
		pnlLogin.Visible = False
		pnlChangePassword.Visible = True
		pnlDefault.Visible = False
	End Sub

	Protected Sub cmdChangePassword_Click(sender As Object, e As System.EventArgs) Handles cmdChangePassword.Click
		'ClearSessions()
		Dim sWinUsername As String = HttpContext.Current.User.Identity.Name

		Dim sOldPassword As String = ""
		If lblTemp.Text.Trim <> "" Then
			sOldPassword = lblTemp.Text.Trim
		Else
			sOldPassword = txtOldPassword.Text.Trim
		End If

		If (sOldPassword.Trim <> "") And (txtNewPassword1.Text.Trim <> "") And (txtNewPassword2.Text.Trim <> "") Then

			If (txtNewPassword1.Text = txtNewPassword2.Text) Then
				' run the change password routine
				' and run the login sproc
				Dim sResult As String = ""
				Dim sResultFK As String = ""

				If WindowsPassword.Trim <> "" Then

					Using lib1 As New clsDB(PasswordAdminEdConnectionString)
						sResultFK = lib1.ChangePassword_Impersonate(
						  sResult,
						  sOldPassword.Trim,
						  txtNewPassword1.Text.Trim,
						  txtNewPassword2.Text.Trim
						 )
					End Using
				Else
					Using lib1 As New clsDB(BIWindowsConnectionString)
						sResultFK = lib1.ChangePassword(
						  sResult,
						  sOldPassword.Trim,
						  txtNewPassword1.Text.Trim,
						  txtNewPassword2.Text.Trim
						 )
					End Using
				End If




				If sResultFK = "11" Then
					lblStatus.Text = "Password changed successfully"

					txtPassword.Text = txtNewPassword1.Text
					DoLogin(True)

					ClearSessions()

				Else
					lblStatus.Text = "ERROR: " & sResult
				End If

			Else
				lblStatus.Text = "ERROR: Passwords do not match"
			End If

		Else
			lblStatus.Text = "ERROR: Please fill in all fields"
        End If

    End Sub


    Protected Sub lnkbackToLoginPage_Click(sender As Object, e As System.EventArgs) Handles lnkbackToLoginPage.Click
        pnlLogin.Visible = True
        pnlChangePassword.Visible = False
        'Dim iChangePassword As Integer = 0
        'Dim inApplicationFK As Integer = 0
        'Try
        '    If Request("appfk") <> "" Then
        '        inApplicationFK = CInt(Request("appfk"))
        '    End If
        '    If Request("cp") <> "" Then
        '        iChangePassword = CInt(Request("cp"))
        '    End If
        'Catch ex As Exception
        'End Try
        'Dim sStringToRemove As String = "cp=" & iChangePassword
        'Response.Redirect(Request.Url.ToString.Replace(sStringToRemove, ""), True)
    End Sub

	Protected Sub lnkcmdFogottenPassword_Click(sender As Object, e As EventArgs) Handles lnkcmdFogottenPassword.Click
		'ClearSessions()
		' NOTE: this is a different page, but the action session value is same as NewPassword, so the copy of code is used, but changed a little
		Session("action") = "NewPasswordStep2"
		' make sure that when they return to the login page with a temporary password, that they are forced to change their password
		' hence they have to be shown the change password screen

		Dim sURL As String = Request.Url.ToString
		If sURL.Trim.ToLower.Contains("cp=") = False Then
			If sURL.Trim.ToLower.Contains("?") Then
				sURL &= "&cp=1"
			Else
				sURL &= "?cp=1"
			End If

		End If
		Session("LoginPage") = Request.Url.ToString
		Session("AutoLoginLinkChangePassword") = sURL
		Response.Redirect("ForgottenPassword.aspx")
	End Sub


	Protected Sub lnkcmdRequestNewPassword_Click(sender As Object, e As EventArgs) Handles lnkcmdRequestNewPassword.Click
		'ClearSessions()
		' NOTE: this is a different page, but the action session value is same as NewPassword, so the copy of code is used, but changed a little
		Session("action") = "NewPasswordStep2"
		' make sure that when they return to the login page with a temporary password, that they are forced to change their password
		' hence they have to be shown the change password screen

		Dim sURL As String = Request.Url.ToString
		If sURL.Trim.ToLower.Contains("cp=") = False Then
			If sURL.Trim.ToLower.Contains("?") Then
				sURL &= "&cp=1"
			Else
				sURL &= "?cp=1"
			End If

		End If
		Session("LoginPage") = Request.Url.ToString
		Session("AutoLoginLinkChangePassword") = sURL
		Response.Redirect("ForgottenPassword.aspx")

		'Session("action") = "NewPasswordStep1"
		'Session("LoginPage") = Request.Url.ToString
		'Session("AutoLoginLinkChangePassword") = Request.Url.ToString
		'Response.Redirect("ForgottenPassword.aspx")
	End Sub


	Protected Sub lnkcmdManageMyAccount_Click(sender As Object, e As EventArgs) Handles lnkcmdManageMyAccount.Click
		'ClearSessions()
		' NOTE: this is a different page, but the action session value is same as NewPassword, so the copy of code is used, but changed a little
		Session("action") = "NewPasswordStep2"
		' make sure that when they return to the login page with a temporary password, that they are forced to change their password
		' hence they have to be shown the change password screen

		Dim sURL As String = Request.Url.ToString
		If sURL.Trim.ToLower.Contains("cp=") = False Then
			If sURL.Trim.ToLower.Contains("?") Then
				sURL &= "&cp=1"
			Else
				sURL &= "?cp=1"
			End If

		End If
		Session("LoginPage") = "From_Manage_My_Account" 'Request.Url.ToString
		Session("AutoLoginLinkChangePassword") = sURL
		Response.Redirect("ForgottenPassword.aspx")
	End Sub


	Protected Sub lnkcmdMyReports_Click(sender As Object, e As EventArgs) Handles lnkcmdMyReports.Click
		Response.Redirect("Report.aspx")
	End Sub
	Protected Sub lnkcmdLogin_Click(sender As Object, e As EventArgs) Handles lnkcmdLogin.Click

		Dim sURL As String = ""

		Try
			If Session("AutoLoginLinkChangePassword") IsNot Nothing Then
				sURL = Session("AutoLoginLinkChangePassword")
			End If
		Catch ex As Exception
		End Try

		Dim sReplaceCPString As String = "&cp=1"
		Try
			If Session("SendNewPassword") IsNot Nothing Then
				' musy change password, so it needs the &cp=1 parameter, hence don;t remove it!!!!
				sReplaceCPString = ""
			End If
		Catch ex As Exception
		End Try

		If sURL.Trim <> "" Then

		Else
			Dim sCurrentURL As String = lblCurrentURL.Text
			If sCurrentURL.Trim <> "" Then
				sURL = sCurrentURL

			Else
				sURL = Request.Url.ToString

			End If

		End If

		If sURL.Trim <> "" Then
			If sReplaceCPString.Trim <> "" Then
				sURL = sURL.Replace(sReplaceCPString, "")
			End If
		End If



		Response.Redirect(sURL)
	End Sub

End Class
