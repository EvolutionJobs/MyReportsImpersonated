
Imports functions
Imports System.Net
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports libEvolutionIntranet
Imports System.Globalization
Imports DevExpress.Web.ASPxGridView
Imports System.Drawing





Partial Class UserControls_ucContactDetails
	Inherits System.Web.UI.UserControl


	Public Property ExistingMobileNumber As String = ""
	Public Property ExistingEmailAddress As String = ""
	Public Property WindowsUsername As String = ""

	Protected Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init



	End Sub


	Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
		' +44 9999 999999
		' leaving off the first zero if there is one
		' 61 --- Australia
		' 49 --- Germany
		' 65 --- Singapore
		' 44 --- United Kingdom

		'lblEditIntro.Text = "Please enter you new mobile and/or your new mobile number below"

		lblErrors.Text = ""


		If Not Me.IsPostBack Then

			ddlstCodes.Items.Clear()
			Dim li0 As New ListItem("----SELECT-----", "")
			Dim li1 As New ListItem("Australia: +61", "+61")
			Dim li2 As New ListItem("Germany: +49", "+49")
			Dim li3 As New ListItem("Singapore: +65", "+65")
			Dim li4 As New ListItem("UK: +44", "+44")

			ddlstCodes.Items.Add(li0)
			ddlstCodes.Items.Add(li1)
			ddlstCodes.Items.Add(li2)
			ddlstCodes.Items.Add(li3)
			ddlstCodes.Items.Add(li4)

			Try
				ddlstCodes.SelectedIndex = 0
			Catch ex As Exception
			End Try


			If ExistingMobileNumber.Trim <> "" Then

				Dim sMobileWithoutIntCode As String = ""
				' extract the int code and show it in the drop down
				Dim iIndex As Integer = 0
				For ix As Integer = 0 To ddlstCodes.Items.Count - 1
					Dim li As ListItem = ddlstCodes.Items(ix)
					If li.Value <> "" Then
						If ExistingMobileNumber.Contains(li.Value) Then
							sMobileWithoutIntCode = ExistingMobileNumber.Replace(li.Value, "")
							iIndex = ix
						End If
					End If
				Next

				txtEditPhone.Text = sMobileWithoutIntCode.Trim

				Try
					ddlstCodes.SelectedIndex = iIndex
				Catch ex As Exception
				End Try

			End If
			If ExistingEmailAddress.Trim <> "" Then
				txtEditEmail.Text = ExistingEmailAddress.Trim
			End If

		End If

	End Sub




	Protected Sub cmdUpdatePhone_Click(sender As Object, e As EventArgs) Handles cmdUpdatePhone.Click
		' go to the relevant verify page
		lblErrors.Text = ""

		Dim sLoginError As String = ""
		sLoginError = DoLogin(txtPasswordPhone.Text)

		If sLoginError.Trim <> "" Then
			lblErrors.Text = sLoginError
		Else
			Dim sCode As String = ""
			If ddlstCodes.SelectedIndex <> -1 Then
				If ddlstCodes.SelectedItem.Value <> "" Then
					sCode = ddlstCodes.SelectedItem.Value
				End If
			End If

			If sCode.Trim <> "" Then

				Dim sNumberEntered As String = txtEditPhone.Text.Trim
				If sNumberEntered.Substring(0, 1) = "0" Then
					sNumberEntered = sNumberEntered.Substring(1)
				End If

				Dim sPhone As String = sCode & sNumberEntered

				Dim sUPDError As String = UpdateUserPhone(WindowsUsername, sPhone)

				If sUPDError.Trim <> "" Then
					lblErrors.Text = "ERROR: Updating phone: " & sUPDError
				Else
					' because thii is an update, send new verify code
					Session("PhoneFinalVerificationCode") = ""

					SaveToLog(WindowsUsername & "; cmdUpdatePhone_Click - SendVerificationSMS:")

					Dim sSMSError As String = SendVerificationSMS(sPhone, WindowsUsername, True)
					' 4. redirect to code entry page
					If sSMSError.Trim <> "" Then
						lblErrors.Text = sSMSError.Trim
					Else
						Session("action") = "NewPasswordStep3updated"
						Response.Redirect(Request.Url.ToString)
					End If
				End If


			Else
				lblErrors.Text = "ERROR: Please select an International Code"
			End If
		End If


	End Sub


	Protected Sub cmdUpdateEmail_Click(sender As Object, e As EventArgs) Handles cmdUpdateEmail.Click
		' go to the relevant verify page

		lblErrors.Text = ""


		If IsValidEmail(txtEditEmail.Text.Trim) = False Then
			lblErrors.Text = "ERROR: Please enter a valid email address"
			Exit Sub
		End If

        'If IsBackup() Then
        '	' backup
        'Else
        '	' only enforce this on live
        '	' make sure we don't allow any company domains
        '	Dim lDomainID As Long = 0
        '	Using lib1 As New clsDB(BIWindowsConnectionString)
        '		lDomainID = lib1.CheckIfDomainIsUsedInEmail(txtEditEmail.Text.Trim)
        '	End Using

        '	If lDomainID > 0 Then
        '		lblErrors.Text = "ERROR: You cannot use your work email address (for security)"
        '		Exit Sub
        '	End If
        'End If

        EdConnectionString = ConfigurationManager.ConnectionStrings("EdConnectionString").ConnectionString

        ' only enforce this on live
        ' make sure we don't allow any company domains
        Dim lDomainID As Long = 0
        Using lib1 As New clsDB(EdConnectionString)
            lDomainID = lib1.CheckIfDomainIsUsedInEmail(txtEditEmail.Text.Trim)
        End Using

        If lDomainID > 0 Then
			lblErrors.Text = "ERROR: You cannot use your work email address (for security)"
			Exit Sub
		End If


		Dim sLoginError As String = ""
		sLoginError = DoLogin(txtPasswordEmail.Text)


		If sLoginError.Trim <> "" Then
			lblErrors.Text = sLoginError
		Else
			Dim sUPDError As String = UpdateUserEmail(WindowsUsername, txtEditEmail.Text.Trim)
			If sUPDError.Trim <> "" Then
				lblErrors.Text = "ERROR: Updating email: " & sUPDError
			Else
				' because thii is an update, send new verify code
				Session("EmailFinalVerificationCode") = ""

				Dim sSMSError As String = SendVerificationEmail(txtEditEmail.Text.Trim, WindowsUsername, True)
				' 4. redirect to code entry page
				If sSMSError.Trim <> "" Then
					lblErrors.Text = sSMSError.Trim
				Else
					Session("action") = "NewPasswordStep3updated"
					Response.Redirect(Request.Url.ToString)
				End If
			End If

		End If


	End Sub

	Public Function DoLogin(ByVal inPassword As String) As String
		Dim sRetError As String = ""
		If inPassword.Trim = "" Then
			sRetError = "ERROR: Please enter a password"
			Return sRetError
		End If
		Dim inApplicationFK As Integer = 22 ' hardcode the wages pages appfk to use
		Try
			If Request("appfk") <> "" Then
				inApplicationFK = CInt(Request("appfk"))
			End If
		Catch ex As Exception
		End Try

		SaveToLog("Login_DoLogin_usContactDetails [" & WindowsUsername & "]: inApplicationFK=" & inApplicationFK)

		Dim PasswordStatus As String = ""
		Dim sResultFK As String = ""
		Dim sReportURL As String = ""
		Dim sSeed As String = ""

		Try

			Dim impersonationContext As System.Security.Principal.WindowsImpersonationContext
			Dim currentWindowsIdentity As System.Security.Principal.WindowsIdentity

			currentWindowsIdentity = CType(HttpContext.Current.User.Identity, System.Security.Principal.WindowsIdentity)
			impersonationContext = currentWindowsIdentity.Impersonate()


			Dim sWinUsername As String = HttpContext.Current.User.Identity.Name

			Dim dtCons As New System.Data.DataTable
			'Using lib1 As New clsDB(BIWindowsConnectionString)
			'	dtCons = lib1.CheckPassword(inPassword, inApplicationFK)
			'End Using


			If WindowsPassword.Trim.Length > 0 Then

				SaveToLog(sWinUsername & " ucContactDetails-WindowsPassword_____WindowsPassword=" & WindowsPassword)



				' ---- impoersonate user -------
				' ---- impoersonate user -------
				' ---- impoersonate user was PasswordAdminEdConnectionString-------
				SaveToLog(sWinUsername & " ucContactDetails-CheckPassword_Impersonation_____PRE_ImpersonateValidUser")
				Using lib1 As New clsDB(PasswordAdminEdConnectionString)
					dtCons = lib1.CheckPassword_Impersonation(inPassword, inApplicationFK)
				End Using

				SaveToLog(sWinUsername & " ucContactDetails-CheckPassword_Impersonation_____POST_ImpersonateValidUser")
				'SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____POST_ImpersonateValidUser_____ [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; Request.Url.ToString=" & Request.Url.ToString)

				' ---- impoersonate user -------
				' ---- impoersonate user -------
				' ---- impoersonate user -------


				'' --- with manual impersonation ------
				'' --- with manual impersonation ------
				'' --- with manual impersonation ------
				'Try
				'	' IF YOU USE IMPERSONATION HERE, THEN YOU must USE INLINE sql TO "Execute As Login = 'evolutionjobs\patels';"
				'	Dim objImpersonation As New UserImpersonation()
				'	If (objImpersonation.ImpersonateValidUser(WindowsUsername, WindowsPassword, WindowsDomain)) Then

				'		' ---- impoersonate user -------
				'		' ---- impoersonate user -------
				'		' ---- impoersonate user was PasswordAdminEdConnectionString-------
				'		SaveToLog(sWinUsername & " ucContactDetails-CheckPassword_Impersonation_____PRE_ImpersonateValidUser")
				'		Using lib1 As New clsDB(PasswordAdminEdConnectionString)
				'			dtCons = lib1.CheckPassword_Impersonation(inPassword, inApplicationFK)
				'		End Using

				'		SaveToLog(sWinUsername & " ucContactDetails-CheckPassword_Impersonation_____POST_ImpersonateValidUser")
				'		'SaveToLog(sWinUsername & " Login-CheckPassword_Impersonation_____POST_ImpersonateValidUser_____ [" & sWinUsername & "]: inApplicationFK=" & inApplicationFK & "; Request.Url.ToString=" & Request.Url.ToString)

				'		' ---- impoersonate user -------
				'		' ---- impoersonate user -------
				'		' ---- impoersonate user -------




				'		'objImpersonation.UndoImpersonation()
				'	Else
				'		'objImpersonation.UndoImpersonation()
				'		Response.Redirect("Error.aspx?e=User+does+not+have+enough+permissions+to+perform+the+task")

				'		'Throw New Exception("User does not have enough permissions to perform the task")
				'	End If
				'Catch ex As Exception
				'	' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'	' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'	' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'	'Response.Redirect("Error.aspx?e=" & ex.Message.Replace(" ", "+"))
				'	'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				'End Try
				'' --- with manual impersonation ------
				'' --- with manual impersonation ------
				'' --- with manual impersonation ------

			Else

				Using lib1 As New clsDB(BIWindowsConnectionString)
					dtCons = lib1.CheckPassword(inPassword, inApplicationFK)
				End Using


			End If


			If dtCons IsNot Nothing Then

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

		Catch ex As Exception
			sRetError = ("ERROR: checking password: " & ex.Message)
		Finally
		End Try

		If sResultFK = "5" Then
			' sResultFK = "5"
			'OK
		Else

			If PasswordStatus.Trim <> "" Then
				sRetError = PasswordStatus
			Else
				sRetError = "ERROR: Authentication failed (No Report URL returned)"
			End If

		End If

		SaveToLog("Login_DoLogin_usContactDetails [" & WindowsUsername & "]: PasswordStatus=" & PasswordStatus & "; sResultFK=" & sResultFK & "; sSeed=" & sSeed & "; sReportURL=" & sReportURL & "; sRetError=" & sRetError)

		Return sRetError

	End Function






	Protected Sub lnkSMSTest_Click(sender As Object, e As EventArgs) Handles lnkSMSTest.Click
		Dim sCode As String = ""
		If ddlstCodes.SelectedIndex <> -1 Then
			If ddlstCodes.SelectedItem.Value <> "" Then
				sCode = ddlstCodes.SelectedItem.Value
			End If
		End If

		If sCode.Trim <> "" Then

			Dim sNumberEntered As String = txtEditPhone.Text.Trim
			If sNumberEntered.Substring(0, 1) = "0" Then
				sNumberEntered = sNumberEntered.Substring(1)
			End If

			Dim PhoneEntered As String = sCode & sNumberEntered

			Dim sSMSError As String = ""
			sSMSError = SendSMS(PhoneEntered.Trim.Replace(" ", "").Replace("+", "00") & "@sms.textapp.net", "TEST SMS From MyReports")

			SaveToLog("lnkSMSTest_Click___@sms.textapp.net___ [" & WindowsUsername & "]: PhoneEntered=" & PhoneEntered & "; sSMSError=" & sSMSError)

			If sSMSError.Trim <> "" Then
				lblSMSTestError.Text = sSMSError.Trim
			Else
				lblSMSTestError.Text = "Please check if a TEST SMS has been sent to your phone (it may take a few minutes)"
			End If

		Else
			lblSMSTestError.Text = "ERROR: Please select an International Code"
		End If

	End Sub







	'Protected Sub UpdateBoth_Click()

	'	Dim sLoginError As String = ""
	'	sLoginError = DoLogin()

	'	If sLoginError.Trim <> "" Then
	'		lblErrors.Text = sLoginError
	'	Else
	'		' do the updates

	'		Dim sUpdPhoneError As String = ""
	'		Dim sUpdEmailError As String = ""

	'		' --------------- phone ----------------------
	'		Dim sCode As String = ""
	'		If ddlstCodes.SelectedIndex <> -1 Then
	'			If ddlstCodes.SelectedItem.Value <> "" Then
	'				sCode = ddlstCodes.SelectedItem.Value
	'			End If
	'		End If

	'		If sCode.Trim <> "" Then

	'			Dim sPhone As String = sCode & txtEditPhone.Text.Trim

	'			Dim sUPDErrorPH As String = UpdateUserPhone(WindowsUsername, sPhone)

	'			If sUPDErrorPH.Trim <> "" Then
	'				sUpdPhoneError &= "ERROR: Updating phone: " & sUPDErrorPH & "<br />"
	'			Else
	'				Dim sSMSError As String = SendVerificationSMS(sPhone, WindowsUsername)
	'				' 4. redirect to code entry page
	'				If sSMSError.Trim <> "" Then
	'					sUpdPhoneError &= sSMSError.Trim & "<br />"
	'				Else
	'					Session("action") = "NewPasswordStep3updated"
	'					Response.Redirect(Request.Url.ToString)
	'				End If
	'			End If

	'		Else
	'			sUpdPhoneError &= "ERROR: Please select an International Code"
	'		End If
	'		' --------------- phone ----------------------


	'		' --------------- email ----------------------
	'		Dim sUPDErrorEM As String = UpdateUserEmail(WindowsUsername, txtEditEmail.Text.Trim)
	'		If sUPDErrorEM.Trim <> "" Then
	'			sUpdEmailError = "ERROR: Updating email: " & sUPDErrorEM & "<br />"
	'		Else
	'			Dim sSMSError As String = SendVerificationEmail(txtEditEmail.Text.Trim, WindowsUsername)
	'			' 4. redirect to code entry page
	'			If sSMSError.Trim <> "" Then
	'				sUpdEmailError = sSMSError.Trim & "<br />"
	'			Else
	'				Session("action") = "NewPasswordStep3updated"
	'				Response.Redirect(Request.Url.ToString)
	'			End If
	'		End If
	'		' --------------- email ----------------------


	'		Dim sErrorMessage As String = ""
	'		If sUpdPhoneError.Trim <> "" Then
	'			sErrorMessage &= "<b>Errors updating phone</b>:" & "<br />" & sUpdPhoneError & "<br />"
	'		End If
	'		If sUpdEmailError.Trim <> "" Then
	'			sErrorMessage &= "<b>Errors updating email</b>:" & "<br />" & sUpdEmailError & "<br />"
	'		End If


	'	End If

	'End Sub



	'Public Sub RefreshData()
	'	dgCats()
	'End Sub


	'Public Function getEdDSN() As String
	'	Return ConnectionString
	'End Function

	'Private Sub dgCats()
	'	CatBindGrid(dtDataTable)
	'End Sub


	'Public Sub CatBindGrid(ByVal dt As DataTable)
	'	gv1.DataSource = dt

	'	Try
	'		Me.DataBind()

	'	Catch err As Exception
	'		lblErrors.Text &= "ERROR: binding grid - " & err.Message
	'	Finally

	'	End Try
	'End Sub



	'Protected Sub gv1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

	'End Sub


	'Protected Sub gv1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)

	'End Sub


	'Protected Sub gv1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowCreated

	'	If (e.Row.RowState And DataControlRowState.Edit) <> 0 OrElse (e.Row.RowState And DataControlRowState.Insert) <> 0 Then

	'	End If


	'End Sub



	'Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound


	'	'Try
	'	'	Dim ddlist As DropDownList = CType(e.Row.FindControl("ddlEditConsultants"), DropDownList)
	'	'	ddlist.DataSource = GetConsultants()
	'	'	ddlist.DataTextField = "Consultant"
	'	'	ddlist.DataValueField = "UserID"
	'	'	ddlist.DataBind()
	'	'	If Not IsNothing(ddlist.Items.FindByText(DataBinder.Eval(e.Row.DataItem, "ConsultantName").ToString())) Then
	'	'		ddlist.Items.FindByText(DataBinder.Eval(e.Row.DataItem, "ConsultantName")).Selected = True
	'	'	End If
	'	'Catch ex As Exception
	'	'End Try

	'	'lblDatakeys.Text &= "<b>RowDataBound 1: </b>" & e.Row.Cells(2).Text & "<br />"

	'End Sub


	'' add to asxc: OnPageIndexChanging="gv1_PageIndexChanging" 
	'Protected Sub gv1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
	'	gv1.SelectedIndex = -1
	'	gv1.PageIndex = e.NewPageIndex
	'	dgCats()
	'End Sub


	'' add to asxc: OnSorting="gv1_Sorting"
	'Protected Sub gv1_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
	'	gv1.SelectedIndex = -1
	'	dgCats()
	'End Sub





	'' add to asxc: not required
	'Protected Sub gv1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles gv1.RowCancelingEdit
	'	gv1.EditIndex = -1
	'	dgCats()
	'End Sub



	'' add to asxc:  not required
	'Protected Sub gv1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles gv1.RowEditing
	'	gv1.EditIndex = e.NewEditIndex
	'	'lblDatakeys.Text &= "<b>e.NewEditIndex: </b>" & e.NewEditIndex & "<br />"
	'	dgCats()

	'	' For updating - set the update command and the parameters on the SQL datasource
	'End Sub



	'' The following catches exceptions on updating data (not needed unless the databinding on ascx is used)
	'' add to asxc: OnRowUpdated="gv1_RowUpdated"
	'Protected Sub gv1_RowUpdated(ByVal sender As Object,
	'   ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs)

	'	If (Not IsDBNull(e.Exception)) Then
	'		lblErrors.Text = e.Exception.Message
	'	End If
	'End Sub


	'Protected Sub gv1_RowDeleted(ByVal sender As Object,
	'   ByVal e As GridViewDeletedEventArgs)

	'	If (Not IsDBNull(e.Exception)) Then
	'		lblErrors.Text = e.Exception.Message
	'		e.ExceptionHandled = True
	'	End If
	'End Sub


	'Protected Sub gv1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gv1.RowDeleting


	'End Sub



	'' -----------------------------------------------------------------------------------
	'' DO NOT add to asxc: OnRowUpdating="gv1_RowUpdating" -  IT UPDATES DATA TWICE!!! The 2nd time without any data
	'' -----------------------------------------------------------------------------------
	'Protected Sub gv1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles gv1.RowUpdating


	'	'Dim iCompanyFK As Integer = getCompanyFK()

	'	'Dim dteDateToday As Date = Date.Now.Date

	'	'Dim sConsultantName As String = ""
	'	'Dim lTVPresentationLunchRotaID As Long = CLng(e.Keys("TVPresentationLunchRotaID").ToString())
	'	'Dim DelIndRecord As Integer = 0
	'	'Dim outTVPresentationLunchRotaID As Long = 0

	'	'Dim row As GridViewRow = gv1.Rows(e.RowIndex)
	'	'Dim ddlistCons As DropDownList = CType(row.FindControl("ddlEditConsultants"), DropDownList)
	'	'' row.Cells(iy).Controls(iz)
	'	''CType(e.Row.FindControl("ddlEditConsultants"), DropDownList)
	'	'Dim sConsTemp As String = ""
	'	'sConsTemp = ddlistCons.SelectedItem.Text()
	'	''lblDatakeys.Text &= "____ddlistCons = " & sConsultantName & "____<br />"

	'	'If sConsTemp.Trim.Contains("[") Then
	'	'	' it's in this form: Gemma Higgins [Ed]
	'	'	' hence remove the [Ed] or [Sales], etc...

	'	'	Dim sSplit() As String = sConsTemp.Split(CChar("["))
	'	'	If sSplit IsNot Nothing Then
	'	'		If sSplit.Length > 0 Then
	'	'			sConsultantName = sSplit(0).Trim
	'	'		End If
	'	'	End If
	'	'Else
	'	'	sConsultantName = sConsTemp
	'	'End If


	'	'Dim iret As Integer = 0
	'	'Using lib1 As New EvoTVCLientLibrary.clsDBAccess(getEdDSN())

	'	'	iret = lib1.InsertUpdateLunchRotaByCompany(
	'	'	Date.Now.Date,
	'	'	iCompanyFK,
	'	'	sConsultantName,
	'	'	lTVPresentationLunchRotaID,
	'	'	0,
	'	'	outTVPresentationLunchRotaID
	'	'   )
	'	'End Using



	'	'If iret > 0 Then
	'	'	lblErrors.Text &= "The lunch rota was updated<br />"

	'	'	' Cancel edit mode. refresh
	'	'	gv1.EditIndex = -1
	'	'	dgCats()

	'	'Else
	'	'	lblErrors.Text &= "ERROR: The lunch rota was not updated<br />"
	'	'End If
	'	''lblDatakeys.Text &= "updated: " & iret & "; outTVPresentationLunchRotaID=" & outTVPresentationLunchRotaID & "<br />"


	'End Sub







End Class
