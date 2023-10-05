
Imports functions
Imports System.Net
Imports System.text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls

Partial Class MyReportsAdmin
    Inherits System.Web.UI.Page


    ' <asp:CommandField ShowDeleteButton="True"  />

    Public pCompanyFK As Integer = 0

    Public Property BIWindowsConnectionString() As String = ConfigurationManager.ConnectionStrings("BIWindowsConnectionString").ConnectionString
    Public Property EdDSN() As String = ConfigurationManager.ConnectionStrings("EdConnectionString").ConnectionString
    '    Get
    '        Dim sBIWindowsConnectionString As String = ""
    '        If HttpContext.Current.Session("BIWindowsConnectionString") IsNot Nothing Then
    '            If CStr(HttpContext.Current.Session("BIWindowsConnectionString")) <> "" Then
    '                sBIWindowsConnectionString = HttpContext.Current.Session("BIWindowsConnectionString")
    '            End If
    '        End If
    '        Return sBIWindowsConnectionString
    '    End Get
    '    Set(ByVal value As String)
    '        HttpContext.Current.Session("BIWindowsConnectionString") = value
    '    End Set
    'End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        'BIWindowsConnectionString = UseBIWondowsConn()

        ' http://intranet.evolutionjobs.local:81/Reporting/MyReportsAdmin.aspx?comp=1
        ' http://intranet.evolutionjobs.local:8022/Reporting/MyReportsAdmin.aspx?comp=1
        ' http://localhost:54809/EvolutionDevelopment/Reporting/MyReportsAdmin.aspx?comp=1

        lblHeader.Text = "My Reports Admin"
        lblErrors.Text = ""



        '' ------------- this is for standalone version ---------
        '' ------------- this is for standalone version ---------
        '' ------------- this is for standalone version ---------
        '' ------------- this is for standalone version ---------
        'If Request("comp") <> "" Then
        '    SelectCompanyID = CInt(Request("comp"))
        'End If

        '' use EvolutionStrict
        'Dim SQLUser As String = sSQLUserEvoStrict
        'Try
        '    ' the userIP is used to get the team Fk for being able to select the Team from drop down
        '    ' it will be passed in to thre user review to allow authentication
        '    lblUserIP.Text = Request.ServerVariables("REMOTE_ADDR")
        '    If Request("TestIP") <> "" Then
        '        lblUserIP.Text = Request("TestIP")
        '    End If

        '    lblUserIP.Text = lblUserIP.Text.Replace("::1", "")

        'Catch ex As Exception
        'End Try

        '' LiveServer
        '' BackUpServer
        'Dim strEdConn As String = ""
        'strEdConn = UseConn("ED", ConfigurationManager.AppSettings("LiveServer").ToString())
        'ConnectionString = strEdConn



        '' --------- check if German ---------
        'Dim sDE As String = ""
        'Dim sHTTP_ACCEPT_LANGUAGE As String = ""
        'Try
        '    sHTTP_ACCEPT_LANGUAGE = Request.ServerVariables("HTTP_ACCEPT_LANGUAGE")
        '    If sHTTP_ACCEPT_LANGUAGE.Trim <> "" Then
        '        Dim sLang As String = ""
        '        If sHTTP_ACCEPT_LANGUAGE.Trim.Contains(",") Then
        '            Dim sSplit() As String = sHTTP_ACCEPT_LANGUAGE.Split(",")
        '            ' just get the first one
        '            If sSplit IsNot Nothing Then
        '                If sSplit(0) IsNot Nothing Then
        '                    sLang = sSplit(0)
        '                End If
        '            End If

        '        Else
        '            sLang = sHTTP_ACCEPT_LANGUAGE.Trim
        '        End If
        '        If sLang.Trim.ToLower = GermanLocaleCode.ToLower Then
        '            sDE = GermanLocaleCode
        '        End If
        '    End If

        'Catch ex As Exception
        'End Try
        'If sDE.Trim = "" Then
        '    If SQLUser.ToLower.Contains("german") Then
        '        sDE = GermanLocaleCode
        '    End If
        'End If
        'LocaleCode = sDE.ToLower
        '' --------- check if German ---------

        'Dim lib1 As New libEvolutionIntranet.clsIntranet(BIWindowsConnectionString)
        '' Dim libEvoResources1 As New libEvolutionIntranet.clsEvoDev(getEdResourcesDSN())
        'Try
        '    UserFKViewstate = lib1.GetConsultantIDFromIP(lblUserIP.Text)
        'Catch ex As Exception
        '    ProcessError(ex.Message)
        'Finally
        '    lib1.Dispose()
        '    ' libEvoResources1.Dispose()
        'End Try

        '' ------------- this is for standalone version ---------
        '' ------------- this is for standalone version ---------
        '' ------------- this is for standalone version ---------
        '' ------------- this is for standalone version ---------


        '' ----- test -----
        '' ----- test -----
        '' ----- test -----
        '' ----- test -----
        'If Request.Url.ToString.ToLower.Contains("localhost") Then
        '    UserFKViewstate = 378
        'End If
        '' ----- test -----
        '' ----- test -----
        '' ----- test -----
        '' ----- test -----



        Dim sWindowsUserName As String = HttpContext.Current.User.Identity.Name
		Dim lWinUsernameBIUserFK As Long = 0
		Dim sWinUsernameBIUserFK As String = ""

        If sWindowsUserName.Trim.Length > 0 Then
            Using lib1 As New clsDB(EdDSN)
                sWinUsernameBIUserFK = lib1.RunSQLScalar("SELECT ed.[dbo].[udf_SP_WIN_CheckWindowsUsername]('" & sWindowsUserName & "')")
            End Using
        Else
            lblErrors.Text = "ERROR: You are not authenticated (username=" & sWindowsUserName & ")"
            Exit Sub
        End If

        If sWinUsernameBIUserFK.Trim <> "" Then
			Long.TryParse(sWinUsernameBIUserFK, lWinUsernameBIUserFK)
		End If

		Dim mpMaster As SiteMaster = DirectCast(Page.Master, SiteMaster)

		'2-Dim myPage As System.Web.UI.Page = DirectCast(Me.Page, System.Web.UI.Page)
		'2-Dim mpMaster As MasterPageEditor = DirectCast(myPage.Master, MasterPageEditor)
		'1-Dim mpMaster As MasterPageEditor = CType(Me.Page.Master, MasterPageEditor)

		pnlUpdateStatus.Visible = False
        Dim UserStatusEditorsList As String = ConfigurationManager.AppSettings("MyReportsEditors")
        Dim users() As String = UserStatusEditorsList.Trim.Split(",")
        If users IsNot Nothing Then
            For iUser As Integer = users.GetLowerBound(0) To users.GetUpperBound(0)
                If users(iUser) IsNot Nothing Then
                    If users(iUser) <> "" Then
                        Dim lAllowedUserFK As Long = 0
                        Long.TryParse(users(iUser), lAllowedUserFK)
                        If lAllowedUserFK > 0 Then
							If lAllowedUserFK = lWinUsernameBIUserFK Then
								pnlUpdateStatus.Visible = True
							End If
						End If
                    End If
                End If
            Next
        End If

        If BIWindowsConnectionString.ToUpper.Contains("SVR21") Or BIWindowsConnectionString.Contains("24.21") Or
            EdDSN.ToUpper.Contains("SVR21") Or EdDSN.Contains("24.21") Then
            lblDatabase.Text = "************** BACKUP ****************"
        Else
            lblDatabase.Text = ""
        End If

        mpMaster.Dispose()


        If Not Me.IsPostBack Then

            lblGVCellEmailError.Text = ""
            lblGVCellPhoneError.Text = ""

            RefreshData()

        End If



    End Sub




	'Public Function getCompanyFK() As Integer
 '       Dim iComp As Integer = 1
 '       Dim mpMaster As MasterPageEditor = DirectCast(Page.Master, MasterPageEditor)
 '       'Dim ser As Services = TheMaster.Services

 '       '2-Dim myPage As System.Web.UI.Page = DirectCast(Me.Page, System.Web.UI.Page)
 '       '2-Dim mpMaster As MasterPageEditor = DirectCast(myPage.Master, MasterPageEditor)
 '       '1-Dim mpMaster As MasterPageEditor = CType(Me.Page.Master, MasterPageEditor)
 '       iComp = mpMaster.CompanyFKViewstate
 '       pCompanyFK = iComp
 '       mpMaster.Dispose()
 '       Return iComp
 '   End Function


    Protected Sub cmdGetData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdGetData.Click
        RefreshData()
    End Sub

    Public Sub RefreshData()
        dgCats()
    End Sub

    Private Sub dgCats()
        Dim ds As DataTable = CatGetDataSet()
        CatBindGrid(ds)
    End Sub


    Public Function CatGetDataSet() As DataTable

        Dim dt As New DataTable
        Using lib1 As New libEvolutionIntranet.clsEvoDev(BIWindowsConnectionString)
            ' set the DelINRecord to -1 to show all records; 0 to show non delinded records and 1 to shoqw only delinded records

            ' PASS  A -1 AS USERFK TO GET ALL USERS
            dt = lib1.MyReportsData(
            "Return Details",
            "",
            -1,
            "",
            ""
            )

            'dt = lib1.MyReportsData(
            '"Return Details",
            'WindowsLogin,
            'UserFK,
            'Phone,
            'Email
            ')

        End Using
        Return dt

    End Function


    Public Sub CatBindGrid(ByVal dt As DataTable)

        Dim dv As New DataView(dt)
        dv.RowFilter = "isnull(user_login, '') <> '' "

        gv1.DataSource = dv

        Try
            Me.DataBind()

        Catch err As Exception
            lblErrors.Text &= "ERROR: binding grid - " & err.Message
        Finally

        End Try
    End Sub


    ' add to asxc: OnSelectedIndexChanged="gv1_SelectedIndexChanged" 
    Protected Sub gv1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'lblDatakeys.Text = "<b>SelectedDataKey.Value: </b>" & Server.HtmlEncode(gv1.SelectedDataKey.Value) & "<br />"
        'lblDatakeys.Text &= "<b>DataKey Field 1: </b>" & Server.HtmlEncode(gv1.SelectedDataKey.Values("MCID")) & "<br />"
        'lblDatakeys.Text &= "<b>Selected MainCat: </b>" & Server.HtmlEncode(gv1.SelectedRow.Cells(1).Text) & "<br />"

        'string selectedCategory = myGridView.SelectedRow.Cells[1].Text;
    End Sub

    ' add to asxc: OnDataBound="gv1_DataBound"
    Protected Sub gv1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim key As DataKey
        'lblDatakeys.Text = "<b>GridView DataKeys:</b>" & "<br />"
        'For Each key In gv1.DataKeys
        '    lblDatakeys.Text &= Server.HtmlEncode(key.Values(0)) & ", "
        'Next

    End Sub


    ' ********** Example of using drop down list **********
    Protected Sub gv1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowCreated

        If (e.Row.RowState And DataControlRowState.Edit) <> 0 OrElse (e.Row.RowState And DataControlRowState.Insert) <> 0 Then
            'Try
            '    Dim ddlist As DropDownList = CType(e.Row.FindControl("ddlEditConsultants"), DropDownList)
            '    ddlist.DataSource = GetConsultants()
            '    ddlist.DataTextField = "Consultant"
            '    ddlist.DataValueField = "UserID"
            '    ddlist.DataBind()
            '    If Not IsNothing(ddlist.Items.FindByText(DataBinder.Eval(e.Row.DataItem, "ConsultantName").ToString())) Then
            '        ddlist.Items.FindByText(DataBinder.Eval(e.Row.DataItem, "ConsultantName")).Selected = True
            '    End If
            'Catch ex As Exception
            'End Try
        End If


    End Sub

    Function GetInternationCodesTable() As DataTable
        ' Create new DataTable instance.
        Dim table As New DataTable
        table.Columns.Add("Country", GetType(String))
        table.Columns.Add("Code", GetType(String))


        ' Add five rows with those columns filled in the DataTable.
        table.Rows.Add("----SELECT-----", "")
        table.Rows.Add("Australia: +61", "+61")
        table.Rows.Add("Germany: +49", "+49")
        table.Rows.Add("Singapore: +65", "+65")
        table.Rows.Add("UK: +44", "+44")
        Return table
    End Function

    Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound

        'If e.Row.RowType = DataControlRowType.DataRow AndAlso gv1.EditIndex = e.Row.RowIndex Then
        '    Dim ddlstCodes As DropDownList = DirectCast(e.Row.FindControl("ddlstEditInternationalCodes"), DropDownList)
        '    Dim query As String = "Customers_GetCities"
        '    Dim cmd As New SqlCommand(query)
        '    cmd.CommandType = CommandType.StoredProcedure
        '    ddlCities.DataSource = GetData(cmd)
        '    ddlCities.DataTextField = "CityName"
        '    ddlCities.DataValueField = "CityName"
        '    ddlCities.DataBind()
        '    ddlCities.Items.FindByValue(TryCast(e.Row.FindControl("lblCity"), Label).Text).Selected = True
        'End If

        If e.Row.RowType = DataControlRowType.DataRow AndAlso gv1.EditIndex = e.Row.RowIndex Then
			Try

				Dim lblPhoneError As Label = CType(e.Row.FindControl("lblPhoneError"), Label)
				Dim lblEmailError As Label = CType(e.Row.FindControl("lblEmailError"), Label)

				' get the errors when trying to update - see gv1_RowUpdating
				If lblGVCellEmailError.Text.Trim <> "" Then
					lblEmailError.Text = lblGVCellEmailError.Text
				End If
				If lblGVCellPhoneError.Text.Trim <> "" Then
					lblPhoneError.Text = lblGVCellPhoneError.Text
				End If

				Dim txtPhone_Number As TextBox = CType(e.Row.FindControl("txtPhone_Number"), TextBox)
				Dim ddlstCodes As DropDownList = CType(e.Row.FindControl("ddlstEditInternationalCodes"), DropDownList)
				'ddlstCodes.Items.Clear()

				ddlstCodes.DataSource = GetInternationCodesTable()
				ddlstCodes.DataTextField = "Country"
				ddlstCodes.DataValueField = "Code"
				ddlstCodes.DataBind()

				'Dim li0 As New ListItem("----SELECT-----", "")
				'Dim li1 As New ListItem("Australia: +61", "+61")
				'Dim li2 As New ListItem("Germany: +49", "+49")
				'Dim li3 As New ListItem("Singapore: +65", "+65")
				'Dim li4 As New ListItem("UK: +44", "+44")

				'ddlstCodes.Items.Add(li0)
				'ddlstCodes.Items.Add(li1)
				'ddlstCodes.Items.Add(li2)
				'ddlstCodes.Items.Add(li3)
				'ddlstCodes.Items.Add(li4)

				'Dim ddlstCodes As DropDownList = CType(e.Row.FindControl("ddlInternationalCodes"), DropDownList)
				'ddlist.DataSource = GetConsultants()
				'ddlist.DataTextField = "Consultant"
				'ddlist.DataValueField = "UserID"
				'ddlist.DataBind()
				'If Not IsNothing(ddlist.Items.FindByText(DataBinder.Eval(e.Row.DataItem, "ConsultantName").ToString())) Then
				'    ddlist.Items.FindByText(DataBinder.Eval(e.Row.DataItem, "ConsultantName")).Selected = True
				'End If

				Dim sPhoneWithoutCode As String = ""
				Dim ExistingMobileNumber As String = DataBinder.Eval(e.Row.DataItem, "Phone_Number")
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

				txtPhone_Number.Text = sMobileWithoutIntCode

				Try
					ddlstCodes.SelectedIndex = iIndex
				Catch ex As Exception
				End Try

			Catch ex As Exception
			End Try
        End If


        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.RowState = DataControlRowState.Edit Then

            End If
        End If






        'Dim tbtxtMCMAINCATEGORY As TextBox = e.Row.FindControl("txtMCMAINCATEGORY")
        'If Not tbtxtMCMAINCATEGORY Is Nothing Then
        '    lblDatakeys.Text &= "RowDataBound MCMAINCATEGORY = " & tbtxtMCMAINCATEGORY.Text & "<br />"
        '    tbtxtMCMAINCATEGORY.Text = "XXXXX"
        'End If

        'Dim strCategoryID As String = e.Row.Cells(0).Text
        'Dim myDropDown As DropDownList
        'myDropDown = CType(e.Item.FindControl("cboCategoryID"), DropDownList)
        'myDropDown.SelectedIndex = myDropDown.Items.IndexOf(myDropDown.Items.FindByValue(strCategoryID))

        'lblDatakeys.Text &= "<b>RowDataBound 1: </b>" & e.Row.Cells(2).Text & "<br />"

    End Sub


    ' add to asxc: OnPageIndexChanging="gv1_PageIndexChanging" 
    Protected Sub gv1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gv1.SelectedIndex = -1
        gv1.PageIndex = e.NewPageIndex
        dgCats()
    End Sub


    ' add to asxc: OnSorting="gv1_Sorting"
    Protected Sub gv1_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        gv1.SelectedIndex = -1
        dgCats()
    End Sub





	' add to asxc: not required
	Protected Sub gv1_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles gv1.RowCancelingEdit
		lblGVCellEmailError.Text = ""
		lblGVCellPhoneError.Text = ""
		gv1.EditIndex = -1
		dgCats()
	End Sub



	' add to asxc:  not required
	Protected Sub gv1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles gv1.RowEditing
		lblGVCellEmailError.Text = ""
		lblGVCellPhoneError.Text = ""

		gv1.EditIndex = e.NewEditIndex
		'lblDatakeys.Text &= "<b>e.NewEditIndex: </b>" & e.NewEditIndex & "<br />"
		dgCats()


        ' notes: need to have the following in the gridview code:
        ' DataKeyNames="au_id,title_id"
        ' AutoGenerateEditButton ="true"
        ' Also set ReadOnly="True" for columns that do not nedd editing

        ' For updating - set the update command and the parameters on the SQL datasource
    End Sub



	' The following catches exceptions on updating data (not needed unless the databinding on ascx is used)
	' add to asxc: OnRowUpdated="gv1_RowUpdated"
	Protected Sub gv1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs)

		e.KeepInEditMode = True

		If (Not IsDBNull(e.Exception)) Then
			lblErrors.Text = e.Exception.Message
		End If
	End Sub


	' The following checks for the deleted function errors
	' For deleting add the following in the gridview code:
	' -- AutoGenerateDeleteButton ="true"
	' -- DeleteCommand="DELETE From [titleauthor] WHERE au_id=@au_id AND title_id=@title_id"
	'<DeleteParameters>
	'     <asp:Parameter Type="String" Name="au_id"></asp:Parameter>
	'     <asp:Parameter Type="String" Name="title_id"></asp:Parameter>
	'</DeleteParameters>
	' --

	' Not needed - not databinding on grid directly
	' add to asxc: OnRowDeleted="gv1_RowDeleted"
	Protected Sub gv1_RowDeleted(ByVal sender As Object, ByVal e As GridViewDeletedEventArgs)

		If (Not IsDBNull(e.Exception)) Then
			lblErrors.Text = e.Exception.Message
			e.ExceptionHandled = True
		End If
	End Sub


	Protected Sub gv1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gv1.RowDeleting
        'Dim row As GridViewRow = gv1.Rows(e.RowIndex)
        'Dim UserFK As Long = CType(row.Cells(1).Text, Long)


        ''lblDatakeys.Text = "<b>SelectedDataKey.Value: </b>" & e.Keys("UserFK").ToString() & "<br />"
        ''lblDatakeys.Text &= "<b>Selected TVID: </b>" & Server.HtmlEncode(row.Cells(1).Text) & "<br />"
        'Dim iCompanyFK As Integer = getCompanyFK()
        'Dim outUserFK As Long = 0
        'Dim iret As Integer = 0
        'Using lib1 As New EvoTVCLientLibrary.clsDBAccess(BIWindowsConnectionString)

        '    ' iret = lib1.InsertUpdateLunchRotaByCompany( _
        '    ' Date.Now.Date, _
        '    ' iCompanyFK, _
        '    ' "", _
        '    ' UserFK, _
        '    ' 1, _
        '    ' outUserFK _
        '    ')

        'End Using

        ''lblDatakeys.Text &= "iret = " & iret & "<br />"
        'If iret = 1 Then
        '    lblErrors.Text &= "The consultant has been removed from the lunch rota<br />"

        '    ' Cancel edit mode. refresh
        '    gv1.EditIndex = -1
        '    dgCats()
        'Else
        '    lblErrors.Text &= "ERROR: The consultant was not removed from the lunch rota<br />"
        'End If

    End Sub



    ' -----------------------------------------------------------------------------------
    ' DO NOT add to asxc: OnRowUpdating="gv1_RowUpdating" -  IT UPDATES DATA TWICE!!! The 2nd time without any data
    ' -----------------------------------------------------------------------------------
    Protected Sub gv1_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles gv1.RowUpdating
		'Dim row As GridViewRow = gv1.Rows(e.RowIndex)
		lblGVCellEmailError.Text = ""
		lblGVCellPhoneError.Text = ""

		Dim sPhoneError As String = ""
		Dim sEmailError As String = ""

		Dim lUserFK As Long = CLng(e.Keys("UserFK").ToString())


		Dim row As GridViewRow = gv1.Rows(e.RowIndex)
        Dim ddlistCode As DropDownList = CType(row.FindControl("ddlstEditInternationalCodes"), DropDownList)
        Dim txtEditPhone As TextBox = CType(row.FindControl("txtPhone_Number"), TextBox)
        Dim txtEditEmail As TextBox = CType(row.FindControl("txtEmail_Address"), TextBox)
        Dim lblPhoneError As Label = CType(row.FindControl("lblPhoneError"), Label)
		Dim lblEmailError As Label = CType(row.FindControl("lblEmailError"), Label)

		lblPhoneError.Text = ""
		lblEmailError.Text = ""

		Dim sCode As String = ""
        If ddlistCode.SelectedIndex <> -1 Then
            sCode = ddlistCode.SelectedItem.Value()
        End If

		' ------------- for phone -----------------
		If sCode.Trim <> "" Then

			If txtEditPhone.Text.Trim <> "" Then

				Dim sNumberEntered As String = txtEditPhone.Text.Trim
				If sNumberEntered.Substring(0, 1) = "0" Then
					sNumberEntered = sNumberEntered.Substring(1)
				End If

				Dim sPhone As String = sCode & sNumberEntered

				If sPhone.Trim <> "" Then

					Try
						Using lib1 As New libEvolutionIntranet.clsEvoDev(BIWindowsConnectionString)
							' to clear the phone number or emails enter a -1

							' PASS  A -1 AS USERFK TO GET ALL USERS
							Dim dtOutput As New DataTable
							dtOutput = lib1.MyReportsData(
							"Update user details",
							"",
							lUserFK,
							sPhone,
							""
							)

							'dt = lib1.MyReportsData(
							'"Return Details",
							'WindowsLogin,
							'UserFK,
							'Phone,
							'Email
							')

						End Using
					Catch ex As Exception
						sPhoneError = "ERROR: Saving phone number: " & ex.Message
					End Try

					If sPhoneError <> "" Then


					Else
						'lblErrors.Text &= "The lunch rota was updated<br />"

						' Cancel edit mode. refresh
						gv1.EditIndex = -1
						dgCats()

						lblGVCellPhoneError.Text = ""
					End If

				Else
					sPhoneError = "ERROR: Please enter a phone number<br />"
				End If

			Else
				sPhoneError = "ERROR: Please enter a phone number<br />"

			End If


		Else
			sPhoneError = "ERROR: Please select an international code<br />"
		End If


		If txtEditEmail.Text.Trim <> "" Then

            Dim bIsEmailValid As Boolean = IsValidEmail(txtEditEmail.Text.Trim)
			If bIsEmailValid = False Then
				sEmailError = "ERROR: Please enter a valid email address"

			Else


				Dim lDomainID As Long = 0
                Using lib1 As New clsDB(EdDSN)
                    lDomainID = lib1.CheckIfDomainIsUsedInEmail(txtEditEmail.Text.Trim)
                End Using

                If lDomainID > 0 Then
					sEmailError = "ERROR: You cannot use your work email address (for security)"

				Else

					Try
						Using lib1 As New libEvolutionIntranet.clsEvoDev(BIWindowsConnectionString)
							' to clear the phone number or emails enter a -1

							' PASS  A -1 AS USERFK TO GET ALL USERS
							Dim dtOutput As New DataTable
							dtOutput = lib1.MyReportsData(
							"Update user details",
							"",
							lUserFK,
							"",
							txtEditEmail.Text.Trim
							)

							'dt = lib1.MyReportsData(
							'"Return Details",
							'WindowsLogin,
							'UserFK,
							'Phone,
							'Email
							')

						End Using
					Catch ex As Exception
						sEmailError = "ERROR: Saving email: " & ex.Message
					End Try

					If sEmailError <> "" Then

					Else
						'lblErrors.Text &= "The lunch rota was updated<br />"

						' Cancel edit mode. refresh
						gv1.EditIndex = -1
						dgCats()

						lblGVCellEmailError.Text = ""
					End If

				End If



			End If


		End If

		' this keeps the row in edit mode allowing the user to fix the error
		' if you set gv1.EditIndex = -1 and call dgCats() it will cancel editing
		If sEmailError.Trim <> "" Or sPhoneError.Trim <> "" Then
			' we must show these errors when rebinding the grid - see gv1_RowDataBound

			If sEmailError.Trim <> "" Then
				lblGVCellEmailError.Text = sEmailError
			End If
			If sPhoneError.Trim <> "" Then
				lblGVCellPhoneError.Text = sPhoneError
			End If

			lblErrors.Text = sEmailError & "<br />" & sPhoneError

			gv1.EditIndex = e.RowIndex
			dgCats()
		End If



		'lblDatakeys.Text &= "updated: " & iret & "; outTVPresentationLunchRotaID=" & outTVPresentationLunchRotaID & "<br />"





		'lblDatakeys.Text = "<b>SelectedDataKey.Value: </b>" & Server.HtmlEncode(e.Keys("UserFK").ToString()) & "<br />"
		''lblDatakeys.Text &= "<b>DataKey Field 1: </b>" & Server.HtmlEncode(gv1.SelectedDataKey.Values("UserFK")) & "<br />"
		'lblDatakeys.Text &= "<b>Selected TVID: </b>" & Server.HtmlEncode(row.Cells(1).Text) & "<br />"



		'' Original - for reference
		'Dim tbtxtMCMAINCATEGORY As TextBox = row.FindControl("txtMCMAINCATEGORY")
		'If Not tbtxtMCMAINCATEGORY Is Nothing Then
		'    lblMain.Text &= "Text entered = " & tbtxtMCMAINCATEGORY.Text & "<br />"
		'End If
		'Dim artb(iControls) As TextBox
		'artb(0) = row.FindControl("txtMCMAINCATEGORY")

		'' For the following use a Templated column
		'If Not row Is Nothing Then
		'    Dim iControls As Integer = row.Cells.Count - 1
		'    Dim arValues(iControls) As String

		'    ' Iterate through all the cells in the row; get values if control = textbox
		'    ' save them in array
		'    Dim iy As Integer = 0
		'    For iy = 0 To iControls
		'        If row.Cells(iy).HasControls Then
		'            Dim iz As Integer = 0
		'            For iz = 0 To row.Cells(iy).Controls.Count - 1
		'                Dim contype As String = "System.Web.UI.WebControls.TextBox"
		'                Dim contypechk As String = "System.Web.UI.WebControls.CheckBox"
		'                Dim contypeddlst As String = "System.Web.UI.WebControls.DropDownList"
		'                If row.Cells(iy).Controls(iz).GetType.ToString.ToLower = contype.ToLower Then
		'                    Dim tb1 As TextBox = row.Cells(iy).Controls(iz)
		'                    arValues(iy) = tb1.Text
		'                    lblDatakeys.Text &= iy & "; tb1VALUE = " & arValues(iy) & "<br />"
		'                End If
		'                If row.Cells(iy).Controls(iz).GetType.ToString.ToLower = contypechk.ToLower Then
		'                    Dim chk1 As CheckBox = row.Cells(iy).Controls(iz)
		'                    If chk1.Checked = True Then
		'                        arValues(iy) = 1
		'                    Else
		'                        arValues(iy) = 0
		'                    End If
		'                    lblDatakeys.Text &= iy & "; chk1VALUE = " & arValues(iy) & "<br />"
		'                End If
		'                If row.Cells(iy).Controls(iz).GetType.ToString.ToLower = contypeddlst.ToLower Then
		'                    Dim ddlist As DropDownList = CType(row.Cells(iy).Controls(iz), DropDownList)
		'                    ' row.Cells(iy).Controls(iz)
		'                    'CType(e.Row.FindControl("ddlEditConsultants"), DropDownList)
		'                    arValues(iy) = ddlist.SelectedItem.Text()
		'                    lblDatakeys.Text &= iy & "; ddlstVALUE = " & arValues(iy) & "<br />"
		'                End If

		'            Next
		'        End If
		'    Next

		'    ' Check that the array has captured the data
		'    Dim ix As Integer = 0
		'    For ix = arValues.GetLowerBound(0) To arValues.GetUpperBound(0)
		'        If Not arValues(ix) Is Nothing Then
		'            lblDatakeys.Text &= ix & "; Text entered = " & arValues(ix) & "<br />"
		'        End If
		'    Next

		' End if



	End Sub






    'Private Function AddSubCategory() As Long
    '    Dim iret As Long = 0
    '    Dim Prod1 As New sst.Commerce.Products

    '    iret = Prod1.AddSubCategory
    '    If iret > 0 Then
    '        lblDatakeys.Text &= "Sub cateogry added (needs to be edited) (" & iret & ")<br />"

    '        ' Cancel edit mode. refresh
    '        gv1.EditIndex = -1
    '        dgCats()
    '    Else
    '        lblDatakeys.Text &= "error: category not added, (" & iret & ")<br />"
    '    End If

    '    Return iret
    'End Function





End Class
