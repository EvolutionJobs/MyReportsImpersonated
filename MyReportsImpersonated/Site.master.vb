


Partial Class SiteMaster
	Inherits System.Web.UI.MasterPage

	'[ <a href="~/Account/Login.aspx" ID="HeadLoginStatus" runat="server">Log In</a> ]
	'[ <asp:LoginStatus ID="HeadLoginStatus" runat="server" LogoutAction="Redirect" LogoutText="Log Out" LogoutPageUrl="~/"/> ]

	'Public Property SQLServer As String = "" ' change to 10.11.24.5 for LIVE



	Public Property UserIPAddress() As String
		Get
			Dim sUserIPAddress As String = ""
			If ViewState("UserIPAddress") IsNot Nothing Then
				If CStr(ViewState("UserIPAddress")) <> "" Then
					sUserIPAddress = ViewState("UserIPAddress")
				End If
			End If
			Return sUserIPAddress
		End Get
		Set(ByVal value As String)
			ViewState("UserIPAddress") = value
		End Set
	End Property


	Public Property CompanyFKViewstate() As Integer
		Get
			Dim iCompFK As Integer = 0
			If ViewState("CompanyFK") IsNot Nothing Then
				If CStr(ViewState("CompanyFK")) <> "" Then
					Integer.TryParse(ViewState("CompanyFK"), iCompFK)
				End If
			End If
			Return iCompFK
		End Get
		Set(ByVal value As Integer)
			ViewState("CompanyFK") = value
		End Set
	End Property

	Public Property LocaleCode() As String
		Get
			Dim sCode As String = "en-GB"
			If ViewState("LocaleCode") IsNot Nothing Then
				If ViewState("LocaleCode") <> "" Then
					sCode = ViewState("LocaleCode")
				End If
			End If
			Return sCode
		End Get
		Set(ByVal value As String)
			ViewState("LocaleCode") = value
		End Set
	End Property


	Public Property UserFKViewstate() As Long
		Get
			Dim lUserFK As Long = 0
			If ViewState("UserFK") IsNot Nothing Then
				Long.TryParse(ViewState("UserFK"), lUserFK)
			End If
			Return lUserFK
		End Get
		Set(ByVal value As Long)
			ViewState("UserFK") = value
		End Set
	End Property

	Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
		'litLoggedInUser.Text = HttpContext.Current.User.Identity.Name

		EdConnectionString = ConfigurationManager.ConnectionStrings("EDConnectionString").ConnectionString
		GrantImpersonateConnectionString = ConfigurationManager.ConnectionStrings("GrantImpersonateConnectionString").ConnectionString
		PasswordAdminEdConnectionString = GrantImpersonateConnectionString.Replace("MyBI", "PasswordAdmin")
		BIWindowsConnectionString = ConfigurationManager.ConnectionStrings("BIWindowsConnectionString").ConnectionString

		Dim sBackup As String = ""
		If BIWindowsConnectionString.Trim.ToLower.Contains("10.11.24.21;") Or BIWindowsConnectionString.Trim.ToLower.Contains("svr21;") Then
			sBackup = "&nbsp;&nbsp;&nbsp;<span style='color:yellow;'>***** BACKUP *****</span>"
		End If
		litMainHeading.Text = "My Reports" & sBackup

		Dim iCompanyFKFromUserFK As Integer = 1
		Try
			Using libEd1 As New libEvolutionIntranet.clsIntranet(BIWindowsConnectionString)


				If UserIPAddress.Trim <> "" Then
					UserFKViewstate = libEd1.GetConsultantIDFromIP(UserIPAddress)
				End If


				' ----- for local testing -----
				' ----- for local testing -----
				' ----- for local testing -----
				' ----- for local testing -----
				If Request.Url.ToString.ToLower.Contains("localhost") Then
					If UserFKViewstate <= 0 Then
						UserFKViewstate = 378
					End If
				End If
				' ----- for local testing -----
				' ----- for local testing -----
				' ----- for local testing -----
				' ----- for local testing -----


				iCompanyFKFromUserFK = libEd1.GetUsersCompanyFK(UserFKViewstate)
			End Using

		Catch ex As Exception

			'ProcessError(ex.Message)
		Finally
		End Try


		'Dim ConnectionString As String = "Password=#PASSWORD#;Persist Security Info=True;User ID=#USER#;Initial Catalog=Ed;Data Source=" & SQLServer
		'EdConnectionString = ConnectionString.Replace("#PASSWORD#", "ed030304").Replace("#USER#", "ed")

	End Sub







End Class

