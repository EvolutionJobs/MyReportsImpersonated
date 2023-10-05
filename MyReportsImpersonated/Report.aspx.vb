

Imports System.Configuration
Imports System.Net
Imports Microsoft.Reporting.WebForms

Partial Class Report
	Inherits System.Web.UI.Page





	Public _isSqlTypesLoaded As Boolean = False

	'https://blogs.msdn.microsoft.com/sqlrsteamblog/2016/11/30/report-viewer-2016-control-update-now-available/
	'https://docs.microsoft.com/en-us/sql/reporting-services/application-integration/integrating-reporting-services-using-reportviewer-controls-get-started?view=sql-server-2017
	'https://blogs.msdn.microsoft.com/sqlrsteamblog/2015/10/20/position-report-parameters-the-way-you-want/
	'https://stackoverflow.com/questions/48987600/ssrs-toolbar-style-layout

	Protected Sub Page_Init(ByVal sender As Object,
			ByVal e As System.EventArgs) Handles Me.Init


		If _isSqlTypesLoaded = False Then
			SqlServerTypes.Utilities.LoadNativeAssemblies(Server.MapPath("~"))
			_isSqlTypesLoaded = True
		End If

		If HttpContext.Current.User.Identity.IsAuthenticated = False Then
			Response.Redirect("Error.aspx")
		End If

		BIApplicationFK = 0
		Dim sAppFK As String = ""
		If Request("appfk") <> "" Then
			sAppFK = Request("appfk")
		End If

		If sAppFK.Trim <> "" Then
			BIApplicationFK = CInt(sAppFK)
		End If

		Dim sRandNum As String = ""
		If Request("RandNum") <> "" Then
			sRandNum = Request("RandNum")
		End If

		Dim sAuthenticationType As String = HttpContext.Current.User.Identity.AuthenticationType
		Dim sWinUsername As String = HttpContext.Current.User.Identity.Name
		Dim impersonationLevel As String = ""

		'--------------------------------------
		Dim impersonationContext As System.Security.Principal.WindowsImpersonationContext
		Dim currentWindowsIdentity As System.Security.Principal.WindowsIdentity

		currentWindowsIdentity = CType(User.Identity, System.Security.Principal.WindowsIdentity)

		If currentWindowsIdentity IsNot Nothing Then
			impersonationLevel = currentWindowsIdentity.ImpersonationLevel
		End If

		impersonationContext = currentWindowsIdentity.Impersonate()
		'--------------------------------------


		SaveToLog(sWinUsername & " ******* Report_INIT_START [HttpContext.Current.User.Identity.Name=" & sWinUsername & "][" & sAuthenticationType & "]" & ";WindowsPassword=" & WindowsPassword & "[ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; [Before GetReportDetails] Request.Url.ToString=" & Request.Url.ToString)

		Page.Title = "Report"
		Dim ReportName As String = ""
		Dim PasswordProtect As Boolean = False
		Dim ReportPath As String = ""


		' get the Windows username and password and then if they exist use the impersonation, 
		' otherwise use old method (which will error for broken PCs)

		If WindowsPassword.Trim.Length > 0 Then


			SaveToLog(sWinUsername & " Report_____GetReportDetails_Impersonation_____ [" & sWinUsername & "][" & sAuthenticationType & "] WindowsUsername=" & WindowsUsername & "; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " [ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)

			'SaveToLog(sWinUsername & " Report_____GetReportDetails_Impersonation_____ EdConnectionString=" & EdConnectionString)

			' --- with manual impersonation ------
			' --- with manual impersonation ------
			' --- with manual impersonation ------
			Try
				' IF YOU USE IMPERSONATION HERE, THEN YOU must USE INLINE sql TO "Execute As Login = 'evolutionjobs\patels';"
				Dim objImpersonation As New UserImpersonation()
				If (objImpersonation.ImpersonateValidUser(WindowsUsername, WindowsPassword, WindowsDomain)) Then


					'' grant impersonation only when inserted into the log table
					'' grant impersonation only when inserted into the log table
					'' must use connection string to MASTER database
					'Dim sRetqqq As String = ""
					'Using lib1i As New clsDB(GrantImpersonateConnectionString)
					'	sRetqqq = lib1i.GrantImpersonation()
					'End Using
					'' grant impersonation only when inserted into the log table
					'' grant impersonation only when inserted into the log table




					' ---- impoersonate user -------
					' ---- impoersonate user -------
					' ---- impoersonate user was EdConnectionString-------
					Using lib1 As New clsDB(GrantImpersonateConnectionString)
						PasswordProtect = lib1.GetReportDetails_Impersonation(sAppFK, ReportPath, ReportName)
					End Using
					SaveToLog(sWinUsername & " Report_____PASSED_IMPERSONATION_____ [" & sWinUsername & "][" & sAuthenticationType & "][ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath & "; Request.Url.ToString=" & Request.Url.ToString)
					' ---- impoersonate user -------
					' ---- impoersonate user -------
					' ---- impoersonate user -------



					'SaveToLog(sWinUsername & " Report_____A-PASSED_IMPERSONATION_____")


					objImpersonation.UndoImpersonation()

					'SaveToLog(sWinUsername & " Report_____B-AFTER_UndoImpersonation")

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
				SaveToLog(sWinUsername & " ***** EXCEPTION ******_____Report_____GetReportDetails (ImpersonateValidUser)_____ [" & sWinUsername & "][" & sAuthenticationType & "]; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " [ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath & "; EX=" & ex.Message & ex.StackTrace)

			End Try
			' --- with manual impersonation ------
			' --- with manual impersonation ------
			' --- with manual impersonation ------


		Else
			SaveToLog(sWinUsername & " Report_____GetReportDetails (May fail)_____ [" & sWinUsername & "][" & sAuthenticationType & "]; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " [ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)


			Try
				Using lib1 As New clsDB(BIWindowsConnectionString)
					PasswordProtect = lib1.GetReportDetails(sAppFK, ReportPath, ReportName)
				End Using

				'SaveToLog(sWinUsername & " Report_____C-AFTER_NORMAL_GetReportDetails")

			Catch ex As Exception
				' this will error with NTAuthority/anonymous, so use this to set the session vars for impersonation
				SaveToLog(sWinUsername & " ***** EXCEPTION ******_____Report_____GetReportDetails (NORMAL)_____ [" & sWinUsername & "][" & sAuthenticationType & "]; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " [ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath & "; EX=" & ex.Message & ex.StackTrace)

				Response.Redirect("LoginWindows.aspx?appfk=" & sAppFK)
			End Try

		End If






		'' --- original ----
		'' --- original ----
		'' --- original ----
		'Using lib1 As New clsDB("Data Source=10.11.24.21;Initial Catalog=Ed;Persist Security Info=True;User id=ed;Password=ed030304;Connect Timeout=1800")
		'	PasswordProtect = lib1.GetReportDetails(sAppFK, ReportPath, ReportName)
		'End Using
		'' --- original ----
		'' --- original ----
		'' --- original ----





		' ======================================
		' ======================================
		' ======================================
		'SaveToLog(sWinUsername & " Report_____PostImpersonation; sRandNum=" & sRandNum & "; sAppFK=" & sAppFK & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)



		If ReportName.Trim <> "" Then
			Page.Title = ReportName
		End If

		'If sAppFK = "45" Then
		'    Page.Title = "Client Placements"
		'End If


		SaveToLog(sWinUsername & " Report_____PASSED_____ [" & sWinUsername & "][" & sAuthenticationType & "][ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)


		If PasswordProtect Then
			' redirect to login page
			If sRandNum.Trim = "" Then
				' only redirect if there is no randnum!!!
				SaveToLog(" Report_INIT___REDIRECT____ [" & sWinUsername & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; REDIRECT_TO: login.aspx?appfk=" & sAppFK)
				Response.Redirect("login.aspx?appfk=" & sAppFK)
			End If
		End If

		If sAppFK.Trim = "" Then
			' show MyReports report
			'ReportPath = "/WagesPages/OpeningPage"
			'ReportPath = "/BIPopulationLog/OpeningPage"
			ReportPath = "/MyReports/OpeningPage"
		End If


		'SaveToLog(sWinUsername & " Report_____Pre = Page.IsPostBack; sRandNum=" & sRandNum & "; sAppFK=" & sAppFK)

		' ======================================
		' ======================================
		' ======================================



		If Not Page.IsPostBack Then

			SaveToLog(sWinUsername & " Report_____Page.IsPostBack-START")

			'Set the processing mode for the ReportViewer to Remote  
			ReportViewer1.ProcessingMode = ProcessingMode.Remote

			Dim sLogParam As String = ""

			Try


				' +++++++++++++++++++++++++++++++++++++
				' +++++++++++++++++++++++++++++++++++++
				' +++++++++++++++++++++++++++++++++++++

				'ReportViewer1.ShowPageNavigationControls = True

				ReportViewer1.ShowExportControls = False
				If sWinUsername.Trim.ToLower.Contains(("Elliottn").ToLower) Or
					sWinUsername.Trim.ToLower.Contains(("patels").ToLower) Or
					sWinUsername.Trim.ToLower.Contains(("humphreys").ToLower) Then
					ReportViewer1.ShowExportControls = True
				End If

				Dim serverReport As ServerReport
				serverReport = ReportViewer1.ServerReport

				'Set the report server URL and report path  
				serverReport.ReportServerUrl = New Uri("http://evosvr05/reportserver")
				serverReport.ReportPath = ReportPath
				'serverReport.ReportPath = ReportPath '"/MyBI/OpeningPage"

				' ---- wages pages -----
				'http://evoweb3/Report.aspx?RandNum=815251&appfk=22
				' appfk=22; ReportPath=/WagesPages/OpeningPage
				' ---- wages pages -----


				'SaveToLog(sWinUsername & " Report_____ serverReport.ReportPath=" & serverReport.ReportPath)



				' =============== not working ====================
				'If WindowsPassword.Trim.Length > 0 Then


				'	SaveToLog(sWinUsername & " Report_____GetReportDetails_Impersonation_____ [" & sWinUsername & "][" & sAuthenticationType & "] WindowsUsername=" & WindowsUsername & "; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " [ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath)

				'	SaveToLog(sWinUsername & " Report_____GetReportDetails_Impersonation_____ EdConnectionString=" & EdConnectionString)

				'	' --- with manual impersonation ------
				'	' --- with manual impersonation ------
				'	' --- with manual impersonation ------
				'	Try
				'		' IF YOU USE IMPERSONATION HERE, THEN YOU must USE INLINE sql TO "Execute As Login = 'evolutionjobs\patels';"
				'		Dim objImpersonation As New UserImpersonation()
				'		If (objImpersonation.ImpersonateValidUser(WindowsUsername, WindowsPassword, WindowsDomain)) Then

				'			' ---- impoersonate user -------
				'			' ---- impoersonate user -------
				'			' ---- impoersonate user -------


				'			'Dim myCredentials As New NetworkCredential("", "", "")
				'			'myCredentials.Domain = WindowsDomain
				'			'myCredentials.UserName = WindowsUsername
				'			'myCredentials.Password = WindowsPassword
				'			'serverReport.ReportServerCredentials = myCredentials


				'			'Get a reference to the default credentials  
				'			Dim credentials As System.Net.ICredentials
				'			credentials = System.Net.CredentialCache.DefaultCredentials
				'			serverReport.ReportServerCredentials = credentials


				'			SaveToLog(sWinUsername & " Report_____PASSED_IMPERSONATION_____ [" & sWinUsername & "][" & sAuthenticationType & "][ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath & "; Request.Url.ToString=" & Request.Url.ToString)

				'			' ---- impoersonate user -------
				'			' ---- impoersonate user -------
				'			' ---- impoersonate user -------




				'			objImpersonation.UndoImpersonation()
				'		Else
				'			objImpersonation.UndoImpersonation()
				'			Response.Redirect("Error.aspx?e=User+does+not+have+enough+permissions+to+perform+the+task")

				'			'Throw New Exception("User does not have enough permissions to perform the task")
				'		End If
				'	Catch ex As Exception
				'		' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'		' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'		' DO NOT CAPTURE ERROR HERE AS IT WORKS BY IGNORING IT!!!!
				'		'Response.Redirect("Error.aspx?e=" & ex.Message.Replace(" ", "+"))
				'		'MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				'		SaveToLog(sWinUsername & " ***** EXCEPTION ******_____Report_____GetReportDetails (ImpersonateValidUser)_____ [" & sWinUsername & "][" & sAuthenticationType & "]; WindowsPassword.Trim.Length=" & WindowsPassword.Trim.Length & " [ImpersonationLevel=" & impersonationLevel & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; PasswordProtect=" & PasswordProtect & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportName=" & ReportName & "; ReportPath=" & ReportPath & "; EX=" & ex.Message & ex.StackTrace)

				'	End Try
				'	' --- with manual impersonation ------
				'	' --- with manual impersonation ------
				'	' --- with manual impersonation ------


				'Else



				'End If
				' =============== not working ====================





				''============== this doesn't work (inside impersonation or outside)  ===================
				'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
				'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
				'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
				'Dim mycred As IReportServerCredentials = New CustomReportCredentials(WindowsUsername, WindowsPassword, WindowsDomain)
				''Set Processing Mode (Grabbing Report from Reporting Server)
				''Set the Report Server, Path, ServerCredentials
				'serverReport.ReportServerCredentials = mycred
				'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
				'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
				'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
				''============== this doesn't work (inside impersonation or outside)  ===================





				'SaveToLog(sWinUsername & " Report_____Set ReportPath=" & ReportPath)

				'/MyBI/OpeningPage
				'/WagesPages/OpeningPage

				If sRandNum.Trim = "" And sAppFK.Trim <> "" Then
					sRandNum = sAppFK.Trim
				End If

				SaveToLog("Report_SHOW [" & sWinUsername & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportPath=" & ReportPath)

				'If sRandNum.Trim <> "" Then
				'             'Create the sales order number report parameter  
				'             Dim SubQueryYearMonthName As New ReportParameter()
				'             SubQueryYearMonthName.Name = "RandNum"
				'             SubQueryYearMonthName.Values.Add(sRandNum)

				'	'Set the report parameters for the report  
				'	Dim parameters() As ReportParameter = {SubQueryYearMonthName}
				'	serverReport.SetParameters(parameters)

				'End If




				Dim plist As New List(Of Integer)

				Dim SubQueryYearMonthName As New ReportParameter("RandNum")
				Dim ServerPath As New ReportParameter("ServerPath")
				Dim par_p2 As New ReportParameter("P2")
				Dim par_p3 As New ReportParameter("P3")
				Dim par_p4 As New ReportParameter("P4")
				Dim par_p5 As New ReportParameter("P5")
				Dim par_p6 As New ReportParameter("P6")
				Dim par_p7 As New ReportParameter("P7")
				Dim par_p8 As New ReportParameter("P8")
				Dim par_p9 As New ReportParameter("P9")
				Dim par_p10 As New ReportParameter("P10")
				Dim par_p11 As New ReportParameter("P11")
				Dim par_p12 As New ReportParameter("P12")
				Dim par_p13 As New ReportParameter("P13")
				Dim par_p14 As New ReportParameter("P14")
				Dim par_p15 As New ReportParameter("P15")


				Dim sServerPath As String = "http://evoweb3/Login.aspx?appfk="

				' ---- exclude wages pages -----
				'Parameter 'ServerPath' does not exist on this report
				'http://evoweb3/Report.aspx?RandNum=815251&appfk=22
				' appfk=22; ReportPath=/WagesPages/OpeningPage
				If sAppFK = "22" And sRandNum.Trim <> "" Then
					sServerPath = ""
				End If
				' ---- exclude wages pages -----


				If sRandNum.Trim <> "" Then
					SubQueryYearMonthName.Values.Add(sRandNum)
					plist.Add(0)
					sLogParam &= "sRandNum=" & sRandNum & ";"
				End If
				If sServerPath <> "" Then
					ServerPath.Values.Add(sServerPath)
					plist.Add(1)
					sLogParam &= "ServerPath=" & sServerPath & ";"


				End If
				If Request("p2") <> "" Then
					par_p2.Values.Add(Request("p2"))
					plist.Add(2)
					sLogParam &= "p2=" & Request("p2") & ";"

				End If
				If Request("p3") <> "" Then
					par_p3.Values.Add(Request("p3"))
					plist.Add(3)
					sLogParam &= "p3=" & Request("p3") & ";"
				End If
				If Request("p4") <> "" Then
					par_p4.Values.Add(Request("p4"))
					plist.Add(4)
					sLogParam &= "p4=" & Request("p4") & ";"
				End If
				If Request("p5") <> "" Then
					par_p5.Values.Add(Request("p5"))
					plist.Add(5)
					sLogParam &= "p5=" & Request("p5") & ";"
				End If
				If Request("p6") <> "" Then
					par_p6.Values.Add(Request("p6"))
					plist.Add(6)
					sLogParam &= "p6=" & Request("p6") & ";"
				End If
				If Request("p7") <> "" Then
					par_p7.Values.Add(Request("p7"))
					plist.Add(7)
					sLogParam &= "p7=" & Request("p7") & ";"
				End If
				If Request("p8") <> "" Then
					par_p8.Values.Add(Request("p8"))
					plist.Add(8)
					sLogParam &= "p8=" & Request("p8") & ";"
				End If
				If Request("p9") <> "" Then
					par_p9.Values.Add(Request("p9"))
					plist.Add(9)
					sLogParam &= "p9=" & Request("p9") & ";"
				End If
				If Request("p10") <> "" Then
					par_p10.Values.Add(Request("p10"))
					plist.Add(10)
					sLogParam &= "p10=" & Request("p10") & ";"
				End If
				If Request("p11") <> "" Then
					par_p11.Values.Add(Request("p11"))
					plist.Add(11)
					sLogParam &= "p11=" & Request("p11") & ";"
				End If
				If Request("p12") <> "" Then
					par_p12.Values.Add(Request("p12"))
					plist.Add(12)
					sLogParam &= "p12=" & Request("p12") & ";"
				End If
				If Request("p13") <> "" Then
					par_p13.Values.Add(Request("p13"))
					plist.Add(13)
					sLogParam &= "p13=" & Request("p13") & ";"
				End If
				If Request("p14") <> "" Then
					par_p14.Values.Add(Request("p14"))
					plist.Add(14)
					sLogParam &= "p14=" & Request("p14") & ";"
				End If
				If Request("p15") <> "" Then
					par_p15.Values.Add(Request("p15"))
					plist.Add(15)
					sLogParam &= "p15=" & Request("p15") & ";"
				End If



				'Set the report parameters for the report  
				Dim parameters(plist.Count - 1) As ReportParameter

				If plist.Count > 0 Then


					For iCount As Integer = 0 To plist.Count - 1
						If plist.Item(iCount) = 0 Then
							parameters(iCount) = SubQueryYearMonthName
						ElseIf plist.Item(iCount) = 1 Then
							parameters(iCount) = ServerPath
						ElseIf plist.Item(iCount) = 2 Then
							parameters(iCount) = par_p2
						ElseIf plist.Item(iCount) = 3 Then
							parameters(iCount) = par_p3
						ElseIf plist.Item(iCount) = 4 Then
							parameters(iCount) = par_p4
						ElseIf plist.Item(iCount) = 5 Then
							parameters(iCount) = par_p5
						ElseIf plist.Item(iCount) = 6 Then
							parameters(iCount) = par_p6
						ElseIf plist.Item(iCount) = 7 Then
							parameters(iCount) = par_p7
						ElseIf plist.Item(iCount) = 8 Then
							parameters(iCount) = par_p8
						ElseIf plist.Item(iCount) = 9 Then
							parameters(iCount) = par_p9
						ElseIf plist.Item(iCount) = 10 Then
							parameters(iCount) = par_p10
						ElseIf plist.Item(iCount) = 11 Then
							parameters(iCount) = par_p11
						ElseIf plist.Item(iCount) = 12 Then
							parameters(iCount) = par_p12
						ElseIf plist.Item(iCount) = 13 Then
							parameters(iCount) = par_p13
						ElseIf plist.Item(iCount) = 14 Then
							parameters(iCount) = par_p14
						ElseIf plist.Item(iCount) = 15 Then
							parameters(iCount) = par_p15
						End If
					Next iCount


					SaveToLog(sWinUsername & " Report_____plist.Count=" & plist.Count & "; sLogParam=" & sLogParam)

				End If

				serverReport.SetParameters(parameters)

				serverReport.Refresh()

				' +++++++++++++++++++++++++++++++++++++
				' +++++++++++++++++++++++++++++++++++++
				' +++++++++++++++++++++++++++++++++++++


			Catch ex As Exception
				SaveToLog(sWinUsername & " Report_____***EXCEPTION*** setting parameters; EX=" & ex.Message & "; sLogParam=" & sLogParam)
			End Try















			''============== this doesn't work (inside impersonation or outside)  ===================
			'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
			'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
			'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
			'Dim mycred As IReportServerCredentials = New CustomReportCredentials(WindowsUsername, WindowsPassword, WindowsDomain)
			''Set Processing Mode (Grabbing Report from Reporting Server)
			''Set the Report Server, Path, ServerCredentials
			'ReportViewer1.ServerReport.ReportServerCredentials = mycred
			'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
			'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
			'' -------- not sure if this is needed with or without the ImpersonateValidUser above ------
			''============== this doesn't work (inside impersonation or outside)  ===================











			''-------------- try specific credentials ---------------
			'' this doesn't work without a DataSource (dsCrendtials.Name)
			''DataSourceCredentials dsCrendtials = New DataSourceCredentials();
			''dsCrendtials.Name = "DataSource1";  // Default Is this you may have different name
			''         dsCrendtials.UserId = "MyUser";  // Set this To be a textbox
			''         dsCrendtials.Password = "MyPassword";  // Set this To be a textbox
			''         ReportViewer.ServerReport.SetDataSourceCredentials(New DataSourceCredentials[] { dsCrendtials });

			'Dim dsCrendtials As DataSourceCredentials = New DataSourceCredentials()
			'dsCrendtials.Name = "DataSource1" ' this is required - doesn;t work without this
			'dsCrendtials.UserId = "evolutionjobs\patels"
			'dsCrendtials.Password = "Warringt0n18"
			'serverReport.SetDataSourceCredentials(New DataSourceCredentials() {dsCrendtials})
			''-------------- try specific credentials ---------------



			' this is required for WBV report to pass in appfk=76, but it error for wages pages
			' hence solution is to have same one or two paramters for all
			'If sAppFK.Trim <> "" Then

			'	If sAppFK.Trim <> "" Then
			'		Dim SubQueryAppFK As New ReportParameter()
			'		SubQueryAppFK.Name = "appfk"
			'		SubQueryAppFK.Values.Add(sAppFK)

			'		'Set the report parameters for the report  
			'		Dim parameters() As ReportParameter = {SubQueryAppFK}
			'		serverReport.SetParameters(parameters)
			'	End If

			'End If






		End If 'If Not Page.IsPostBack Then


		SaveToLog("Report_END [" & sWinUsername & "]: sAppFK=" & sAppFK & "; sRandNum=" & sRandNum & "; Request.Url.ToString=" & Request.Url.ToString & "; ReportPath=" & ReportPath)


	End Sub

End Class
