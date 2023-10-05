Imports System.Data.SqlClient
Imports System.Data
Imports System.Runtime.InteropServices

Public Class clsDB
	Implements IDisposable

	Public ErrorOccurred As Boolean = False

	Private moConn As SqlConnection
	Private mstrDSN As String

	Const MODULENAME = "clsDBA"

	''---- for impersonation ------
	'' You should include System.Runtime.InteropServices
	'<DllImport(“c:\Windows\System32\advapi32.dll”)>
	'Public Shared Function LogonUser(lpszUserName As String, lpszDomain As String, lpszPassword As String, dwLogonType As Integer, dwLogonProvider As Integer, ByRef phToken As Integer) As Boolean

	'End Function
	'' ---- for impersonation ------




	Public Sub New(ByVal DSN As String)
		mstrDSN = DSN
		If mstrDSN Is Nothing Then Exit Sub
		If mstrDSN.Length = 0 Then Exit Sub
		OpenSQLConnection()
	End Sub

	Public Property DSN() As String
		Get
			Return mstrDSN
		End Get
		Set(ByVal Value As String)
			mstrDSN = Value
			If mstrDSN Is Nothing Then Exit Property
			If mstrDSN.Length = 0 Then Exit Property
			OpenSQLConnection()
		End Set
	End Property


	Public Property EnteredUsername() As String = ""
	Public Property EnteredPassword() As String = ""

	Public Function ToSecureString(ByVal plainString As String) As System.Security.SecureString
		If plainString Is Nothing Then Return Nothing
		Dim secureString As System.Security.SecureString = New System.Security.SecureString()

		For Each c As Char In plainString.ToCharArray()
			secureString.AppendChar(c)
		Next

		Return secureString
	End Function


	Private Sub OpenSQLConnection()
		'Destroy old object - bit extreme but don't want any 
		'increasing open connections and the connection
		'pool should minimise overheads of re-opening


		'' ----------------- manual impersonation --------------
		'' ----------------- manual impersonation --------------
		'' ----------------- manual impersonation --------------

		'If TypeOf (HttpContext.Current.User) Is WindowsPrincipal Then

		'	' Impersonate the IIS identity.

		'	If HttpContext.Current.User.Identity.IsAuthenticated = False Then
		'		HttpContext.Current.Response.Redirect("Error.aspx?Error=User is Not Authenticated by Windows.")
		'	End If

		'	Dim Id As WindowsIdentity

		'	Id = DirectCast(HttpContext.Current.User.Identity, WindowsIdentity)

		'	Dim ImpersonateContext As WindowsImpersonationContext

		'	ImpersonateContext = Id.Impersonate()

		'	'Dim Principal As WindowsPrincipal = DirectCast(HttpContext.Current.User, WindowsPrincipal)
		'	'Dim Identity As WindowsIdentity = DirectCast(Principal.Identity, WindowsIdentity)
		'	Dim Token As IntPtr = Id.Token

		'	Try
		'		If Not (moConn Is Nothing) Then
		'			If moConn.State = ConnectionState.Open Then
		'				moConn.Close()
		'				moConn.Dispose()
		'				moConn = Nothing
		'			End If
		'		End If
		'	Catch sqlex As SqlException
		'		Throw sqlex

		'		logError("OpenSQLConnection(Close exiting conn); Token=" & Token.ToString, sqlex.Message, sqlex.StackTrace)

		'		ErrorOccurred = True
		'	End Try

		'	moConn = New SqlConnection(mstrDSN)


		'	Try
		'		moConn.Open()
		'	Catch sqlex As SqlException
		'		Throw sqlex
		'		logError("OpenSQLConnection(); Token=" & Token.ToString, sqlex.Message, sqlex.StackTrace)
		'	End Try

		'	' Revert to the original ID as shown here.
		'	ImpersonateContext.Undo()
		'Else

		'	' User isn’t Windows authenticated.
		'	' Throw an error or take other steps.
		'	HttpContext.Current.Response.Redirect("Error.aspx?Error=User is Not Authenticated by Windows")
		'End If
		'' ----------------- manual impersonation --------------
		'' ----------------- manual impersonation --------------
		'' ----------------- manual impersonation --------------



		'' ------------------ try userid and password -------------------------
		'' ------------------ try userid and password -------------------------
		'' ------------------ try userid and password -------------------------
		'Try
		'	If Not (moConn Is Nothing) Then
		'		If moConn.State = ConnectionState.Open Then
		'			moConn.Close()
		'			moConn.Dispose()
		'			moConn = Nothing
		'		End If
		'	End If
		'Catch sqlex As SqlException
		'	Throw sqlex
		'	'LogError(Nothing, sqlex.Message)
		'	ErrorOccurred = True
		'End Try

		'moConn = New SqlConnection(mstrDSN)

		'Try

		'	'SecureString pwd = txtPwd.SecurePassword;  
		'	'pwd.MakeReadOnly();  
		'	'SqlCredential cred = New SqlCredential(txtUserId.Text, pwd);  
		'	'conn.Credential = cred;  
		'	'conn.Open(); 

		'	'SecureString theSecureString = new NetworkCredential("", "myPass").SecurePassword

		'	'Dim pwd As System.Security.SecureString = ToSecureString("Warringt0n18") 'New Net.NetworkCredential("", "password").SecurePassword
		'	'pwd.MakeReadOnly()
		'	'Dim cred As SqlCredential = New SqlCredential("evolutionjobs\patels", pwd)
		'	'moConn.Credential = cred


		'	moConn.Open()
		'Catch sqlex As SqlException
		'	Throw sqlex
		'	logError("OpenSQLConnection()", sqlex.Message, sqlex.StackTrace)
		'End Try
		'' ------------------ try userid and password -------------------------
		'' ------------------ try userid and password -------------------------
		'' ------------------ try userid and password -------------------------



		' ------------------ original -------------------------
		' ------------------ original -------------------------
		' ------------------ original -------------------------
		Try
			If Not (moConn Is Nothing) Then
				If moConn.State = ConnectionState.Open Then
					moConn.Close()
					moConn.Dispose()
					moConn = Nothing
				End If
			End If
		Catch sqlex As SqlException
			Throw sqlex
			'LogError(Nothing, sqlex.Message)
			ErrorOccurred = True
		End Try

		moConn = New SqlConnection(mstrDSN)

		Try
			moConn.Open()
		Catch sqlex As SqlException
			'Throw sqlex
			' probably will be: Login failed for user 'NT AUTHORITY\ANONYMOUS LOGON'.
			'logError("OpenSQLConnection()", sqlex.Message, sqlex.StackTrace)
			If WindowsPassword.Trim.Length = 0 Then
				HttpContext.Current.Response.Redirect("LoginWindows.aspx?appfk=" & BIApplicationFK)
			End If

		End Try
		' ------------------ original -------------------------
		' ------------------ original -------------------------
		' ------------------ original -------------------------







	End Sub




	Private Sub CloseSQLConnection()
		If Not (moConn Is Nothing) Then
			Try
				If moConn.State <> ConnectionState.Closed Then
					moConn.Close()
				End If
			Catch sqlex As SqlException
				Throw sqlex
				'LogError(Nothing, sqlex.Message)
				ErrorOccurred = True
			Finally
				moConn.Dispose()
				moConn = Nothing
			End Try
		End If
	End Sub

	Public Sub Dispose() Implements System.IDisposable.Dispose
		CloseSQLConnection()
	End Sub

	Public Function RunSQL(ByVal sSql As String) As DataTable

		RunSQL = New DataTable
		Dim ocmd As New SqlCommand(sSql, moConn)
		Dim sqladp As New SqlDataAdapter(ocmd)

		'reset the error flag
		ErrorOccurred = False

		Try
			sqladp.Fill(RunSQL)

		Catch ex As Exception
			'System.Web.HttpContext.Current.Response.Write(ex.Message)
			Throw ex
			'LogError(ex, sSql)
			ErrorOccurred = True
		Finally
			sqladp.Dispose()
			ocmd.Dispose()
		End Try
	End Function


	Public Function RunSQLScalar(ByVal inSQL As String) As String

		Dim objResult As New Object
		Dim sResult As String = ""

		Using oCMD As New SqlCommand(inSQL, moConn)
			oCMD.CommandType = CommandType.Text
			objResult = oCMD.ExecuteScalar
		End Using

		If objResult IsNot Nothing AndAlso Not IsDBNull(objResult) Then
			sResult = CStr(objResult)
		End If

		Return sResult
	End Function






	Public Function InsertUpdateImpersonatedUser() As String

		Dim outStatus As String = ""
		Dim oRet As New Object
		Dim sRet As String = ""

		If WindowsPassword.Trim.Length > 0 Then
			Dim sSQL As String = "ed..usp_MyReports_InsertUpdateImpersonatedUser"
			Dim oCmd As New SqlCommand(sSQL, moConn)
			oCmd.CommandType = CommandType.StoredProcedure

			Try
				oCmd.Parameters.AddWithValue("@SQLAccount", ImpersonatingAccount)
				oCmd.Parameters.AddWithValue("@WindowsUsername", WindowsUsername)
				oCmd.Parameters.Add("@outWindowsUsername", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output

				oRet = oCmd.ExecuteScalar()

				If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
					sRet = (oRet).ToString
				End If

				If Not IsDBNull(oCmd.Parameters("@outWindowsUsername").Value) Then
					outStatus = CStr(oCmd.Parameters("@outWindowsUsername").Value)
				End If

			Catch ex As Exception
				logError("InsertUpdateImpersonatedUser(EVOWEB3)", "ERROR: InsertUpdateImpersonatedUser:" & "; " & ex.Message, ex.StackTrace)
				Throw ex
			Finally
				oCmd.Dispose()
			End Try

		End If

		Return outStatus
	End Function



	Public Function GrantImpersonation() As String

		Dim sRetaaa As String = ""

		Dim ImpersonatedUserUpdateInsert As String = ""
		ImpersonatedUserUpdateInsert = InsertUpdateImpersonatedUser()

		If ImpersonatedUserUpdateInsert.Trim.ToLower = "inserted" Then

			Dim oRet As New Object
			Dim sSQL As String = ""

			If WindowsPassword.Trim.Length > 0 Then

				sSQL = "GRANT IMPERSONATE ON LOGIN:: [EVOLUTIONJOBS\" & WindowsUsername & "] TO [" & ImpersonatingAccount & "]; "

				Dim oCmd As New SqlCommand(sSQL, moConn)
				oCmd.CommandType = CommandType.Text

				Try

					oRet = oCmd.ExecuteScalar()

					If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
						sRetaaa = (oRet).ToString
					End If

				Catch ex As Exception
					logError("GrantImpersonation(EVOWEB3)", "ERROR: GrantImpersonation:" & "; " & ex.Message, ex.StackTrace)

					Throw ex
				Finally
					oCmd.Dispose()
				End Try

			End If

		End If

		Return sRetaaa
	End Function




	' if you use impersonation, then the inline SQL is required!!!
	' runs under the SQL login for impersonation (e.g. Ed or ReportAccess)
	Public Function GetReportDetails_Impersonation(
	 ByVal AppFK As String,
	 ByRef ReportPath As String,
	 ByRef ReportName As String
	 ) As Boolean

		SaveToLog(WindowsUsername & "; clsDB (START) - GetReportDetails_Impersonation_____GetAppDetail_____ AppFK=" & AppFK)
		'SaveToLog(WindowsUsername & "; clsDB (START) - mstrDSN=" & mstrDSN)


		Dim iAppfk As Integer = 0
		If AppFK.Trim <> "" Then
			iAppfk = CInt(AppFK.Trim)
		End If

		Dim PasswordProtect As Boolean = False
		Dim dtOutput As New DataTable
		Dim lRet As Long = -1

		' if you use impersonation, then the inline SQL is required!!!
		' if you use impersonation, then the inline SQL is required!!!
		' if you use impersonation, then the inline SQL is required!!!
		' if you use impersonation, then the inline SQL is required!!!
		' if you use impersonation, then the inline SQL is required!!!

		' "GRANT IMPERSONATE ON USER:: [evolutionjobs\patels] TO [ed]" &
		' "exec bi..usp_BI_MyBI_Populate @ActionType='GetAppDetail', @ViewAsUser=null, @AppFK=" & iAppfk & "; " &

		Dim sSQL As String = "" &
			"" &
			"create table #Data " &
			"( " &
			"    PasswordProtect int, " &
			"	SSRS_ParamName varchar(1000), " &
			"	ReportName varchar(100) " &
			"); " &
			"" &
			"" &
			"" &
			"Execute As Login = 'EVOLUTIONJOBS\" & WindowsUsername & "'; " &
			"" &
			"   insert into #Data " &
			"	( " &
			"		PasswordProtect, " &
			"		SSRS_ParamName, " &
			"		ReportName " &
			"	) " &
			"" &
			"exec bi..usp_BI_MyBI_Populate @ActionType='GetAppDetail', @ViewAsUser='EVOLUTIONJOBS\" & WindowsUsername & "', @AppFK=" & iAppfk & "; " &
			"" &
			"Revert;" &
			"" &
			"SELECT * FROM #Data " &
			"drop table #Data" &
			""

		' EXEC usp_BI_SSRSReportLogging 'LogReportAccess', 31 
		' usp_BI_MyBI_PopulateXXXKILLME

		'Dim sSQL As String = "exec bi..usp_BI_MyBI_PopulateXXXKILLME @ActionType='GetAppDetail', @ViewAsUser='EVOLUTIONJOBS\" & WindowsUsername & "', @AppFK=" & iAppfk & "; "

		'Dim sSQL As String = "bi..usp_BI_MyBI_PopulateXXXKILLME"

		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.Text

		Try
			oCmd.Parameters.AddWithValue("@ActionType", "GetAppDetail")
			oCmd.Parameters.AddWithValue("@ViewAsUser", "EVOLUTIONJOBS\" & WindowsUsername)
			oCmd.Parameters.AddWithValue("@AppFK", iAppfk)

			Using daLoader As New SqlDataAdapter(oCmd)
				daLoader.Fill(dtOutput)
			End Using
		Catch ex As Exception
			SaveToLog(WindowsUsername & "; EXCEPTION --- clsDB (END) - GetReportDetails_Impersonation_____GetAppDetail_____ [iAppfk=" & iAppfk & "][PasswordProtect=" & PasswordProtect & "][ReportName=" & ReportName & "]; ReportPath=" & ReportPath & "; EX=" & ex.Message & ex.StackTrace)

			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		'SaveToLog(WindowsUsername & "; clsDB - GetReportDetails_Impersonation (After getting dtOutput)")

		Try
			If dtOutput IsNot Nothing Then
				If dtOutput.Rows.Count > 0 Then

					'SaveToLog(WindowsUsername & "; clsDB (After getting dtOutput) dtOutput.Rows.Count=" & dtOutput.Rows.Count)

					For Each row As DataRow In dtOutput.Rows
						If Not IsDBNull(row("PasswordProtect")) Then
							PasswordProtect = CBool(row("PasswordProtect"))
						End If
						If Not IsDBNull(row("SSRS_ParamName")) Then
							ReportPath = CStr(row("SSRS_ParamName"))
						End If
						If Not IsDBNull(row("ReportName")) Then
							ReportName = CStr(row("ReportName"))
						End If
					Next
				End If
			End If
		Catch ex As Exception
			' an error with 1 unknow dield being returned, e.g. "Incorrect Application name"
			SaveToLog(WindowsUsername & "; clsDB - ***** EXCEPTION **** getting values from dtOutput: " & ex.Message)
		End Try


		SaveToLog(WindowsUsername & "; clsDB (END) - GetReportDetails_Impersonation_____GetAppDetail_____ [iAppfk=" & iAppfk & "][PasswordProtect=" & PasswordProtect & "][ReportName=" & ReportName & "]; ReportPath=" & ReportPath)


		Return PasswordProtect
	End Function



	' original - runs with WindowsConnectionString
	Public Function GetReportDetails(
	 ByVal AppFK As String,
	 ByRef ReportPath As String,
	 ByRef ReportName As String
	 ) As Boolean

		Dim iAppfk As Integer = 0
		If AppFK.Trim <> "" Then
			iAppfk = CInt(AppFK.Trim)
		End If

		Dim PasswordProtect As Boolean = False
		Dim dtOutput As New DataTable
		Dim lRet As Long = -1
		Dim sSQL As String = "bi..usp_BI_MyBI_Populate" '"bi..usp_BI_MyBI_Populate_Impersonate"
		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.StoredProcedure

		Try
			'create   proc [dbo].[usp_BI_MyBI_Populate](
			'    @ActionType varchar(100) = null, --     <SH>   2019-03-14
			'    @ViewAsUser varchar(100) = null, --     <SH>   2019-04-01
			'    @AppFK      int          = null  --     <SH>   2019-04-05
			oCmd.Parameters.AddWithValue("@ActionType", "GetAppDetail")
			oCmd.Parameters.AddWithValue("@ViewAsUser", DBNull.Value)
			oCmd.Parameters.AddWithValue("@AppFK", iAppfk)
			'oCmd.Parameters.Add("@PasswordProtect", SqlDbType.Bit).Direction = ParameterDirection.Output
			'oCmd.Parameters.Add("@SSRS_ParamName", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output

			Using daLoader As New SqlDataAdapter(oCmd)
				daLoader.Fill(dtOutput)
			End Using

			'If Not IsDBNull(oCmd.Parameters("@PasswordProtect").Value) Then
			'    PasswordProtect = CBool(oCmd.Parameters("@PasswordProtect").Value)
			'End If
			'If Not IsDBNull(oCmd.Parameters("@SSRS_ParamName").Value) Then
			'    ReportPath = CStr(oCmd.Parameters("@SSRS_ParamName").Value)
			'End If
		Catch ex As Exception
			SaveToLog(WindowsUsername & "; EXCEPTION --- clsDB (END) - GetReportDetails_____GetAppDetail(NORMAL)_____ [iAppfk=" & iAppfk & "][PasswordProtect=" & PasswordProtect & "][ReportName=" & ReportName & "]; ReportPath=" & ReportPath & "; EX=" & ex.Message & ex.StackTrace)

			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		'SaveToLog(WindowsUsername & "; clsDB - GetReportDetails (After getting dtOutput)")

		If dtOutput IsNot Nothing Then
			If dtOutput.Rows.Count > 0 Then
				For Each row As DataRow In dtOutput.Rows
					If Not IsDBNull(row("PasswordProtect")) Then
						PasswordProtect = CBool(row("PasswordProtect"))
					End If
					If Not IsDBNull(row("SSRS_ParamName")) Then
						ReportPath = CStr(row("SSRS_ParamName"))
					End If
					If Not IsDBNull(row("ReportName")) Then
						ReportName = CStr(row("ReportName"))
					End If
				Next
			End If
		End If


		'SaveToLog(WindowsUsername & "; clsDB (END) - GetReportDetails_____GetAppDetail(NORMAL)_____ [iAppfk=" & iAppfk & "][PasswordProtect=" & PasswordProtect & "][ReportName=" & ReportName & "]; ReportPath=" & ReportPath)


		Return PasswordProtect
	End Function




	Public Function CheckPassword_Impersonation(
	 ByVal sPassword As String,
	 ByVal inApplicationFK As Integer
	 ) As DataTable

		' unlike GetCompanies, this gets all companies

		Dim dtOutput As New DataTable

		'Dim sSQL As String = "bi..usp_BI_PWPort_Manager"
		'old---Dim sSQL As String = "bi..usp_BI_PWPort_Manager @GiveItAGo,@OnwardApplicationFK,@Password"

		Dim sSQL As String = "" &
			"" &
			"create table #Data " &
			"( " &
			"    PasswordStatus varchar(100), " &
			"	ResultFK int, " &
			"	ReportURL varchar(1000), " &
			"	Seed varchar(100) " &
			"); " &
			"" &
			"" &
			"" &
			"Execute As Login = 'EVOLUTIONJOBS\" & WindowsUsername & "'; " &
			"" &
			"   insert into #Data " &
			"	( " &
			"		PasswordStatus, " &
			"		ResultFK, " &
			"		ReportURL, " &
			"		Seed " &
			"	) " &
			"" &
			"exec bi..usp_BI_PWPort_Manager @ActionType='GiveItAGo',@Param1=" & inApplicationFK & ",@Param2='" & sPassword & "'; " &
			"" &
			"Revert;" &
			"" &
			"SELECT * FROM #Data " &
			"drop table #Data" &
			"" &
			""
		'Select Case@Result as PasswordStatus,
		'      @ResultFK as ResultFK,
		'      @url as ReportURL,
		'      @RandomValue as Seed

		SaveToLog(WindowsUsername & "; clsDB (Pre data -  CheckPassword_Impersonation)")

		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.Text

		Try
			oCmd.Parameters.AddWithValue("@ActionType", "GiveItAGo")
			oCmd.Parameters.AddWithValue("@Param1", inApplicationFK)
			oCmd.Parameters.AddWithValue("@Param2", sPassword)

			Using daLoader As New SqlDataAdapter(oCmd)
				daLoader.Fill(dtOutput)
			End Using
		Catch ex As Exception
			SaveToLog(WindowsUsername & "; clsDB ***EXCEPTION****(POST data -  CheckPassword_Impersonation) Ex=" * ex.Message & ex.StackTrace)
			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		SaveToLog(WindowsUsername & "; clsDB (POST data -  CheckPassword_Impersonation)")

		Return dtOutput
	End Function




	Public Function CheckPassword(
	 ByVal sPassword As String,
	 ByVal inApplicationFK As Integer
	 ) As DataTable

		' unlike GetCompanies, this gets all companies

		Dim dtOutput As New DataTable

		Dim sSQL As String = "bi..usp_BI_PWPort_Manager"
		'Dim sSQL As String = "bi..usp_BI_PWPort_Manager @GiveItAGo,@OnwardApplicationFK,@Password"
		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.StoredProcedure

		Try
			oCmd.Parameters.AddWithValue("@ActionType", "GiveItAGo")
			oCmd.Parameters.AddWithValue("@Param1", inApplicationFK)
			oCmd.Parameters.AddWithValue("@Param2", sPassword)

			Using daLoader As New SqlDataAdapter(oCmd)
				daLoader.Fill(dtOutput)
			End Using
		Catch ex As Exception
			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		Return dtOutput
	End Function


	' original
	Public Function ChangePassword(
	 ByRef Result As String,
	 ByVal CurrentPW As String,
	 ByVal NewPW1 As String,
	 ByVal NewPW2 As String
	 ) As Integer

		Dim ResultFK As Integer = 0

		'declare @Result      varchar(1000),@ResultFK int

		'exec bi..usp_BI_SSRSSec 
		'    @Result output,
		'    @ResultFK output ,
		'    @CurrentPW,
		'        'Usr_Change',
		'    @NewPW1,
		'    @NewPW2

		'select @Result as Result,@ResultFK as ResultFK

		Dim oRet As New Object
		Dim lRet As Long = -1
		Dim sSQL As String = "bi..usp_BI_SSRSSec"
		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.StoredProcedure

		Try

			'@Result      varchar(100) output,
			'@ResultFK    int          output,
			'@Param1      varchar(100),
			'@RequestType varchar(100),
			'@Param2      varchar(100) = null,
			'@Param3      varchar(100) = null,
			'@Param4      varchar(100) = null
			oCmd.Parameters.Add("@Result", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
			oCmd.Parameters.Add("@ResultFK", SqlDbType.Int).Direction = ParameterDirection.Output
			oCmd.Parameters.AddWithValue("@Param1", CurrentPW)
			oCmd.Parameters.AddWithValue("@RequestType", "Usr_Change")
			oCmd.Parameters.AddWithValue("@Param2", NewPW1)
			oCmd.Parameters.AddWithValue("@Param3", NewPW2)
			oCmd.Parameters.AddWithValue("@Param4", DBNull.Value)

			oRet = oCmd.ExecuteScalar()

			If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
				lRet = CLng(oRet)
			End If

			If Not IsDBNull(oCmd.Parameters("@Result").Value) Then
				Result = CStr(oCmd.Parameters("@Result").Value)
			End If
			If Not IsDBNull(oCmd.Parameters("@ResultFK").Value) Then
				ResultFK = CInt(oCmd.Parameters("@ResultFK").Value)
			End If
		Catch ex As Exception
			logError("ChangePassword", "ERROR: change password:" & "; " & ex.Message, ex.StackTrace)

			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		Return ResultFK
	End Function



	Public Function ChangePassword_Impersonate(
	 ByRef Result As String,
	 ByVal CurrentPW As String,
	 ByVal NewPW1 As String,
	 ByVal NewPW2 As String
	 ) As String

		Dim strResultFK As String = ""

		'declare @Result      varchar(1000),@ResultFK int

		'exec bi..usp_BI_SSRSSec 
		'    @Result output,
		'    @ResultFK output ,
		'    @CurrentPW,
		'        'Usr_Change',
		'    @NewPW1,
		'    @NewPW2

		'select @Result as Result,@ResultFK as ResultFK

		Dim oRet As New Object
		Dim sRet As String = ""
		Dim sSQL As String = ""

		sSQL = "" &
				"" &
				"" &
				"Execute As Login = 'evolutionjobs\" & WindowsUsername & "'; " &
				"" &
				" " &
				"EXEC	bi..[usp_BI_SSRSSec] " &
				"		@Result = @Result OUTPUT, " &
				"		@ResultFK = @ResultFK OUTPUT, " &
				"		@Param1 = '" & CurrentPW & "', " &
				"		@RequestType = 'Usr_Change', " &
				"		@Param2 = '" & NewPW1 & "', " &
				"		@Param3 = '" & NewPW2 & "', " &
				"		@Param4 = '" & DBNull.Value & "' " &
				" " &
				" " &
				" " &
				"Revert;" &
				"" &
				" " &
				"SELECT	@Result as N'@Result', " &
				"		@ResultFK as N'@ResultFK' " &
				" " &
				" " &
				""



		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.Text

		Try

			'@Result      varchar(100) output,
			'@ResultFK    int          output,
			'@Param1      varchar(100),
			'@RequestType varchar(100),
			'@Param2      varchar(100) = null,
			'@Param3      varchar(100) = null,
			'@Param4      varchar(100) = null
			oCmd.Parameters.Add("@Result", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
			oCmd.Parameters.Add("@ResultFK", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
			oCmd.Parameters.AddWithValue("@Param1", CurrentPW)
			oCmd.Parameters.AddWithValue("@RequestType", "Usr_Change")
			oCmd.Parameters.AddWithValue("@Param2", NewPW1)
			oCmd.Parameters.AddWithValue("@Param3", NewPW2)
			oCmd.Parameters.AddWithValue("@Param4", DBNull.Value)

			oRet = oCmd.ExecuteScalar()

			If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
				sRet = (oRet).ToString
			End If

			If Not IsDBNull(oCmd.Parameters("@Result").Value) Then
				Result = CStr(oCmd.Parameters("@Result").Value)
			End If
			If Not IsDBNull(oCmd.Parameters("@ResultFK").Value) Then
				strResultFK = CStr(oCmd.Parameters("@ResultFK").Value)
			End If
		Catch ex As Exception
			logError("ChangePassword(EVOWEB3)", "ERROR: change password:" & "; " & ex.Message, ex.StackTrace)

			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		Return strResultFK
	End Function




	Public Function RunVerification(
	 ByVal Please As String,
	 ByVal Login As String,
	 ByVal UserFK As Integer,
	 ByVal Phone As String,
	 ByVal Email As String
	 ) As DataTable

		CloseSQLConnection()

		Dim dtOutput As New DataTable

		If WindowsPassword.Trim <> "" Then

			Using lib1 As New clsDB(PasswordAdminEdConnectionString)
				' 1. get Stored info for comparison
				dtOutput = lib1.RunVerification_Impersonation(Please, Login, UserFK, Phone, Email)
			End Using

		Else

			Using lib1 As New clsDB(BIWindowsConnectionString)
				' 1. get Stored info for comparison
				dtOutput = lib1.RunVerification_Windows(Please, Login, UserFK, Phone, Email)
			End Using

		End If


		Return dtOutput
	End Function




	' original - windows authentication
	Public Function RunVerification_Windows(
	 ByVal Please As String,
	 ByVal Login As String,
	 ByVal UserFK As Integer,
	 ByVal Phone As String,
	 ByVal Email As String
	 ) As DataTable


		Dim dtOutput As New DataTable
		Dim sSQL As String = "bi..usp_BI_PW_PhoneEmailManagement"
		Dim oCmd As New SqlCommand(sSQL, moConn)
		oCmd.CommandType = CommandType.StoredProcedure

		Try
			'@Please   varchar(100) = 'Return details',
			'@Login as varchar(100) = null,
			'@UserFK   int          = null,
			'@Phone    varchar(100) = '',
			'@Email    varchar(100) = ''
			oCmd.Parameters.AddWithValue("@Please", Please)
			' If Login.Trim <> "" Then
			'	oCmd.Parameters.AddWithValue("@Login", UserFK)
			'Else
			'	oCmd.Parameters.AddWithValue("@Login", DBNull.Value)
			'End If
			If UserFK > 0 Then
				oCmd.Parameters.AddWithValue("@EmployeeFK", UserFK)
			Else
				oCmd.Parameters.AddWithValue("@EmployeeFK", DBNull.Value)
			End If

			' -------- IMPORTANT -------------
			' -------- IMPORTANT -------------
			' If you want the phone or email unmodified, pass through a NULL, so pass in an EMPTY STRING 
			' if you want to clear the phone or email, then pass in a "-1"
			' -------- IMPORTANT -------------
			' -------- IMPORTANT -------------

			If Phone.Trim <> "" Then
				If Phone.Trim = "-1" Then
					' CLEAR the value: pass through an empty string
					oCmd.Parameters.AddWithValue("@Phone", "")
				Else
					oCmd.Parameters.AddWithValue("@Phone", Phone)
				End If
			Else
				' NULL means do not change anything
				oCmd.Parameters.AddWithValue("@Phone", DBNull.Value)
			End If

			If Email.Trim <> "" Then
				If Email.Trim = "-1" Then
					' CLEAR the value: pass through an empty string
					oCmd.Parameters.AddWithValue("@Email", "")
				Else
					oCmd.Parameters.AddWithValue("@Email", Email)
				End If
			Else
				oCmd.Parameters.AddWithValue("@Email", DBNull.Value)
			End If

			'oCmd.Parameters.Add("@PasswordProtect", SqlDbType.Bit).Direction = ParameterDirection.Output
			'oCmd.Parameters.Add("@SSRS_ParamName", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output

			Using daLoader As New SqlDataAdapter(oCmd)
				daLoader.Fill(dtOutput)
			End Using


		Catch ex As Exception
			logError("RunVerification", "ERROR: Retrieving temporary code: Login=" & Login & "; UserFK=" & UserFK & "; Phone=" & Phone & "; Email=" & Email & "; " & ex.Message, ex.StackTrace)
			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		Return dtOutput
	End Function





	Public Function RunVerification_Impersonation(
	 ByVal Please As String,
	 ByVal Login As String,
	 ByVal UserFK As Integer,
	 ByVal Phone As String,
	 ByVal Email As String
	 ) As DataTable


		Dim dtOutput As New DataTable
		'Dim sSQL As String = "" '"bi..usp_BI_PW_PhoneEmailManagement"

		Dim sSQL As New System.Text.StringBuilder()
		sSQL.Append("  ")
		sSQL.Append("Execute As Login = 'evolutionjobs\" & WindowsUsername & "'; ")
		sSQL.Append("  ")
		sSQL.Append("exec bi..usp_BI_PW_PhoneEmailManagement   ")
		sSQL.Append("  ")
		sSQL.Append("@Please='" & Please & "' ")
		sSQL.Append(",  ")
		If UserFK > 0 Then
			sSQL.Append("@EmployeeFK=" & UserFK & "  ")
		Else
			sSQL.Append("@EmployeeFK=null  ")
		End If
		sSQL.Append(",  ")
		If Phone.Trim <> "" Then
			If Phone.Trim = "-1" Then
				' CLEAR the value: pass through an empty string
				sSQL.Append("@Phone='' ")
			Else
				sSQL.Append("@Phone='" & Phone & "'  ")
			End If
		Else
			' NULL means do not change anything
			sSQL.Append("@Phone=null  ")
		End If
		sSQL.Append(",  ")
		If Email.Trim <> "" Then
			If Email.Trim = "-1" Then
				' CLEAR the value: pass through an empty string
				sSQL.Append("@Email=''  ")
			Else
				sSQL.Append("@Email='" & Email & "'  ")
			End If
		Else
			sSQL.Append("@Email=null  ")
		End If
		sSQL.Append(" ; ")
		sSQL.Append("  ")
		sSQL.Append("Revert;  ")
		sSQL.Append("  ")
		sSQL.Append("  ")

		'sSQL = "" &
		'		"" &
		'		"" &
		'		"" &
		'		"Execute As Login = 'evolutionjobs\" & WindowsUsername & "'; " &
		'		"" &
		'		"exec bi..usp_BI_PW_PhoneEmailManagement @Please=@Please, @EmployeeFK=@EmployeeFK,@Phone=@Phone,@Email=@Email ; " &
		'		"" &
		'		"" &
		'		"Revert;" &
		'		""

		SaveToLog(WindowsUsername & "; clsDB - RunVerification: sSQL=" & sSQL.ToString)

		Dim oCmd As New SqlCommand(sSQL.ToString, moConn)
		oCmd.CommandType = CommandType.Text
		'oCmd.CommandType = CommandType.StoredProcedure

		Try
			'@Please   varchar(100) = 'Return details',
			'@Login as varchar(100) = null,
			'@UserFK   int          = null,
			'@Phone    varchar(100) = '',
			'@Email    varchar(100) = ''
			oCmd.Parameters.AddWithValue("@Please", Please)
			' If Login.Trim <> "" Then
			'	oCmd.Parameters.AddWithValue("@Login", UserFK)
			'Else
			'	oCmd.Parameters.AddWithValue("@Login", DBNull.Value)
			'End If
			If UserFK > 0 Then
				oCmd.Parameters.AddWithValue("@EmployeeFK", UserFK)
			Else
				oCmd.Parameters.AddWithValue("@EmployeeFK", DBNull.Value)
			End If

			' -------- IMPORTANT -------------
			' -------- IMPORTANT -------------
			' If you want the phone or email unmodified, pass through a NULL, so pass in an EMPTY STRING 
			' if you want to clear the phone or email, then pass in a "-1"
			' -------- IMPORTANT -------------
			' -------- IMPORTANT -------------

			If Phone.Trim <> "" Then
				If Phone.Trim = "-1" Then
					' CLEAR the value: pass through an empty string
					oCmd.Parameters.AddWithValue("@Phone", "")
				Else
					oCmd.Parameters.AddWithValue("@Phone", Phone)
				End If
			Else
				' NULL means do not change anything
				oCmd.Parameters.AddWithValue("@Phone", DBNull.Value)
			End If

			If Email.Trim <> "" Then
				If Email.Trim = "-1" Then
					' CLEAR the value: pass through an empty string
					oCmd.Parameters.AddWithValue("@Email", "")
				Else
					oCmd.Parameters.AddWithValue("@Email", Email)
				End If
			Else
				oCmd.Parameters.AddWithValue("@Email", DBNull.Value)
			End If

			'oCmd.Parameters.Add("@PasswordProtect", SqlDbType.Bit).Direction = ParameterDirection.Output
			'oCmd.Parameters.Add("@SSRS_ParamName", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output

			Using daLoader As New SqlDataAdapter(oCmd)
				daLoader.Fill(dtOutput)
			End Using


		Catch ex As Exception
			SaveToLog(WindowsUsername & "; clsDB - RunVerification(EVOWEB3): ***EXCEPTION***: usp_BI_PW_PhoneEmailManagement: Login=" & Login & "; UserFK=" & UserFK & "; Phone=" & Phone & "; Email=" & Email & "; " & ex.Message & ex.StackTrace)
			logError("RunVerification(EVOWEB3)", "Error: Retrieving temporary code: Login=" & Login & "; UserFK=" & UserFK & "; Phone=" & Phone & "; Email=" & Email & "; " & ex.Message, ex.StackTrace)
			Throw ex
		Finally
			oCmd.Dispose()
		End Try

		SaveToLog(WindowsUsername & "; clsDB - RunVerification: END")

		Return dtOutput
	End Function



	Public Function GetVerificationCodes(
	 ByVal windowslogin As String,
	 ByRef outPhoneCode As String,
	 ByRef outEmailCode As String
	 ) As String

		Dim sRetError As String = ""
		Dim sUserLogin As String = ""
		'exec usp_BI_PW_PhoneEmailManagement 'Show Verify Status',
		''evolutionjobs\humphreys'


		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE

		Try

			SaveToLog(WindowsUsername & "; clsDB - GetVerificationCodes-A: windowslogin=" & windowslogin)

			Dim dtOutput As New DataTable
			dtOutput = RunVerification(
			"Return Details",
			"",
			0,
			"",
			""
			)

			SaveToLog(WindowsUsername & "; clsDB - GetVerificationCodes-B: windowslogin=" & windowslogin)

			'Employee        userFK	Phone_Number	Email_Address
			'Stuart Humphrey	507	    FC58284	        Verified


			If dtOutput IsNot Nothing Then
				If dtOutput.Rows.Count > 0 Then

					SaveToLog(WindowsUsername & "; clsDB - GetVerificationCodes-C: dtOutput.Rows.Count=" & dtOutput.Rows.Count)

					For Each row As DataRow In dtOutput.Rows
						If Not IsDBNull(row("User_Login")) Then
							sUserLogin = CStr(row("User_Login"))
						End If
						If Not IsDBNull(row("Phone_VeriCode")) Then
							outPhoneCode = CStr(row("Phone_VeriCode")).Replace("Verified", "").Replace("n/a", "")
						End If
						If Not IsDBNull(row("Email_VeriCode")) Then
							outEmailCode = CStr(row("Email_VeriCode")).Replace("Verified", "").Replace("n/a", "")
						End If
					Next
				End If
			End If




		Catch ex As Exception
			sRetError = "ERROR: getting phone and email verification codes: " & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - GetVerificationCodes: ***EXCEPTION***: getting phone and email verification codes; windowslogin=)")

			If windowslogin.Trim <> "" Then
				logError("GetVerificationCodes(EVOWEB3)", "ERROR: getting phone and email verification codes; windowslogin=" & windowslogin & ": " & ex.Message, ex.StackTrace)

			End If

		End Try

		SaveToLog(WindowsUsername & "; clsDB - GetVerificationCodes-D: sUserLogin=" & sUserLogin & "; outPhoneCode=" & outPhoneCode & "; outEmailCode=" & outEmailCode)

		Return sRetError
	End Function



	Public Function GetPhoneOrEmail(
	 ByVal windowslogin As String,
	 ByRef outPhone As String,
	 ByRef outEmail As String
	 ) As String

		Dim sRetError As String = ""

		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE


		Try

			SaveToLog(WindowsUsername & "; clsDB - GetPhoneOrEmail-A: getting phone and email verification codes; windowslogin=" & windowslogin)

			Dim dtOutput As New DataTable
			dtOutput = RunVerification(
			"Return Details",
			"",
			0,
			"",
			""
			)

			SaveToLog(WindowsUsername & "; clsDB - GetPhoneOrEmail-B: getting phone and email verification codes; windowslogin=" & windowslogin)

			'Employee        userFK	User_Login	                Phone_Number	Email_Address
			'Stuart Humphrey	507	    evolutionjobs\humphreys		+111             ccc

			If dtOutput IsNot Nothing Then
				If dtOutput.Rows.Count > 0 Then
					For Each row As DataRow In dtOutput.Rows
						If Not IsDBNull(row("Phone_Number")) Then
							outPhone = CStr(row("Phone_Number"))
						End If
						If Not IsDBNull(row("Email_Address")) Then
							outEmail = CStr(row("Email_Address"))
						End If
					Next
				End If
			End If

			If outPhone.Trim <> "" Then
				outPhone = outPhone.Replace("No phone registered", "").Replace("n/a", "")
			End If
			If outEmail.Trim <> "" Then
				outEmail = outEmail.Replace("No email registered", "").Replace("n/a", "")
			End If

		Catch ex As Exception
			sRetError = "ERROR: Getting phone and email information:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - GetPhoneOrEmail: ***EXCEPTION***: Getting phone and email information:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message)

			If windowslogin.Trim <> "" Then
				' ERROR: Getting phone and email information: windowslogin=EVOLUTIONJOBS\WoodT; Login failed for user 'NT AUTHORITY\ANONYMOUS LOGON'.
				logError("GetPhoneOrEmailX(EVOWEB3)", "ERROR: Getting phone and email information: WindowsUsername=" & WindowsUsername & "; windowslogin=" & windowslogin & "; " & ex.Message, ex.StackTrace)
			End If
		End Try

		Return sRetError
	End Function





	Public Function CheckVerificationCode(
	 ByVal windowslogin As String,
	 ByRef inPhoneCode As String,
	 ByRef inEmailCode As String
	 ) As String

		Dim sRetError As String = ""
		Dim sResultMessage As String = ""
		Try

			Dim sPLease As String = "Enter Verification Code"
			If inPhoneCode.Trim <> "" Then
				inEmailCode = ""
			Else
				inPhoneCode = ""
			End If

			' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
			' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
			' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
			' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
			' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE

			SaveToLog(WindowsUsername & "; clsDB - CheckVerificationCode-A: getting phone and email verification codes; windowslogin=" & windowslogin)

			Dim dtOutput As New DataTable
			dtOutput = RunVerification(
			sPLease,
			"",
			0,
			inPhoneCode,
			inEmailCode
			)

			SaveToLog(WindowsUsername & "; clsDB - CheckVerificationCode-B: getting phone and email verification codes; windowslogin=" & windowslogin)

			'Return outEmailMessage IF SUCCESSFUL = Phone Verified
			'Return outEmailMessage IF SUCCESSFUL = Phone Verified
			'Return outEmailMessage IF SUCCESSFUL = Phone Verified

			If dtOutput IsNot Nothing Then
				If dtOutput.Rows.Count > 0 Then
					For Each row As DataRow In dtOutput.Rows
						If Not IsDBNull(row("ResultMessage")) Then
							sResultMessage = CStr(row("ResultMessage"))
						End If
					Next
					'For Each row As DataRow In dtOutput.Rows
					'	If Not IsDBNull(row("Phone_Number")) Then
					'		outPhoneMessage = CStr(row("Phone_Number"))
					'	End If
					'	If Not IsDBNull(row("Email_Address")) Then
					'		outEmailMessage = CStr(row("Email_Address"))
					'	End If
					'Next
				End If
			End If
		Catch ex As Exception
			sRetError = "ERROR: Getting phone and email information: " & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - CheckVerificationCode: ***EXCEPTION***: Getting phone and email information:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message)


			If windowslogin.Trim <> "" Then
				logError("CheckVerificationCodeX(EVOWEB3)", "ERROR: Checking verification code: windowslogin=" & windowslogin & "; " & ex.Message, ex.StackTrace)
			End If
		End Try

		If sResultMessage.Contains("Verified") Then
			' OK
			sRetError = ""
		Else
			If inPhoneCode.Trim <> "" Then
				sRetError = "ERROR: Phone number was not verified"
			Else
				sRetError = "ERROR: Email was not verified"
			End If
		End If

		Return sRetError
	End Function




	Public Function UpdateVerificationCode(
	 ByVal windowslogin As String,
	 ByVal InPhoneCode As String,
	 ByVal InEmailCode As String
	 ) As String

		Dim sRetError As String = ""

		'exec usp_BI_PW_PhoneEmailManagement 'Show Verify Status',
		''evolutionjobs\humphreys'


		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE

		Try
			Dim sGetError As String = ""
			Dim sOutPhoneCode As String = ""
			Dim sOutEmailCode As String = ""

			If InPhoneCode.Trim <> "" Then

				SaveToLog(WindowsUsername & "; clsDB - UpdateVerificationCode-A: getting phone and email verification codes; windowslogin=" & windowslogin)

				Dim dtOutput As New DataTable
				dtOutput = RunVerification(
				"Update Verification Code",
				"",
				0,
				InPhoneCode,
				""
				)

				SaveToLog(WindowsUsername & "; clsDB - UpdateVerificationCode-B: getting phone and email verification codes; windowslogin=" & windowslogin)


				If WindowsPassword.Trim <> "" Then
					Using lib1 As New clsDB(PasswordAdminEdConnectionString)
						' 1. get Stored info for comparison
						sGetError = lib1.GetVerificationCodes("", sOutPhoneCode, sOutEmailCode)
					End Using
				Else
					Using lib1 As New clsDB(BIWindowsConnectionString)
						' 1. get Stored info for comparison
						sGetError = lib1.GetVerificationCodes("", sOutPhoneCode, sOutEmailCode)
					End Using
				End If


				If sOutPhoneCode.Trim <> "" Then
					If sOutPhoneCode.Trim <> InPhoneCode Then
						sRetError = "ERROR: Saving verification code"
					End If
				End If

			Else

				Dim dtOutput As New DataTable
				dtOutput = RunVerification(
				"Update Verification Code",
				"",
				0,
				"",
				InEmailCode
				)

				If WindowsPassword.Trim <> "" Then
					Using lib1 As New clsDB(PasswordAdminEdConnectionString)
						' 1. get Stored info for comparison
						sGetError = lib1.GetVerificationCodes("", sOutPhoneCode, sOutEmailCode)
					End Using
				Else
					Using lib1 As New clsDB(BIWindowsConnectionString)
						' 1. get Stored info for comparison
						sGetError = lib1.GetVerificationCodes("", sOutPhoneCode, sOutEmailCode)
					End Using
				End If



				If sOutEmailCode.Trim <> "" Then
					If sOutEmailCode.Trim <> InEmailCode Then
						sRetError = "ERROR: Saving verification code"
					End If
				End If

			End If


		Catch ex As Exception
			sRetError = "ERROR: getting phone and email verification codes: " & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - UpdateVerificationCode: ***EXCEPTION***: Getting phone and email information:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message)


			If windowslogin.Trim <> "" Then
				logError("UpdateVerificationCode(EVOWEB3)", "ERROR: update verification code: windowslogin=" & windowslogin & "; " & ex.Message, ex.StackTrace)
			End If

		End Try

		Return sRetError
	End Function




	Public Function SaveUserPhone(
	 ByVal windowslogin As String,
	 ByVal inPhone As String
	 ) As String

		Dim sRetError As String = ""


		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE

		' -------- IMPORTANT -------------
		' -------- IMPORTANT -------------
		' If you want the phone or email unmodified, pass through a NULL, so pass in an EMPTY STRING 
		' if you want to clear the phone or email, then pass in a "-1"
		' -------- IMPORTANT -------------
		' -------- IMPORTANT -------------

		Try
			Dim dtOutput As New DataTable
			dtOutput = RunVerification(
			"Update user details",
			"",
			0,
			inPhone,
			""
			)

			SaveToLog(WindowsUsername & "; clsDB - SaveUserPhone:")

		Catch ex As Exception
			sRetError = "ERROR: Saving phone information: " & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - SaveUserPhone: ***EXCEPTION***: Saving phone information:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message)

			If windowslogin.Trim <> "" Then
				logError("SaveUserPhone(EVOWEB3)", "ERROR: save user phone: windowslogin=" & windowslogin & "; " & ex.Message, ex.StackTrace)
			End If

		End Try

		Return sRetError
	End Function



	Public Function SaveUserEmail(
	 ByVal windowslogin As String,
	 ByVal inEmail As String
	 ) As String

		Dim sRetError As String = ""

		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE

		' -------- IMPORTANT -------------
		' -------- IMPORTANT -------------
		' If you want the phone or email unmodified, pass through a NULL, so pass in an EMPTY STRING 
		' if you want to clear the phone or email, then pass in a "-1"
		' -------- IMPORTANT -------------
		' -------- IMPORTANT -------------

		Try
			Dim dtOutput As New DataTable
			dtOutput = RunVerification(
			"Update user details",
			"",
			0,
			"",
			inEmail
			)


		Catch ex As Exception
			sRetError = "ERROR: Saving email information: " & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - SaveUserEmail: ***EXCEPTION***: Saving email information:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message)

			If windowslogin.Trim <> "" Then
				logError("SaveUserEmail(EVOWEB3)", "ERROR: save user email: windowslogin=" & windowslogin & "; " & ex.Message, ex.StackTrace)
			End If

		End Try

		Return sRetError
	End Function







	Public Function GetNewPassword(
	 ByVal windowslogin As String,
	 ByRef outPassword As String
	 ) As String

		Dim sRetError As String = ""
		outPassword = ""

		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE
		' DO NOT PASS IN WINDOWS LOGIN NAME BECAUSE IT WILL USE IT WHEN CONNECTING TO DATABASE

		Try
			Dim dtOutput As New DataTable
			dtOutput = RunVerification(
			"Get New Temporary Password",
			"",
			0,
			"",
			""
			)

			If dtOutput IsNot Nothing Then
				If dtOutput.Rows.Count > 0 Then
					For Each row As DataRow In dtOutput.Rows
						' must be 4 characters
						If Not IsDBNull(row("TempPW")) Then
							outPassword = CStr(row("TempPW"))
						End If
					Next
				End If
			End If

			If outPassword.Trim <> "" Then
				If outPassword.Length > 4 Then
					sRetError = outPassword
				End If
			Else
				sRetError = "ERROR: Retrieving temporary code"

				If windowslogin.Trim <> "" Then
					logError("GetNewPassword1(EVOWEB3)", "ERROR: Retrieving temporary code: windowslogin=" & windowslogin & "; ", Nothing)

				End If


			End If


		Catch ex As Exception
			sRetError = "ERROR: Retrieving temporary code: " & ex.Message

			SaveToLog(WindowsUsername & "; clsDB - GetNewPassword: ***EXCEPTION***: Retrieving temporary code:WindowsUsername=" & WindowsUsername & "; ex=" & ex.Message)

			If windowslogin.Trim <> "" Then
				logError("GetNewPassword2(EVOWEB3)", "ERROR: Retrieving temporary code: windowslogin=" & windowslogin & "; " & ex.Message, ex.StackTrace)

			End If


		End Try

		Return sRetError
	End Function





	Public Function CheckIfDomainIsUsedInEmail(ByVal InEmail As String) As Long
		Dim lDomainID As Long = 0
		Dim sReturn As String = ""
		Dim sEmailDomain As String = ""
		Dim sSplit() As String = InEmail.Split("@")
		If sSplit IsNot Nothing Then
			If sSplit.Length > 0 Then
				sEmailDomain = sSplit(1)
			End If
		End If
		If sEmailDomain.Trim <> "" Then
			sReturn = RunSQLScalar("select ed.[dbo].[udf_MyReports_CheckIfDomainIsUsedInEmail]('" & sEmailDomain & "')")

			If sReturn.Trim <> "" Then
				Long.TryParse(sReturn.Trim, lDomainID)
			End If

		End If

		Return lDomainID
	End Function




	'Public Function GetNewPasswordForAUser(
	' ByVal sAdminPassword As String,
	' ByVal EmployeeFK As Integer
	' ) As DataTable

	'    Dim dtOutput As New DataTable

	'    Dim sSQL As String = "bi..usp_BI_PasswordAdmin_Populate"

	'    'First you will need to add Application Name=PasswordAdmin to your connection string.
	'    'The @CurrentPW needs to be your SSRS password And @employeeFK Is the user who needs a temp password.
	'    'Please get back to me if there are any issues.

	'    'exec bi..usp_BI_PasswordAdmin_Populate 'ResetPassword',@CurrentPW,null,null,@EmployeeFK

	'    '@SelectType varchar(100) = '',
	'    '@CurrentPW  varchar(100) = null,
	'    '@NewPw1     varchar(100) = null,
	'    '@NewPw2     varchar(100) = null,
	'    '@EmployeeFK int          = null

	'    Dim oCmd As New SqlCommand(sSQL, moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    Try
	'        oCmd.Parameters.AddWithValue("@SelectType", "ResetPassword")
	'        oCmd.Parameters.AddWithValue("@CurrentPW", sAdminPassword)
	'        oCmd.Parameters.AddWithValue("@NewPw1", DBNull.Value)
	'        oCmd.Parameters.AddWithValue("@NewPw2", DBNull.Value)
	'        oCmd.Parameters.AddWithValue("@EmployeeFK", EmployeeFK)

	'        Using daLoader As New SqlDataAdapter(oCmd)
	'            daLoader.Fill(dtOutput)
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try

	'    Return dtOutput
	'End Function

	' =============================================================================
	' =============================================================================
	' =============================================================================
	' =============================================================================





	'Public Function GetCompaniesAll(
	' ByVal CompanyFK As Long,
	' ByVal UserFK As Long
	' ) As DataTable

	'    ' unlike GetCompanies, this gets all companies

	'    Dim dtOutput As New DataTable

	'    Dim sSQL As String = "dbo.usp_SP_Intranet_GetCompaniesAll"
	'    Dim oCmd As New SqlCommand(sSQL, moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    Try
	'        oCmd.Parameters.Add("@UserFK", SqlDbType.Int).Value = UserFK
	'        oCmd.Parameters.Add("@CompanyFK", SqlDbType.Int).Value = CompanyFK

	'        Using daLoader As New SqlDataAdapter(oCmd)
	'            daLoader.Fill(dtOutput)
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try

	'    Return dtOutput
	'End Function



	'Public Function GetUserEmailAddress(
	' ByVal UserFK As Long,
	' ByRef sOutFirstName As String
	' ) As String
	'    Dim sEmail As String = ""
	'    sOutFirstName = ""

	'    Dim dtOutput As New DataTable

	'    Dim sSQL As String = "dbo.usp_NS_GetUserEmailAddress"
	'    Dim oCmd As New SqlCommand(sSQL, moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    Try
	'        oCmd.Parameters.Add("@SignerUserFK", SqlDbType.Int).Value = UserFK

	'        Using daLoader As New SqlDataAdapter(oCmd)
	'            daLoader.Fill(dtOutput)
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try


	'    If dtOutput IsNot Nothing Then
	'        If dtOutput.Rows.Count > 0 Then
	'            For Each rowX As DataRow In dtOutput.Rows
	'                If Not IsDBNull(rowX("EmailAddress")) Then
	'                    sEmail = CStr(rowX("EmailAddress"))
	'                End If
	'                If Not IsDBNull(rowX("firstname")) Then
	'                    sOutFirstName = CStr(rowX("firstname"))
	'                End If
	'                Exit For
	'            Next
	'        End If
	'    End If

	'    Return sEmail
	'End Function


	'Public Function InsertUpdateNewStarterData(
	'ByVal NSNewStarterID As Integer,
	'ByVal CompanyFK As Integer,
	'ByVal NewStarterUserFK As Long,
	'ByVal NewStarterName As String,
	'ByVal DateInformedIT As String,
	'ByVal StartDate As String,
	'ByVal IRContactUserFK As Long,
	'ByVal LineManagerUserFK As Long,
	'ByVal DatabaseFK As Long,
	'ByVal TeamFK As Long,
	'ByVal CarReg As String,
	'ByVal JobTitle As String,
	'ByVal JobTitleForTV As String,
	'ByVal FOBArranged As Integer,
	'ByVal ReferencesToHR As Integer,
	'ByVal HolidayBooked As Integer,
	'ByVal HolidayStartDate As String,
	'ByVal HolidayEndDate As String,
	'ByVal KnownAs As String,
	'ByVal EmailjobTitle As String,
	'ByVal CopyDataFrom As String,
	'ByVal RoleTypeCode As Integer,
	'ByVal NeedDeskMoves As Integer,
	'ByVal ReportingToThem As Integer,
	'ByVal NewVerticalMarket As String,
	'ByVal CutomisedRegionPatch As Boolean,
	'ByVal Notes As String,
	'ByRef outNSNewStarterID As Long
	') As Long

	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_InsertUpdateNewStarterData", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure
	'    oCmd.Parameters.AddWithValue("@NSNewStarterID", NSNewStarterID)
	'    oCmd.Parameters.AddWithValue("@CompanyFK", CompanyFK)
	'    oCmd.Parameters.AddWithValue("@NewStarterUserFK", NewStarterUserFK)
	'    oCmd.Parameters.AddWithValue("@NewStarterName", NewStarterName)
	'    If DateInformedIT.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@DateInformedIT", CDate(DateInformedIT))
	'    Else
	'        oCmd.Parameters.AddWithValue("@DateInformedIT", DBNull.Value)
	'    End If
	'    If StartDate.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@StartDate", CDate(StartDate))
	'    Else
	'        oCmd.Parameters.AddWithValue("@StartDate", DBNull.Value)
	'    End If
	'    oCmd.Parameters.AddWithValue("@IRContactUserFK", IRContactUserFK)
	'    oCmd.Parameters.AddWithValue("@LineManagerUserFK", LineManagerUserFK)
	'    oCmd.Parameters.AddWithValue("@DatabaseFK", DatabaseFK)
	'    oCmd.Parameters.AddWithValue("@TeamFK", TeamFK)
	'    oCmd.Parameters.AddWithValue("@CarReg", CarReg)
	'    oCmd.Parameters.AddWithValue("@JobTitle", JobTitle)
	'    oCmd.Parameters.AddWithValue("@JobTitleForTV", JobTitleForTV)
	'    oCmd.Parameters.AddWithValue("@FOBArranged", FOBArranged)
	'    oCmd.Parameters.AddWithValue("@ReferencesToHR", ReferencesToHR)
	'    oCmd.Parameters.AddWithValue("@HolidayBooked", HolidayBooked)
	'    If HolidayStartDate.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@HolidayStartDate", CDate(HolidayStartDate))
	'    Else
	'        oCmd.Parameters.AddWithValue("@HolidayStartDate", DBNull.Value)
	'    End If
	'    If HolidayEndDate.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@HolidayEndDate", CDate(HolidayEndDate))
	'    Else
	'        oCmd.Parameters.AddWithValue("@HolidayEndDate", DBNull.Value)
	'    End If
	'    oCmd.Parameters.AddWithValue("@KnownAs", KnownAs)
	'    oCmd.Parameters.AddWithValue("@EmailjobTitle", EmailjobTitle)
	'    oCmd.Parameters.AddWithValue("@CopyDataFrom", CopyDataFrom)
	'    oCmd.Parameters.AddWithValue("@RoleTypeCode", RoleTypeCode)
	'    oCmd.Parameters.AddWithValue("@NeedDeskMoves", NeedDeskMoves)
	'    oCmd.Parameters.AddWithValue("@ReportingToThem", ReportingToThem)
	'    oCmd.Parameters.AddWithValue("@NewVerticalMarket", NewVerticalMarket)
	'    oCmd.Parameters.AddWithValue("@CutomisedRegionPatch", IIf(CutomisedRegionPatch = True, 1, 0))
	'    oCmd.Parameters.AddWithValue("@Notes", Notes)
	'    oCmd.Parameters.AddWithValue("@DateUpdated", Date.Now)

	'    oCmd.Parameters.Add("@outNSNewStarterID", SqlDbType.Int).Direction = ParameterDirection.Output

	'    Try

	'        oRet = oCmd.ExecuteScalar()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@outNSNewStarterID").Value) Then
	'            outNSNewStarterID = CLng(oCmd.Parameters("@outNSNewStarterID").Value)
	'        End If

	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return lRet
	'End Function




	'Public Function AddTVPingHistory(ByVal UserFK As Long, ByVal HistoryActionFK As Long) As Long
	'    Dim ReturnIDOut As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("[dbo].[usp_NS_AddHistory]", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@UserFK", UserFK)
	'    oCmd.Parameters.AddWithValue("@HistoryActionText", "New Starter Form TV Ping")
	'    oCmd.Parameters.AddWithValue("@HistoryActionFK", HistoryActionFK)

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return ReturnIDOut
	'End Function


	'Public Function AddSignerToList(
	'ByVal NSSignOffGroupFK As Long,
	'ByVal CompanyFK As Long,
	'ByVal SignerUserFK As Long
	') As Long
	'    Dim ReturnIDOut As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_AddSignerToList", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@NSSignOffGroupFK", NSSignOffGroupFK)
	'    oCmd.Parameters.AddWithValue("@CompanyFK", CompanyFK)
	'    oCmd.Parameters.AddWithValue("@SignerUserFK ", SignerUserFK)

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return ReturnIDOut
	'End Function




	'Public Function AddSignOff(
	'    ByVal NSSignOffGroupFK As Long,
	'    ByVal CompanyFK As Long,
	'    ByVal SignedByUserFK As Long,
	'    ByVal NSNewStarterFK As Long
	'    ) As Long
	'    Dim ReturnIDOut As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("[dbo].[usp_NS_AddSignOff]", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@NSSignOffGroupFK", NSSignOffGroupFK)
	'    oCmd.Parameters.AddWithValue("@CompanyFK", CompanyFK)
	'    oCmd.Parameters.AddWithValue("@SignedByUserFK ", SignedByUserFK)
	'    oCmd.Parameters.AddWithValue("@NSNewStarterFK ", NSNewStarterFK)

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return ReturnIDOut
	'End Function



	'Public Function Get_AllOrIR_Signed(
	'    ByVal companyfk As Long,
	'    ByVal NSNewStarterFK As Long,
	'    ByRef outIsAllSigned As Boolean
	'    ) As Boolean

	'    Dim outIsHRSigned As Boolean = False
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_Get_AllOrIR_Signed", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@companyfk", companyfk)
	'    oCmd.Parameters.AddWithValue("@NSNewStarterFK", NSNewStarterFK)

	'    oCmd.Parameters.Add("@outIsAllSigned", SqlDbType.Int).Direction = ParameterDirection.Output
	'    oCmd.Parameters.Add("@outIsHRSigned", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@outIsAllSigned").Value) Then
	'            outIsAllSigned = CBool(oCmd.Parameters("@outIsAllSigned").Value)
	'        End If
	'        If Not IsDBNull(oCmd.Parameters("@outIsHRSigned").Value) Then
	'            outIsHRSigned = CBool(oCmd.Parameters("@outIsHRSigned").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return outIsHRSigned
	'End Function


	'  Public Function UpdateHistoryFK(
	'   ByVal NSNewStarterID As Long,
	'   ByVal HistoryFK As Long
	'   ) As Long
	'      Dim outHistoryFK As Long = -1
	'      Dim oRet As New Object
	'      Dim lRet As Long = -1

	'      Dim oCmd As New SqlCommand("dbo.usp_NS_UpdateHistoryFK", moConn)
	'      oCmd.CommandType = CommandType.StoredProcedure

	'      oCmd.Parameters.AddWithValue("@NSNewStarterID", NSNewStarterID)
	'      oCmd.Parameters.AddWithValue("@HistoryFK", HistoryFK)

	'      oCmd.Parameters.Add("@outHistoryFK", SqlDbType.Int).Direction = ParameterDirection.Output
	'      Try
	'          oRet = oCmd.ExecuteNonQuery()

	'          If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'              lRet = CLng(oRet)
	'          End If

	'          If Not IsDBNull(oCmd.Parameters("@outHistoryFK").Value) Then
	'              outHistoryFK = CLng(oCmd.Parameters("@outHistoryFK").Value)
	'          End If
	'      Catch ex As Exception
	'          Throw ex
	'      Finally
	'          oCmd.Dispose()
	'      End Try
	'      Return outHistoryFK
	'  End Function

	'  Public Function GetNSSignOffGroupIDByWorkflowOrder(ByVal NSSignOffGroupCode As String) As Integer
	'      Dim NSSignOffGroupID As Integer = RunSQLScalar("SELECT NSSignOffGroupID FROM dbo._vtblNSSignOffGroup WHERE NSSignOffGroupCode = '" & NSSignOffGroupCode & "' ")
	'      Return NSSignOffGroupID
	'  End Function

	'  Public Function IsAdminUser(ByVal inUserID As Integer) As Boolean
	'      Dim bIsAdmin As Boolean = False
	'      Dim oRet As Object = Nothing
	'      oRet = RunSQLScalar("SELECT [dbo].[udf_NS_IsAdminUser] (" & inUserID & ") ")
	'      If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'          bIsAdmin = CBool(oRet)
	'      End If
	'      Return bIsAdmin
	'  End Function


	'  Public Function SSRSCheckPassword(
	'ByVal sPassword As String,
	'ByVal inApplicationFK As Integer
	') As DataTable

	'      ' unlike GetCompanies, this gets all companies

	'      'Dim impersonationContext As System.Security.Principal.WindowsImpersonationContext
	'      'Dim currentWindowsIdentity As System.Security.Principal.WindowsIdentity

	'      'currentWindowsIdentity = CType(HttpContext.Current.User.Identity, System.Security.Principal.WindowsIdentity)
	'      'impersonationContext = currentWindowsIdentity.Impersonate()

	'      Dim dtOutput As New DataTable

	'      Dim sSQL As String = "bi..usp_BI_PWPort_Manager"
	'      'Dim sSQL As String = "bi..usp_BI_PWPort_Manager @GiveItAGo,@OnwardApplicationFK,@Password"
	'      Dim oCmd As New SqlCommand(sSQL, moConn)
	'      oCmd.CommandType = CommandType.StoredProcedure

	'      Try
	'          oCmd.Parameters.AddWithValue("@ActionType", "GiveItAGo")
	'          oCmd.Parameters.AddWithValue("@Param1", inApplicationFK)
	'          oCmd.Parameters.AddWithValue("@Param2", sPassword)

	'          Using daLoader As New SqlDataAdapter(oCmd)
	'              daLoader.Fill(dtOutput)
	'          End Using
	'      Catch ex As Exception
	'          Throw ex
	'      Finally
	'          oCmd.Dispose()
	'      End Try

	'      Return dtOutput
	'  End Function












	'Public Function GetUserEmailAddress( _
	' ByVal UserFK As Long, _
	' ByRef sOutFirstName As String _
	' ) As String
	'    Dim sEmail As String = ""
	'    sOutFirstName = ""

	'    Dim dtOutput As New DataTable

	'    Dim sSQL As String = "dbo.usp_NS_GetUserEmailAddress"
	'    Dim oCmd As New SqlCommand(sSQL, moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    Try
	'        oCmd.Parameters.Add("@SignerUserFK", SqlDbType.Int).Value = UserFK

	'        Using daLoader As New SqlDataAdapter(oCmd)
	'            daLoader.Fill(dtOutput)
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try


	'    If dtOutput IsNot Nothing Then
	'        If dtOutput.Rows.Count > 0 Then
	'            For Each rowX As DataRow In dtOutput.Rows
	'                If Not IsDBNull(rowX("EmailAddress")) Then
	'                    sEmail = CStr(rowX("EmailAddress"))
	'                End If
	'                If Not IsDBNull(rowX("firstname")) Then
	'                    sOutFirstName = CStr(rowX("firstname"))
	'                End If
	'                Exit For
	'            Next
	'        End If
	'    End If

	'    Return sEmail
	'End Function


	'Public Function InsertUpdateNewStarterData( _
	'ByVal NSNewStarterID As Integer, _
	'ByVal CompanyFK As Integer, _
	'ByVal NewStarterUserFK As Long, _
	'ByVal NewStarterName As String, _
	'ByVal DateInformedIT As String, _
	'ByVal StartDate As String, _
	'ByVal IRContactUserFK As Long, _
	'ByVal LineManagerUserFK As Long, _
	'ByVal DatabaseFK As Long, _
	'ByVal TeamFK As Long, _
	'ByVal CarReg As String, _
	'ByVal JobTitle As String, _
	'ByVal FOBArranged As Integer, _
	'ByVal ReferencesToHR As Integer, _
	'ByVal HolidayBooked As Integer, _
	'ByVal HolidayStartDate As String, _
	'ByVal HolidayEndDate As String, _
	'ByVal KnownAs As String, _
	'ByVal EmailjobTitle As String, _
	'ByVal CopyDataFrom As String, _
	'ByVal RoleTypeCode As Integer, _
	'ByVal NeedDeskMoves As Integer, _
	'ByVal ReportingToThem As Integer, _
	'ByVal NewVerticalMarket As String, _
	'ByVal CutomisedRegionPatch As Boolean, _
	'ByVal Notes As String, _
	'ByRef outNSNewStarterID As Long _
	') As Long

	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_InsertUpdateNewStarterData", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure
	'    oCmd.Parameters.AddWithValue("@NSNewStarterID", NSNewStarterID)
	'    oCmd.Parameters.AddWithValue("@CompanyFK", CompanyFK)
	'    oCmd.Parameters.AddWithValue("@NewStarterUserFK", NewStarterUserFK)
	'    oCmd.Parameters.AddWithValue("@NewStarterName", NewStarterName)
	'    If DateInformedIT.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@DateInformedIT", CDate(DateInformedIT))
	'    Else
	'        oCmd.Parameters.AddWithValue("@DateInformedIT", DBNull.Value)
	'    End If
	'    If StartDate.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@StartDate", CDate(StartDate))
	'    Else
	'        oCmd.Parameters.AddWithValue("@StartDate", DBNull.Value)
	'    End If
	'    oCmd.Parameters.AddWithValue("@IRContactUserFK", IRContactUserFK)
	'    oCmd.Parameters.AddWithValue("@LineManagerUserFK", LineManagerUserFK)
	'    oCmd.Parameters.AddWithValue("@DatabaseFK", DatabaseFK)
	'    oCmd.Parameters.AddWithValue("@TeamFK", TeamFK)
	'    oCmd.Parameters.AddWithValue("@CarReg", CarReg)
	'    oCmd.Parameters.AddWithValue("@JobTitle", JobTitle)
	'    oCmd.Parameters.AddWithValue("@FOBArranged", FOBArranged)
	'    oCmd.Parameters.AddWithValue("@ReferencesToHR", ReferencesToHR)
	'    oCmd.Parameters.AddWithValue("@HolidayBooked", HolidayBooked)
	'    If HolidayStartDate.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@HolidayStartDate", CDate(HolidayStartDate))
	'    Else
	'        oCmd.Parameters.AddWithValue("@HolidayStartDate", DBNull.Value)
	'    End If
	'    If HolidayEndDate.Trim <> "" Then
	'        oCmd.Parameters.AddWithValue("@HolidayEndDate", CDate(HolidayEndDate))
	'    Else
	'        oCmd.Parameters.AddWithValue("@HolidayEndDate", DBNull.Value)
	'    End If
	'    oCmd.Parameters.AddWithValue("@KnownAs", KnownAs)
	'    oCmd.Parameters.AddWithValue("@EmailjobTitle", EmailjobTitle)
	'    oCmd.Parameters.AddWithValue("@CopyDataFrom", CopyDataFrom)
	'    oCmd.Parameters.AddWithValue("@RoleTypeCode", RoleTypeCode)
	'    oCmd.Parameters.AddWithValue("@NeedDeskMoves", NeedDeskMoves)
	'    oCmd.Parameters.AddWithValue("@ReportingToThem", ReportingToThem)
	'    oCmd.Parameters.AddWithValue("@NewVerticalMarket", NewVerticalMarket)
	'    oCmd.Parameters.AddWithValue("@CutomisedRegionPatch", IIf(CutomisedRegionPatch = True, 1, 0))
	'    oCmd.Parameters.AddWithValue("@Notes", Notes)
	'    oCmd.Parameters.AddWithValue("@DateUpdated", Date.Now)

	'    oCmd.Parameters.Add("@outNSNewStarterID", SqlDbType.Int).Direction = ParameterDirection.Output

	'    Try

	'        oRet = oCmd.ExecuteScalar()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@outNSNewStarterID").Value) Then
	'            outNSNewStarterID = CLng(oCmd.Parameters("@outNSNewStarterID").Value)
	'        End If

	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return lRet
	'End Function




	'Public Function AddTVPingHistory(ByVal UserFK As Long, ByVal HistoryActionFK As Long) As Long
	'    Dim ReturnIDOut As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("[dbo].[usp_NS_AddHistory]", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@UserFK", UserFK)
	'    oCmd.Parameters.AddWithValue("@HistoryActionText", "New Starter Form TV Ping")
	'    oCmd.Parameters.AddWithValue("@HistoryActionFK", HistoryActionFK)

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return ReturnIDOut
	'End Function


	'Public Function AddSignerToList( _
	'ByVal NSSignOffGroupFK As Long, _
	'ByVal CompanyFK As Long, _
	'ByVal SignerUserFK As Long _
	') As Long
	'    Dim ReturnIDOut As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_AddSignerToList", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@NSSignOffGroupFK", NSSignOffGroupFK)
	'    oCmd.Parameters.AddWithValue("@CompanyFK", CompanyFK)
	'    oCmd.Parameters.AddWithValue("@SignerUserFK ", SignerUserFK)

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return ReturnIDOut
	'End Function




	'Public Function AddSignOff( _
	'    ByVal NSSignOffGroupFK As Long, _
	'    ByVal CompanyFK As Long, _
	'    ByVal SignedByUserFK As Long, _
	'    ByVal NSNewStarterFK As Long _
	'    ) As Long
	'    Dim ReturnIDOut As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("[dbo].[usp_NS_AddSignOff]", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@NSSignOffGroupFK", NSSignOffGroupFK)
	'    oCmd.Parameters.AddWithValue("@CompanyFK", CompanyFK)
	'    oCmd.Parameters.AddWithValue("@SignedByUserFK ", SignedByUserFK)
	'    oCmd.Parameters.AddWithValue("@NSNewStarterFK ", NSNewStarterFK)

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return ReturnIDOut
	'End Function



	'Public Function Get_AllOrIR_Signed( _
	'    ByVal companyfk As Long, _
	'    ByVal NSNewStarterFK As Long, _
	'    ByRef outIsAllSigned As Boolean _
	'    ) As Boolean

	'    Dim outIsHRSigned As Boolean = False
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_Get_AllOrIR_Signed", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@companyfk", companyfk)
	'    oCmd.Parameters.AddWithValue("@NSNewStarterFK", NSNewStarterFK)

	'    oCmd.Parameters.Add("@outIsAllSigned", SqlDbType.Int).Direction = ParameterDirection.Output
	'    oCmd.Parameters.Add("@outIsHRSigned", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@outIsAllSigned").Value) Then
	'            outIsAllSigned = CBool(oCmd.Parameters("@outIsAllSigned").Value)
	'        End If
	'        If Not IsDBNull(oCmd.Parameters("@outIsHRSigned").Value) Then
	'            outIsHRSigned = CBool(oCmd.Parameters("@outIsHRSigned").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return outIsHRSigned
	'End Function


	'Public Function UpdateHistoryFK( _
	' ByVal NSNewStarterID As Long, _
	' ByVal HistoryFK As Long _
	' ) As Long
	'    Dim outHistoryFK As Long = -1
	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_NS_UpdateHistoryFK", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure

	'    oCmd.Parameters.AddWithValue("@NSNewStarterID", NSNewStarterID)
	'    oCmd.Parameters.AddWithValue("@HistoryFK", HistoryFK)

	'    oCmd.Parameters.Add("@outHistoryFK", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteNonQuery()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@outHistoryFK").Value) Then
	'            outHistoryFK = CLng(oCmd.Parameters("@outHistoryFK").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return outHistoryFK
	'End Function

	'Public Function GetNSSignOffGroupIDByWorkflowOrder(ByVal NSSignOffGroupCode As String) As Integer
	'    Dim NSSignOffGroupID As Integer = RunSQLScalar("SELECT NSSignOffGroupID FROM dbo._vtblNSSignOffGroup WHERE NSSignOffGroupCode = '" & NSSignOffGroupCode & "' ")
	'    Return NSSignOffGroupID
	'End Function

	'Public Function IsAdminUser(ByVal inUserID As Integer) As Boolean
	'    Dim bIsAdmin As Boolean = False
	'    Dim oRet As Object = Nothing
	'    oRet = RunSQLScalar("SELECT [dbo].[udf_NS_IsAdminUser] (" & inUserID & ") ")
	'    If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'        bIsAdmin = CBool(oRet)
	'    End If
	'    Return bIsAdmin
	'End Function

	'Private Function GetConsultantEmail(ByVal inUserID As Long) As String
	'    Dim sEmail As String = ""
	'    Try
	'        Dim sSQL As String = ""
	'        Dim dt As New DataTable

	'        sSQL = "select top 1 * "
	'        sSQL &= "FROM tblUser INNER JOIN"
	'        sSQL &= " tblDomain ON tblUser.DomainFK = tblDomain.DomainID"
	'        sSQL &= " where UserID = " & inUserID
	'        dt = RunSQL(sSQL)

	'        If dt IsNot Nothing Then
	'            If dt.Rows.Count > 0 Then
	'                Dim iCount As Integer = 0
	'                For Each row As DataRow In dt.Rows
	'                    Dim Email As String = ""
	'                    Dim DomainName As String = ""
	'                    If Not IsDBNull(row("Email")) Then
	'                        Email = CStr(row("Email"))
	'                    End If
	'                    If Not IsDBNull(row("DomainName")) Then
	'                        DomainName = CStr(row("DomainName"))
	'                    End If
	'                    sEmail = Email & "@" & DomainName

	'                Next
	'            End If
	'        End If

	'        'Return dt.Rows(0)("Email") & "@" & dt.Rows(0)("DomainName")

	'    Catch ex As Exception
	'        Throw ex
	'    Finally

	'    End Try
	'    Return sEmail
	'End Function


	'Public Function GetPaymentTerms(ByVal inClient As Long) As DataTable
	'    Dim dtOutput As New DataTable
	'    Dim oCmd As New SqlCommand("dbo.usp_SP_AC_GetPaymentTerms", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure
	'    oCmd.Parameters.Add("@ClientID", SqlDbType.Int).Value = inClient
	'    Try
	'        Using daLoader As New SqlDataAdapter(oCmd)
	'            daLoader.Fill(dtOutput)
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try

	'    Return dtOutput
	'End Function



	'Public Function sssssssssssssss( _
	'ByVal InvoiceListItemID As Long, _
	'ByVal InvoiceHeaderFK As Long, _
	'ByVal Salary As Decimal, _
	'ByVal Total As Decimal, _
	'ByVal Notes As String, _
	'ByVal CandidateFK As Long, _
	'ByVal JobFK As Long, _
	'ByVal HasVAT As Integer, _
	'ByRef ReturnIDOut As Long _
	') As Long

	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_AC_Invoices_InsertUpdateDetailsRow", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure
	'    oCmd.Parameters.Add("@InvoiceListItemID", SqlDbType.Int).Value = InvoiceListItemID
	'    oCmd.Parameters.Add("@InvoiceHeaderFK", SqlDbType.Int).Value = InvoiceHeaderFK
	'    oCmd.Parameters.Add("@Salary", SqlDbType.Money).Value = Salary
	'    oCmd.Parameters.Add("@Total", SqlDbType.Money).Value = Total
	'    oCmd.Parameters.Add("@Notes", SqlDbType.VarChar, 8000).Value = Notes
	'    oCmd.Parameters.Add("@CandidateFK", SqlDbType.Int).Value = CandidateFK
	'    oCmd.Parameters.Add("@JobFK", SqlDbType.Int).Value = JobFK
	'    oCmd.Parameters.Add("@HasVAT", SqlDbType.Int).Value = HasVAT

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteScalar()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return lRet
	'End Function


	'Public Function ssssssssssssssss( _
	'ByVal InvoiceHeaderID As Long, _
	'ByVal InvoiceNumber As String, _
	'ByVal ClientFK As Long, _
	'ByVal UserFK As Long, _
	'ByVal OwnerFK As Long, _
	'ByVal Initials As String, _
	'ByVal ResourcerInitials As String, _
	'ByVal Total As Decimal, _
	'ByVal VAT As Decimal, _
	'ByVal GrandTotal As Decimal, _
	'ByVal TaxDate As Date, _
	'ByRef ReturnIDOut As Long _
	') As Long

	'    Dim oRet As New Object
	'    Dim lRet As Long = -1

	'    Dim oCmd As New SqlCommand("dbo.usp_SP_Invoices_InsertUpdateHeaderRow", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure
	'    oCmd.Parameters.Add("@InvoiceHeaderID", SqlDbType.Int).Value = InvoiceHeaderID
	'    oCmd.Parameters.Add("@InvoiceNumber", SqlDbType.VarChar, 20).Value = InvoiceNumber
	'    oCmd.Parameters.Add("@ClientFK", SqlDbType.Int).Value = ClientFK
	'    oCmd.Parameters.Add("@UserFK", SqlDbType.Int).Value = UserFK
	'    oCmd.Parameters.Add("@OwnerFK", SqlDbType.Int).Value = OwnerFK
	'    oCmd.Parameters.Add("@Initials", SqlDbType.VarChar, 5).Value = Initials
	'    oCmd.Parameters.Add("@ResourcerInitials", SqlDbType.VarChar, 5).Value = ResourcerInitials
	'    oCmd.Parameters.Add("@Total", SqlDbType.Money).Value = Total
	'    oCmd.Parameters.Add("@VAT", SqlDbType.Money).Value = VAT
	'    oCmd.Parameters.Add("@GrandTotal", SqlDbType.Money).Value = GrandTotal
	'    oCmd.Parameters.Add("@TaxDate", SqlDbType.DateTime).Value = TaxDate

	'    oCmd.Parameters.Add("@ReturnID", SqlDbType.Int).Direction = ParameterDirection.Output
	'    Try
	'        oRet = oCmd.ExecuteScalar()

	'        If oRet IsNot Nothing AndAlso Not IsDBNull(oRet) Then
	'            lRet = CLng(oRet)
	'        End If

	'        If Not IsDBNull(oCmd.Parameters("@ReturnID").Value) Then
	'            ReturnIDOut = CLng(oCmd.Parameters("@ReturnID").Value)
	'        End If
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try
	'    Return lRet
	'End Function



	'' ----- backup ----
	'Public Function ssssssssssssssssssss(ByVal Intranet As Integer) As DataTable
	'    Dim dtOutput As New DataTable
	'    Dim oCmd As New SqlCommand("dbo.usp_JF_AM_AccountsManagement_GetInvoiceListPaid_v5", moConn)
	'    oCmd.CommandType = CommandType.StoredProcedure
	'    oCmd.Parameters.Add("@Intranet", SqlDbType.Int).Value = Intranet
	'    oCmd.CommandTimeout = 360
	'    Try
	'        Using daLoader As New SqlDataAdapter(oCmd)
	'            daLoader.Fill(dtOutput)
	'        End Using
	'    Catch ex As Exception
	'        Throw ex
	'    Finally
	'        oCmd.Dispose()
	'    End Try

	'    Return dtOutput
	'End Function
























End Class
