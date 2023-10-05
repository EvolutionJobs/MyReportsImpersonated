Imports System.Data
Imports System.Web

Partial Class WagesLoginWindows
	Inherits System.Web.UI.Page


	Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

		' live site
		'http://EvoWeb3/LoginWindows.aspx?appfk=22


		' Test site
		'http://EvoWeb3:91/LoginWindows.aspx?appfk=22

		lblStatus.Text = ""

		Dim sWinUsername As String = HttpContext.Current.User.Identity.Name

		txtWindowsUsername.Text = sWinUsername


		SaveToLog("LOGIN_LOAD [" & sWinUsername & "]: sWinUsername=" & sWinUsername & "; WindowsDomain=" & WindowsDomain)

		If Me.IsPostBack = False Then



		End If

	End Sub

	Protected Sub cmdSubmit_Click(sender As Object, e As System.EventArgs) Handles cmdSubmit.Click
		DoLogin(False)
	End Sub

	Public Sub DoLogin(ByVal inFromPassword As Boolean)

		If txtWindowsPassword.Text.Trim = "" Then
			lblStatus.Text = "ERROR: Please enter a password"
			Exit Sub
		End If

		Dim sAppFK As String = ""
		If Request("appfk") <> "" Then
			sAppFK = Request("appfk")
		End If

		Dim sWinUsername As String = HttpContext.Current.User.Identity.Name


		SaveToLog("Login_DoLogin - Before_CheckWindowsUsername - [" & sWinUsername & "]: inPasswordLength=" & txtWindowsPassword.Text.Trim.Length & "; EdConnectionString=" & EdConnectionString)

		' change - testing only
		Dim lWinUsernameBIUserFK As Long = 9999999999999

		'Dim sWinObj As New Object

		'Using lib1 As New clsDB(EdConnectionString)
		'	sWinObj = lib1.RunSQLScalar("SELECT Ed.[dbo].[udf_SP_WIN_CheckWindowsUsername]('" & sWinUsername & "')")
		'End Using

		'If sWinObj IsNot Nothing AndAlso Not IsDBNull(sWinObj) Then
		'	Long.TryParse(sWinObj, lWinUsernameBIUserFK)
		'End If

		SaveToLog("Login_DoLogin - After_CheckWindowsUsername - [" & sWinUsername & "]: inPasswordLength=" & txtWindowsPassword.Text.Trim.Length & "; lWinUsernameBIUserFK=" & lWinUsernameBIUserFK)



		If lWinUsernameBIUserFK > 0 Then


			Dim sLoginStatus As String = "0"

			Try

				Dim impersonationContext As System.Security.Principal.WindowsImpersonationContext
				Dim currentWindowsIdentity As System.Security.Principal.WindowsIdentity

				currentWindowsIdentity = CType(User.Identity, System.Security.Principal.WindowsIdentity)
				impersonationContext = currentWindowsIdentity.Impersonate()

				'' this doesn't work - no need to do that here anyway
				'sLoginStatus = PC.ADWrapper.Login(txtWindowsUsername.Text, txtWindowsPassword.Text)

				lblStatus.Text = "Login Status = " & sLoginStatus

				SaveToLog("Login_DoLogin (AfterWindowsLoginCheck) [" & sWinUsername & "]: sLoginStatus=" & sLoginStatus & "; txtWindowsUsername.Text=" & txtWindowsUsername.Text & "; txtWindowsPassword.Text.length=" & txtWindowsPassword.Text.Trim.Length & "; sAppFK=" & sAppFK)

				If sLoginStatus = "0" Then
					' *** set the WindowsUsername session ***
					WindowsUsername = sWinUsername.Trim.ToLower.Replace(WindowsDomain, "").Replace("\", "")

					If txtWindowsPassword.Text.Trim.Length > 0 Then

						Dim sEncryptPass As String = libEdContracts.Utils.CryptoEncrypt(txtWindowsPassword.Text.Trim, "support@evolutionjobs.co.uk")

						WindowsPassword = sEncryptPass

						SaveToLog("Login_DoLogin (REDIRECT) [" & sWinUsername & "]: sLoginStatus=" & sLoginStatus & "; txtWindowsUsername.Text=" & txtWindowsUsername.Text & "; WindowsPassword(encrypted)=" & WindowsPassword & "; sAppFK=" & sAppFK)

						Response.Redirect("Report.aspx?appfk=" & sAppFK)
					End If

				End If


			Catch ex As Exception
				lblStatus.Text = ("ERROR: checking password: " & ex.Message)
				SaveToLog("ERROR: Login_DoLogin [" & sWinUsername & "]: sLoginStatus=" & sLoginStatus & "; txtWindowsUsername.Text=" & txtWindowsUsername.Text & "; txtWindowsPassword.Text.length=" & txtWindowsPassword.Text.Trim.Length & "; sAppFK=" & sAppFK & "; EX=" & ex.Message)

			Finally
			End Try


			SaveToLog("Login_DoLogin (END) [" & sWinUsername & "]: sLoginStatus=" & sLoginStatus & "; txtWindowsUsername.Text=" & txtWindowsUsername.Text & "; txtWindowsPassword.Text.length=" & txtWindowsPassword.Text.Trim.Length & "; sAppFK=" & sAppFK)



		Else

			Response.Redirect("Error.aspx?e=Your+windows+username+was+not+recognised.+Please+login+to+the+relevant+intranet+below+first+,+then+try+again")
			Exit Sub

		End If



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



End Class
