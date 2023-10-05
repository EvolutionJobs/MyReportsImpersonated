Imports System.Data

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        '     <add name="BIWindowsConnectionString" connectionString="Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;application name=WHUA"/>

        'Initial catalog=BI;data source=10.11.24.21;Integrated Security=SSPI;persist security info=True;
        'Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;Persist security info=True;application name=WHUA

        ' http://localhost:52354/SSRSLayer/Default.aspx?appfk=22

        Dim inApplicationFK As Integer = 0
        Try
            If Request("appfk") <> "" Then
                inApplicationFK = CInt(Request("appfk"))
            End If
        Catch ex As Exception
        End Try


		lblWindowsUsername.Text = HttpContext.Current.User.Identity.Name

        If Me.IsPostBack = False Then

            If inApplicationFK > 0 Then
                pnlLogin.Visible = True
                pnlChangePassword.Visible = False
            Else
                pnlLogin.Visible = False
                pnlChangePassword.Visible = False
                lblStatus.Text = "ERROR: Application value missing"
            End If

        End If

    End Sub

    Protected Sub cmdSubmit_Click(sender As Object, e As System.EventArgs) Handles cmdSubmit.Click

        If txtPassword.Text.Trim = "" Then
            lblStatus.Text = "ERROR: Please enter a password"
            Exit Sub
        End If

        Dim inApplicationFK As Integer = 0
        Try
            If Request("appfk") <> "" Then
                inApplicationFK = CInt(Request("appfk"))
            End If
        Catch ex As Exception
        End Try


        If inApplicationFK <= 0 Then
            lblStatus.Text = "ERROR: no Report URL returned"
            Exit Sub
        End If


        Dim PasswordStatus As String = ""
        Dim sResultFK As String = ""
        Dim sReportURL As String = ""
        Dim sSeed As String = ""

        Try

			Dim sWinUsername As String = HttpContext.Current.User.Identity.Name

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
						SaveToLog(sWinUsername & " Default-CheckPassword_Impersonation_____PRE_ImpersonateValidUser")
						Using lib1 As New clsDB(PasswordAdminEdConnectionString)
							dtCons = lib1.CheckPassword_Impersonation(txtPassword.Text, inApplicationFK)
						End Using

						SaveToLog(sWinUsername & " Default-CheckPassword_Impersonation_____POST_ImpersonateValidUser")
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
            lblStatus.Text = ("ERROR: checking password: " & ex.Message)
        Finally
        End Try


        If sReportURL.Trim <> "" Then
            If sResultFK = "5" Then

                Dim sURL As String = sReportURL & "&RandNum=" & sSeed
                lblURL.Text = "<a href='" & sURL & "'>Click Here</a>"
                Response.Redirect(sURL)
                ' http://evosvr05/ReportServer/Pages/ReportViewer.aspx?%2fWagesPages%2fOpeningPage&rs:Command=Render&RandNum=517251
                '"http://evosvr05/ReportServer/Pages/ReportViewer.aspx?%2fWagesPages%2fOpeningPage&rs:Command=Render&RandNum="
            Else
                lblStatus.Text = PasswordStatus
            End If
        Else
            lblStatus.Text = "ERROR: no Report URL returned"

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
        pnlLogin.Visible = False
        pnlChangePassword.Visible = True
    End Sub

    Protected Sub cmdChangePassword_Click(sender As Object, e As System.EventArgs) Handles cmdChangePassword.Click

        If (txtOldPassword.Text <> "") And (txtNewPassword1.Text <> "") And (txtNewPassword2.Text <> "") Then

            If (txtNewPassword1.Text = txtNewPassword2.Text) Then
                ' run the change password routine
                ' and run the login sproc

            Else
                lblStatus.Text = "ERROR: Passwords do not match"
            End If

        Else
            lblStatus.Text = "ERROR: Please fill in all fields"
        End If

    End Sub


End Class
