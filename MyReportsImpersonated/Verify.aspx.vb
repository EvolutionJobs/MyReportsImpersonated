Imports System.Data

Partial Class _Verify
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        '     <add name="BIWindowsConnectionString" connectionString="Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;application name=WHUA"/>

        'Initial catalog=BI;data source=10.11.24.21;Integrated Security=SSPI;persist security info=True;
        'Initial Catalog=BI;Data Source=10.11.24.21;Integrated Security=SSPI;Persist security info=True;application name=WHUA

        lblWindowsUsername.Text = HttpContext.Current.User.Identity.Name

        'BIWindowsConnectionString = ConfigurationManager.ConnectionStrings("EDConnectionString").ConnectionString
        BIWindowsConnectionString = ConfigurationManager.ConnectionStrings("BIWindowsConnectionString").ConnectionString

        Dim sPanel As String = ""
        If Request("t") <> "" Then
            sPanel = Request("t")
        End If
        lblType.Text = sPanel


        If sPanel = "p" Then
            lblEditIntro.Text = "A verification code has been sent to your phone. Please enter it here:"

        ElseIf sPanel = "e" Then
            lblEditIntro.Text = "A verification code has been sent to your email address. Please enter it here:"

        End If


        If Me.IsPostBack = False Then




        End If

    End Sub



    Protected Sub cmdVerify_Click(sender As Object, e As EventArgs) Handles cmdVerify.Click

        If txtCode.Text.Trim <> "" Then

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
                If lblType.Text = "p" Then

                    If txtCode.Text.Trim = outPhoneCode.Trim Then
                        ' OK = Email Verified or Phone Verified or Email Verification Failed or Phone Verification Failed
                        Using lib1 As New clsDB(BIWindowsConnectionString)
                            ' 1. get Stored info for comparison
                            sGetError = lib1.CheckVerificationCode(lblWindowsUsername.Text, txtCode.Text.Trim, "")
                        End Using

                        If sGetError.Trim.ToLower.Contains(("Failed").ToLower) Then
                            ' failed - return error
                            lblStatus.Text = "ERROR: " & sGetError
                        Else
                            ' OK: check session to see if they clicked the forgotten password and if they did, send SMS with new password
                            ' if they didn't, then redirect them to the login page

                            Dim bSendNewPassword As Boolean = False
                            Try
                                If Session("action") = "RequestNewPassword" Or Session("action") = "FogottenPassword" Then
                                    bSendNewPassword = True
                                End If
                            Catch ex As Exception
                            End Try

                            If bSendNewPassword Then
                                ' 1. get new password
                                ' 2. send via SMS / email

                                lblStatus.Text = "TODO: send new password"
                                SendNewPassword("ViaPhone")

                            Else
                                lblStatus.Text = "Phone verified successfully"
                            End If

                        End If
                    Else
                        ' error - 
                        lblStatus.Text = "ERROR: Wrong code entered"
                    End If

                ElseIf lblType.Text = "e" Then


                    If txtCode.Text.Trim = outEmailCode.Trim Then
                        ' OK = Email Verified or Phone Verified or Email Verification Failed or Phone Verification Failed
                        Using lib1 As New clsDB(BIWindowsConnectionString)
                            ' 1. get Stored info for comparison
                            sGetError = lib1.CheckVerificationCode(lblWindowsUsername.Text, "", txtCode.Text.Trim)
                        End Using

                        If sGetError.Trim.ToLower.Contains(("Failed").ToLower) Then
                            ' failed - return error
                            lblStatus.Text = "ERROR: " & sGetError
                        Else
                            ' OK: check session to see if they clicked the forgotten password and if they did, send SMS with new password
                            ' if they didn't, then redirect them to the login page

                            Dim bSendNewPassword As Boolean = False
                            Try
                                If Session("action") = "RequestNewPassword" Or Session("action") = "FogottenPassword" Then
                                    bSendNewPassword = True
                                End If
                            Catch ex As Exception
                            End Try

                            If bSendNewPassword Then
                                ' 1. get new password
                                ' 2. send via SMS / email

                                lblStatus.Text = "TODO: send new password"
                                SendNewPassword("ViaEmail")


                            Else
                                lblStatus.Text = "Email verified successfully"
                            End If

                        End If
                    Else
                        ' error - 
                        lblStatus.Text = "ERROR: Wrong code entered"
                    End If


                End If
            End If

        Else
            lblStatus.Text = "Please enter a verification code"
        End If


    End Sub


    Public Function SendNewPassword(ByVal Method As String) As String
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

        ' TODO
        ' TODO
        ' TODO
        ' TODO
        ' TODO
        Dim dtCons As New System.Data.DataTable
        'Using lib1 As New clsDB(BIWindowsConnectionString)
        '    dtCons = lib1.GetNewPasswordForAUser("Evolution2", 517)
        'End Using
        ' TODO
        ' TODO
        ' TODO
        ' TODO
        ' TODO

        Dim sCode As String = ""
        If dtCons IsNot Nothing Then

            Dim iCount As Integer = 0
            Dim row As DataRow
            For Each row In dtCons.Rows
                If Not IsDBNull(row(0)) Then
                    sCode = row(0)
                    Exit For
                End If
                iCount += 1
            Next

            Dim sSMSMessage As String = ""
            If Method = "ViaPhone" Then
                ' 3. Send SMS 
                ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
                ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
                ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
                ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
                ' IMPORTANT : THE SMS ONLY WORKS ON A SERVER BECAUSE IT HAS TO HAVE RIGHTS TO RELAY TO EXTERNAL DOMAIN
                sSMSMessage = String.Format(sSMSNewPasswordMessage, sCode)
				sGetError = SendSMS(sDBPhone.Trim.Replace(" ", "").Replace("+", "00") & SMSDomain, sSMSMessage)
				If sGetError.Trim <> "" Then
                    lblStatus.Text = sGetError
                Else
                    lblStatus.Text = "A temporary password has been sent to your mobile.<br />Click the link below to go to the login page."
                End If

            Else

                sSMSMessage = String.Format(sEmailNewPasswordMessage, sCode)
                sGetError = SendSMS(sDBEmail, sSMSMessage)
                If sGetError.Trim <> "" Then
                    lblStatus.Text = sGetError
                Else
                    lblStatus.Text = "A temporary password has been sent to your email.<br />Click the link below to go to the login page."
                End If

            End If

        End If

        Return sGetError
    End Function



    Protected Sub lnkcmdLoginPage_Click(sender As Object, e As EventArgs) Handles lnkcmdLoginPage.Click

    End Sub
    Protected Sub lnkcmdManageMyAccount_Click(sender As Object, e As EventArgs) Handles lnkcmdManageMyAccount.Click

    End Sub

End Class
