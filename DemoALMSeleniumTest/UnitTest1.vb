Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Edge

<TestClass()> Public Class UnitTest1

    'Helix ALM variables for automation result(s).
    Private ALMRESTAPI As String = "http://localhost:8900/helix-alm/api/v0/"
    Private ALMProject As String = "Sample Project"
    Private ALMUser As String = "Administrator"
    Private ALMPassword As String = ""


    'Private objPAGEDriver As New ChromeDriver
    'Private objPAGEDriver As New InternetExplorerDriver
    Private objPAGEDriver As New EdgeDriver


    'Helix ALM variables for automation result(s).
    Private objHelixALMAPI As New HelixALM(ALMRESTAPI, ALMProject, ALMUser, ALMPassword)

    <TestMethod()> Public Sub TestInvalidAuthError()

        'Set Test Case RecordId found in Helix ALM
        Dim TCid As Long = 97

        Try

            'Set Browser Settings...
            objPAGEDriver.Manage().Window.Maximize()
            objPAGEDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30)

            'Driver Test using Web Driver
            objPAGEDriver.Navigate().GoToUrl("http://core.floodgate.co.za/UI_v2/login.aspx")

            'Find Elements
            objPAGEDriver.FindElementById("user_name").SendKeys("test")
            objPAGEDriver.FindElementById("user_password").SendKeys("test")

            objPAGEDriver.FindElementById("btLogin").Click()

            'Check for Error Div
            'lblErrorMessage, read value in div lblErrorMessage
            Dim objWebElement As IWebElement

            objWebElement = objPAGEDriver.FindElementById("lblErrorMessage")

            If objWebElement.Displayed = False Then
                'Test Failed
                Assert.Fail("Auth error message not visible on error.")
            Else
                'Verify Display Value
                If objWebElement.Text <> "Invalid login details" Then
                    'Test Failed
                    Assert.Fail("Auth error message is incorrect.")
                Else
                    'Load Result(s)
                    objHelixALMAPI.TestPassed(TCid, "Test Auth Login Page")
                End If
            End If

        Catch eFailed As AssertFailedException
            'Load Result(s)
            objHelixALMAPI.TestFailed(TCid, "Test Auth Login Page - " & eFailed.Message)

        Catch ex As Exception
            'Test Failed with Exception
            Assert.Fail()
            'Load Result(s)
            objHelixALMAPI.TestFailed(TCid, "Test Auth Login Page")

        Finally
            objPAGEDriver.Close()

        End Try

    End Sub

    <TestMethod()> Public Sub TestCopyRight()

        'Set Test Case RecordId found in Helix ALM
        Dim TCid As Long = 98

        Try

            'Set Browser Settings
            objPAGEDriver.Manage().Window.Maximize()
            objPAGEDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30)

            'Driver Test using Web Driver
            objPAGEDriver.Navigate().GoToUrl("http://core.floodgate.co.za/UI_v2/login.aspx")

            'Check for CopyRight Div
            'copyright, read div value
            Dim objWebElement As IWebElement

            objWebElement = objPAGEDriver.FindElementByClassName("copyright")

            If objWebElement.Displayed = False Then
                'Test Failed
                Assert.Fail("Copyright not visible.")
            Else
                'Verify Display Value
                If objWebElement.Text <> Date.Now.Year.ToString & " © Floodgate Agenciess " Then
                    'Test Failed
                    Assert.Fail("Copyright does not match required value.")
                Else
                    'Load Result(s)
                    objHelixALMAPI.TestPassed(TCid, "Test Copy Right")
                End If
            End If

        Catch eFailed As AssertFailedException
            'Load Result(s)
            objHelixALMAPI.TestFailed(TCid, "Test Copy Right - " & eFailed.Message)

        Catch ex As Exception
            'Test Failed with Exception
            Assert.Fail()
            'Load Result(s)
            objHelixALMAPI.TestFailed(TCid, "Test Copy Right")

        Finally
            objPAGEDriver.Close()

        End Try

    End Sub

End Class