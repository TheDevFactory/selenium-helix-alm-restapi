Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting

Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Edge

<TestClass()> Public Class UnitTest1

    'Private objPAGEDriver As New ChromeDriver
    'Private objPAGEDriver As New InternetExplorerDriver
    Private objPAGEDriver As New EdgeDriver


    <TestMethod()> Public Sub TestInvalidAuthError()

        Try

            'Set Browser Settings
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
                Assert.Fail()
            Else
                'Verify Display Value
                If objWebElement.Text <> "Invalid login details" Then
                    'Test Failed
                    Assert.Fail()
                End If
            End If

        Catch ex As Exception
            Assert.Fail()
        Finally
            objPAGEDriver.Close()
        End Try

    End Sub

    <TestMethod()> Public Sub TestCopyRight()

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
                Assert.Fail()
            Else
                'Verify Display Value
                If objWebElement.Text <> Date.Now.Year.ToString & " © Floodgate Agencies " Then
                    'Test Failed
                    Assert.Fail()
                End If
            End If

        Catch ex As Exception
            Assert.Fail()
        Finally
            objPAGEDriver.Close()
        End Try

    End Sub

End Class