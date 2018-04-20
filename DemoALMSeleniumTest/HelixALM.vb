Imports System.Net
Imports RestSharp
Imports RestSharp.Authenticators

Public Class HelixALM

    'Base URL
    Private Property RESTBaseURL As String = ""
    'Auth Token that is needed to Auth and execute actions
    Private Property AuthToken As String = ""
    Private Property ProjectId As String = "0"

    'Connect and get Auth Token
    Sub New(BaseURL As String, ProjectName As String, AuthUser As String, AuthPassword As String)
        Try

            'Ignore https
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            'Set Variables
            Me.RESTBaseURL = BaseURL
            Me.AuthToken = ""

            'If no valid Auth token exists we need to connect to get one
            If Me.AuthToken = "" Then
                'Get the project ID to work with
                Me.ProjectId = FindProjectID(ProjectName, AuthUser, AuthPassword)
                Me.AuthToken = ProjectAuthToken(Me.ProjectId, AuthUser, AuthPassword)
            Else
                'Auth Token already exists, no need to get one.
            End If

        Catch ex As Exception
            Me.ProjectId = "0"
            Me.AuthToken = ""
        End Try
    End Sub

    'Functions To Access REST API
    Private Function FindProjectID(ProjectName As String, AuthUser As String, AuthPassword As String) As String

        'REQUEST
        'GET https://localhost:8443/helix-alm/api/v0/projects
        'Authorization: Basic

        Dim strReturnValue As String = "0"

        Try
            'Basic Auth against REST API to get project list
            Dim objRestClient As New RestSharp.RestClient(Me.RESTBaseURL)
            objRestClient.Authenticator = New HttpBasicAuthenticator(AuthUser, AuthPassword)

            Dim objRestRequest = New RestSharp.RestRequest("projects", RestSharp.Method.GET)
            Dim objRestResponse As New RestSharp.RestResponse

            objRestResponse = objRestClient.Execute(objRestRequest)

            'Only perform the actions if the response is completed and status 300
            If objRestResponse.ResponseStatus = RestSharp.ResponseStatus.Completed And objRestResponse.StatusCode = HttpStatusCode.OK Then

                Dim objJS As New System.Web.Script.Serialization.JavaScriptSerializer
                Dim jsonObject As Dictionary(Of String, Object) = objJS.Deserialize(Of Object)(objRestResponse.Content)

                For Each objObject In jsonObject

                    If objObject.Key = "projects" Then

                        For Each objProjects In objObject.Value
                            strReturnValue = objProjects.ToString

                            If Trim(ProjectName) = Trim(objProjects("name").ToString) Then
                                strReturnValue = objProjects("id").ToString
                                Exit For
                            End If

                        Next

                    End If

                Next

            Else
                strReturnValue = "0"
            End If

            Return strReturnValue

        Catch ex As Exception
            Return "0"
        End Try
    End Function

    'Functions To Access Auth Token in REST API
    Private Function ProjectAuthToken(ProjectId As String, AuthUser As String, AuthPassword As String) As String

        'REQUEST
        'GET https://localhost:8443/helix-alm/api/v0/4/token

        'Subsequents Calls, Example:
        'Authorization: Bearer eyJhbGciOI6IkpXVCJ9.eyJleHAi3Q6jLTQibGljIjoicH

        Dim strReturnValue As String = "0"

        Try
            'Return Auth Token that can be used (Ensure we pass the projectid we want to login to).
            Dim objRestClient As New RestSharp.RestClient(Me.RESTBaseURL & ProjectId & "/")
            objRestClient.Authenticator = New HttpBasicAuthenticator(AuthUser, AuthPassword)

            Dim objRestRequest = New RestSharp.RestRequest("token", RestSharp.Method.GET)
            Dim objRestResponse As New RestSharp.RestResponse

            objRestResponse = objRestClient.Execute(objRestRequest)

            'Only perform the actions if the response is completed and status 300
            If objRestResponse.ResponseStatus = RestSharp.ResponseStatus.Completed And objRestResponse.StatusCode = HttpStatusCode.OK Then

                Dim objJS As New System.Web.Script.Serialization.JavaScriptSerializer
                Dim jsonObject As Dictionary(Of String, Object) = objJS.Deserialize(Of Object)(objRestResponse.Content)

                For Each objObject In jsonObject

                    If objObject.Key = "accessToken" Then
                        strReturnValue = objObject.Value.ToString
                        Exit For
                    End If

                Next

            Else
                strReturnValue = "0"
            End If

            Return strReturnValue

        Catch ex As Exception
            Return ""
        End Try
    End Function








    'Functions that can be used by all Methods
    'Test passed, ensure we mark the test execution (Test Run) as passed.
    Public Function TestPassed(TCid As Long, TCDetail As String) As String

        Dim strJSON As String = ""

        Try

            ''Build JSON Object.. in Loop and parse and send to server
            'strJSON = "{" &
            '            " ""testCaseIDs"" : [" & TCid & "], " &
            '            " ""eventsData"" : [ { ""name"" : ""Pass"", ""fields"" : [ { ""label"" : ""Notes"", ""String"" : ""Generated through REST API - " & TCDetail & """ } ] } ] " &
            '            "}"

            strJSON = "{" &
                        " ""testCaseIDs"" : [" & TCid & "], " &
                        " ""eventsData"" : [ { ""name"" : ""Pass"", ""fields"" : [ { ""label"" : ""Notes"", ""string"" : ""Generated through REST API"" } ] } ] " &
                        "}"

            ''{
            ''	"testCaseIDs" : [1, 2, 3],
            ''	"folder" : { "path" : "/Public/a" },
            ''	"testRunSet" : { "label" : "RunSet1" },
            ''	"variants" : [
            ''		{
            ''			"label" : "Alphabet",
            ''			"menuItemArray" : [
            ''				{ "label" : "a" },
            ''				{ "label" : "n" }, 
            ''				{ "label" : "d" }, 
            ''				{ "label" : "y" },
            ''				{ "label" : "r" },
            ''				{ "label" : "o" },
            ''				{ "label" : "c" },
            ''				{ "label" : "k" },
            ''				{ "label" : "s" },
            ''				{ "label" : "z" }
            ''			]
            ''		}

            ''	],
            ''	"eventsData" : [ 
            ''		{ "name" : "Comment", "fields" : [ { "label" : "Notes", "string" : "Generated through REST API" } ] },
            ''		{ "name" : "Pass" } ]
            ''}


            Dim objJS As New System.Web.Script.Serialization.JavaScriptSerializer
            Dim jsonObject As Dictionary(Of String, Object) = objJS.Deserialize(Of Object)(strJSON)


            'Use the Helix ALM REST API
            If Me.AuthToken <> "" And Me.ProjectId <> "0" Then

                'Use AuthToken
                Dim objRestClient As New RestSharp.RestClient(Me.RESTBaseURL & Me.ProjectId & "/testruns/")
                objRestClient.Authenticator = New HttpBasicAuthenticator("Bearer", Me.AuthToken)

                Dim objRestRequest = New RestSharp.RestRequest("generate", RestSharp.Method.POST)
                Dim objRestResponse As New RestSharp.RestResponse

                objRestRequest.AddParameter("Authorization", String.Format("Bearer " + Me.AuthToken), ParameterType.HttpHeader)
                objRestRequest.AddJsonBody(jsonObject) 'JSON Object to Send

                'Create Test Run for Test Case Id
                objRestResponse = objRestClient.Execute(objRestRequest) 'Execute Json POST Request

                'Status Code 429 Indicates to many requests.. wait and try again
                If objRestResponse.StatusCode = 429 Then
                    'Way too many requests.. pause and continue in a 30seconds again
                    'Delay for 30 seconds
                    Threading.Thread.Sleep(30000)
                    objRestResponse = objRestClient.Execute(objRestRequest)
                End If




                'Mark the new Test Run as Passed



            End If

        Catch ex As Exception

        End Try
    End Function

    'Test failed, ensure we mark the test execution (Test Run) as failed and create a defect record.
    Public Function TestFailed(TCid As Long, TCDetail As String) As String

        Dim strJSON As String = ""

        Try

            ''Build JSON Object.. in Loop and parse and send to server
            'strJSON = "{" &
            '            " ""testCaseIDs"" : [" & TCid & "], " &
            '            " ""eventsData"" : [ { ""name"" : ""Fail"", ""fields"" : [ { ""label"" : ""Notes"", ""String"" : ""Generated through REST API - " & TCDetail & """ } ] } ] " &
            '            "}"

            strJSON = "{" &
                        " ""testCaseIDs"" : [" & TCid & "], " &
                        " ""eventsData"" : [ { ""name"" : ""Fail"", ""fields"" : [ { ""label"" : ""Notes"", ""string"" : ""Generated through REST API"" } ] } ] " &
                        "}"


            Dim objJS As New System.Web.Script.Serialization.JavaScriptSerializer
            Dim jsonObject As Dictionary(Of String, Object) = objJS.Deserialize(Of Object)(strJSON)


            'Use the Helix ALM REST API
            If Me.AuthToken <> "" And Me.ProjectId <> "0" Then

                'Use AuthToken
                Dim objRestClient As New RestSharp.RestClient(Me.RESTBaseURL & Me.ProjectId & "/testruns/")
                objRestClient.Authenticator = New HttpBasicAuthenticator("Bearer", Me.AuthToken)

                Dim objRestRequest = New RestSharp.RestRequest("generate", RestSharp.Method.POST)
                Dim objRestResponse As New RestSharp.RestResponse

                objRestRequest.AddParameter("Authorization", String.Format("Bearer " + Me.AuthToken), ParameterType.HttpHeader)
                objRestRequest.AddJsonBody(jsonObject) 'JSON Object to Send

                'Create Test Run for Test Case Id
                objRestResponse = objRestClient.Execute(objRestRequest) 'Execute Json POST Request

                'Status Code 429 Indicates to many requests.. wait and try again
                If objRestResponse.StatusCode = 429 Then
                    'Way too many requests.. pause and continue in a 30seconds again
                    'Delay for 30 seconds
                    Threading.Thread.Sleep(30000)
                    objRestResponse = objRestClient.Execute(objRestRequest)
                End If


                'Mark the new Test Run as Passed



            End If

        Catch ex As Exception

        End Try
    End Function


End Class
