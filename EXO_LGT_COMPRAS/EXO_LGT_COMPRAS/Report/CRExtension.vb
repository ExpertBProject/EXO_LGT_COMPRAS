Imports System.Runtime.CompilerServices
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports

Namespace Extensions
    Public Module CrystalReportExtensions
        <Extension()>
        Public Sub ApplyNewServer(ByVal report As ReportDocument, driver As String, serverName As String, username As String, password As String, databasename As String)

            For Each subReport As ReportDocument In report.Subreports
                For Each crTable As Table In subReport.Database.Tables
                    Dim loi As TableLogOnInfo = crTable.LogOnInfo
                    loi.ConnectionInfo.AllowCustomConnection = True
                    'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Connection String").Value.ToString <> "" Then
                    '    Dim x As String() = Split(loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Connection String").Value.ToString, ";")
                    '    Dim Texto As String = loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Connection String").Value.ToString.Trim
                    '    For i = 0 To x.Length - 1
                    '        Dim m As String() = Split(x(0), "=")
                    '        If m.Length > 1 Then
                    '            Select Case m(0).Trim
                    '                Case "DRIVER"
                    '                    If Texto.Length > 0 Then
                    '                        Texto = Texto & ";"
                    '                    End If
                    '                    Texto = Texto & m(0).Trim & "=" & driver.Trim
                    '                Case "UID"
                    '                    If Texto.Length > 0 Then
                    '                        Texto = Texto & ";"
                    '                    End If
                    '                    Texto = Texto & m(0).Trim & "=" & username.Trim
                    '                Case "PWD"
                    '                    If Texto.Length > 0 Then
                    '                        Texto = Texto & ";"
                    '                    End If
                    '                    Texto = Texto & m(0).Trim & "=" & password.Trim
                    '                Case "SERVERNODE"
                    '                    If Texto.Length > 0 Then
                    '                        Texto = Texto & ";"
                    '                    End If
                    '                    Texto = Texto & m(0).Trim & "=" & serverName.Trim
                    '                Case "DATABASE"
                    '                    If Texto.Length > 0 Then
                    '                        Texto = Texto & ";"
                    '                    End If
                    '                    Texto = Texto & m(0).Trim & "=" & databasename.Trim
                    '            End Select
                    '        End If
                    '    Next
                    '    loi.ConnectionInfo.LogonProperties.Set("Connection String", Texto)
                    'End If
                    'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Provider").Value.ToString.Trim <> driver Then
                    '    loi.ConnectionInfo.LogonProperties.Set("Provider", driver)
                    'End If
                    'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("PreQEServerName").Value.ToString.Trim <> serverName Then
                    '    loi.ConnectionInfo.LogonProperties.Set("PreQEServerName", serverName)
                    'End If
                    'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Server").Value.ToString.Trim <> serverName Then
                    '    loi.ConnectionInfo.LogonProperties.Set("Server", serverName)
                    'End If
                    loi.ConnectionInfo.DatabaseName = databasename
                    loi.ConnectionInfo.ServerName = serverName
                    loi.ConnectionInfo.UserID = username
                    loi.ConnectionInfo.Password = password
                    crTable.ApplyLogOnInfo(loi)
                Next
            Next

            'Loop through each table in the report and apply the new login information (in our case, a DSN)
            For Each crTable As Table In report.Database.Tables
                Dim loi As TableLogOnInfo = crTable.LogOnInfo
                loi.ConnectionInfo.AllowCustomConnection = True
                'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Connection String").Value.ToString <> "" Then
                '    Dim x As String() = Split(loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Connection String").Value.ToString, ";")
                '    Dim Texto As String = "" 'loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Connection String").Value.ToString.Trim
                '    For i = 0 To x.Length - 1
                '        If x(i) = "" Then Continue For
                '        Dim m As String() = Split(x(i), "=")
                '        If m.Length > 1 Then
                '            Select Case m(0).Trim
                '                Case "DRIVER"
                '                    If Texto.Length > 0 Then
                '                        Texto = Texto & ";"
                '                    End If
                '                    Texto = Texto & m(0).Trim & "=" & driver.Trim
                '                Case "UID"
                '                    If Texto.Length > 0 Then
                '                        Texto = Texto & ";"
                '                    End If
                '                    Texto = Texto & m(0).Trim & "=" & username.Trim
                '                Case "PWD"
                '                    If Texto.Length > 0 Then
                '                        Texto = Texto & ";"
                '                    End If
                '                    Texto = Texto & m(0).Trim & "=" & password.Trim
                '                Case "SERVERNODE"
                '                    If Texto.Length > 0 Then
                '                        Texto = Texto & ";"
                '                    End If
                '                    Texto = Texto & m(0).Trim & "=" & serverName.Trim
                '                Case "DATABASE"
                '                    If Texto.Length > 0 Then
                '                        Texto = Texto & ";"
                '                    End If
                '                    Texto = Texto & m(0).Trim & "=" & databasename.Trim
                '            End Select
                '        End If
                '    Next
                '    loi.ConnectionInfo.LogonProperties.Set("Connection String", Texto)
                'End If
                'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Provider").Value.ToString.Trim <> driver Then
                '    loi.ConnectionInfo.LogonProperties.Set("Provider", driver)
                'End If

                ''error no sabmeos que es esta variable
                ''If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("PreQEServerName").Value.ToString.Trim <> serverName Then
                ''loi.ConnectionInfo.LogonProperties.Set("PreQEServerName", serverName)
                ''End If



                'If loi.ConnectionInfo.LogonProperties.LookupNameValuePair("Server").Value.ToString.Trim <> serverName Then
                '    loi.ConnectionInfo.LogonProperties.Set("Server", serverName)
                'End If
                loi.ConnectionInfo.DatabaseName = databasename
                loi.ConnectionInfo.ServerName = serverName
                loi.ConnectionInfo.UserID = username
                loi.ConnectionInfo.Password = password
                crTable.ApplyLogOnInfo(loi)
            Next

        End Sub
        ''' <summary>
        ''' Applies a new server name to the ReportDocument.  This method is SQL Server specific if integratedSecurity is True.
        ''' </summary>
        ''' <param name="report"></param>
        ''' <param name="serverName">The name of the new server.</param>
        ''' <param name="integratedSecurity">Whether or not to apply integrated security to the ReportDocument.</param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub ApplyNewServer(report As ReportDocument, serverName As String, integratedSecurity As Boolean)

            For Each subReport As ReportDocument In report.Subreports
                For Each crTable As Table In subReport.Database.Tables
                    Dim loi As TableLogOnInfo = crTable.LogOnInfo
                    loi.ConnectionInfo.ServerName = serverName

                    If integratedSecurity = True Then
                        loi.ConnectionInfo.IntegratedSecurity = True
                    End If

                    crTable.ApplyLogOnInfo(loi)
                Next
            Next

            'Loop through each table in the report and apply the new login information (in our case, a DSN)
            For Each crTable As Table In report.Database.Tables
                Dim loi As TableLogOnInfo = crTable.LogOnInfo
                loi.ConnectionInfo.ServerName = serverName

                If integratedSecurity = True Then
                    loi.ConnectionInfo.IntegratedSecurity = True
                End If

                crTable.ApplyLogOnInfo(loi)
                'If your DatabaseName is changing at runtime, specify the table location. 
                'crTable.Location = ci.DatabaseName & ".dbo." & crTable.Location.Substring(crTable.Location.LastIndexOf(".") + 1)
            Next

        End Sub
        <Extension()>
        Public Function DoesParameterExist(ByVal report As ReportDocument, ByVal paramName As String) As Boolean

            If report Is Nothing Or report.ParameterFields Is Nothing Then
                Return False
            End If

            For Each param As ParameterField In report.ParameterFields
                If paramName = param.Name Then
                    Return True
                End If
            Next

            Return False
        End Function
        ''' <summary>
        ''' Takes a parameter string and places them in the corresponding parameters for the report.  The parameter string must 
        ''' be semi-colon delimited with the parameter inside of that delimited with an equal sign.  E.g.<br /><br />
        ''' 
        ''' <code>
        ''' lastName=Pell;startDate=1/1/2012;endDate=1/7/2012
        ''' </code>
        ''' 
        ''' </summary>
        ''' <param name="report">The Crystal Reports ReportDocument object.</param>
        ''' <param name="parameters">A parameter string representing name/values.  See the summary for usage.</param>
        ''' <remarks></remarks>
        <Extension()>
        Public Sub ApplyParameters(report As ReportDocument, parameters As String)
            ApplyParameters(report, parameters, True)
        End Sub
        <Extension()>
        Public Sub ApplyParameters(report As ReportDocument, parameters As String, removeInvalidParameters As Boolean)

            ' No parameters (or valid parameters) were provided.
            If String.IsNullOrEmpty(parameters) = True Or parameters.Contains("=") = False Then
                Exit Sub
            End If

            ' Get rid of any trailing or leading semi-colons that would mess up the splitting.
            parameters = parameters.Trim(";".ToCharArray)

            ' The list of parameters split out by the semi-colon delimiter
            Dim parameterList As String() = parameters.Split(Chr(Asc(";")))

            For Each parameter As String In parameterList
                ' nameValue(0) = Parameter Name, nameValue(0) = Value
                Dim nameValue As String() = parameter.Split(Chr(Asc("=")))

                ' Validate that the parameter exists and throw a legit exception that describes it as opposed to the
                ' Crystal Report COM Exception that gives you little detail.  
                If report.DoesParameterExist(nameValue(0)) = False And removeInvalidParameters = False Then
                    Throw New Exception(String.Format("The parameter '{0}' does not exist in the Crystal Report.", nameValue(0)))
                ElseIf report.DoesParameterExist(nameValue(0)) = False And removeInvalidParameters = True Then
                    Continue For
                End If

                ' The ParameterFieldDefinition MUST be disposed of otherwise memory issues will occur, that's why
                ' we're going the "using" route.  Using should Dispose of it even if an Exception occurs.
                Using pfd As ParameterFieldDefinition = report.DataDefinition.ParameterFields.Item(nameValue(0))
                    Dim pValues As ParameterValues
                    Dim parm As ParameterDiscreteValue
                    pValues = New ParameterValues

                    parm = New ParameterDiscreteValue
                    parm.Value = nameValue(1)

                    pValues.Add(parm)
                    pfd.ApplyCurrentValues(pValues)
                End Using
            Next

        End Sub

    End Module

End Namespace




