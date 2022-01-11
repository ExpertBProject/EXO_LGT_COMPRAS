Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class EXO_OPOR
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try

            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "142"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Form_Load(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oItem As SAPbouiCOM.Item

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Visible = False

            'Buscar XML de update
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

#Region "Botones"
            oItem = oForm.Items.Add("btnENV", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oItem.Width = oForm.Items.Item("2").Width + 30
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Enviar Doc."
            oItem.TextStyle = 1
            oItem.LinkTo = "2"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
#End Region

            oForm.Visible = True

            EventHandler_Form_Load = True

        Catch ex As Exception
            oForm.Visible = True
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnENV"
                    If pVal.ActionSuccess = True Then
                        If EnviarDoc(oForm) = False Then
                            Exit Function
                        End If
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function EnviarDoc(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim sDocEntry As String = "" : Dim sDocNum As String = "" : Dim sFecha As String = ""
        Dim sRutaFicheros As String = "" : Dim rutaCrystal As String = ""
        Dim sContacto As String = "" : Dim sIC As String = "" : Dim sMailProveedor As String = ""
        Dim sMensaje As String = ""
        Dim sCrystal As String = ""
        Dim sReport As String = ""
        Dim sSQL As String = ""
#End Region

        EnviarDoc = False

        Try
            sDocEntry = oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0).ToUpper
            sDocNum = oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocNum", 0).ToUpper
            sFecha = oForm.DataSources.DBDataSources.Item("OPOR").GetValue("DocDate", 0).ToUpper
            rutaCrystal = objGlobal.path : rutaCrystal = objGlobal.pathCrystal
            sCrystal = objGlobal.funcionesUI.refDi.OGEN.valorVariable("RPT_PED_COMPRAS")
            sRutaFicheros = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\PEDCOMPRAS\"
            If System.IO.Directory.Exists(sRutaFicheros) = False Then
                System.IO.Directory.CreateDirectory(sRutaFicheros)
            Else
                'Borramos Hco.
                Dim Fecha As DateTime = DateTime.Now
                For Each archivo As String In My.Computer.FileSystem.GetFiles(sRutaFicheros, FileIO.SearchOption.SearchTopLevelOnly)
                    Dim Fecha_Archivo As DateTime = My.Computer.FileSystem.GetFileInfo(archivo).LastWriteTime
                    Dim diferencia = (CType(Fecha, DateTime) - CType(Fecha_Archivo, DateTime)).TotalDays

                    If diferencia >= 30 Then ' Nº de días
                        File.Delete(archivo)
                    End If
                Next
            End If
            If CType(oForm.Items.Item("85").Specific, SAPbouiCOM.ComboBox).Selected IsNot Nothing Then
                sContacto = CType(oForm.Items.Item("85").Specific, SAPbouiCOM.ComboBox).Selected.Value
                sIC = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value.ToString
            End If
            If sContacto <> "" Then
                sSQL = "SELECT ""E_MailL"" FROM ""OCPR"" WHERE ""CardCode""='" & sIC & "' and ""CntctCode""='" & sContacto & "' "
                sMailProveedor = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                If sMailProveedor.Trim <> "" Then
                    GenerarCrystal(rutaCrystal, sCrystal, sDocEntry, sDocNum, sFecha, sRutaFicheros, sReport)
                    EnviarMail(sRutaFicheros, sReport, sDocNum, sMailProveedor)
                Else
                    sMensaje = "El contacto no tiene asignado un Mail. No se puede enviar el documento."
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox(sMensaje)
                End If
            Else
                sMensaje = "El documento no tiene asignado Contacto. No se puede enviar el documento."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
            End If
            EnviarDoc = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
    Public Sub GenerarCrystal(ByVal rutaCrystal As String, ByVal sCrystal As String, ByVal sDocEntry As String, ByVal sDocNum As String, ByVal sFecha As String, ByVal sFileName As String, ByRef sReport As String)

        Dim oCRReport As ReportDocument = Nothing
        Dim oFileDestino As DiskFileDestinationOptions = Nothing
        Dim sServer As String = ""
        Dim sDriver As String = ""
        Dim sBBDD As String = ""
        Dim sUser As String = ""
        Dim sPwd As String = ""
        Dim sConnection As String = ""
        Dim oLogonProps As NameValuePairs2 = Nothing

        Dim conrepor As DataSourceConnections = Nothing
        Try
            oCRReport = New ReportDocument()

            oCRReport.Load(rutaCrystal & "\" & sCrystal)

            oCRReport.DataSourceConnections.Clear()

            'Establecemos las conexiones a la BBDD
            sServer = "hana:30015" ' objGlobal.compañia.Server
            'sServer = objGlobal.refDi.SQL.dameCadenaConexion.ToString
            sBBDD = objGlobal.compañia.CompanyDB
            sUser = objGlobal.refDi.SQL.usuarioSQL
            sPwd = objGlobal.refDi.SQL.claveSQL

            sDriver = "HDBODBC"
            sConnection = "DRIVER={" & sDriver & "};UID=" & sUser & ";PWD=" & sPwd & ";SERVERNODE=" & sServer & ";DATABASE=" & sBBDD & ";"
            'sConnection = "DRIVER={" & sDriver & "};" & sServer & ";DATABASE=" & sBBDD & ";"
            objGlobal.SBOApp.StatusBar.SetText("Conectando: " & sConnection, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
            oLogonProps = oCRReport.DataSourceConnections(0).LogonProperties
            oLogonProps.Set("Provider", sDriver)
            oLogonProps.Set("Connection String", sConnection)

            oCRReport.DataSourceConnections(0).SetLogonProperties(oLogonProps)
            oCRReport.DataSourceConnections(0).SetConnection(sServer, sBBDD, False)

            For Each oSubReport As ReportDocument In oCRReport.Subreports
                For Each oConnection As IConnectionInfo In oSubReport.DataSourceConnections
                    oConnection.SetConnection(sServer, sBBDD, False)
                    oConnection.SetLogon(sUser, sPwd)
                Next
            Next
            'Establecemos los parámetros para el report.
            oCRReport.SetParameterValue("DocKey@", sDocEntry)
            oCRReport.SetParameterValue("ObjectId@", "22")
            oCRReport.SetParameterValue("Schema@", sBBDD)

            'Preparamos para la exportación
            sReport = sFileName & "PEDIDO_" & sDocNum & "_" & sFecha & ".pdf"
            'Compruebo si existe y lo borro
            If IO.File.Exists(sReport) Then
                IO.File.Delete(sReport)
            End If
            objGlobal.SBOApp.StatusBar.SetText("Generando pdf para envio impresión...Espere por favor", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)

            oCRReport.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

            oFileDestino = New CrystalDecisions.Shared.DiskFileDestinationOptions
            oFileDestino.DiskFileName = sReport

            'Le pasamos al reporte el parámetro destino del reporte (ruta)
            oCRReport.ExportOptions.DestinationOptions = oFileDestino

            'Le indicamos que el reporte no es para mostrarse en pantalla, sino, que es para guardar en disco
            oCRReport.ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile

            'Finalmente exportamos el reporte a PDF
            oCRReport.Export()
            '            oCRReport.ExportToDisk(ExportFormatType.PortableDocFormat, sReport)


            'Cerramos
            oCRReport.Close()
            oCRReport.Dispose()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oCRReport = Nothing
            oFileDestino = Nothing
        End Try
    End Sub
    Public Sub EnviarMail(ByRef sFileName As String, ByRef sReport As String, ByVal sNumPedido As String, ByVal sMailProveedor As String)
        Dim correo As New System.Net.Mail.MailMessage()
        Dim adjunto As System.Net.Mail.Attachment

        Dim StrFirma As String = ""
        Dim htmbody As New System.Text.StringBuilder()
        Dim cuerpo As String = ""

        Dim sMailSaliente As String = ""
        Dim sPrioridad As String = "" : Dim sNotificacion As String = "" : Dim sMensaje As String = ""
        Dim sUs As String = "" : Dim sPass As String = ""
        Try
            sMailSaliente = objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_MAIL")
            correo.From = New System.Net.Mail.MailAddress(sMailSaliente, "Lingotes Especiales, S.A.")
            If sReport <> "" Then
                adjunto = New System.Net.Mail.Attachment(sReport)
                correo.Attachments.Add(adjunto)
            End If

            Dim FicheroCab As String = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\mail.htm"
            Dim srCAB As StreamReader = New StreamReader(FicheroCab)

            cuerpo = srCAB.ReadToEnd()
            correo.Subject = "Confirmación Pedido Nº " & sNumPedido.ToString()
            correo.Body = cuerpo
            correo.IsBodyHtml = True
            sPrioridad = objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_PRIORIDAD")
            Select Case sPrioridad.Trim
                Case "0" : correo.Priority = System.Net.Mail.MailPriority.Normal
                Case "1" : correo.Priority = System.Net.Mail.MailPriority.Low
                Case "2" : correo.Priority = System.Net.Mail.MailPriority.High
                Case Else : correo.Priority = System.Net.Mail.MailPriority.Normal
            End Select
            sNotificacion = objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_NOTIFICACION")
            Select Case sNotificacion.Trim
                Case "0" : correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.None
                Case "1" : correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.OnSuccess
                Case "2" : correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.OnFailure
                Case Else : correo.DeliveryNotificationOptions = Net.Mail.DeliveryNotificationOptions.Never
            End Select

            correo.To.Add(sMailProveedor)

            Dim smtp As New System.Net.Mail.SmtpClient

            smtp.Host = objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_HOST")
            smtp.Port = CInt(objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_PORT"))
            smtp.UseDefaultCredentials = True
            sUs = objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_MAIL_US")
            sPass = objGlobal.funcionesUI.refDi.OGEN.valorVariable("ENV_MAIL_PASS")
            smtp.Credentials = New System.Net.NetworkCredential(sUs, sPass)
            smtp.EnableSsl = True

            smtp.Send(correo)
            correo.Dispose()
            sMensaje = "Mensaje Enviado con el Pedido Nº " & sNumPedido & "."
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox(sMensaje)
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub
End Class
