Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Variables Globales"
    Public Shared _sArticulo As String = ""
    Public Shared _sCardCode As String = ""
    Public Shared _sCatalogo As String = ""
    Public Shared _sLineaSel As String = ""
#End Region

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)

        If actualizar Then
            cargaDatos()
            Cambiar_Nombre_Propiedades()
            ParametrizacionGeneral()
            CargaFirma()
        End If
        cargamenu()
    End Sub
    Private Sub cargaDatos()
        Dim sXML As String = ""
        Dim res As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_SUPLEMENTO.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_SUPLEMENTO", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OCRD", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OPDN.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OPDN", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_PDN1.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_PDN1", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OSCN.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs_OSCN", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_PPINDICE.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_PPINDICE", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_TICKET.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_TICKET", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If
    End Sub
    Private Sub Cambiar_Nombre_Propiedades()
        Dim sSQL As String = ""

        If objGlobal.refDi.comunes.esAdministrador Then
            sSQL = "UPDATE ""OITG"" SET ""ItmsGrpNam""='Precio Provisional' WHERE ""ItmsTypCod""=6"
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = False Then
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar la propiedad 6 del artículo como ""Precio Provisional"" ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la propiedad 6 del artículo como ""Precio Provisional"" ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

            sSQL = "UPDATE ""OCQG"" SET ""GroupName""='Cerrar Albarán de Proveedor' WHERE ""GroupCode""=6"
            If objGlobal.refDi.SQL.executeNonQuery(sSQL) = False Then
                objGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar la propiedad 6 del Interlocutor como ""Cerrar Albarán de Proveedor"" ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                objGlobal.SBOApp.StatusBar.SetText("Se ha actualizado la propiedad 6 del Interlocutor como ""Cerrar Albarán de Proveedor"" ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        End If
    End Sub
    Private Sub ParametrizacionGeneral()

        If Not objGlobal.refDi.OGEN.existeVariable("Tarifa_Compra") Then
            objGlobal.refDi.OGEN.fijarValorVariable("Tarifa_Compra", "Lista de precios 01")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Tarifa_Compra"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        If Not objGlobal.refDi.OGEN.existeVariable("RPT_PED_COMPRAS") Then
            objGlobal.refDi.OGEN.fijarValorVariable("RPT_PED_COMPRAS", "PEDIDOCOMPRAS.rpt")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Report de Pedido de compras"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_MAIL") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_MAIL", "compras.mprimas@lingotes.com")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Mail para envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_PRIORIDAD") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_PRIORIDAD", "2")
            '0 --> Prioridad Normal
            '1 --> Prioridad Baja
            '2 --> Prioridad alta
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Prioridad de Envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_NOTIFICACION") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_NOTIFICACION", "2")
            '0 --> Sin notificación
            '1 --> Si la entrega es correcta
            '2 --> Si la entrega falla
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Prioridad de Envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_HOST") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_HOST", "smtp.outlook.com")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""HOST para envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_PORT") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_PORT", "587")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""PORT para envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_MAIL_US") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_MAIL_US", "compras.mprimas@lingotes.com")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Usuario Mail para envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
        If Not objGlobal.refDi.OGEN.existeVariable("ENV_MAIL_PASS") Then
            objGlobal.refDi.OGEN.fijarValorVariable("ENV_MAIL_PASS", "Woy58956")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""Password Mail para envío"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        If Not objGlobal.refDi.OGEN.existeVariable("RPT_TICKET") Then
            objGlobal.refDi.OGEN.fijarValorVariable("RPT_TICKET", "PDN10001")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""RPT_TICKET"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        If Not objGlobal.refDi.OGEN.existeVariable("DIR_CTRL_REF") Then
            objGlobal.refDi.OGEN.fijarValorVariable("DIR_CTRL_REF", "\\xper-rdpdes02\compartidaB1\Lingotes\Ficheros\08.Historico\CTRL_REF\")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""DIR_CTRL_REF"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        If Not objGlobal.refDi.OGEN.existeVariable("DIR_PESOS") Then
            objGlobal.refDi.OGEN.fijarValorVariable("DIR_PESOS", "\\xper-rdpdes02\compartidaB1\Lingotes\Ficheros\08.Historico\PESOS")
            objGlobal.SBOApp.StatusBar.SetText("Creado Variable ""DIR_PESOS"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub
    Private Sub CargaFirma()
        Dim path As String = objGlobal.refDi.OGEN.pathDLL & "\08.Historico\"
        If System.IO.Directory.Exists(path) = False Then
            System.IO.Directory.CreateDirectory(path)
        End If
        If objGlobal.refDi.comunes.esAdministrador Then
            EXO_GLOBALES.CopiarRecurso(Reflection.Assembly.GetExecutingAssembly(), "mail.htm", path & "\mail.htm")
        End If
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        'objGlobal.SBOApp.StatusBar.SetText(res, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Public Overrides Function filtros() As Global.SAPbouiCOM.EventFilters
        Dim fXML As String = ""
        Try
            fXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml")
            Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
            filtro.LoadFromXML(fXML)
            Return filtro
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        Finally

        End Try
    End Function

    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "993"
                    Clase = New EXO_993(objGlobal)
                    Return CType(Clase, EXO_993).SBOApp_ItemEvent(infoEvento)
                Case "UDO_FT_EXO_SUPLEMENTO"
                    Clase = New EXO_SUPLEMENTO(objGlobal)
                    Return CType(Clase, EXO_SUPLEMENTO).SBOApp_ItemEvent(infoEvento)
                Case "EXO_HOMO"
                    Clase = New EXO_HOMO(objGlobal)
                    Return CType(Clase, EXO_HOMO).SBOApp_ItemEvent(infoEvento)
                Case "UDO_FT_EXO_PPINDICE"
                    Clase = New EXO_PPINDICE(objGlobal)
                    Return CType(Clase, EXO_PPINDICE).SBOApp_ItemEvent(infoEvento)
                Case "EXO_CPPIND"
                    Clase = New EXO_CPPIND(objGlobal)
                    Return CType(Clase, EXO_CPPIND).SBOApp_ItemEvent(infoEvento)
                Case "142"
                    Clase = New EXO_OPOR(objGlobal)
                    Return CType(Clase, EXO_OPOR).SBOApp_ItemEvent(infoEvento)
                Case "143"
                    Clase = New EXO_143(objGlobal)
                    Return CType(Clase, EXO_143).SBOApp_ItemEvent(infoEvento)
                Case "UDO_FT_EXO_TICKET"
                    Clase = New EXO_TICKET(objGlobal)
                    Return CType(Clase, EXO_TICKET).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(infoEvento As BusinessObjectInfo) As Boolean
        Dim Res As Boolean = True
        Dim Clase As Object = Nothing
        Try
            Select Case infoEvento.FormTypeEx
                Case "UDO_FT_EXO_SUPLEMENTO"
                    Clase = New EXO_SUPLEMENTO(objGlobal)
                    Return CType(Clase, EXO_SUPLEMENTO).SBOApp_FormDataEvent(infoEvento)
                Case "143"
                    Clase = New EXO_143(objGlobal)
                    Return CType(Clase, EXO_143).SBOApp_FormDataEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try

    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim Clase As Object = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case ""
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnHoIC"
                        Clase = New EXO_HOMO(objGlobal)
                        Return CType(Clase, EXO_HOMO).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnPPerIn"
                        Clase = New EXO_PPINDICE(objGlobal)
                        Return CType(Clase, EXO_PPINDICE).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnRPRE"
                        Clase = New EXO_CPPIND(objGlobal)
                        Return CType(Clase, EXO_CPPIND).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnSTKT"
                        Clase = New EXO_TICKET(objGlobal)
                        Return CType(Clase, EXO_TICKET).SBOApp_MenuEvent(infoEvento)
                    Case "EXO_MNPESOE"
                        Clase = New EXO_143(objGlobal)
                        Return CType(Clase, EXO_143).SBOApp_MenuEvent(infoEvento)
                    Case "EXO_MNPESOS"
                        Clase = New EXO_143(objGlobal)
                        Return CType(Clase, EXO_143).SBOApp_MenuEvent(infoEvento)
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
    Public Overrides Function SBOApp_RightClickEvent(infoEvento As ContextMenuInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim Clase As Object = Nothing

        Try

            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            Select Case oForm.TypeEx
                Case "143"
                    Clase = New EXO_143(objGlobal)
                    Return CType(Clase, EXO_143).SBOApp_RightClickEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_RightClickEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            If objGlobal.SBOApp.ClientType = BoClientType.ct_Desktop Then
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            End If
            Clase = Nothing
        End Try
    End Function
End Class
