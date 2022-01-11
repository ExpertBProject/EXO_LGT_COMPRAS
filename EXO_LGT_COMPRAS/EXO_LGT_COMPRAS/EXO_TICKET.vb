Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_TICKET
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean

        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnSTKT"
                        If CargarUDO() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally

        End Try
    End Function
    Public Function CargarUDO() As Boolean
        CargarUDO = False

        Try
            objGlobal.funcionesUI.cargaFormUdoBD("EXO_TICKET")

            CargarUDO = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
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
                        Case "UDO_FT_EXO_TICKET"
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
                        Case "UDO_FT_EXO_TICKET"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TICKET"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_TICKET"
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
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btnPE"
                    LeerFichero("PE", oForm)
                Case "btnPS"
                    LeerFichero("PS", oForm)
                Case "btnImp"
                    Imprimir(oForm)
            End Select

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub LeerFichero(ByVal sBoton As String, ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sArchivo As String = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\DOC_CARGADOS\" & objGlobal.compañia.CompanyDB & "\COMPRAS\PESOS\"
        Dim sTipoArchivo As String = "Ficheros CSV|*.csv|Texto|*.txt"
        Dim sArchivoOrigen As String = objGlobal.funcionesUI.refDi.OGEN.valorVariable("DIR_PESOS") & "\Peso.csv"
        Dim sNomFICH As String = ""
#End Region
        Try
            If System.IO.Directory.Exists(sArchivo) = False Then
                System.IO.Directory.CreateDirectory(sArchivo)
            End If

            ''Tenemos que controlar que es cliente o web
            'If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
            '    sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
            'Else
            '    'Controlar el tipo de fichero que vamos a abrir según campo de formato
            '    sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
            'End If
            If IO.File.Exists(sArchivoOrigen) = False Then
                objGlobal.SBOApp.MessageBox("No existe fichero """ & sArchivoOrigen & """ a importar.")
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No existe fichero """ & sArchivoOrigen & """ a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                Exit Sub
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fichero: " & sArchivoOrigen, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                sArchivo = sArchivo & sNomFICH
                'Hacemos copia de seguridad para tratarlo
                EXO_GLOBALES.Copia_Seguridad(sArchivoOrigen, sArchivo, objGlobal)
                'Ahora abrimos el fichero para tratarlo
                TratarFichero(sArchivo, sBoton, oForm)
                IO.File.Delete(sArchivoOrigen)
            End If

        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub
    Private Sub TratarFichero(ByVal sArchivo As String, ByVal sBoton As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sDelimitador As String = "2"
        Try
            objGlobal.SBOApp.StatusBar.SetText("Cargando datos...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)


#Region "TXT|CSV"
            If File.Exists(sArchivo) Then
                Using MyReader As New Microsoft.VisualBasic.
                    FileIO.TextFieldParser(sArchivo, System.Text.Encoding.UTF7)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    Select Case sDelimitador
                        Case "1" : MyReader.SetDelimiters(vbTab)
                        Case "2" : MyReader.SetDelimiters(";")
                        Case "3" : MyReader.SetDelimiters(",")
                        Case "4" : MyReader.SetDelimiters("-")
                        Case Else : MyReader.SetDelimiters(vbTab)
                    End Select

                    Dim currentRow As String()
                    Dim bPrimeraLinea As Boolean = True

                    While Not MyReader.EndOfData
                        Try
                            'If bPrimeraLinea = True Then
                            '    currentRow = MyReader.ReadFields() : currentRow = MyReader.ReadFields()
                            '    bPrimeraLinea = False
                            'Else
                            currentRow = MyReader.ReadFields()
                            'End If

                            Dim currentField As String
                            Dim scampos(1) As String
                            Dim iCampo As Integer = 0
                            For Each currentField In currentRow
                                iCampo += 1
                                ReDim Preserve scampos(iCampo)
                                scampos(iCampo) = currentField
                                'SboApp.MessageBox(scampos(iCampo))
                            Next


                            'Grabamos la línea
                            Select Case sBoton
                                Case "PE"
                                    oForm.DataSources.DBDataSources.Item("@EXO_TICKET").SetValue("U_EXO_PESO_E", 0, scampos(1).Replace(",", "."))
                                    Dim sHora As String = ""
                                    Try
                                        Dim sTiempo As String() = scampos(2).Split(CChar(":"))
                                        For Each item As String In sTiempo
                                            sHora &= item.ToString
                                        Next
                                        oForm.DataSources.DBDataSources.Item("@EXO_TICKET").SetValue("U_EXO_HORA_E", 0, sHora)
                                    Catch ex As Exception
                                        sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                                        oForm.DataSources.DBDataSources.Item("@EXO_TICKET").SetValue("U_EXO_HORA_E", 0, sHora)
                                    End Try
                                Case "PS"
                                    oForm.DataSources.DBDataSources.Item("@EXO_TICKET").SetValue("U_EXO_PESO_S", 0, scampos(1).Replace(",", "."))
                                    Dim sHora As String = ""
                                    Try
                                        Dim sTiempo As String() = scampos(2).Split(CChar(":"))
                                        For Each item As String In sTiempo
                                            sHora &= item.ToString
                                        Next
                                        oForm.DataSources.DBDataSources.Item("@EXO_TICKET").SetValue("U_EXO_HORA_S", 0, sHora)
                                    Catch ex As Exception
                                        sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                                        oForm.DataSources.DBDataSources.Item("@EXO_TICKET").SetValue("U_EXO_HORA_S", 0, sHora)
                                    End Try
                            End Select

                        Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Línea " & ex.Message & " no es válida y se omitirá.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objGlobal.SBOApp.MessageBox("Línea " & ex.Message & " no es válida y se omitirá.")
                        End Try
                    End While
                End Using
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
#End Region

            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se ha leido correctamente el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Private Sub Imprimir(ByRef oForm As SAPbouiCOM.Form)
        Try
            Dim oCmpSrv As SAPbobsCOM.CompanyService = objGlobal.compañia.GetCompanyService()
            Dim oReportLayoutService As SAPbobsCOM.ReportLayoutsService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService), SAPbobsCOM.ReportLayoutsService)
            Dim oPrintParam As SAPbobsCOM.ReportLayoutPrintParams = CType(oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams), SAPbobsCOM.ReportLayoutPrintParams)
            Dim sLayout As String = objGlobal.refDi.OGEN.valorVariable("RPT_TICKET")
            Dim iDocentry As Integer = CType(oForm.DataSources.DBDataSources.Item("@EXO_TICKET").GetValue("DocNum", 0).Trim, Integer)
            oPrintParam.LayoutCode = sLayout 'codigo del formato importado en SAP
            oPrintParam.DocEntry = iDocentry 'parametro que se envia al crystal, DocEntry de la transaccion

            oReportLayoutService.Print(oPrintParam)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class
