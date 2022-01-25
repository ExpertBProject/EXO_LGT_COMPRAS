Imports System.IO
Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_143
    Private objGlobal As EXO_UIAPI.EXO_UIAPI
#Region "Variables"
    Private Shared _iLineNumRightClick As Integer = -1
#End Region
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
                        Case "143"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_VALIDATE_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "143"
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
                        Case "143"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "143"
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
    Private Function EventHandler_VALIDATE_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_VALIDATE_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "38" Then
                If (pVal.ColUID = "U_EXO_PESO_E" Or pVal.ColUID = "U_EXO_PESO_S") And pVal.ItemChanged = True Then
                    Dim dPesoEntrada As Double = 0 : Dim dPesoSalida As Double = 0
                    dPesoEntrada = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_PESO_E").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)
                    dPesoSalida = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_PESO_S").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.ToString)
                    Dim dPeso As Double = dPesoEntrada - dPesoSalida
                    If dPeso > 0 Then
                        Try
                            CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("11").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dPeso, EXO_GLOBALES.FuenteInformacion.Otros)
                        Catch ex As Exception
                            objGlobal.SBOApp.StatusBar.SetText("La cantidad no puede ser actualizada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If
                End If
            End If
            EventHandler_VALIDATE_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_RightClickEvent(ByVal infoEvento As ContextMenuInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

            If infoEvento.BeforeAction = False Then
                Select Case oForm.TypeEx
                    Case "143"
                        If objGlobal.SBOApp.Menus.Exists("EXO_mnsep") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO_mnsep")
                        End If

                        If objGlobal.SBOApp.Menus.Exists("EXO_MNPESOE") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO_MNPESOE")
                        End If
                        If objGlobal.SBOApp.Menus.Exists("EXO_MNPESOS") Then
                            objGlobal.SBOApp.Menus.RemoveEx("EXO_MNPESOS")
                        End If
                End Select
            Else
                Select Case oForm.TypeEx
                    Case "143"
                        'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If infoEvento.ItemUID = "38" Then
                                If infoEvento.Row > 0 Then
                                    _iLineNumRightClick = infoEvento.Row
                                End If
                                oCreationPackage = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams), SAPbouiCOM.MenuCreationParams)
                                Dim oMenuItem As SAPbouiCOM.MenuItem = objGlobal.SBOApp.Menus.Item("1280") 'Data'
                                Dim oMenus As SAPbouiCOM.Menus = oMenuItem.SubMenus
                                If Not objGlobal.SBOApp.Menus.Exists("EXO_mnsep") Then
                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_SEPERATOR
                                    oCreationPackage.Position = oMenuItem.SubMenus.Count + 1
                                    oCreationPackage.UniqueID = "EXO_mnsep"
                                    oCreationPackage.Enabled = True
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)
                                End If

                                If Not objGlobal.SBOApp.Menus.Exists("EXO_MNPESOE") Then
                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.Position = oMenuItem.SubMenus.Count + 1
                                    oCreationPackage.UniqueID = "EXO_MNPESOE"
                                    oCreationPackage.String = "Peso de Entrada"
                                    oCreationPackage.Enabled = True
                                    oMenus = oMenuItem.SubMenus
                                    oMenus.AddEx(oCreationPackage)
                                End If

                                If Not objGlobal.SBOApp.Menus.Exists("EXO_MNPESOS") Then
                                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                    oCreationPackage.Position = oMenuItem.SubMenus.Count + 1
                                    oCreationPackage.UniqueID = "EXO_MNPESOS"
                                    oCreationPackage.String = "Peso de Salida"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            End If
                        ' End If
                End Select
            End If

            Return True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim sMensaje As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.ActiveForm
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO_MNPESOE"
                        If _iLineNumRightClick > 0 Then
                            LeerFichero("PE", oForm)
                        Else
                            sMensaje = "Tiene que seleccionar una línea."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                        End If

                    Case "EXO_MNPESOS"
                        If _iLineNumRightClick > 0 Then
                            LeerFichero("PS", oForm)
                        Else
                            sMensaje = "Tiene que seleccionar una línea."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox(sMensaje)
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
                                    CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_PESO_E").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = scampos(1).Replace(",", ".")
                                    Dim sHora As String = ""
                                    Try
                                        Dim sTiempo As String() = scampos(2).Split(CChar(":"))
                                        For Each item As String In sTiempo
                                            sHora &= item.ToString
                                        Next
                                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_HORA_E").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = sHora
                                    Catch ex As Exception
                                        sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_HORA_E").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = sHora
                                    End Try
                                Case "PS"
                                    CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_PESO_S").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = scampos(1).Replace(",", ".")
                                    Dim sHora As String = ""
                                    Try
                                        Dim sTiempo As String() = scampos(2).Split(CChar(":"))
                                        For Each item As String In sTiempo
                                            sHora &= item.ToString
                                        Next
                                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_HORA_S").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = sHora
                                    Catch ex As Exception
                                        sHora = Now.Hour.ToString("00") & Now.Minute.ToString("00")
                                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_HORA_S").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = sHora
                                    End Try
                            End Select
                            Dim dPesoEntrada As Double = 0 : Dim dPesoSalida As Double = 0
                            dPesoEntrada = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_PESO_E").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value.ToString)
                            dPesoSalida = EXO_GLOBALES.DblTextToNumber(objGlobal.compañia, CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_PESO_S").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value.ToString)
                            Dim dPeso As Double = dPesoEntrada - dPesoSalida
                            If dPeso > 0 Then
                                CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("11").Cells.Item(_iLineNumRightClick).Specific, SAPbouiCOM.EditText).Value = EXO_GLOBALES.DblNumberToText(objGlobal.compañia, dPeso, EXO_GLOBALES.FuenteInformacion.Otros)
                            End If

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
            'objGlobal.SBOApp.MessageBox("Se ha leido correctamente el fichero. Fin del proceso")
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
    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sDocEntry As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "143"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                        End Select
                End Select
            Else
                Select Case infoEvento.FormTypeEx
                    Case "143"
                        Select Case infoEvento.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess = True Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText

                                    If ControldeFrecuencia(sDocEntry) = False Then
                                        Return False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess = True Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText

                                    If ControldeFrecuencia(sDocEntry) = False Then
                                        Return False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select
                End Select
            End If

            Return True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function ControldeFrecuencia(ByVal sDocEntry As String) As Boolean
#Region "Variables"
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsCuenta As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sHomo As String = ""
        Dim sMensaje As String = ""
        Dim sCardCode As String = "" : Dim sItemCode As String = "" : Dim sCatalogo As String = ""
        Dim sRef As String = ""
        Dim sNomFich As String = "" : Dim sRutaFich As String = "" : Dim sLinea As String = ""
        Dim sACTFRECUENCIA As String = "N"
#End Region
        ControldeFrecuencia = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT C.""CardCode"", IC.""U_EXO_HOMO"", L.* FROM ""PDN1"" L  "
            sSQL &= " INNER JOIN ""OPDN"" C ON L.""DocEntry""=C.""DocEntry"" "
            sSQL &= " INNER JOIN ""OCRD"" IC ON IC.""CardCode""=C.""CardCode"" "
            sSQL &= " WHERE C.""DocEntry""=" & sDocEntry & " and L.""U_EXO_CRTLF""='N'"
            oRs.DoQuery(sSQL)
            For i = 0 To oRs.RecordCount - 1
                sHomo = oRs.Fields.Item("U_EXO_HOMO").Value.ToString.Trim
                sCardCode = oRs.Fields.Item("CardCode").Value.ToString.Trim
                sItemCode = oRs.Fields.Item("ItemCode").Value.ToString.Trim
                sCatalogo = oRs.Fields.Item("SubCatNum").Value.ToString.Trim

                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Homo: " & sHomo & " - IC: " & sCardCode & " - Art: " & sItemCode & " - Cat: " & sCatalogo, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                sSQL = "SELECT ""U_EXO_FREC"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' and ""Substitute""='" & sCatalogo & "'"
                sACTFRECUENCIA = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                If sACTFRECUENCIA = "Y" Then
                    If sHomo = "-" Then
                        sMensaje = "El proveedor no tiene asignado una Homologación para calcular el control de frecuencias. Por favor, revise los datos."
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox(sMensaje)
                        Exit Function
                    Else

#Region "Buscamos el valor de Referencia"
                        Select Case sHomo
                            Case "A" : sSQL = "SELECT ""U_EXO_REFA"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' and ""Substitute""='" & sCatalogo & "'"
                            Case "B" : sSQL = "SELECT ""U_EXO_REFB"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' and ""Substitute""='" & sCatalogo & "'"
                            Case "C" : sSQL = "SELECT ""U_EXO_REFC"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' and ""Substitute""='" & sCatalogo & "'"
                        End Select
                        sRef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sRef.Trim = "" Then
                            Select Case sHomo
                                Case "A" : sSQL = "SELECT TOP 1 ""U_EXO_REFA"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' "
                                Case "B" : sSQL = "SELECT TOP 1 ""U_EXO_REFB"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' "
                                Case "C" : sSQL = "SELECT TOP 1 ""U_EXO_REFC"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' "
                            End Select
                            sRef = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                            If sRef.Trim = "" Then
                                sMensaje = "No se encuentra el catálogo para el artículo " & sItemCode & " y el proveedor " & sCardCode & ", revise los datos para calcular el control de frecuencias."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Function
                            End If
                        End If
#End Region
                        If sRef.Trim <> "" Then
                            Dim iRef As Integer = CType(sRef, Integer)
                            Dim iCuenta As Integer = 0

                            Dim sPath As String = objGlobal.funcionesUI.refDi.OGEN.valorVariable("DIR_CTRL_REF")
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Creando Documento de frecuencia en " & sPath, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            If IO.Directory.Exists(sPath) = False Then
                                IO.Directory.CreateDirectory(sPath)
                            End If

                            oRsCuenta = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                            sSQL = "SELECT COUNT(*) FROM "
                            sSQL &= " (SELECT DISTINCT C.""DocEntry"" FROM ""OPDN"" C "
                            sSQL &= " INNER JOIN ""PDN1"" L ON C.""DocEntry"" =L.""DocEntry"" "
                            sSQL &= " WHERE C.""DocEntry""<>" & sDocEntry & " And (C.""CardCode""='" & sCardCode & "' and L.""ItemCode""='" & sItemCode & "' "
                            sSQL &= " And (L.""SubCatNum""='' or L.""SubCatNum""='" & sCatalogo & "') and L.""U_EXO_CRTLF""='N') ) "
                            iCuenta = CType(objGlobal.refDi.SQL.sqlNumericaB1(sSQL), Integer)
                            If (iRef <= iCuenta + 1) Then
                                sMensaje = "Ref: " & sRef.ToString.Trim & " y existe(n) " & iCuenta.ToString.Trim & " recepciones. Se crea fichero y se guarda en el directorio " & sPath.ToString.Trim
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                'Guardamos CSV a la carpeta indicada
#Region "Guardamos CSV a la carpeta indicada"
                                sNomFich = "CNTRL_FR_RECEP_" & sCardCode & "_" & sItemCode & "_" & sCatalogo
                                sRutaFich = Path.Combine(sPath & sNomFich & ".csv")
                                If IO.File.Exists(sRutaFich) = False Then
                                    IO.File.Delete(sRutaFich)
                                End If
                                FileOpen(1, sRutaFich, OpenMode.Output)
                                sMensaje = "Generando fichero - " & sRutaFich
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                sLinea = sCardCode & ";" & sItemCode
                                PrintLine(1, sLinea)
                                FileClose(1)
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fichero Creado...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                                'Actualizamos las líneas de las recepciones
#Region "Actualizamos línea de la recepción actual "
                                sSQL = "UPDATE ""PDN1"" SET ""U_EXO_CRTLF""='Y' "
                                sSQL &= " WHERE ""DocEntry""=" & sDocEntry
                                sSQL &= "  and ""LineNum""=" & oRs.Fields.Item("LineNum").Value.ToString.Trim
                                If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                    sMensaje = "Se ha actualizado la recepción con Nº Interno " & sDocEntry
                                    sMensaje &= " y línea Nº" & oRs.Fields.Item("LineNum").ToString.Trim
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Else
                                    sMensaje = "No se ha podido actualizar la recepción con Nº Interno " & sDocEntry
                                    sMensaje &= " y línea Nº" & oRs.Fields.Item("LineNum").ToString.Trim
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
#End Region
#Region "Actualizamos las líneas de las recepciones"
                                sSQL = "SELECT DISTINCT C.""DocEntry"", L.""LineNum"" "
                                sSQL &= " FROM ""OPDN"" C "
                                sSQL &= " INNER JOIN ""PDN1"" L ON C.""DocEntry"" =L.""DocEntry"" "
                                sSQL &= " WHERE C.""DocEntry""<>" & sDocEntry & " And (C.""CardCode""='" & sCardCode & "' and L.""ItemCode""='" & sItemCode & "' "
                                sSQL &= " And (L.""SubCatNum""='' or L.""SubCatNum""='" & sCatalogo & "') and L.""U_EXO_CRTLF""='N') "
                                oRsCuenta.DoQuery(sSQL)
                                For a = 0 To oRsCuenta.RecordCount - 1
                                    sSQL = "UPDATE ""PDN1"" SET ""U_EXO_CRTLF""='Y' "
                                    sSQL &= " WHERE ""DocEntry""=" & oRsCuenta.Fields.Item("DocEntry").Value.ToString.Trim
                                    sSQL &= "  and ""LineNum""=" & oRsCuenta.Fields.Item("LineNum").Value.ToString.Trim
                                    If objGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                                        sMensaje = "Se ha actualizado la recepción con Nº Interno " & oRsCuenta.Fields.Item("DocEntry").Value.ToString.Trim
                                        sMensaje &= " y línea Nº" & oRsCuenta.Fields.Item("LineNum").Value.ToString.Trim
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Else
                                        sMensaje = "No se ha podido actualizar la recepción con Nº Interno " & oRsCuenta.Fields.Item("DocEntry").Value.ToString.Trim
                                        sMensaje &= " y línea Nº" & oRsCuenta.Fields.Item("LineNum").Value.ToString.Trim
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                    oRsCuenta.MoveNext()
                                Next
#End Region
                            Else
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra datos para generar el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        Else
                            sMensaje = "No se encuentra el catálogo para el artículo " & sItemCode & " y el proveedor " & sItemCode & ", revise los datos para calcular el control de frecuencias."
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objGlobal.SBOApp.MessageBox(sMensaje)
                            Exit Function
                        End If
                    End If
                End If
                oRs.MoveNext()
            Next
            ControldeFrecuencia = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
            Return False
        Catch ex As Exception
            Throw ex
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCuenta, Object))
        End Try
    End Function
End Class
