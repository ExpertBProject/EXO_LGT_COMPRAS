Imports SAPbouiCOM
Public Class EXO_CPPIND
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
                    Case "EXO-MnRPRE"
                        If CargarForm() = False Then
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
    Public Function CargarForm() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_UIAPI.EXO_XML(objGlobal)

        CargarForm = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CPPIND.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            'sSQL = "SELECT DISTINCT ""Category"" ""COD"", ""Category"" ""ANNO"" FROM ""OFPR"" order by ""Category"" "
            'objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            'CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueOnly
            'CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).Select(Now.Year.ToString, BoSearchKey.psk_ByValue)
            oForm.Items.Item("btn_Cal").Enabled = False
            CargarForm = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
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
                        Case "EXO_CPPIND"
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
                        Case "EXO_CPPIND"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                    If EventHandler_et_DOUBLE_CLICK_Before(infoEvento) = False Then
                                        Return False
                                    End If
                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CPPIND"
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
                        Case "EXO_CPPIND"
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
    Private Function EventHandler_et_DOUBLE_CLICK_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_et_DOUBLE_CLICK_Before = False

        Try
            If pVal.ActionSuccess = False And pVal.ColUID = "Sel" Then
                oForm.Freeze(True)
                For iRow = 0 To oForm.DataSources.DataTables.Item("DT_DOC").Rows.Count - 1
                    If oForm.DataSources.DataTables.Item("DT_DOC").GetValue("Sel", iRow).ToString = "Y" Then
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Sel", iRow, "N")
                    Else
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Sel", iRow, "Y")
                    End If
                Next
            End If

            EventHandler_et_DOUBLE_CLICK_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByVal pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btn_Cal" Then
                If pVal.ActionSuccess = True Then
                    If objGlobal.SBOApp.MessageBox("¿Calculamos Precios Fin de Mes?", 1, "Sí", "No") = 1 Then
                        If ComprobarDOC(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btn_Cal").Enabled = False
                            'Calculando datos
                            objGlobal.SBOApp.StatusBar.SetText("Calculando datos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If Calcular_Precio(oForm, "DT_DOC", objGlobal) = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                        End If
                    End If
                End If
            ElseIf pVal.ItemUID = "btn_Fich" Then

                Cargar_Grid(oForm)
                oForm.Items.Item("btn_Cal").Enabled = True
            End If

            EventHandler_ItemPressed_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Sub Cargar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sDFecha As String = "" : Dim sHFecha As String = ""
        Try
            oForm.Freeze(True)
            sDFecha = CType(oForm.Items.Item("txtDFceha").Specific, SAPbouiCOM.EditText).Value.ToString
            sHFecha = CType(oForm.Items.Item("txtHFceha").Specific, SAPbouiCOM.EditText).Value.ToString
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"", '     ' as ""Estado"", CAB.""DocEntry"" ""Nº Interno"", CAB.""DocNum"" ""Nº Documento"", CAB.""CardCode"" ""Código"", CAB.""CardName"" ""Nombre"", CAB.""DocDate"" ""Fecha Contable"" , "
            sSQL &= " LIN.""LineNum"" ""Línea"", LIN.""ItemCode"" ""Cód. Artículo"",LIN.""Dscription"" ""Artículo"",LIN.""Quantity"" ""Cantidad"",LIN.""Price"" ""Precio"", LIN.""SubCatNum"" ""Catálogo"", CAB.""U_EXO_CERRAR"" ""Cerrar Doc."" "
            sSQL &= ",  IC.""QryGroup6""  ""Cerrar Doc. IC"" , ITM.""QryGroup7"" ""Buscar Precio Tarifa"", CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " FROM ""OPDN"" CAB "
            sSQL &= " INNER JOIN ""PDN1"" LIN On CAB.""DocEntry""=LIN.""DocEntry"" "
            sSQL &= " INNER JOIN ""OITM"" ITM On LIN.""ItemCode""=ITM.""ItemCode"" "
            sSQL &= " INNER JOIN ""OCRD"" IC ON IC.""CardCode""=CAB.""CardCode"" "
            sSQL &= "WHERE CAB.""DocStatus""<>'C' and CAB.""CANCELED""='N' "
            If sDFecha.Trim <> "" Then
                sSQL &= " And CAB.""DocDate"">='" & sDFecha & "' "
            End If
            If sHFecha.Trim <> "" Then
                sSQL &= " And CAB.""DocDate""<='" & sHFecha & "' "
            End If
            sSQL &= " And ITM.""QryGroup6""='Y' "
            sSQL &= " ORDER BY CAB.""CardCode"", CAB.""DocNum"",LIN.""LineNum"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
            objGlobal.SBOApp.StatusBar.SetText("Datos Cargados con éxito.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            oForm.Items.Item("btn_Cal").Enabled = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 16
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                If i = 2 Then
                    oColumnTxt.LinkedObjectType = "20"
                ElseIf i = 4 Then
                    oColumnTxt.LinkedObjectType = "2"
                ElseIf i = 8 Then
                    oColumnTxt.LinkedObjectType = "4"
                ElseIf i = 7 Or i = 10 Or i = 11 Then
                    oColumnTxt.RightJustified = True
                End If
            Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(15).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function Calcular_Precio(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Calcular_Precio = False
#Region "VARIABLES"
        Dim sDocEntry As String = "" : Dim sLineNum As String = "" : Dim sDocNum As String = ""
        Dim sCardCode As String = ""
        Dim sItemCode As String = ""
        Dim sCatalogo As String = ""
        Dim oDoc As SAPbobsCOM.Documents = Nothing : Dim sMensaje As String = ""
        Dim sSQL As String = ""
        Dim sIndice As String = ""
        Dim dFecha As Date = Now : Dim sFecha As String = "" : Dim sMes As String = "" : Dim sAnno As String = ""
        Dim dPrecioIndice As Double = 0 : Dim dPrecioSuplemento As Double = 0 : Dim dPrecio As Double = 0
        Dim sCodTarifa As String = "" : Dim sNomTarifa As String = ""
        Dim sBuscarPrecioTarifa As String = "N"
        Dim sCerrarDoc As String = "" : Dim sCerrarDocIC As String = ""
        Dim sCerrar() As String = {"", ""}

        Dim oOSPP As SAPbobsCOM.SpecialPrices = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim iContPrecios As Integer = 0
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
#End Region
        Try
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                oDoc = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes), SAPbobsCOM.Documents)
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    If sCerrar(0) = "" Then
                        sCerrar(0) = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Interno", i).ToString()
                        sCerrar(1) = ""
                    Else
                        If sCerrar(0) <> oForm.DataSources.DataTables.Item(sData).GetValue("Nº Interno", i).ToString() Then
                            If sCerrar(1) = "OK" Then
                                If CType(oForm.Items.Item("chkCerrar").Specific, SAPbouiCOM.CheckBox).Checked Then
                                    If sCerrarDoc = "Y" And sCerrarDocIC = "Y" Then
                                        oDoc.GetByKey(CType(sCerrar(0), Integer))
                                        If oDoc.Close() <> 0 Then
                                            sMensaje = oobjGlobal.compañia.GetLastErrorCode.ToString & " / " & oobjGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                                            oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Else
                                            sMensaje = "Se ha Cerrado correctamente el documento Nº " & sDocNum & " y Nº interno " & sDocEntry
                                            oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        End If
                                    End If
                                End If
                            End If
                            sCerrar(0) = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Interno", i).ToString()
                            sCerrar(1) = ""
                        End If
                    End If
                    sDocEntry = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Interno", i).ToString()
                    sDocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString()
                    sLineNum = oForm.DataSources.DataTables.Item(sData).GetValue("Línea", i).ToString()
                    sCardCode = oForm.DataSources.DataTables.Item(sData).GetValue("Código", i).ToString()
                    sItemCode = oForm.DataSources.DataTables.Item(sData).GetValue("Cód. Artículo", i).ToString()
                    sCatalogo = oForm.DataSources.DataTables.Item(sData).GetValue("Catálogo", i).ToString()
                    sFecha = oForm.DataSources.DataTables.Item(sData).GetValue("Fecha Contable", i).ToString()
                    sCerrarDoc = oForm.DataSources.DataTables.Item(sData).GetValue("Cerrar Doc.", i).ToString()
                    sCerrarDocIC = oForm.DataSources.DataTables.Item(sData).GetValue("Cerrar Doc. IC", i).ToString()
                    sBuscarPrecioTarifa = oForm.DataSources.DataTables.Item(sData).GetValue("Buscar Precio Tarifa", i).ToString()
                    If oDoc.GetByKey(CType(sDocEntry, Integer)) = True Then
#Region "Buscamos el precio"
                        'Buscamos el índice del Catálogo
                        sSQL = "SELECT ""U_EXO_INDI"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' and ""Substitute""='" & sCatalogo & "'"
                        sIndice = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                        If sIndice.Trim = "" Then
                            sSQL = "SELECT TOP 1 ""U_EXO_INDI"" FROM ""OSCN"" WHERE ""CardCode""='" & sCardCode & "' and ""ItemCode""='" & sItemCode & "' "
                            sIndice = oobjGlobal.refDi.SQL.sqlStringB1(sSQL)
                            If sIndice.Trim = "" Then
                                oobjGlobal.SBOApp.StatusBar.SetText("No se encuentra el catálogo para el artículo " & sItemCode & " y el proveedor " & sItemCode & ", revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                        If sBuscarPrecioTarifa = "N" Then
                            If sIndice.Trim <> "" Then
                                dFecha = CDate(sFecha)
                                sMes = dFecha.Month.ToString("00")
                                sAnno = dFecha.Year.ToString("0000")
#Region "Precio Indice"
                                Select Case sIndice
                                    Case "1" : sSQL = "SELECT TOP 1 ""U_EXO_I1"" FROM ""@EXO_PPINDICEL"" WHERE ""Code""='" & sAnno & "' and ""U_EXO_MES""<='" & sMes & "' order by ""U_EXO_MES"" desc"
                                    Case "2" : sSQL = "SELECT TOP 1 ""U_EXO_I2"" FROM ""@EXO_PPINDICEL"" WHERE ""Code""='" & sAnno & "' and ""U_EXO_MES""<='" & sMes & "' order by ""U_EXO_MES"" desc"
                                    Case "3" : sSQL = "SELECT TOP 1 ""U_EXO_I3"" FROM ""@EXO_PPINDICEL"" WHERE ""Code""='" & sAnno & "' and ""U_EXO_MES""<='" & sMes & "' order by ""U_EXO_MES"" desc"
                                    Case "4" : sSQL = "SELECT TOP 1 ""U_EXO_I4"" FROM ""@EXO_PPINDICEL"" WHERE ""Code""='" & sAnno & "' and ""U_EXO_MES""<='" & sMes & "' order by ""U_EXO_MES"" desc"
                                    Case "5" : sSQL = "SELECT TOP 1 ""U_EXO_I5"" FROM ""@EXO_PPINDICEL"" WHERE ""Code""='" & sAnno & "' and ""U_EXO_MES""<='" & sMes & "' order by ""U_EXO_MES"" desc"
                                    Case Else
                                        oobjGlobal.SBOApp.StatusBar.SetText("El índice """ & sIndice & """ para el catálogo para el artículo " & sItemCode & " y el proveedor " & sItemCode & " no es correcto, revise los datos. Se para el proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Function
                                End Select
                                dPrecioIndice = oobjGlobal.refDi.SQL.sqlNumericaB1(sSQL)
#End Region
#Region "Precio Suplemento"
                                sSQL = "SELECT TOP 1 L.""U_EXO_SUP"" FROM ""@EXO_SUPLEMENTOL"" L INNER JOIN ""@EXO_SUPLEMENTO"" C ON L.""Code""=C.""Code"" "
                                sSQL &= " WHERE ""U_EXO_IC""='" & sCardCode & "' and ""U_EXO_ART""='" & sItemCode & "' and ""U_EXO_CAT"" ='" & sCatalogo & "' and ( MONTH(L.""U_EXO_FECHA"")<='" & sMes & "' and YEAR(L.""U_EXO_FECHA"")='" & sAnno & "') "
                                sSQL &= " ORDER BY ""U_EXO_FECHA"" desc"
                                dPrecioSuplemento = oobjGlobal.refDi.SQL.sqlNumericaB1(sSQL)
#End Region
                                dPrecio = dPrecioIndice + dPrecioSuplemento

                            Else
                                oobjGlobal.SBOApp.StatusBar.SetText("No se ha podido Encontrar el índice en el documento Nº" & sDocNum & " y la línea """ & sLineNum & """ , revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            End If
                        Else
#Region "Buscamos el precio de la Tarifa"
                            Dim oRsPrecio As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                            Dim Buscar_Precio As String = ""
                            sSQL = "select t0.""ItemCode"",t0.""ItemName"",T0.""UserText"",COALESCE(t0.""SHeight1"",0) ""SHeight1"", " +
                                        " COALESCE(T0.""SLength1"",0) ""SLength1"" ,COALESCE(T0.""SWidth1"",0) ""SWidth1"", " +
                                        " COALESCE(T0.""SWeight1"",0) ""SWeight1"",  " +
                                        " Case when coalesce(t3.""Price"",0)>0 then (100-t3.""Discount"")*t1.""Price""/100 When coalesce(t2.""Price"",0) >0 Then (100-t2.""Discount"")*t1.""Price""/100 Else t1.""Price"" End As Price, coalesce(A3.""Rate"",0) As IVA " +
                                        " ,t1.""Price"" as PrecioOriginal, case when coalesce(t3.""Discount"",0)>0 then t3.""Discount"" when coalesce(t2.""Discount"",0)>0 then t2.""Discount"" else 0 end as Discount, " +
                                        " case when t0.""UgpEntry""=-1 then 'N' ELSE 'Y' END AS Divisible"
                            sSQL &= " from ""OITM"" t0  inner join ""ITM1"" t1  On t0.""ItemCode""=t1.""ItemCode"" And coalesce(t0.""frozenFor"",'N')='N' and t1.""PriceList""= (select ""ListNum"" from ""OCRD"" where ""CardCode""='" + sCardCode + "') " +
                                        " left join ""OSPP"" t2 on t1.""ItemCode""=t2.""ItemCode"" and t2.""CardCode""='" & sCardCode & "' " +
                                        " LEFT join ""OVTG"" A3  on t0.""VatGourpSa""=A3.""Code"" " +
                                        "LEFT JOIN ""SPP1"" t3  on t3.""ItemCode""=t2.""ItemCode"" and t2.""CardCode""=t3.""CardCode"" and t3.""FromDate""<='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "' " +
                                        " And ifnull(T3.""ToDate"",'20991231')>='" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00") & "'  "
                            sSQL &= " where  (t0.""ItemCode"" like '%" & sItemCode & "%' )"
                            oRsPrecio.DoQuery(sSQL)
                            If oRsPrecio.RecordCount > 0 Then
                                Buscar_Precio = oRsPrecio.Fields.Item("Price").Value.ToString
                            End If
                            If Buscar_Precio.Trim <> "" Then
                                dPrecio = EXO_GLOBALES.DblTextToNumber(oobjGlobal.compañia, Buscar_Precio)
                            Else
                                dPrecio = 0
                            End If
#End Region
                        End If
#Region "Actualizamos la línea"
                        For lin = 0 To oDoc.Lines.Count - 1
                            oDoc.Lines.SetCurrentLine(lin)
                            If oDoc.Lines.LineNum = CType(sLineNum, Integer) Then
                                oDoc.Lines.UnitPrice = dPrecio

                                If oDoc.Update() <> 0 Then
                                    sMensaje = oobjGlobal.compañia.GetLastErrorCode.ToString & " / " & oobjGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                                    oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "ERROR")
                                    oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sMensaje)
                                    sCerrar(1) = "ERROR"
                                Else
                                    sMensaje = "Se ha Actualizado correctamente el documento Nº " & sDocNum & " y Nº interno " & sDocEntry & " la línea " & sLineNum
                                    oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "OK")
                                    oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, "Precio: " & dPrecio.ToString)
                                    If sCerrar(1) <> "ERROR" Then
                                        sCerrar(1) = "OK"
                                    End If
#Region "Actualizar tarifa"
                                    If CType(oForm.Items.Item("chkTarifa").Specific, SAPbouiCOM.CheckBox).Checked And sBuscarPrecioTarifa = "N" Then
                                        oRs = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                                        sNomTarifa = oobjGlobal.funcionesUI.refDi.OGEN.valorVariable("Tarifa_Compra")
                                        sCodTarifa = oobjGlobal.refDi.SQL.sqlStringB1("SELECT ""ListNum"" FROM ""OPLN"" WHERE ""ListName""='" & sNomTarifa & "' ")
                                        If sCodTarifa.Trim <> "" Then
                                            'sSQL = "UPDATE ""ITM1"" "
                                            'sSQL &= " SET ""Price""=" & EXO_GLOBALES.DblNumberToText(oobjGlobal.compañia, dPrecio, EXO_GLOBALES.FuenteInformacion.Otros)
                                            'sSQL &= " WHERE ""ItemCode""='" & sItemCode & "' and ""PriceList""=" & sCodTarifa
                                            'oobjGlobal.refDi.SQL.sqlUpdB1(sSQL)
                                            'Son precios especiales
                                            Dim dInicioMesActual As Date = New Date(dFecha.Year, dFecha.Month, 1)
                                            Dim dFinMesActual As Date = New Date(dFecha.Year, dFecha.Month, DateSerial(dFecha.Year, dFecha.Month + 1, 0).Day)
                                            oOSPP = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSpecialPrices), SAPbobsCOM.SpecialPrices)
                                            If oOSPP.GetByKey(sItemCode, sCardCode) = True Then
                                                oOSPP.Valid = SAPbobsCOM.BoYesNoEnum.tYES

                                                oRs.DoQuery("SELECT COUNT(t2.""LINENUM"") ""CONTADOR"" " &
                                                "FROM ""OSPP"" t1 INNER JOIN " &
                                                """SPP1"" t2 ON t1.""ItemCode"" = t2.""ItemCode"" AND " &
                                                "t1.""CardCode"" = t2.""CardCode"" " &
                                                "WHERE COALESCE(t1.""Valid"", 'N') = 'Y' " &
                                                "AND t2.""ItemCode"" = '" & sItemCode & "' " &
                                                "AND t2.""CardCode"" = '" & sCardCode & "' and t1.""ListNum""='0' ")

                                                If oRs.RecordCount > 0 Then
                                                    iContPrecios = CInt(oRs.Fields.Item("CONTADOR").Value.ToString)
                                                Else
                                                    iContPrecios = 0
                                                End If
                                                sSQL = "SELECT t2.""LINENUM"" " &
                                                "FROM ""OSPP"" t1 INNER JOIN " &
                                                """SPP1"" t2 ON t1.""ItemCode"" = t2.""ItemCode"" AND " &
                                                "t1.""CardCode"" = t2.""CardCode"" " &
                                                "WHERE COALESCE(t1.""Valid"", 'N') = 'Y' " &
                                                "AND t2.""ItemCode"" = '" & sItemCode & "' " &
                                                "AND t2.""CardCode"" = '" & sCardCode & "'  and t2.""ListNum""='0' " &
                                                "AND ((CASE WHEN TO_CHAR(COALESCE(t2.""FromDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE TO_CHAR(COALESCE(t2.""FromDate"", ''), 'YYYYMMDD') END <= CASE WHEN TO_CHAR(COALESCE(t2.""FromDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE '" & Right("000" & dFecha.Year.ToString, 4) & Right("0" & dFecha.Month.ToString, 2) & "01' END " &
                                                "AND CASE WHEN TO_CHAR(COALESCE(t2.""ToDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE TO_CHAR(COALESCE(t2.""ToDate"", ''), 'YYYYMMDD') END >= CASE WHEN TO_CHAR(COALESCE(t2.""ToDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE '" & Right("000" & dFecha.Year.ToString, 4) & Right("0" & dFecha.Month.ToString, 2) & "01' END) " &
                                                "OR (CASE WHEN TO_CHAR(COALESCE(t2.""FromDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE TO_CHAR(COALESCE(t2.""FromDate"", ''), 'YYYYMMDD') END <= CASE WHEN TO_CHAR(COALESCE(t2.""FromDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE '" & Right("000" & dFecha.Year.ToString, 4) & Right("0" & dFecha.Month.ToString, 2) & Right("0" & DateSerial(dFecha.Year, dFecha.Month + 1, 0).Day.ToString, 2) & "' END " &
                                                "AND CASE WHEN TO_CHAR(COALESCE(t2.""ToDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE TO_CHAR(COALESCE(t2.""ToDate"", ''), 'YYYYMMDD') END >= CASE WHEN TO_CHAR(COALESCE(t2.""ToDate"", ''), 'YYYYMMDD') = '' THEN '' ELSE '" & Right("000" & dFecha.Year.ToString, 4) & Right("0" & dFecha.Month.ToString, 2) & Right("0" & DateSerial(dFecha.Year, dFecha.Month + 1, 0).Day.ToString, 2) & "' END)) " &
                                                "ORDER BY t2.""LINENUM"" "
                                                oRs.DoQuery(sSQL)

                                                oXml.LoadXml(oRs.GetAsXML())
                                                oNodes = oXml.SelectNodes("//row")

                                                If oRs.RecordCount > 0 Then
                                                    For j As Integer = oNodes.Count - 1 To 0 Step -1
                                                        oNode = oNodes.Item(j)

                                                        oOSPP.SpecialPricesDataAreas.SetCurrentLine(CInt(oNode.SelectSingleNode("LINENUM").InnerText))
                                                        oOSPP.SpecialPricesDataAreas.Delete()
                                                    Next
                                                End If

                                                iContPrecios -= oRs.RecordCount

                                                If iContPrecios > 0 Then
                                                    oOSPP.SpecialPricesDataAreas.Add()
                                                End If

                                                oOSPP.SpecialPricesDataAreas.AutoUpdate = SAPbobsCOM.BoYesNoEnum.tNO
                                                oOSPP.SpecialPricesDataAreas.PriceListNo = 0
                                                oOSPP.SpecialPricesDataAreas.DateFrom = dInicioMesActual
                                                oOSPP.SpecialPricesDataAreas.Dateto = dFinMesActual
                                                oOSPP.SpecialPricesDataAreas.SpecialPrice = dPrecio

                                                If oOSPP.Update() <> 0 Then
                                                    Throw New Exception(oobjGlobal.compañia.GetLastErrorCode.ToString & " / " & oobjGlobal.compañia.GetLastErrorDescription.Replace("'", ""))
                                                End If
                                            Else
                                                oOSPP.Valid = SAPbobsCOM.BoYesNoEnum.tYES
                                                oOSPP.ItemCode = sItemCode
                                                oOSPP.CardCode = sCardCode
                                                oOSPP.PriceListNum = 0 'CInt(sCodTarifa)

                                                oOSPP.SpecialPricesDataAreas.AutoUpdate = SAPbobsCOM.BoYesNoEnum.tNO
                                                oOSPP.SpecialPricesDataAreas.PriceListNo = 0
                                                oOSPP.SpecialPricesDataAreas.DateFrom = dInicioMesActual
                                                oOSPP.SpecialPricesDataAreas.Dateto = dFinMesActual
                                                oOSPP.SpecialPricesDataAreas.SpecialPrice = dPrecio

                                                If oOSPP.Add() <> 0 Then
                                                    Throw New Exception(oobjGlobal.compañia.GetLastErrorCode.ToString & " / " & oobjGlobal.compañia.GetLastErrorDescription.Replace("'", ""))
                                                End If
                                            End If
                                            oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha actualizado en la Tarifa de compra el artículo " & sItemCode & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            If oOSPP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOSPP)
                                            oOSPP = Nothing
                                        Else
                                            oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra la Tarifa de compra. No se puede actualizar el artículo " & sItemCode & ". Revise la parametrización ""Tarifa_Compra"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                End If
#End Region
                                Exit For
                            End If
                        Next
#End Region

#End Region

                    Else
                        oobjGlobal.SBOApp.StatusBar.SetText("No se ha podido Encontrar el documento Nº" & sDocNum & ", revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                End If
            Next
            'cerramos el último documento tratado
            If sCerrar(1) = "OK" Then
                If CType(oForm.Items.Item("chkCerrar").Specific, SAPbouiCOM.CheckBox).Checked Then
                    If sCerrarDoc = "Y" And sCerrarDocIC = "Y" Then
                        oDoc.GetByKey(CType(sCerrar(0), Integer))
                        If oDoc.Close() <> 0 Then
                            sMensaje = oobjGlobal.compañia.GetLastErrorCode.ToString & " / " & oobjGlobal.compañia.GetLastErrorDescription.Replace("'", "")
                            oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            sMensaje = "Se ha Cerrado correctamente el documento Nº " & sDocNum & " y Nº interno " & sDocEntry
                            oobjGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                End If
            End If

            Calcular_Precio = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oDoc IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDoc)
            oDoc = Nothing
            If oOSPP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOSPP)
            oOSPP = Nothing
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
            oRs = Nothing
        End Try
    End Function
End Class
