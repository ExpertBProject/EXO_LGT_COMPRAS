Imports SAPbouiCOM
Public Class EXO_HOMO
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
                        Case "EXO_HOMO"
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
                        Case "EXO_HOMO"
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
                        Case "EXO_HOMO"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_HOMO"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        Return False
                                    End If
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
    Private Function EventHandler_Choose_FromList_Before(ByVal pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sItemCode As String
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim sGroupCode As String = ""
        Dim bEsADR As Boolean = False
        Dim bDebeSerADR As Boolean = False
        Dim h As Integer = 1

        EventHandler_Choose_FromList_Before = False

        Try
            If pVal.ItemUID = "txtDPROV" Or pVal.ItemUID = "txtHPROV" Then 'Rango Proveedores
                oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions
                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "S"
                'oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByVal pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EventHandler_Choose_FromList_After = False

        Try
            oForm = Me.objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            oDataTable = oCFLEvento.SelectedObjects
            If Not oDataTable Is Nothing Then
                Select Case oCFLEvento.ChooseFromListUID
                    Case "CFLDPR"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "txtDPROV" Then
                                Try
                                    CType(oForm.Items.Item("txtDPROV").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    Case "CFLHPR"
                        oDataTable = oCFLEvento.SelectedObjects

                        If oDataTable IsNot Nothing Then
                            If pVal.ItemUID = "txtHPROV" Then
                                Try
                                    CType(oForm.Items.Item("txtHPROV").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                End Select
            End If

            EventHandler_Choose_FromList_After = True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.FormDatatable(oDataTable)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
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
                    If objGlobal.SBOApp.MessageBox("¿Calculamos datos de Homologación?", 1, "Sí", "No") = 1 Then
                        If ComprobarDOC(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btn_Cal").Enabled = False
                            'Calculando datos
                            objGlobal.SBOApp.StatusBar.SetText("Calculando datos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If Calcular_Homo(oForm, "DT_DOC", objGlobal) = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                            oForm.Items.Item("btn_Cal").Enabled = True
                        End If
                    End If
                End If
            ElseIf pVal.ItemUID = "btn_Fich" Then
                Cargar_Grid(oForm, CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm.DataSources.UserDataSources.Item("UDDPR").Value.ToString, oForm.DataSources.UserDataSources.Item("UDHPR").Value.ToString)


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
    Public Shared Function Calcular_Homo(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByRef oobjGlobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        Calcular_Homo = False
#Region "VARIABLES"
        Dim oRsCab As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLin As SAPbobsCOM.Recordset = CType(oobjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sProveedor As String = ""
        Dim iD3Cal As Integer = 0 : Dim iD15Cal As Integer = 0 : Dim iD45Cal As Integer = 0 : Dim iD200Cal As Integer = 0
        Dim iD3Log As Integer = 0 : Dim iD15Log As Integer = 0 : Dim iD45Log As Integer = 0 : Dim iD200Log As Integer = 0
        Dim iRecep As Long = 1
        Dim IL As Double = 0 : Dim IC As Double = 0 : Dim TOTAL As Double = 0
        Dim PESOL As Integer = 20 : Dim PESOC As Integer = 80
        Dim sHomo As String = ""
#End Region


        Try

            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    If sProveedor <> "" And sProveedor <> oForm.DataSources.DataTables.Item(sData).GetValue("Código", i).ToString Then
                        'Grabamos la homologación
#Region "Grabamos Homologación"
                        'HOMOLOGADO TIPO A Proveedor con índice de rendimiento TOTAL superior al 90%.
                        'HOMOLOGADO TIPO B: Proveedor con índice de rendimiento TOTAL comprendido entre el 80 y 90%.
                        'HOMOLOGADO TIPO C: Proveedor con índice de rendimiento en Calidad o Logística inferior al 80%.
                        sHomo = ""
                        If TOTAL < 80 Then
                            sSQL = "UPDATE ""OCRD"" SET ""U_EXO_HOMO""='C' where ""CardCode""='" & sProveedor & "' "
                            sHomo = "C"
                        ElseIf TOTAL >= 80 And TOTAL <= 90 Then
                            sSQL = "UPDATE ""OCRD"" SET ""U_EXO_HOMO""='B' where ""CardCode""='" & sProveedor & "' "
                            sHomo = "B"
                        ElseIf TOTAL > 90 Then
                            sSQL = "UPDATE ""OCRD"" SET ""U_EXO_HOMO""='A' where ""CardCode""='" & sProveedor & "' "
                            sHomo = "A"
                        End If
                        If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                            oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Homologación " & sHomo & " al Proveedor " & sProveedor, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Else
                            oobjGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar Homologación " & sHomo & " al Proveedor " & sProveedor, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
#End Region
                    End If
                    If sProveedor <> oForm.DataSources.DataTables.Item(sData).GetValue("Código", i).ToString Then
                        'Ponemos a cero
                        sProveedor = oForm.DataSources.DataTables.Item(sData).GetValue("Código", i).ToString
                        iD3Cal = 0 : iD15Cal = 0 : iD45Cal = 0 : iD200Cal = 0
                        iD3Log = 0 : iD15Log = 0 : iD45Log = 0 : iD200Log = 0
                        iRecep = 0
                        IL = 0 : IC = 0 : TOTAL = 0
                    End If
                    iRecep += 1

                    'Calculamos
                    sSQL = "SELECT * FROM ""PDN1"" WHere ""DocEntry""=" & oForm.DataSources.DataTables.Item(sData).GetValue("Nº Interno", i).ToString
                    oRsLin.DoQuery(sSQL)
                    For lin = 0 To oRsLin.RecordCount - 1
                        Select Case oRsLin.Fields.Item("U_EXO_DEMECAL").Value.ToString
                            Case "3" : iD3Cal += 1
                            Case "15" : iD15Cal += 1
                            Case "45" : iD45Cal += 1
                            Case "200" : iD200Cal += 1
                        End Select
                        Dim dFecha As Date = CDate(oForm.DataSources.DataTables.Item(sData).GetValue("Fecha Contable", i).ToString)
                        Dim sFechaLin As String = oRsLin.Fields.Item("U_EXO_FECHAENTPROV").Value.ToString
                        Dim dFechaLin As Date = CDate(oForm.DataSources.DataTables.Item(sData).GetValue("Fecha Contable", i).ToString)
                        Dim iDiasDif As Long = 0
                        If sFechaLin <> "" Then
                            dFechaLin = CDate(sFechaLin)
                            If dFechaLin.Year <= 2000 Then
                                dFechaLin = CDate(oForm.DataSources.DataTables.Item(sData).GetValue("Fecha Contable", i).ToString)
                            End If
                        End If
                        iDiasDif = DateDiff(DateInterval.Day, dFechaLin, dFecha)
                        oobjGlobal.SBOApp.StatusBar.SetText("Dif. Días F. Contable con U_EXO_FECHAENTPROV: " & iDiasDif.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If iDiasDif = -1 Or iDiasDif = 2 Then
                            iD3Log += 1
                        ElseIf iDiasDif = -2 Or iDiasDif = 3 Then
                            iD15Log += 1
                        ElseIf iDiasDif = -3 Then ' Gastos
                            iD45Log += 1
                        Else
                            iD200Log += 1 'Falta ver cómo se ve la parada en línea y falta resto
                        End If
                        oRsLin.MoveNext()
                    Next
                    Dim X As Integer = 0
                    If iRecep <= 30 Then
                        X = 45
                    ElseIf iRecep > 30 Then
                        X = 3
                    End If
                    oobjGlobal.SBOApp.StatusBar.SetText("Recepciones: " & iRecep.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    IL = (1 - (3 * iD3Log + 5 * iD15Log + 45 * iD45Log + 200 * iD200Log) / (iRecep * X)) * PESOL
                    IC = (1 - (3 * iD3Cal + 5 * iD15Cal + 45 * iD45Cal + 200 * iD200Cal) / (iRecep * X)) * PESOC
                    If IC = 0 Then
                        IC = 80
                    End If
                    If IL = 0 Then
                        IL = 20
                    End If

                    TOTAL = IL + IC
                    oobjGlobal.SBOApp.StatusBar.SetText("IL:  " & IL & " - IC: " & IC & " - TOTAL: " & TOTAL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If
            Next
#Region "Grabamos Homologación"
            'HOMOLOGADO TIPO A Proveedor con índice de rendimiento TOTAL superior al 90%.
            'HOMOLOGADO TIPO B: Proveedor con índice de rendimiento TOTAL comprendido entre el 80 y 90%.
            'HOMOLOGADO TIPO C: Proveedor con índice de rendimiento en Calidad o Logística inferior al 80%.
            sHomo = ""
            If TOTAL < 80 Then
                sSQL = "UPDATE ""OCRD"" SET ""U_EXO_HOMO""='C' where ""CardCode""='" & sProveedor & "' "
                sHomo = "C"
            ElseIf TOTAL >= 80 And TOTAL <= 90 Then
                sSQL = "UPDATE ""OCRD"" SET ""U_EXO_HOMO""='B' where ""CardCode""='" & sProveedor & "' "
                sHomo = "B"
            ElseIf TOTAL > 90 Then
                sSQL = "UPDATE ""OCRD"" SET ""U_EXO_HOMO""='A' where ""CardCode""='" & sProveedor & "' "
                sHomo = "A"
            End If
            If oobjGlobal.refDi.SQL.executeNonQuery(sSQL) = True Then
                oobjGlobal.SBOApp.StatusBar.SetText("Actualizado Homologación " & sHomo & " al Proveedor " & sProveedor, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oobjGlobal.SBOApp.StatusBar.SetText("No se ha podido actualizar Homologación " & sHomo & " al Proveedor " & sProveedor, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
#End Region

            Calcular_Homo = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCab, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLin, Object))
        End Try
    End Function
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
    Private Sub Cargar_Grid(ByRef oForm As SAPbouiCOM.Form, ByVal sPeriodo As String, ByVal sDProv As String, ByVal sHProv As String)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)

            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"", ""DocEntry"" ""Nº Interno"", ""DocNum"" ""Nº Documento"", ""CardCode"" ""Código"", ""CardName"" ""Nombre"", ""DocDate"" ""Fecha Contable"" "
            sSQL &= ", ""DocTotal"" ""Importe"" "
            sSQL &= " From ""OPDN"" "
            sSQL &= " WHERE ""CANCELED""='N' and YEAR(""DocDate"")='" & sPeriodo & "' and ""CardCode"">='" & sDProv & "' and ""CardCode""<='" & sHProv & "' "
            sSQL &= " ORDER BY ""CardCode"", ""DocNum"" "
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

            For i = 1 To 6
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                If i = 1 Then
                    oColumnTxt.LinkedObjectType = "20"
                ElseIf i = 3 Then
                    oColumnTxt.LinkedObjectType = "2"
                ElseIf i = 6 Then
                    oColumnTxt.RightJustified = True
                End If

            Next



            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean

        Dim sSQL As String = ""
        Try
            If infoEvento.BeforeAction = True Then

            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnHoIC"
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
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_HOMO.srf")

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
            sSQL = "SELECT DISTINCT ""Category"" ""COD"", ""Category"" ""ANNO"" FROM ""OFPR"" order by ""Category"" "
            objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).ExpandType = BoExpandType.et_ValueOnly
            CType(oForm.Items.Item("cbPER").Specific, SAPbouiCOM.ComboBox).Select(Now.Year.ToString, BoSearchKey.psk_ByValue)

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
End Class
