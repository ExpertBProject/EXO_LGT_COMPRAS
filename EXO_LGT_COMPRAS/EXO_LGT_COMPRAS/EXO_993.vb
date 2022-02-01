Imports SAPbouiCOM
Public Class EXO_993
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
                        Case "993"
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
                        Case "993"
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
                        Case "993"
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
                        Case "993"
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
            oItem = oForm.Items.Add("btnSUP", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oItem.Width = oForm.Items.Item("2").Width * 2
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Suplemento"
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
                Case "btnSUP"
                    If pVal.ActionSuccess = True Then
                        If CargarUDOSuplemento(oForm) = False Then
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
    Public Function CargarUDOSuplemento(ByRef oForm As SAPbouiCOM.Form) As Boolean
#Region "Variables"
        Dim sArticulo As String = ""
        Dim sCardCode As String = ""
        Dim sCatalogo As String = ""
        Dim iLineaSel As Integer = 0 : Dim bLineaSel As Boolean = False
        Dim sExiste As String = "" : Dim sMensaje As String = ""
        Dim sSQL As String = ""
        Dim sMatrix As String = ""
#End Region

        CargarUDOSuplemento = False

        Try
            If oForm.PaneLevel = 1 Then
                sMatrix = "17"
            Else
                sMatrix = "28"
            End If
            For i = 1 To CType(oForm.Items.Item(sMatrix).Specific, SAPbouiCOM.Matrix).RowCount
                If CType(oForm.Items.Item(sMatrix).Specific, SAPbouiCOM.Matrix).IsRowSelected(i) = True Then
                    iLineaSel = i : bLineaSel = True
                    Exit For
                End If
            Next

            If bLineaSel = True Then
                If oForm.PaneLevel = 1 Then
                    sCardCode = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.EditText).Value.ToString
                    sArticulo = CType(CType(oForm.Items.Item("17").Specific, SAPbouiCOM.Matrix).Columns.Item("1").Cells.Item(iLineaSel).Specific, SAPbouiCOM.EditText).Value.ToString
                    sCatalogo = CType(CType(oForm.Items.Item("17").Specific, SAPbouiCOM.Matrix).Columns.Item("3").Cells.Item(iLineaSel).Specific, SAPbouiCOM.EditText).Value.ToString
                Else
                    sArticulo = CType(oForm.Items.Item("21").Specific, SAPbouiCOM.EditText).Value.ToString
                    sCardCode = CType(CType(oForm.Items.Item("28").Specific, SAPbouiCOM.Matrix).Columns.Item("1").Cells.Item(iLineaSel).Specific, SAPbouiCOM.EditText).Value.ToString
                    sCatalogo = CType(CType(oForm.Items.Item("28").Specific, SAPbouiCOM.Matrix).Columns.Item("3").Cells.Item(iLineaSel).Specific, SAPbouiCOM.EditText).Value.ToString
                End If


                If oForm.Mode = BoFormMode.fm_OK_MODE Then
                    'Si no existe, creamos el artículo
                    sSQL = "SELECT ""Code"" FROM ""@EXO_SUPLEMENTO"" WHERE ""Code""='" & sArticulo & "_" & sCardCode & "_" & sCatalogo & "' "
                    sExiste = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                    If sExiste = "" Then
                        'Presentamos UDO Y escribimos los datos de la cabecera
                        INICIO._sArticulo = sArticulo
                        INICIO._sCardCode = sCardCode
                        INICIO._sCatalogo = sCatalogo
                        INICIO._sLineaSel = iLineaSel.ToString
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_SUPLEMENTO")
                    Else
                        INICIO._sArticulo = ""
                        INICIO._sCardCode = ""
                        INICIO._sCatalogo = ""
                        INICIO._sLineaSel = ""
                        objGlobal.funcionesUI.cargaFormUdoBD_Clave("EXO_SUPLEMENTO", sArticulo & "_" & sCardCode & "_" & sCatalogo)
                    End If
                Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, guarde primero los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objGlobal.SBOApp.MessageBox("Por favor, guarde primero los datos")
                End If
            Else
                sMensaje = "Tiene que seleccionar una línea."
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objGlobal.SBOApp.MessageBox(sMensaje)
            End If



            CargarUDOSuplemento = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally

        End Try
    End Function
End Class
