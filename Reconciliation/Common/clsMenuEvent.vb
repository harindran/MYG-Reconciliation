Imports SAPbouiCOM
Namespace Reconciliation

    Public Class clsMenuEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods

        Public Sub MenuEvent_For_StandardMenu(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    'Case "BPREC"
                    '    DataFromFinCustomer_MenuEvent(pVal, BubbleEvent)
                    'Case "385"
                    '    ProcessExternal_MenuEvent(pVal, BubbleEvent)
                    'Case "60800"
                    '    ReconciliationBankStatement_MenuEvent(pVal, BubbleEvent)
                    Case "CBRS"
                        ExternalBankRecoConsolidated_MenuEvent(pVal, BubbleEvent)
                End Select
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Default_Sample_MenuEvent(ByVal pval As SAPbouiCOM.MenuEvent, ByVal BubbleEvent As Boolean)
            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                Else
                    Select Case pval.MenuUID
                        Case "1281"
                        Case Else
                    End Select
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            End Try
        End Sub

#Region "Reconciliation"

        Private Sub ProcessExternal_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1292"
                        Case "1293"
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291"
                        Case "1293"
                        Case "CLF"
                            objform.Items.Item("txtFName").Specific.String = ""
                        Case "1304" 'Refresh

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub ReconciliationBankStatement_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1292"
                        Case "1293"
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                        Case "1282" ' Add Mode
                        Case "1288", "1289", "1290", "1291"
                        Case "1293"
                        'Case "GTS"
                        '    If Not objaddon.FormExist("MULDATA") Then
                        '        objMultiLoad = objaddon.objapplication.Forms.ActiveForm
                        '        If objMultiLoad.Items.Item("txtdepdata").Specific.String <> "" Then
                        '            objMultiLoad.Items.Item("txtdepdata").Specific.String = ""
                        '            Dim odbdsDetails As SAPbouiCOM.DBDataSource
                        '            odbdsDetails = objMultiLoad.DataSources.DBDataSources.Item("JDT1")
                        '            Dim Matrix0 As SAPbouiCOM.Matrix
                        '            Matrix0 = objMultiLoad.Items.Item("30").Specific
                        '            For RecRow As Integer = 0 To odbdsDetails.Size - 1
                        '                If Matrix0.Columns.Item("27").Cells.Item(RecRow + 1).Specific.Checked = True Then
                        '                    Matrix0.Columns.Item("27").Cells.Item(RecRow + 1).Specific.Checked = False
                        '                End If
                        '            Next
                        '        End If
                        '        Dim activeform As New FrmMultiData
                        '        activeform.Show()
                        '        activeform.UIAPIRawForm.Left = objform.Left + 100
                        '        activeform.UIAPIRawForm.Top = objform.Top + 100
                        '    End If
                        Case "1304" 'Refresh

                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub DataFromFinCustomer_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)

            Try
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283", "1284" 'Remove & Cancel
                            objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1292"
                        Case "1293"

                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("txtentry").Enabled = True
                            objform.Items.Item("txtdate").Enabled = True

                        Case "1282" ' Add Mode

                            objform.Items.Item("txtentry").Specific.string = objaddon.objglobalmethods.GetNextNumber("BPREC") 'objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OREC")
                            objform.Items.Item("txtdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                        Case "1288", "1289", "1290", "1291"

                        Case "1293"

                        Case "1292"

                        Case "1304" 'Refresh


                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub ExternalBankRecoConsolidated_MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
            Try
                Dim odbdsHeader As SAPbouiCOM.DBDataSource
                objform = objaddon.objapplication.Forms.ActiveForm
                If pval.BeforeAction = True Then
                    Select Case pval.MenuUID
                        Case "1283" 'Remove 
                            objaddon.objapplication.SetStatusBarMessage("Remove is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False
                        Case "1284" ' Cancel
                            If objform.Items.Item("cstat").Specific.Selected.Value = "C" And objform.Items.Item("trecono").Specific.String <> "" Then
                                If objaddon.objapplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") <> 1 Then BubbleEvent = False : Exit Sub
                                If CancelBankReconciliation(objform.Items.Item("tacctcode").Specific.String, objform.Items.Item("trecono").Specific.String) = False Then
                                    objaddon.objapplication.StatusBar.SetText("Failed to cancel the Bank Reconciliation: ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                            Else
                                BubbleEvent = False
                            End If
                    End Select
                Else
                    Select Case pval.MenuUID
                        Case "1281" 'Find Mode
                            objform.Items.Item("tclrbal").Enabled = True
                            objform.Items.Item("tdiff").Enabled = True
                            objform.Items.Item("tentry").Enabled = True

                        Case "1282" ' Add Mode
                            odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_OBRS")
                            objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_OBRS")
                            objform.Items.Item("tdocdate").Specific.String = "A"
                            objform.Items.Item("trem").Specific.String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss")
                    End Select
                End If
            Catch ex As Exception
                objform.Freeze(False)
                ' objaddon.objapplication.SetStatusBarMessage("Error in Menu Event" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Function CancelBankReconciliation(ByVal AcctCode As String, ByVal RecoNo As String) As Boolean
            Try
                Dim oCompanyService As SAPbobsCOM.CompanyService = objaddon.objcompany.GetCompanyService()
                Dim ExtReconSvc As SAPbobsCOM.ExternalReconciliationsService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ExternalReconciliationsService)
                Dim ExtReconciliation As SAPbobsCOM.ExternalReconciliation = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliation)
                ExtReconciliation.ReconciliationAccountType = SAPbobsCOM.ReconciliationAccountTypeEnum.rat_GLAccount

                Dim ExtReconParam As SAPbobsCOM.ExternalReconciliationParams = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliationParams)
                ExtReconParam.AccountCode = AcctCode ' "10602010"
                ExtReconParam.ReconciliationNo = RecoNo ' "3"
                ExtReconciliation = ExtReconSvc.GetReconciliation(ExtReconParam)
                ExtReconSvc.CancelReconciliation(ExtReconParam)
                objaddon.objapplication.StatusBar.SetText("Bank Reconciliation cancelled successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("CancelBankReconciliation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
        End Function

        Sub DeleteRow(ByVal objMatrix As SAPbouiCOM.Matrix, ByVal TableName As String)
            Try
                Dim DBSource As SAPbouiCOM.DBDataSource
                'objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource()
                DBSource = objform.DataSources.DBDataSources.Item(TableName) '"@MIREJDET1"
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.GetLineData(i)
                    DBSource.Offset = i - 1
                    DBSource.SetValue("LineId", DBSource.Offset, i)
                    objMatrix.SetLineData(i)
                    objMatrix.FlushToDataSource()
                Next
                DBSource.RemoveRecord(DBSource.Size - 1)
                objMatrix.LoadFromDataSource()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Finally
            End Try
        End Sub

#End Region

    End Class
End Namespace