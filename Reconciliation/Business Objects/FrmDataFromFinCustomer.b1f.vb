Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SAPbobsCOM

Namespace Reconciliation
    <FormAttribute("BPREC", "Business Objects/FrmDataFromFinCustomer.b1f")>
    Friend Class FrmDataFromFinCustomer
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public WithEvents odbdsHeader, odbdsDetails As SAPbouiCOM.DBDataSource
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.EditText0 = CType(Me.GetItem("tentry").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lDate").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tdocdate").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("mcont").Specific, SAPbouiCOM.Matrix)
            Me.EditText2 = CType(Me.GetItem("txtFName").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lexcel").Specific, SAPbouiCOM.StaticText)
            Me.Button2 = CType(Me.GetItem("bexcel").Specific, SAPbouiCOM.Button)
            Me.Button3 = CType(Me.GetItem("Clear").Specific, SAPbouiCOM.Button)
            Me.EditText3 = CType(Me.GetItem("txtdocnum").Specific, SAPbouiCOM.EditText)
            Me.Folder0 = CType(Me.GetItem("fcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldr").Specific, SAPbouiCOM.Folder)
            Me.StaticText0 = CType(Me.GetItem("lno").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("cseries").Specific, SAPbouiCOM.ComboBox)
            Me.EditText4 = CType(Me.GetItem("tdocnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("lrem").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("trem").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub


        Private Sub OnCustomInitialize()

            Try
                objform.Left = (objaddon.objapplication.Desktop.Width - objform.MaxWidth) / 2
                objform.Top = (objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 4

                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_OBRS")
                'odbdsDetails = objform.DataSources.DBDataSources.Item("@AT_BRS1")
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_OBRS")
                objform.Items.Item("tdocdate").Specific.String = "A"
                objform.Items.Item("trem").Specific.String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss")
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "series", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocnum", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocdate", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "mcont", True, False, True)
                objform.EnableMenu("1283", False) 'Remove
                Folder0.Item.Click()

            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Button3 As SAPbouiCOM.Button
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText

#End Region

        Private Sub LoadExcel(ByVal FileName As String)
            Try
                objaddon.objapplication.SetStatusBarMessage("Looking out for Excel Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
                Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim excelRng As Microsoft.Office.Interop.Excel.Range
                Dim j As Integer = 1
                Try
                    FileName = objform.Items.Item("txtFName").Specific.string
                    Dim RowIndex As Integer
                    ExcelWorkbook = ExcelApp.Workbooks.Open(FileName)
                    ExcelWorkSheet = ExcelWorkbook.ActiveSheet
                    'excelRng = ExcelWorkSheet.Range("A1")
                    excelRng = ExcelWorkSheet.UsedRange
                    objaddon.objapplication.SetStatusBarMessage("Excel Loading please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    objform.Freeze(True)
                    Matrix0.Clear()
                    If objform.Items.Item("txtFName").Specific.String <> "" Then
                        If ExcelWorkSheet.Cells(1, 1).Value = "Customer Name" And ExcelWorkSheet.Cells(1, 2).Value = "Date" And ExcelWorkSheet.Cells(1, 3).Value = "Deal ID 1" And ExcelWorkSheet.Cells(1, 4).Value = "Deal ID 2" And ExcelWorkSheet.Cells(1, 5).Value = "Amount" Then
                            For RowIndex = 2 To excelRng.Rows.Count
                                Matrix0.AddRow()
                                'Dim AttDate As Date = Date.ParseExact(ExcelWorkSheet.Cells(RowIndex, 2).Value, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                Matrix0.Columns.Item("#").Cells.Item(j).Specific.String = j
                                Matrix0.Columns.Item("CustName").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 1).Value).ToUpper
                                Matrix0.Columns.Item("Date").Cells.Item(j).Specific.String = Format(ExcelWorkSheet.Cells(RowIndex, 2).Value, "yyyyMMdd") 'ExcelWorkSheet.Cells(RowIndex, 2).Value 
                                Matrix0.Columns.Item("DealID1").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 3).Value).ToUpper
                                Matrix0.Columns.Item("DealID2").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 4).Value).ToUpper
                                Matrix0.Columns.Item("Amount").Cells.Item(j).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 5).Value)
                                j += 1
                            Next RowIndex
                        Else
                            objaddon.objapplication.StatusBar.SetText("Expected ColumnName Not found...Please check the excel format", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                        objform.Freeze(False)

                    End If
                    objaddon.objapplication.Menus.Item("1300").Activate()
                    ExcelApp.ActiveWorkbook.Close()
                    objaddon.objapplication.StatusBar.SetText("Excel Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'objaddon.objapplication.SetStatusBarMessage("Excel Loaded..." & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Catch ex As Exception
                    objform.Freeze(False)
                    ExcelApp.ActiveWorkbook.Close()
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub InternRec()
            Try
                Dim service As IInternalReconciliationsService = objaddon.objcompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim openTrans As InternalReconciliationOpenTrans = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans)
                Dim reconParams As IInternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard
                openTrans.InternalReconciliationOpenTransRows.Add()
                openTrans.InternalReconciliationOpenTransRows.Item(0).Selected = BoYesNoEnum.tYES
                openTrans.InternalReconciliationOpenTransRows.Item(0).TransId = 196
                openTrans.InternalReconciliationOpenTransRows.Item(0).TransRowId = 0
                openTrans.InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = 50
                'openTrans.InternalReconciliationOpenTransRows.Add()
                'openTrans.InternalReconciliationOpenTransRows.Item(1).Selected = BoYesNoEnum.tYES
                'openTrans.InternalReconciliationOpenTransRows.Item(1).TransId = 195
                'openTrans.InternalReconciliationOpenTransRows.Item(1).TransRowId = 1
                'openTrans.InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = 100
                reconParams = service.Add(openTrans)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button2.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If EditText2.Value = "" Then
                    objaddon.objapplication.StatusBar.SetText("Please select a excel file...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub InternalReconciliation()
            Try
                Dim service As InternalReconciliationsService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.InternalReconciliationsService)
                Dim transParams As InternalReconciliationOpenTransParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTransParams)
                Dim reconcileOnAccount As Boolean = True
                Dim transId1 As Integer = 190, transRowId1 As Integer = 1, transId2 As Integer = 191, transRowId2 As Integer = 0
                transParams.ReconDate = DateTime.Today
                transParams.DateType = ReconSelectDateTypeEnum.rsdtPostDate
                transParams.FromDate = Now.Date ' New DateTime(2017, 10, 11)
                transParams.ToDate = Now.Date ' New DateTime(2017, 11, 11)

                If reconcileOnAccount = False Then
                    transParams.CardOrAccount = CardOrAccountEnum.coaAccount
                    transParams.AccountNo = "100101"
                Else
                    transParams.CardOrAccount = CardOrAccountEnum.coaCard
                    transParams.InternalReconciliationBPs.Add()
                    transParams.InternalReconciliationBPs.Item(0).BPCode = "C00002"
                End If
                Dim openTrans As InternalReconciliationOpenTrans = service.GetOpenTransactions(transParams)
                Dim row As InternalReconciliationOpenTransRows = openTrans.InternalReconciliationOpenTransRows

                For i As Integer = 1 To Matrix0.VisualRowCount
                    For Each row In openTrans.InternalReconciliationOpenTransRows
                        If Matrix0.Columns.Item("CustName").Cells.Item(i).Specific.String = openTrans.InternalReconciliationOpenTransRows.Item(row).ShortName And (Matrix0.Columns.Item("DealID1").Cells.Item(i).Specific.String = openTrans.InternalReconciliationOpenTransRows.Item(row).TransId Or Matrix0.Columns.Item("DealID2").Cells.Item(i).Specific.String = "1") Then
                            row.Selected = BoYesNoEnum.tYES
                            row.ReconcileAmount = CDbl(Matrix0.Columns.Item("Amount").Cells.Item(i).Specific.String)

                        End If
                        If (transId1 = row.TransId And transRowId1 = row.TransRowId) Then
                            row.Selected = BoYesNoEnum.tYES
                            row.ReconcileAmount = 50
                            'row.CashDiscount = 1
                        ElseIf (transId2 = row.TransId And transRowId2 = row.TransRowId) Then
                            row.Selected = BoYesNoEnum.tYES
                            row.ReconcileAmount = 50
                        End If
                    Next
                Next

                Dim reconParams As InternalReconciliationParams = service.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                reconParams = service.Add(openTrans)
                Dim ii As Integer = reconParams.ReconNum
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    LoadExcel(objform.Items.Item("txtFName").Specific.string)
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Button3_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button3.ClickBefore
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    EditText2.Value = ""
                    Matrix0.Clear()
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                'InternalReconciliation()
                InternRec()
                BubbleEvent = False : Exit Sub
                If Matrix0.VisualRowCount = 0 Then
                    objaddon.objapplication.StatusBar.SetText("Please fill line data..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Dim objMatrix As SAPbouiCOM.Matrix
                Dim objRecForm As SAPbouiCOM.Form
                Dim chkSelect As SAPbouiCOM.CheckBox
                Dim status As Boolean = False
                Dim currency, DocDate As String
                objRecForm = objaddon.objapplication.Forms.GetForm("120060805", 1)
                objMatrix = objRecForm.Items.Item("120000039").Specific
                currency = objaddon.objglobalmethods.getSingleValue("Select ""MainCurncy"" from OADM")
                Dim Amount As String = currency + "  "
                Dim custName, DealID, Price, PostingDate As String
                objaddon.objapplication.StatusBar.SetText("Validating Data from UDO to Internal Reconciliation Please wait..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                For i As Integer = 1 To Matrix0.VisualRowCount
                    Amount += Matrix0.Columns.Item("Amount").Cells.Item(i).Specific.String
                    DocDate = Matrix0.Columns.Item("Date").Cells.Item(i).Specific.String
                    For RecMatrix As Integer = 1 To objMatrix.VisualRowCount
                        chkSelect = objMatrix.Columns.Item("120000002").Cells.Item(RecMatrix).Specific
                        custName = objMatrix.Columns.Item("120000030").Cells.Item(RecMatrix).Specific.String
                        DealID = objMatrix.Columns.Item("120000013").Cells.Item(RecMatrix).Specific.String
                        Price = objMatrix.Columns.Item("120000014").Cells.Item(RecMatrix).Specific.String
                        PostingDate = objMatrix.Columns.Item("120000008").Cells.Item(RecMatrix).Specific.String
                        If Matrix0.Columns.Item("CustName").Cells.Item(i).Specific.String <> "" Then
                            'If custName.ToUpper = Matrix0.Columns.Item("CustName").Cells.Item(i).Specific.String And PostingDate = DocDate And (DealID.ToUpper = Matrix0.Columns.Item("DealID1").Cells.Item(i).Specific.String Or DealID.ToUpper = Matrix0.Columns.Item("DealID2").Cells.Item(i).Specific.String) And Price = Amount Then
                            If (DealID.ToUpper = Matrix0.Columns.Item("DealID1").Cells.Item(i).Specific.String Or DealID.ToUpper = Matrix0.Columns.Item("DealID2").Cells.Item(i).Specific.String) And Price = Amount Then
                                chkSelect.Checked = True
                                status = True
                                Exit For
                            End If
                        End If
                    Next
                    Amount = currency + "  "
                    If status = True Then
                        Matrix0.Columns.Item("Remarks").Cells.Item(i).Specific.String = "Found"
                        status = False
                    Else
                        Matrix0.Columns.Item("Remarks").Cells.Item(i).Specific.String = "Not Found"
                    End If
                Next
                objaddon.objapplication.StatusBar.SetText("Validated Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Button0_PressedAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.ActionSuccess = True Then
                    EditText0.Value = objaddon.objglobalmethods.GetNextNumber("BPREC") 'objaddon.objglobalmethods.GetNextDocEntry_Value("@MIPL_OREC")
                    objform.Items.Item("txtdate").Specific.string = Now.Date.ToString("dd/MM/yy")
                End If
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("BPREC", pVal.FormTypeCount)
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
