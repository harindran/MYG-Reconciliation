Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System.Text.RegularExpressions
Imports SAPbobsCOM

Namespace Reconciliation
    <FormAttribute("120060805", "Business Objects/FrmInternalReconciliation.b1f")>
    Friend Class FrmInternalReconciliation
        Inherits SystemFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public WithEvents objMatrix As SAPbouiCOM.Matrix

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("AutoSel").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("readexcel").Specific, SAPbouiCOM.Button)
            Me.EditText0 = CType(Me.GetItem("tfname").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub


        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("120060805", 1)
                'objform = objaddon.objapplication.Forms.ActiveForm
                EditText0.Item.Left = objform.Items.Item("120000002").Left + objform.Items.Item("120000002").Width + 15
                EditText0.Item.Top = objform.Items.Item("120000002").Top
                Button0.Item.Top = objform.Items.Item("120000002").Top
                Button0.Item.Left = EditText0.Item.Left + EditText0.Item.Width + 5 'objform.Items.Item("120000002").Left + objform.Items.Item("120000002").Width + 3
                Button0.Item.Height = objform.Items.Item("120000002").Height

                'Button1.Item.Top = objform.Items.Item("120000002").Top
                'Button1.Item.Left = Button0.Item.Left + Button0.Item.Width + 3
                'Button1.Item.Height = objform.Items.Item("120000002").Height
                Button1.Item.Visible = False
            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents EditText0 As SAPbouiCOM.EditText

#End Region

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                ReadExcelData(EditText0.Value)
                'Dim StrQuery As String, Ref2 As String, Ref3 As String
                ''Dim price() As String
                'Dim Flag As Boolean = False
                'Dim Amount, RecAmount As Double
                'Dim objRecordset, objRS As SAPbobsCOM.Recordset
                ''Dim objedit As SAPbouiCOM.EditText
                'objRecordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRS = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRecordset.DoQuery(GetRecData)
                'If objRecordset.RecordCount > 0 Then
                '    objaddon.objapplication.StatusBar.SetText("Validating First level auto reconciliation Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    objMatrix = objform.Items.Item("120000039").Specific
                '    'objedit = objform.Items.Item("120000015").Specific
                '    StrQuery = "SELECT ROW_NUMBER() OVER (order by T0.""RefDate"",T0.""TransId"") as ""#"",T0.""BalDueDeb"" , T0.""BalDueCred"" ,T1.""CardCode"" AS ""Cust Num"", T0.""TransId"","
                '    StrQuery += vbCrLf + "T1.""CardName"" AS ""Cust Name"",T0.""SYSDeb"" AS ""Debit Amt"",T0.""SYSCred"" * -1 AS ""Credit Amt"",T0.""LineMemo"","
                '    StrQuery += vbCrLf + "CASE WHEN T0.""TransType"" = 13 THEN 'AR Invoice' "
                '    StrQuery += vbCrLf + "WHEN T0.""TransType"" = 14 THEN 'AR Cred Memo'"
                '    StrQuery += vbCrLf + "WHEN T0.""TransType"" = 24 THEN 'Payment'"
                '    StrQuery += vbCrLf + "WHEN T0.""TransType"" = 30 Then 'JE' ELSE 'Other' END AS ""Trans Type"","
                '    StrQuery += vbCrLf + "T0.""Ref1"" AS ""Ref1"",T0.""Ref2"" AS ""Ref2"",T0.""Ref3Line"" as ""Ref3"",T0.""RefDate"""
                '    StrQuery += vbCrLf + "FROM JDT1 T0 INNER JOIN OCRD T1 ON T0.""ShortName"" = T1.""CardCode"" " 'AND T1.""CardType"" = 'C'
                '    StrQuery += vbCrLf + "WHERE T0.""RefDate"" between '" & RecFromDate.ToString("yyyyMMdd") & "' and '" & RecToDate.ToString("yyyyMMdd") & "' "
                '    StrQuery += vbCrLf + " and t1.""CardCode"" in (" & Trim(CardCode) & ") and T0.""IntrnMatch"" = '0' and ((T0.""BalDueDeb"") + ( T0.""BalDueCred""))<>0 order by T0.""RefDate"",T0.""TransId"" " 'ORDER BY T0.""TransId"""
                '    objRS.DoQuery(StrQuery)
                '    objaddon.objglobalmethods.WriteErrorLog("Internal Rec Query: " + StrQuery)
                '    objaddon.objglobalmethods.WriteErrorLog("Upload Query: " + GetRecData)
                '    'objaddon.objapplication.StatusBar.SetText("Validating 1...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '    If objRS.RecordCount > 0 Then
                '        Dim Row As Integer
                '        While Not objRecordset.EoF  'Customized data
                '            objRS.MoveFirst()
                '            RecAmount = CDbl(objRecordset.Fields.Item("DebAmount").Value)
                '            For Rec As Integer = 0 To objRS.RecordCount - 1 ' Reconcil screen data
                '                Amount = IIf(Val(objRS.Fields.Item("BalDueDeb").Value) = 0, -CDbl(objRS.Fields.Item("BalDueCred").Value), CDbl(objRS.Fields.Item("BalDueDeb").Value))
                '                Ref2 = IIf(objRS.Fields.Item("Ref2").Value.ToString.ToUpper() = "", "", objRS.Fields.Item("Ref2").Value.ToString.ToUpper())
                '                Ref3 = IIf(objRS.Fields.Item("Ref3").Value.ToString.ToUpper() = "", "", objRS.Fields.Item("Ref3").Value.ToString.ToUpper())
                '                ''objaddon.objglobalmethods.WriteErrorLog("Row: " + CStr(Rec) + "Amount: " + CStr(Amount) + "RecAmount: " + RecAmount + "Ref2: " + Ref2 + "Ref3: " + Ref3)
                '                ''Dim ass As String = IIf(objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper = "", 0, objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper)
                '                ''If (IIf(objRS.Fields.Item("Ref1").Value.ToString.ToUpper = "", 0, objRS.Fields.Item("Ref1").Value.ToString.ToUpper) = objRecordset.Fields.Item("Ref").Value.ToString.ToUpper Or IIf(objRS.Fields.Item("Ref3").Value.ToString.ToUpper = "", 0, objRS.Fields.Item("Ref3").Value.ToString.ToUpper) = objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper) And objRS.Fields.Item("LineMemo").Value.ToString.ToUpper = objRecordset.Fields.Item("Memo").Value.ToString.ToUpper And Amount = RecAmount Then
                '                ''If (IIf(objRS.Fields.Item("Ref3").Value.ToString.ToUpper = "", "0", objRS.Fields.Item("Ref3").Value.ToString.ToUpper) = IIf(objRecordset.Fields.Item("Ref").Value.ToString.ToUpper = "", "0", objRecordset.Fields.Item("Ref").Value.ToString.ToUpper) Or IIf(objRS.Fields.Item("Ref3").Value.ToString.ToUpper = "", "0", objRS.Fields.Item("Ref3").Value.ToString.ToUpper) = objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper) And objRS.Fields.Item("LineMemo").Value.ToString.ToUpper = objRecordset.Fields.Item("Memo").Value.ToString.ToUpper And Amount = RecAmount Then
                '                ''If IIf(objRS.Fields.Item("Ref1").Value.ToString.ToUpper = "", "0", objRS.Fields.Item("Ref1").Value.ToString.ToUpper) = IIf(objRecordset.Fields.Item("Ref").Value.ToString.ToUpper = "", "0", objRecordset.Fields.Item("Ref").Value.ToString.ToUpper) And IIf(objRS.Fields.Item("Ref2").Value.ToString.ToUpper = "", "0", objRS.Fields.Item("Ref2").Value.ToString.ToUpper) = IIf(objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper = "0", "0", objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper) And objRS.Fields.Item("LineMemo").Value.ToString.ToUpper = objRecordset.Fields.Item("Memo").Value.ToString.ToUpper And Amount = RecAmount Then
                '                'StrQuery = objRecordset.Fields.Item("Ref").Value.ToString.ToUpper()
                '                'StrQuery = objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper()
                '                'StrQuery = objRecordset.Fields.Item("Memo").Value.ToString.ToUpper()

                '                If (Ref2 = objRecordset.Fields.Item("Ref").Value.ToString.ToUpper() And Ref3 = objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper() Or objRS.Fields.Item("LineMemo").Value.ToString.ToUpper() = objRecordset.Fields.Item("Memo").Value.ToString.ToUpper()) And Amount = RecAmount Then
                '                    'objaddon.objapplication.StatusBar.SetText("Validating 2...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '                    Row = CInt(objRS.Fields.Item("#").Value)
                '                    objaddon.objglobalmethods.WriteErrorLog("Match Row : " + CStr(Row))
                '                    If objMatrix.Columns.Item("120000002").Cells.Item(Row).Specific.checked = True Then Continue For
                '                    objMatrix.Columns.Item("120000002").Cells.Item(Row).Specific.checked = True
                '                    Flag = True
                '                    Exit For
                '                End If
                '                objRS.MoveNext()
                '            Next
                '            objRecordset.MoveNext()
                '        End While
                '    End If
                '    'While Not objRecordset.EoF
                '    '    For i = 1 To objMatrix.VisualRowCount
                '    '        custName = objMatrix.Columns.Item("120000030").Cells.Item(i).Specific.String
                '    '        DealID = objMatrix.Columns.Item("120000013").Cells.Item(i).Specific.String
                '    '        'GetPrice = Regex.Replace(objMatrix.Columns.Item("120000014").Cells.Item(i).Specific.String, "[^0-9]", "")
                '    '        price = Split(objMatrix.Columns.Item("120000014").Cells.Item(i).Specific.String, Left(objMatrix.Columns.Item("120000014").Cells.Item(i).Specific.String, 5))
                '    '        Amount = CDbl(price(1).ToString.Replace(",", "").Replace("(", "").Replace(")", ""))
                '    '        RecAmount = objRecordset.Fields.Item("DebAmount").Value
                '    '        If (DealID.ToUpper = objRecordset.Fields.Item("Memo").Value.ToString.ToUpper Or DealID.ToUpper = objRecordset.Fields.Item("Memo2").Value.ToString.ToUpper) And custName.ToUpper = objRecordset.Fields.Item("Ref").Value.ToString.ToUpper And Amount = RecAmount Then
                '    '            objMatrix.Columns.Item("120000002").Cells.Item(i).Specific.checked = True
                '    '        End If
                '    '    Next i
                '    '    objRecordset.MoveNext()
                '    'End While
                'Else
                '    objaddon.objapplication.StatusBar.SetText("No Data found for the selected range...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'End If
                'If Flag = True Then
                '    objaddon.objapplication.StatusBar.SetText("First level auto reconciliation completed...Please validate second level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'Else
                '    objaddon.objapplication.StatusBar.SetText("No Matching Found...Please Check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try

        End Sub

        Private Sub InternRec()
            Try
                Dim oParam As SAPbobsCOM.InternalReconciliationParams
                Dim oReconService As SAPbobsCOM.InternalReconciliationsService
                Dim oOposting As SAPbobsCOM.InternalReconciliationOpenTrans
                oReconService = objaddon.objcompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.InternalReconciliationsService)
                oParam = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams)
                oOposting = oReconService.GetDataInterface(SAPbobsCOM.InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans) ' oReconService.Add(oOposting) 
                With oOposting
                    .CardOrAccount = SAPbobsCOM.CardOrAccountEnum.coaCard
                    .coaCard.InternalReconciliationOpenTransRows.Add()
                    .InternalReconciliationOpenTransRows.Item(0).Selected = SAPbobsCOM.BoYesNoEnum.tYES
                    .InternalReconciliationOpenTransRows.Item(0).TransId = 803
                    .InternalReconciliationOpenTransRows.Item(0).TransRowId = 1
                    .InternalReconciliationOpenTransRows.Item(0).ReconcileAmount = 738.38 'gridRecon.DataTable.GetValue("Actual Amount", gridRecon.GetDataTableRowIndex(i))
                    .InternalReconciliationOpenTransRows.Add()
                    .InternalReconciliationOpenTransRows.Item(1).Selected = SAPbobsCOM.BoYesNoEnum.tYES
                    .InternalReconciliationOpenTransRows.Item(1).TransId = 4510
                    .InternalReconciliationOpenTransRows.Item(1).TransRowId = 0
                    .InternalReconciliationOpenTransRows.Item(1).ReconcileAmount = -738.38
                End With
                Try
                    oParam = oReconService.Add(oOposting)
                Catch ex As Exception
                End Try
            Catch ex As Exception
            End Try
        End Sub

        Private Sub TestFunct()
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
                For i As Integer = 1 To objMatrix.VisualRowCount
                    For Each row In openTrans.InternalReconciliationOpenTransRows
                        If objMatrix.Columns.Item("1200000013").Cells.Item(i).Specific.String = "1" Then

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

        Private Sub Button1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            Try
                'If Not objaddon.FormExist("BPREC") Then
                '    Dim activeform As New FrmDataFromFinCustomer
                '    activeform.Show()
                'End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub BankReconciliation()
            Try
                Dim oCompanyService As SAPbobsCOM.CompanyService = objaddon.objcompany.GetCompanyService()
                Dim ExtReconSvc As SAPbobsCOM.ExternalReconciliationsService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ExternalReconciliationsService)
                Dim ExtReconciliation As SAPbobsCOM.ExternalReconciliation = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliation)
                ExtReconciliation.ReconciliationAccountType = SAPbobsCOM.ReconciliationAccountTypeEnum.rat_GLAccount

                Dim bnkPage As SAPbobsCOM.BankPages = objaddon.objcompany.GetBusinessObject(BoObjectTypes.oBankPages)
                bnkPage.AccountCode = "10602010"
                bnkPage.CreditAmount = 500
                bnkPage.DueDate = DateTime.Now
                bnkPage.DocNumberType = BoBpsDocTypes.bpdt_DocNum
                'bnkPage.PaymentReference = "UTR DATE"
                bnkPage.Reference = "Auto Post"
                bnkPage.Add()

                Dim str As String = objaddon.objcompany.GetNewObjectKey()
                Dim sequence As Integer = Convert.ToInt32(str.Split(vbTab)(1))

                Dim bstLine1 As SAPbobsCOM.ReconciliationBankStatementLine = ExtReconciliation.ReconciliationBankStatementLines.Add()
                bstLine1.BankStatementAccountCode = "10602010"
                bstLine1.Sequence = sequence ' 0

                Dim jeLine3 As SAPbobsCOM.ReconciliationJournalEntryLine = ExtReconciliation.ReconciliationJournalEntryLines.Add()
                jeLine3.TransactionNumber = "9425"
                jeLine3.LineNumber = 1
                ExtReconSvc.Reconcile(ExtReconciliation)
                ''''''''''''
                'Dim ExtReconParam As SAPbobsCOM.ExternalReconciliationParams = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliationParams)
                'ExtReconParam.AccountCode = "10602010"
                ''ExtReconParam.ReconciliationNo = "2"
                'ExtReconciliation = ExtReconSvc.GetReconciliation(ExtReconParam)

                'Dim ExtReconsParamsCollection As SAPbobsCOM.ExternalReconciliationsParamsCollection = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliationsParamsCollection)
                'Dim ExtReconFilteredParams As SAPbobsCOM.ExternalReconciliationFilterParams = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliationFilterParams)
                'ExtReconFilteredParams.ReconciliationAccountType = SAPbobsCOM.ReconciliationAccountTypeEnum.rat_GLAccount
                'ExtReconFilteredParams.AccountCodeFrom = "10602010"
                'ExtReconFilteredParams.AccountCodeTo = "10602010"
                'Dim Date1 As Date = Date.ParseExact("20231026", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'Dim Date2 As Date = Date.ParseExact("20231031", "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                'ExtReconFilteredParams.ReconciliationDateFrom = Date2 ' "26/10/2023"
                'ExtReconFilteredParams.ReconciliationDateTo = Date2 '"31/10/2023"
                'ExtReconFilteredParams.ReconciliationNoFrom = 1
                'ExtReconFilteredParams.ReconciliationNoTo = 2
                'ExtReconsParamsCollection = ExtReconSvc.GetReconciliationList(ExtReconFilteredParams)

                'Dim ExtReconciliation1 As SAPbobsCOM.ExternalReconciliation = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliation)
                'ExtReconciliation1.ReconciliationAccountType = SAPbobsCOM.ReconciliationAccountTypeEnum.rat_GLAccount

                'Dim bnkPage As SAPbobsCOM.BankPages = objaddon.objcompany.GetBusinessObject(BoObjectTypes.oBankPages)
                'bnkPage.AccountCode = "10602010"
                'bnkPage.CreditAmount = 500
                'bnkPage.DueDate = DateTime.Now
                'bnkPage.DocNumberType = BoBpsDocTypes.bpdt_DocNum
                'bnkPage.PaymentReference = "UTR DATE"
                'bnkPage.Reference = "AUTO PAYMENT"
                'bnkPage.Add()
                'Dim str As String = objaddon.objcompany.GetNewObjectKey()
                'Dim sequence As Integer = Convert.ToInt32(str.Split(vbTab)(1))

                'Dim bstLine1 As SAPbobsCOM.ReconciliationBankStatementLine = ExtReconciliation1.ReconciliationBankStatementLines.Add()
                'bstLine1.BankStatementAccountCode = "10602010"
                'bstLine1.Sequence = sequence ' 0

                'Dim jeLine3 As SAPbobsCOM.ReconciliationJournalEntryLine = ExtReconciliation1.ReconciliationJournalEntryLines.Add()
                'jeLine3.TransactionNumber = "9425"
                'jeLine3.LineNumber = 1

                'ExtReconSvc.Reconcile(ExtReconciliation1)

                ''For Each ExtReconParam1 As SAPbobsCOM.ExternalReconciliationParams In ExtReconsParamsCollection
                ''    ExtReconSvc.CancelReconciliation(ExtReconParam1)
                ''Next
            Catch ex As Exception

            End Try
        End Sub

        Private Sub ReadExcelData(ByVal FileName As String)
            Try
                objaddon.objapplication.SetStatusBarMessage("Looking for the Excel Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
                Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim excelRng As Microsoft.Office.Interop.Excel.Range
                Dim j As Integer = 1
                Dim Flag As Boolean = False
                Dim DealID, InvoiceNo, CustName, MatDealID, MatInvoiceNo, MatCustName, strquery As String
                Dim Amount, RecAmount As Double
                Dim price() As String
                Try
                    ExcelWorkbook = ExcelApp.Workbooks.Open(FileName)
                    ExcelWorkSheet = ExcelWorkbook.ActiveSheet
                    'excelRng = ExcelWorkSheet.Range("A1")
                    excelRng = ExcelWorkSheet.UsedRange
                    objaddon.objapplication.SetStatusBarMessage("Excel Reading please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    'objform.Freeze(True)
                    objMatrix = objform.Items.Item("120000039").Specific

                    If excelRng.Cells(1, 1).Value.ToString.ToUpper() = "DEALID" And excelRng.Cells(1, 2).Value.ToString.ToUpper() = "INVOICENO" And excelRng.Cells(1, 3).Value.ToString.ToUpper() = "CUSTOMER NAME" And excelRng.Cells(1, 4).Value.ToString.ToUpper() = "CREDIT AMOUNT" Then
                        For ExcIndex = 2 To excelRng.Rows.Count
                            objaddon.objapplication.StatusBar.SetText("Reading Excel Row: " + CStr(ExcIndex), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            If Not excelRng.Cells(ExcIndex, 1).Value Is Nothing Then DealID = CStr(excelRng.Cells(ExcIndex, 1).Value).ToUpper Else DealID = "" 'Ref3 Field in Internal Reconciliation screen
                            If Not excelRng.Cells(ExcIndex, 2).Value Is Nothing Then InvoiceNo = CStr(excelRng.Cells(ExcIndex, 2).Value).ToUpper Else InvoiceNo = "" 'Ref2 Field in Internal Reconciliation screen
                            If Not excelRng.Cells(ExcIndex, 3).Value Is Nothing Then CustName = CStr(excelRng.Cells(ExcIndex, 3).Value).ToUpper Else CustName = "" 'Details Field in Internal Reconciliation screen
                            If Not excelRng.Cells(ExcIndex, 4).Value Is Nothing Then Amount = CDbl(excelRng.Cells(ExcIndex, 4).Value) Else Amount = 0

                            For Matindex = 1 To objMatrix.VisualRowCount
                                strquery = objMatrix.Columns.Item("120000027").Cells.Item(Matindex).Specific.Value
                                price = Split(objMatrix.Columns.Item("120000027").Cells.Item(Matindex).Specific.String, Left(objMatrix.Columns.Item("120000027").Cells.Item(Matindex).Specific.String, 5))
                                RecAmount = CDbl(price(1).ToString.Replace(",", "").Replace("(", "").Replace(")", ""))
                                MatDealID = objMatrix.Columns.Item("120000013").Cells.Item(Matindex).Specific.String.ToUpper()
                                MatInvoiceNo = objMatrix.Columns.Item("120000012").Cells.Item(Matindex).Specific.String.ToUpper()
                                MatCustName = objMatrix.Columns.Item("120000030").Cells.Item(Matindex).Specific.String.ToUpper()
                                If (DealID = MatDealID And InvoiceNo = MatInvoiceNo Or CustName = MatCustName) And Amount = RecAmount Then
                                    If objMatrix.Columns.Item("120000002").Cells.Item(Matindex).Specific.checked = True Then Exit For
                                    objMatrix.Columns.Item("120000002").Cells.Item(Matindex).Specific.checked = True
                                    Flag = True
                                    Exit For
                                End If
                            Next

                        Next

                    Else
                        objaddon.objapplication.StatusBar.SetText("Expected ColumnName Not found...Please check the excel format", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    'objform.Freeze(False)
                    objaddon.objapplication.StatusBar.SetText("Excel Readed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    If Flag = True Then
                        objaddon.objapplication.StatusBar.SetText("First level auto reconciliation completed...Please validate second level.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Else
                        objaddon.objapplication.StatusBar.SetText("No Matching Found...Please Check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                    ExcelApp.ActiveWorkbook.Close()
                Catch ex As Exception
                    'objform.Freeze(False)
                    ExcelApp.ActiveWorkbook.Close()
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End Try
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                If EditText0.Value = "" Then
                    objaddon.objapplication.StatusBar.SetText("Please select a excel file...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try
        End Sub


    End Class
End Namespace
