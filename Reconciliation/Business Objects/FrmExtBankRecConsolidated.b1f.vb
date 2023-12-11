Option Strict Off
Option Explicit On

Imports System.Drawing
Imports SAPbobsCOM
Imports SAPbouiCOM.Framework

Namespace Reconciliation
    <FormAttribute("CBRS", "Business Objects/FrmExtBankRecConsolidated.b1f")>
    Friend Class FrmExtBankRecConsolidated
        Inherits UserFormBase

        Public WithEvents objform As SAPbouiCOM.Form
        Public WithEvents odbdsHeader, odbdsDetails As SAPbouiCOM.DBDataSource
        Private WithEvents pCFL As SAPbouiCOM.ISBOChooseFromListEventArg
        Dim StrQuery As String
        Dim objRs As SAPbobsCOM.Recordset
        Private WithEvents objCheck As SAPbouiCOM.CheckBox

        Public Sub New()
        End Sub


        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lacctcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("tacctcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("lacctname").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("tacctname").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("lno").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox0 = CType(Me.GetItem("series").Specific, SAPbouiCOM.ComboBox)
            Me.EditText2 = CType(Me.GetItem("tdocnum").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("ldocdate").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("tdocdate").Specific, SAPbouiCOM.EditText)
            Me.StaticText4 = CType(Me.GetItem("lendbal").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("tendbal").Specific, SAPbouiCOM.EditText)
            Me.Button2 = CType(Me.GetItem("bgetdata").Specific, SAPbouiCOM.Button)
            Me.Folder0 = CType(Me.GetItem("fcont").Specific, SAPbouiCOM.Folder)
            Me.Folder1 = CType(Me.GetItem("fldr").Specific, SAPbouiCOM.Folder)
            Me.Matrix0 = CType(Me.GetItem("mcont").Specific, SAPbouiCOM.Matrix)
            Me.StaticText5 = CType(Me.GetItem("lrem").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("trem").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lclrbal").Specific, SAPbouiCOM.StaticText)
            Me.EditText6 = CType(Me.GetItem("tclrbal").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("ldiff").Specific, SAPbouiCOM.StaticText)
            Me.EditText7 = CType(Me.GetItem("tdiff").Specific, SAPbouiCOM.EditText)
            Me.Button3 = CType(Me.GetItem("breco").Specific, SAPbouiCOM.Button)
            Me.EditText8 = CType(Me.GetItem("tentry").Specific, SAPbouiCOM.EditText)
            Me.StaticText8 = CType(Me.GetItem("lstat").Specific, SAPbouiCOM.StaticText)
            Me.ComboBox1 = CType(Me.GetItem("cstat").Specific, SAPbouiCOM.ComboBox)
            Me.StaticText9 = CType(Me.GetItem("lenddate").Specific, SAPbouiCOM.StaticText)
            Me.EditText9 = CType(Me.GetItem("tenddate").Specific, SAPbouiCOM.EditText)
            Me.StaticText10 = CType(Me.GetItem("lrecno").Specific, SAPbouiCOM.StaticText)
            Me.EditText10 = CType(Me.GetItem("trecono").Specific, SAPbouiCOM.EditText)
            Me.LinkedButton0 = CType(Me.GetItem("lkcode").Specific, SAPbouiCOM.LinkedButton)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform.Left = (objaddon.objapplication.Desktop.Width - objform.MaxWidth) / 2
                objform.Top = (objaddon.objapplication.Desktop.Height - objform.MaxHeight) / 4

                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_OBRS")
                odbdsDetails = objform.DataSources.DBDataSources.Item("@AT_BRS1")
                objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_OBRS")
                objform.Items.Item("tdocdate").Specific.String = "A"
                objform.Items.Item("trem").Specific.String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss")
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "series", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocnum", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cstat", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tdocdate", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tacctcode", True, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tacctname", False, True, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "bgetdata", True, False, False)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "breco", False, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "mcont", True, False, True)
                objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "trecono", False, True, False)
                Matrix0.Columns.Item("clramt").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                objform.EnableMenu("1283", False) 'Remove
                objform.EnableMenu("1286", False) 'Close
                Folder0.Item.Click()
                objform.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized
                'StrQuery = "Select ""USERID"",""TPLId"" from OUSR Where ""USER_CODE""='" & objaddon.objcompany.UserName & "'"
                'objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'objRs.DoQuery(StrQuery)
                'objaddon.objglobalmethods.Update_UserFormSettings_UDF(objform, "-" + objform.TypeEx, Convert.ToInt32(objRs.Fields.Item("USERID").Value), Convert.ToInt32(objRs.Fields.Item("TPLId").Value))

            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox0 As SAPbouiCOM.ComboBox
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Folder0 As SAPbouiCOM.Folder
        Private WithEvents Folder1 As SAPbouiCOM.Folder
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText7 As SAPbouiCOM.EditText
        Private WithEvents Button3 As SAPbouiCOM.Button
        Private WithEvents EditText8 As SAPbouiCOM.EditText
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText
        Private WithEvents ComboBox1 As SAPbouiCOM.ComboBox
        Private WithEvents StaticText9 As SAPbouiCOM.StaticText
        Private WithEvents EditText9 As SAPbouiCOM.EditText
        Private WithEvents StaticText10 As SAPbouiCOM.StaticText
        Private WithEvents EditText10 As SAPbouiCOM.EditText
        Private WithEvents LinkedButton0 As SAPbouiCOM.LinkedButton

#End Region

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("CBRS", pVal.FormTypeCount)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button2_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button2.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then BubbleEvent = False : Return
                If Button2.Item.Enabled = False Then BubbleEvent = False : Return
                If EditText0.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("Account Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If CDbl(EditText4.Value) = 0 Then
                    objaddon.objapplication.SetStatusBarMessage("Statement ending balance is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If EditText9.Value = "" Then
                    objaddon.objapplication.SetStatusBarMessage("End date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                StrQuery = "SELECT ROW_NUMBER() OVER (ORDER BY A.""DocDate"",A.""TransId"") AS ""LineId"",* FROM"
                StrQuery += vbCrLf + "(SELECT 'N' AS ""Selected"",T1.""TransId"",T1.""Account"",T1.""TransType"" AS ""ObjType"",T1.""LineMemo"" as ""LineMemo"",T1.""BaseRef"" AS ""DocNum"",T1.""CreatedBy"" as ""DocEntry"","
                StrQuery += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'IN' WHEN T1.""TransType""='14' THEN 'CN' WHEN T1.""TransType""='203' or T1.""TransType""='204' THEN 'DT' WHEN T1.""TransType""='18' THEN 'PU' WHEN T1.""TransType""='19' THEN 'PC'"
                StrQuery += vbCrLf + "WHEN T1.""TransType""='24' THEN 'RC' WHEN T1.""TransType""='46' THEN 'PS' WHEN T1.""TransType""='25' THEN 'DP'  Else 'JE' END AS ""Origin"","
                StrQuery += vbCrLf + "CASE WHEN T1.""TransType""='13' THEN 'A/R Invoice' WHEN T1.""TransType""='14' THEN 'A/R Credit Memo' WHEN T1.""TransType""='203' THEN 'A/R DownPayment' WHEN T1.""TransType""='18' THEN 'A/P Invoice'"
                StrQuery += vbCrLf + "WHEN T1.""TransType""='19' THEN 'A/R Credit Memo' WHEN T1.""TransType""='24' THEN 'Incoming Payment' WHEN T1.""TransType""='46' THEN 'Outgoing Payment'  WHEN T1.""TransType""='25' THEN 'Deposit' ELSE 'Journal Entry' END AS ""DocType"","
                StrQuery += vbCrLf + "SUM(CASE WHEN T1.""Credit""<>0  THEN  -T1.""Credit"" ELSE T1.""Debit"" END) AS ""DocTotal"",SUM(CASE WHEN T1.""BalDueCred""<>0  THEN -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END) AS ""Balance"","
                StrQuery += vbCrLf + "To_Varchar(T0.""RefDate"",'yyyyMMdd') AS ""DocDate"",T1.""BPLId"",(SELECT ""BPLName"" FROM OBPL WHERE ""BPLId""=T1.""BPLId"") AS ""BPLName"",T0.""Ref1"",T0.""Ref2"",T0.""Ref3"""
                StrQuery += vbCrLf + "FROM OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" where T1.""DprId"" is null and T1.""ExtrMatch""=0"
                StrQuery += vbCrLf + "group by T1.""TransId"",T1.""Account"",T1.""TransType"",T1.""LineMemo"",T1.""BaseRef"",T1.""CreatedBy"",T0.""RefDate"",T1.""BPLId"",T0.""Ref1"",T0.""Ref2"",T0.""Ref3"",T1.""ExtrMatch"""
                StrQuery += vbCrLf + ") A "
                StrQuery += vbCrLf + "WHERE A.""DocDate""<='" & EditText9.Value & "' and A.""BPLId"" in (Select T0.""BPLId"" from OBPL T0 join USR6 T1 on T0.""BPLId""=T1.""BPLId"" where T1.""UserCode""='" & objaddon.objcompany.UserName & "' and T0.""Disabled""<>'Y') " ' DateTime.Now.Date.ToString("yyyyMMdd")
                StrQuery += vbCrLf + "and A.""Account""='" & EditText0.Value & "' ORDER BY A.""DocDate"",A.""TransId"""
                LoadData(StrQuery)

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button3_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button3.ClickBefore
            Try
                If Button3.Item.Enabled = False Then BubbleEvent = False : Return
                If ComboBox1.Selected.Value = "C" Then BubbleEvent = False : Button3.Item.Enabled = False : Return
                If CDbl(EditText7.Value) <> 0 Then BubbleEvent = False : objaddon.objapplication.StatusBar.SetText("Reconciliation is only possible when the difference is zero...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button3_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try
                If objaddon.objapplication.MessageBox("Do you want to reconcile the selected transactions?", 2, "Yes", "No") <> 1 Then Exit Sub
                BankReconciliation(EditText0.Value, EditText4.Value)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText0_ChooseFromListBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles EditText0.ChooseFromListBefore
            Try
                GL_CFLcondition(pVal, "cflgl")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub EditText0_ChooseFromListAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText0.ChooseFromListAfter
            Try
                If pVal.ActionSuccess = False Then Exit Sub
                pCFL = pVal
                If Not pCFL.SelectedObjects Is Nothing Then
                    Try
                        odbdsHeader.SetValue("U_AcctCode", 0, pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value)
                        odbdsHeader.SetValue("U_AcctName", 0, pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value)
                        Matrix0.Clear()
                        'EditText0.Value = pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value
                    Catch ex As Exception
                    End Try

                End If
            Catch ex As Exception
            Finally
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try

        End Sub

        Private Sub ComboBox0_ComboSelectAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles ComboBox0.ComboSelectAfter
            Try
                If objform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then Exit Sub
                odbdsHeader.SetValue("DocNum", 0, objaddon.objglobalmethods.GetDocNum("AT_OBRS", CInt(ComboBox0.Selected.Value)))
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Matrix0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.PressedAfter
            Try
                objCheck = Matrix0.Columns.Item("cleared").Cells.Item(pVal.Row).Specific
                If pVal.ColUID = "cleared" Then
                    If objCheck.Checked = True Then
                        'Matrix1.SelectRow(pVal.Row, True, True)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                        Matrix0.SetCellWithoutValidation(pVal.Row, "clramt", CDbl(Matrix0.Columns.Item("payment").Cells.Item(pVal.Row).Specific.String))
                    Else
                        'Matrix1.SelectRow(pVal.Row, False, True)
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Matrix0.Item.BackColor)
                        Matrix0.SetCellWithoutValidation(pVal.Row, "clramt", 0)
                    End If
                    Calculate_Total()
                End If
            Catch ex As Exception
            End Try


        End Sub

        Private Sub Button0_PressedAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.PressedAfter
            Try
                If pVal.ActionSuccess = True And objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_OBRS")
                    objform.Items.Item("tdocdate").Specific.String = "A"
                    objform.Items.Item("trem").Specific.String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss")
                End If
            Catch ex As Exception
            End Try


        End Sub

        Private Sub EditText4_ValidateAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles EditText4.ValidateAfter
            Try
                If CDbl(EditText4.Value) = 0 Then Exit Sub
                Calculate_Total()
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                If ComboBox1.Selected.Value = "C" Then
                    Matrix0.Item.Enabled = False
                    If Button3.Item.Enabled = True Then Button3.Item.Enabled = False
                    EditText4.Item.Enabled = False
                    EditText9.Item.Enabled = False
                    DeleteRow()
                End If
                For i As Integer = 1 To Matrix0.VisualRowCount
                    objCheck = Matrix0.Columns.Item("cleared").Cells.Item(i).Specific
                    If objCheck.Checked = True Then
                        If objCheck.Checked = True Then
                            Matrix0.CommonSetting.SetRowBackColor(i, Color.PeachPuff.ToArgb)
                        End If
                    End If
                Next

            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                Dim selectionFlag As Boolean = False
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                If Matrix0.VisualRowCount = 0 Then
                    objaddon.objapplication.SetStatusBarMessage("Row is Missing", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                If CDbl(EditText7.Value) <> 0 Then
                    objaddon.objapplication.SetStatusBarMessage("Difference should be zero...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                    BubbleEvent = False : Exit Sub
                End If
                Matrix0.FlushToDataSource()
                For iRow As Integer = 0 To odbdsDetails.Size - 1
                    If odbdsDetails.GetValue("U_Cleared", iRow) = "Y" Then
                        selectionFlag = True : Exit For
                    End If
                Next
                If selectionFlag = False Then BubbleEvent = False : objaddon.objapplication.SetStatusBarMessage("Select a row...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform.Freeze(True)
                Matrix0.AutoResizeColumns()
                objform.Freeze(False)
            Catch ex As Exception
                objform.Freeze(False)
            End Try
        End Sub

#Region "Functions"

        Private Sub GL_CFLcondition(ByVal pVal As SAPbouiCOM.SBOItemEventArg, ByVal CFLID As String)
            If pVal.ActionSuccess = True Then Exit Sub
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = objform.ChooseFromLists.Item(CFLID)
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()

                oCond = oConds.Add()
                oCond.Alias = "Postable"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "LocManTran"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "N"
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCond = oConds.Add()
                oCond.Alias = "FrozenFor"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "N"
                oCFL.SetConditions(oConds)
            Catch ex As Exception
                'SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Choose FromList Filter Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End Try
        End Sub

        Private Function LoadData(ByVal Query As String) As Boolean
            Try
                If Query = "" Then Exit Function
                objRs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRs.DoQuery(Query)
                Matrix0.Clear()
                odbdsDetails.Clear()
                If objRs.RecordCount > 0 Then
                    objaddon.objapplication.StatusBar.SetText("Loading data Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objform.Freeze(True)
                    While Not objRs.EoF
                        Matrix0.AddRow()
                        Matrix0.GetLineData(Matrix0.VisualRowCount)
                        odbdsDetails.SetValue("LineId", 0, objRs.Fields.Item("LineId").Value.ToString)
                        odbdsDetails.SetValue("U_Cleared", 0, objRs.Fields.Item("Selected").Value.ToString)
                        odbdsDetails.SetValue("U_Type", 0, objRs.Fields.Item("Origin").Value.ToString)
                        odbdsDetails.SetValue("U_TransId", 0, objRs.Fields.Item("TransId").Value.ToString)
                        odbdsDetails.SetValue("U_Date", 0, objRs.Fields.Item("DocDate").Value)
                        odbdsDetails.SetValue("U_Ref1", 0, objRs.Fields.Item("Ref1").Value.ToString)
                        odbdsDetails.SetValue("U_Ref2", 0, objRs.Fields.Item("Ref2").Value.ToString)
                        odbdsDetails.SetValue("U_Ref3", 0, objRs.Fields.Item("Ref3").Value.ToString)
                        odbdsDetails.SetValue("U_PayAmt", 0, objRs.Fields.Item("DocTotal").Value.ToString)
                        odbdsDetails.SetValue("U_ClrAmt", 0, 0)
                        odbdsDetails.SetValue("U_Remarks", 0, objRs.Fields.Item("LineMemo").Value.ToString)
                        Matrix0.CommonSetting.SetRowBackColor(Matrix0.VisualRowCount, Matrix0.Item.BackColor)
                        'objform.DataSources.UserDataSources.Item("UD_0").Value = objRs.Fields.Item("Balance").Value.ToString
                        Matrix0.SetLineData(Matrix0.VisualRowCount)
                        objRs.MoveNext()
                    End While

                    'Matrix0.LoadFromDataSource()
                    objform.Refresh()
                    Matrix0.AutoResizeColumns()
                    objaddon.objapplication.StatusBar.SetText("Data Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    objform.Freeze(False)
                    Return True
                Else
                    objaddon.objapplication.StatusBar.SetText("No records found...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            Catch ex As Exception
                objform.Freeze(False)
                Return False
            End Try
        End Function

        Private Sub Calculate_Total()
            Try
                Dim value, diffAmt As Decimal
                Matrix0.FlushToDataSource()
                For iRow As Integer = 0 To odbdsDetails.Size - 1
                    If odbdsDetails.GetValue("U_Cleared", iRow) = "Y" Then
                        value = value + CDec(odbdsDetails.GetValue("U_PayAmt", iRow))
                    End If
                Next

                odbdsHeader.SetValue("U_ClearTot", 0, value) 'Total
                diffAmt = value - CDec(EditText4.Value)
                odbdsHeader.SetValue("U_DiffAmt", 0, diffAmt) 'Total
            Catch ex As Exception

            End Try
        End Sub

        Private Sub BankReconciliation(ByVal Acctcode As String, ByVal Amount As Double)
            Try
                Dim oCompanyService As SAPbobsCOM.CompanyService = objaddon.objcompany.GetCompanyService()
                Dim ExtReconSvc As SAPbobsCOM.ExternalReconciliationsService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.ExternalReconciliationsService)
                Dim ExtReconciliation As SAPbobsCOM.ExternalReconciliation = ExtReconSvc.GetDataInterface(SAPbobsCOM.ExternalReconciliationsServiceDataInterfaces.ersExternalReconciliation)
                ExtReconciliation.ReconciliationAccountType = SAPbobsCOM.ReconciliationAccountTypeEnum.rat_GLAccount
                objaddon.objapplication.StatusBar.SetText("Reconciling transactions Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim bnkPage As SAPbobsCOM.BankPages = objaddon.objcompany.GetBusinessObject(BoObjectTypes.oBankPages)
                bnkPage.AccountCode = Acctcode '"10602010"
                bnkPage.CreditAmount = Amount ' 500
                bnkPage.DueDate = DateTime.Now
                bnkPage.DocNumberType = BoBpsDocTypes.bpdt_DocNum
                'bnkPage.PaymentReference = "UTR DATE"
                bnkPage.Reference = "Auto Post"
                bnkPage.Add()

                Dim str As String = objaddon.objcompany.GetNewObjectKey()
                Dim sequence As Integer = Convert.ToInt32(str.Split(vbTab)(1))

                Dim bstLine1 As SAPbobsCOM.ReconciliationBankStatementLine = ExtReconciliation.ReconciliationBankStatementLines.Add()
                bstLine1.BankStatementAccountCode = Acctcode ' "10602010"
                bstLine1.Sequence = sequence ' 0
                Dim jeLine As SAPbobsCOM.ReconciliationJournalEntryLine
                For index = 1 To Matrix0.VisualRowCount
                    objCheck = Matrix0.Columns.Item("cleared").Cells.Item(index).Specific
                    If objCheck.Checked = True Then
                        StrQuery = "Select T1.""TransId"",T1.""Line_ID"",CASE WHEN T1.""BalDueCred""<>0  THEN -T1.""BalDueCred"" ELSE T1.""BalDueDeb"" END ""Balance"""
                        StrQuery += vbCrLf + "FROM OJDT T0 join JDT1 T1 ON T0.""TransId""=T1.""TransId"" Where T1.""Account""='" & Acctcode & "' and T1.""TransId""=" & Matrix0.Columns.Item("transid").Cells.Item(index).Specific.String & ""
                        objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        objRs.DoQuery(StrQuery)
                        If objRs.RecordCount > 0 Then
                            For rec = 0 To objRs.RecordCount - 1
                                jeLine = ExtReconciliation.ReconciliationJournalEntryLines.Add()
                                jeLine.TransactionNumber = objRs.Fields.Item(0).Value '"9425"
                                jeLine.LineNumber = objRs.Fields.Item(1).Value
                                objRs.MoveNext()
                            Next
                        End If
                    End If

                Next

                ExtReconSvc.Reconcile(ExtReconciliation)

                StrQuery = "Select ""BankMatch"" from OBNK Where ""AcctCode""='" & Acctcode & "' and ""Sequence""=" & sequence & ""
                Dim RecoNo As String = objaddon.objglobalmethods.getSingleValue(StrQuery)
                If RecoNo <> "" Then EditText10.Value = RecoNo 'Reco No.
                ComboBox1.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                objform.Items.Item("1").Click()
                Matrix0.Item.Enabled = False
                objform.ActiveItem = "trem"
                Button3.Item.Enabled = False
                EditText4.Item.Enabled = False
                EditText9.Item.Enabled = False
                objaddon.objapplication.StatusBar.SetText("Reconciled transactions successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("BankReconciliation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub DeleteRow()
            Try
                Dim Flag As Boolean = False
                Dim objSelect As SAPbouiCOM.CheckBox

                For i As Integer = Matrix0.VisualRowCount To 1 Step -1
                    objSelect = Matrix0.Columns.Item("cleared").Cells.Item(i).Specific
                    If objSelect.Checked = False Then
                        Matrix0.DeleteRow(i)
                        odbdsDetails.RemoveRecord(i - 1)
                        Flag = True
                    End If
                Next
                If Flag = True Then
                    For i As Integer = 1 To Matrix0.VisualRowCount
                        objSelect = Matrix0.Columns.Item("cleared").Cells.Item(i).Specific
                        If objSelect.Checked = True Then
                            Matrix0.Columns.Item("#").Cells.Item(i).Specific.String = i
                        End If
                    Next
                    objform.Freeze(False)
                    If objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'objForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                End If
            Catch ex As Exception
                objform.Freeze(False)
            Finally
            End Try
        End Sub

        Private Sub Matrix0_KeyDownAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.KeyDownAfter
            Try
                objCheck = Matrix0.Columns.Item("cleared").Cells.Item(pVal.Row).Specific
                If pVal.ColUID = "cleared" And pVal.CharPressed = 32 Then 'space key
                    If objCheck.Checked = True Then
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Color.PeachPuff.ToArgb)
                        Matrix0.SetCellWithoutValidation(pVal.Row, "clramt", CDbl(Matrix0.Columns.Item("payment").Cells.Item(pVal.Row).Specific.String))
                    Else
                        Matrix0.CommonSetting.SetRowBackColor(pVal.Row, Matrix0.Item.BackColor)
                        Matrix0.SetCellWithoutValidation(pVal.Row, "clramt", 0)
                    End If
                    Calculate_Total()
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub CancelBankReconciliation(ByVal AcctCode As String, ByVal RecoNo As String)
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

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("CancelBankReconciliation: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub



#End Region
    End Class
End Namespace
