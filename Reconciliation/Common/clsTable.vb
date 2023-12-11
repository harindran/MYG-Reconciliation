﻿Namespace Reconciliation

    Public Class clsTable

        Public Sub FieldCreation()
            'BP_Reconciliation()
            Bank_Reconciliation()
            AddFields("OITR", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo, 100, SAPbobsCOM.BoFldSubTypes.st_Link)
        End Sub



#Region "Document Data Creation"

        Private Sub BP_Reconciliation()
            AddFields("OBNK", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo, 100, SAPbobsCOM.BoFldSubTypes.st_Link)
            AddFields("JDT1", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo, 100, SAPbobsCOM.BoFldSubTypes.st_Link)
            AddTables("MIPL_OREC", "BP Reconciliation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("MIPL_REC1", "BP Reconciliation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("@MIPL_OREC", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_OREC", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo, 200, SAPbobsCOM.BoFldSubTypes.st_Link)

            AddFields("@MIPL_REC1", "CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_REC1", "DealID", "Deal ID1", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_REC1", "DealID2", "Deal ID2", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_REC1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_REC1", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@MIPL_REC1", "Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddUDO("BPREC", "BP_Reconciliation", SAPbobsCOM.BoUDOObjType.boud_Document, "MIPL_OREC", {"MIPL_REC1"}, {"DocEntry", "DocNum", "U_DocDate"}, True, True)
        End Sub

        Private Sub Bank_Reconciliation()
            AddTables("AT_OBRS", "Bank Reconciliation Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("AT_BRS1", "Bank Reconciliation Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("@AT_OBRS", "AcctCode", "Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)
            AddFields("@AT_OBRS", "AcctName", "Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@AT_OBRS", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@AT_OBRS", "EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@AT_OBRS", "EndBal", "Statement Balance", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@AT_OBRS", "ClearTot", "Cleared Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@AT_OBRS", "DiffAmt", "Difference Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@AT_OBRS", "RecoNo", "Reconciliation No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 15)

            AddFields("@AT_BRS1", "Cleared", "Cleared Book", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, , , "N", , {"Y,Yes", "N,No"})
            AddFields("@AT_BRS1", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@AT_BRS1", "TransId", "Transaction ID", SAPbobsCOM.BoFieldTypes.db_Numeric, 10, SAPbobsCOM.BoFldSubTypes.st_None)
            AddFields("@AT_BRS1", "Ref1", "Reference 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@AT_BRS1", "Ref2", "Reference 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@AT_BRS1", "Ref3", "Reference 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 100)
            AddFields("@AT_BRS1", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 100)
            AddFields("@AT_BRS1", "Date", "Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("@AT_BRS1", "PayAmt", "Payment Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("@AT_BRS1", "ClrAmt", "Cleared Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddUDO("AT_OBRS", "Bank Reconciliation", SAPbobsCOM.BoUDOObjType.boud_Document, "AT_OBRS", {"AT_BRS1"}, {"DocEntry", "DocNum"}, True, True, True)
        End Sub

#End Region

#Region "Master Data Creation"

        Private Sub GeneralSettings()
            AddTables("MIPL_GEN", "SubContracting General", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            'Tab 1

            'Tab 2
            AddFields("@MIPL_GEN", "ResEn", "Resources Enable", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            'Tab 3 
            AddFields("@MIPL_GEN", "DatePO", "Date PO", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "ItemBOM", "ItemCode BOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "Costing", "Costing", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "POItem", "PO Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)

            AddFields("@MIPL_GEN", "AutoPO", "Auto ProdOrder", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "RecLoad", "Receipt Load", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "APLoad", "APInvoice Load", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddFields("@MIPL_GEN", "SubScreen", "SubPO Screen", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, , , "N", True)
            AddUDO("GENSET", "Sub-Con General Settings", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIPL_GEN", {""}, {"Code", "Name"}, True, False)
        End Sub
        Private Sub SubContractingBOM()
            AddTables("MIPL_OBOM", "SubContracting BOM Header", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("MIPL_BOM1", "SubContracting BOM Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)

            AddFields("@MIPL_OBOM", "DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 12)
            AddFields("@MIPL_OBOM", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OBOM", "BOMType", "BOM Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OBOM", "WhseCode", "Warehouse Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OBOM", "Project", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_OBOM", "Avgplan", "Average Plan", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_OBOM", "Distrule", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)

            AddFields("@MIPL_BOM1", "Itemcode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_BOM1", "ItemDesc", "Item Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            AddFields("@MIPL_BOM1", "Type", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_BOM1", "Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("@MIPL_BOM1", "UOMName", "UOM Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_BOM1", "Whse", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_BOM1", "Distrule", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_BOM1", "Unitprice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_BOM1", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 6, SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("@MIPL_BOM1", "Comments", "Comments", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            AddUDO("SUBBOM", "SubContractingBOM", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIPL_OBOM", {"MIPL_BOM1"}, {"Code", "Name", "U_DocEntry", "U_Qty", "U_BOMType", "U_WhseCode"}, True, False)
        End Sub
        Private Sub Costing()
            AddTables("MIPL_SBGL", "Sub-GL", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("@MIPL_SBGL", "ItmGrp", "Sub ItemGroup", SAPbobsCOM.BoFieldTypes.db_Alpha, 20)
            AddFields("@MIPL_SBGL", "InvWhse", "Inventory Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            AddFields("@MIPL_SBGL", "WhsCode", "Sub Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "GoodIssue", "Sub GoodsIssue", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "GoodsReceipt", "Sub GoodReceipt", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "JEGLCode", "JE GLCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "JEGLName", "JE GLName", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            AddFields("@MIPL_SBGL", "BranchID", "Branch ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10)
            AddFields("@MIPL_SBGL", "BranchNam", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 40)
            AddUDO("SUBGL", "SubContractingGL", SAPbobsCOM.BoUDOObjType.boud_MasterData, "MIPL_SBGL", {""}, {"Code", "Name", "U_WhsCode"}, True, False)
        End Sub
#End Region

#Region "Table Creation Common Functions"

        Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
            Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
            Try
                oUserTablesMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                'Adding Table
                If Not oUserTablesMD.GetByKey(strTab) Then
                    oUserTablesMD.TableName = strTab
                    oUserTablesMD.TableDescription = strDesc
                    oUserTablesMD.TableType = nType

                    If oUserTablesMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription & strTab)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
                oUserTablesMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Sub AddFields(ByVal strTab As String, ByVal strCol As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoFieldTypes, _
                             Optional ByVal nEditSize As Integer = 10, Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, _
                              Optional ByVal defaultvalue As String = "", Optional ByVal Yesno As Boolean = False, Optional ByVal Validvalues() As String = Nothing)
            Dim oUserFieldMD1 As SAPbobsCOM.UserFieldsMD
            oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                'If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                '    strTab = "@" + strTab
                'End If
                If Not IsColumnExists(strTab, strCol) Then
                    'If Not oUserFieldMD1 Is Nothing Then
                    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    'End If
                    'oUserFieldMD1 = Nothing
                    'oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc
                    oUserFieldMD1.Name = strCol
                    oUserFieldMD1.Type = nType
                    oUserFieldMD1.SubType = nSubType
                    oUserFieldMD1.TableName = strTab
                    oUserFieldMD1.EditSize = nEditSize
                    oUserFieldMD1.Mandatory = Mandatory
                    oUserFieldMD1.DefaultValue = defaultvalue

                    If Yesno = True Then
                        oUserFieldMD1.ValidValues.Value = "Y"
                        oUserFieldMD1.ValidValues.Description = "Yes"
                        oUserFieldMD1.ValidValues.Add()
                        oUserFieldMD1.ValidValues.Value = "N"
                        oUserFieldMD1.ValidValues.Description = "No"
                        oUserFieldMD1.ValidValues.Add()
                    End If

                    Dim split_char() As String
                    If Not Validvalues Is Nothing Then
                        If Validvalues.Length > 0 Then
                            For i = 0 To Validvalues.Length - 1
                                If Trim(Validvalues(i)) = "" Then Continue For
                                split_char = Validvalues(i).Split(",")
                                If split_char.Length <> 2 Then Continue For
                                oUserFieldMD1.ValidValues.Value = split_char(0)
                                oUserFieldMD1.ValidValues.Description = split_char(1)
                                oUserFieldMD1.ValidValues.Add()
                            Next
                        End If
                    End If
                    Dim val As Integer
                    val = oUserFieldMD1.Add
                    If val <> 0 Then
                        objaddon.objapplication.SetStatusBarMessage(objaddon.objcompany.GetLastErrorDescription & " " & strTab & " " & strCol, True)
                    End If
                    'System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                End If
            Catch ex As Exception
                Throw ex
            Finally

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                oUserFieldMD1 = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Sub

        Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim strSQL As String
            Try
                If objaddon.HANA Then
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
                Else
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
                End If

                oRecordSet = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery(strSQL)

                If oRecordSet.Fields.Item(0).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
                oRecordSet = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try
        End Function

        Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
            Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

            Try
                '// The meta-data object must be initialized with a
                '// regular UserKeys object
                oUserKeysMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

                If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                    '// Set the table name and the key name
                    oUserKeysMD.TableName = strTab
                    oUserKeysMD.KeyName = strKey

                    '// Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn
                    oUserKeysMD.Elements.Add()
                    oUserKeysMD.Elements.ColumnAlias = "RentFac"

                    '// Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                    '// Add the key
                    If oUserKeysMD.Add <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
                oUserKeysMD = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub

        Private Sub AddUDO(ByVal strUDO As String, ByVal strUDODesc As String, ByVal nObjectType As SAPbobsCOM.BoUDOObjType, ByVal strTable As String, ByVal childTable() As String, ByVal sFind() As String,
                           Optional ByVal cancel As Boolean = False, Optional ByVal canlog As Boolean = False, Optional ByVal Manageseries As Boolean = False)

            Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
            Dim tablecount As Integer = 0
            Try
                oUserObjectMD = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                If oUserObjectMD.GetByKey(strUDO) = 0 Then

                    oUserObjectMD.Code = strUDO
                    oUserObjectMD.Name = strUDODesc
                    oUserObjectMD.ObjectType = nObjectType
                    oUserObjectMD.TableName = strTable

                    If (cancel) Then
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                    Else
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                    End If

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES

                    If Manageseries Then oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES Else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO

                    If canlog Then
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                        oUserObjectMD.LogTableName = "A" + strTable.ToString
                    Else
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                        oUserObjectMD.LogTableName = ""
                    End If

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO : oUserObjectMD.ExtensionName = ""

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    tablecount = 1
                    If sFind.Length > 0 Then
                        For i = 0 To sFind.Length - 1
                            If Trim(sFind(i)) = "" Then Continue For
                            oUserObjectMD.FindColumns.ColumnAlias = sFind(i)
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount)
                            tablecount = tablecount + 1
                        Next
                    End If

                    tablecount = 0
                    If Not childTable Is Nothing Then
                        If childTable.Length > 0 Then
                            For i = 0 To childTable.Length - 1
                                If Trim(childTable(i)) = "" Then Continue For
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount)
                                oUserObjectMD.ChildTables.TableName = childTable(i)
                                oUserObjectMD.ChildTables.Add()
                                tablecount = tablecount + 1
                            Next
                        End If
                    End If

                    If oUserObjectMD.Add() <> 0 Then
                        Throw New Exception(objaddon.objcompany.GetLastErrorDescription)
                    End If
                    End If

                    Catch ex As Exception
                    Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
                oUserObjectMD = Nothing
                GC.WaitForPendingFinalizers()
                GC.Collect()
            End Try

        End Sub

#End Region

    End Class
End Namespace
