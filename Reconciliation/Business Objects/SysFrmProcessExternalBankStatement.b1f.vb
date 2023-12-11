Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System.IO

Namespace Reconciliation
    <FormAttribute("385", "Business Objects/SysFrmProcessExternalBankStatement.b1f")>
    Friend Class SysFrmProcessExternalBankStatement
        Inherits SystemFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public WithEvents objMatrix As SAPbouiCOM.Matrix
        Private WithEvents objcombo As SAPbouiCOM.ComboBox

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.EditText0 = CType(Me.GetItem("txtFName").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("BtnLoad").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()
            Try

                objform = objaddon.objapplication.Forms.GetForm("385", 1)
                EditText0.Item.Height = objform.Items.Item("4").Height
                'Button0.Item.Height = objform.Items.Item("2").Height
                Button0.Item.Top = EditText0.Item.Top
                objMatrix = objform.Items.Item("5").Specific
                objcombo = objform.Items.Item("14").Specific
                objcombo.Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"

        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents Button0 As SAPbouiCOM.Button

#End Region

        Private Sub ReadExcel_Old(ByVal FileName As String)
            Try
                objaddon.objapplication.SetStatusBarMessage("Looking out for Excel Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
                Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim excelRng As Microsoft.Office.Interop.Excel.Range
                Dim j As Integer = 1
                Try
                    objform = objaddon.objapplication.Forms.GetForm("385", 1)
                    FileName = objform.Items.Item("txtFName").Specific.string
                    Dim RowIndex As Integer
                    ExcelWorkbook = ExcelApp.Workbooks.Open(FileName)
                    ExcelWorkSheet = ExcelWorkbook.ActiveSheet
                    'excelRng = ExcelWorkSheet.Range("A1")
                    excelRng = ExcelWorkSheet.UsedRange
                    objaddon.objapplication.SetStatusBarMessage("Excel Data Loading Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    'objform.Freeze(True)
                    Dim DealID1, DealID2 As String
                    'objMatrix.Clear()
                    If objform.Items.Item("txtFName").Specific.String <> "" Then
                        If ExcelWorkSheet.Cells(1, 1).Value = "Customer Name" And ExcelWorkSheet.Cells(1, 2).Value = "Date" And ExcelWorkSheet.Cells(1, 3).Value = "Deal ID 1" And ExcelWorkSheet.Cells(1, 4).Value = "Deal ID 2" And ExcelWorkSheet.Cells(1, 5).Value = "Amount" Then
                            For RowIndex = 2 To excelRng.Rows.Count
                                If CStr(ExcelWorkSheet.Cells(RowIndex, 1).Value) <> "" Then
                                    If objMatrix.Columns.Item("4").Cells.Item(objMatrix.VisualRowCount).Specific.String <> "" Then
                                        objMatrix.AddRow()
                                    End If
                                    'Dim AttDate As Date = Date.ParseExact(ExcelWorkSheet.Cells(RowIndex, 2).Value, "dd/MM/yy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                    'objMatrix.Columns.Item("0").Cells.Item(j).Specific.String = j
                                    If CStr(ExcelWorkSheet.Cells(RowIndex, 3).Value) = "" Then
                                        DealID1 = "0"
                                    Else
                                        DealID1 = CStr(ExcelWorkSheet.Cells(RowIndex, 3).Value)
                                    End If
                                    If CStr(ExcelWorkSheet.Cells(RowIndex, 4).Value) = "" Then
                                        DealID2 = "0"
                                    Else
                                        DealID2 = CStr(ExcelWorkSheet.Cells(RowIndex, 4).Value)
                                    End If
                                    objMatrix.Columns.Item("3").Cells.Item(objMatrix.VisualRowCount).Specific.String = DealID1.ToUpper  ' DealID 1
                                    objMatrix.Columns.Item("2").Cells.Item(objMatrix.VisualRowCount).Specific.String = Format(ExcelWorkSheet.Cells(RowIndex, 2).Value, "yyyyMMdd") 'ExcelWorkSheet.Cells(RowIndex, 2).Value 
                                    objMatrix.Columns.Item("4").Cells.Item(objMatrix.VisualRowCount).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 1).Value).ToUpper  'Cust Name
                                    objMatrix.Columns.Item("19").Cells.Item(objMatrix.VisualRowCount).Specific.String = DealID2.ToUpper 'DealID 2
                                    objMatrix.Columns.Item("5").Cells.Item(objMatrix.VisualRowCount).Specific.String = CStr(ExcelWorkSheet.Cells(RowIndex, 5).Value)        'Amount
                                    j += 1
                                End If
                            Next RowIndex
                        Else
                            objaddon.objapplication.StatusBar.SetText("Expected ColumnName Not found...Please check the excel format", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                        'objform.Freeze(False)
                        objMatrix.AutoResizeColumns()
                    End If
                    'objaddon.objapplication.Menus.Item("1300").Activate()
                    objaddon.objapplication.StatusBar.SetText("Excel Data Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'objaddon.objapplication.SetStatusBarMessage("Excel Loaded..." & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Catch ex As Exception
                    'objform.Freeze(False)
                    'ExcelApp.ActiveWorkbook.Close()
                    objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Finally
                    'objform.Freeze(False)
                    ExcelApp.ActiveWorkbook.Close()
                End Try
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub ReadExcel(ByVal FileName As String)
            Try
                objaddon.objapplication.SetStatusBarMessage("Looking out for Excel Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Dim ExcelApp As New Microsoft.Office.Interop.Excel.Application
                Dim ExcelWorkbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Dim ExcelWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim excelRng As Microsoft.Office.Interop.Excel.Range
                Dim j As Integer = 1
                Try
                    objform = objaddon.objapplication.Forms.GetForm("385", 1)
                    FileName = objform.Items.Item("txtFName").Specific.string
                    Dim RowIndex As Integer
                    ExcelWorkbook = ExcelApp.Workbooks.Open(FileName)
                    ExcelWorkSheet = ExcelWorkbook.ActiveSheet
                    'excelRng = ExcelWorkSheet.Range("A1")
                    excelRng = ExcelWorkSheet.UsedRange
                    objaddon.objapplication.SetStatusBarMessage("Excel Data Loading Please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
                    'objform.Freeze(True)
                    Dim DealID1, InvoiceNo As String
                    'objMatrix.Clear()

                    If objform.Items.Item("txtFName").Specific.String <> "" Then

                        If excelRng.Cells(1, 1).Value.ToString.ToUpper() = "DEALID" And excelRng.Cells(1, 2).Value.ToString.ToUpper() = "INVOICENO" And excelRng.Cells(1, 3).Value.ToString.ToUpper() = "CUSTOMER NAME" And excelRng.Cells(1, 4).Value.ToString.ToUpper() = "CREDIT AMOUNT" Then
                            For RowIndex = 2 To excelRng.Rows.Count
                                If CStr(excelRng.Cells(RowIndex, 1).Value) <> "" Then
                                    'For i As Integer = 2 To excelRng.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeVisible).EntireRow.Count
                                    'Next
                                    If excelRng.Rows(RowIndex).Hidden = True Then Continue For
                                    If objMatrix.Columns.Item("4").Cells.Item(objMatrix.VisualRowCount).Specific.String <> "" Then
                                        objMatrix.AddRow()
                                    End If
                                    objMatrix.Columns.Item("3").Cells.Item(objMatrix.VisualRowCount).Specific.String = CStr(excelRng.Cells(RowIndex, 1).Value) ' DealID1 '.ToUpper  ' DealID 1
                                    objMatrix.Columns.Item("2").Cells.Item(objMatrix.VisualRowCount).Specific.String = "A" ' Format(DateTime.Now, "yyyyMMdd") ' Format(ExcelWorkSheet.Cells(RowIndex, 2).Value, "yyyyMMdd") 'ExcelWorkSheet.Cells(RowIndex, 2).Value 
                                    objMatrix.Columns.Item("4").Cells.Item(objMatrix.VisualRowCount).Specific.String = CStr(excelRng.Cells(RowIndex, 3).Value) 'Cust Name
                                    objMatrix.Columns.Item("19").Cells.Item(objMatrix.VisualRowCount).Specific.String = CStr(excelRng.Cells(RowIndex, 2).Value) ' InvoiceNo '.ToUpper 'DealID 2
                                    objMatrix.Columns.Item("5").Cells.Item(objMatrix.VisualRowCount).Specific.String = CStr(excelRng.Cells(RowIndex, 4).Value)        'Amount
                                    'j += 1
                                End If
                            Next RowIndex
                        Else
                            objaddon.objapplication.StatusBar.SetText("Expected ColumnName Not found...Please check the excel format", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                        'objform.Freeze(False)
                        objMatrix.AutoResizeColumns()
                    End If
                    'objaddon.objapplication.Menus.Item("1300").Activate()
                    objaddon.objapplication.StatusBar.SetText("Excel Data Loaded Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    'objaddon.objapplication.SetStatusBarMessage("Excel Loaded..." & DocEntry, SAPbouiCOM.BoMessageTime.bmt_Long, False)
                Catch ex As Exception
                    'objform.Freeze(False)
                    'ExcelApp.ActiveWorkbook.Close()
                    objaddon.objapplication.SetStatusBarMessage("Load Excel: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
                Finally
                    'objform.Freeze(False)
                    ExcelApp.ActiveWorkbook.Close()
                End Try
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Items.Item("txtFName").Specific.string = "" Then Exit Sub
                ReadExcel(objform.Items.Item("txtFName").Specific.string)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                If objform.Items.Item("4").Specific.String = "" Then
                    objaddon.objapplication.SetStatusBarMessage("BP Code is Missing", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
                If objform.Items.Item("txtFName").Specific.String = "" Then
                    objaddon.objapplication.SetStatusBarMessage("File Name is Missing.Please choose a file", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                    BubbleEvent = False : Exit Sub
                End If
            Catch ex As Exception

            End Try

        End Sub


    End Class
End Namespace
