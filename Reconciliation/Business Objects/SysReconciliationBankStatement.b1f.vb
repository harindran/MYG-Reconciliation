Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Reconciliation
    <FormAttribute("60800", "Business Objects/SysReconciliationBankStatement.b1f")>
    Friend Class SysReconciliationBankStatement
        Inherits SystemFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Dim odbdsDetails As SAPbouiCOM.DBDataSource

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.EditText0 = CType(Me.GetItem("txtdepdata").Specific, SAPbouiCOM.EditText)
            Me.Button0 = CType(Me.GetItem("btndep").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("30").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler LoadAfter, AddressOf Me.Form_LoadAfter

        End Sub

        Private Sub OnCustomInitialize()
            Try

            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"

        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

#End Region

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                objaddon.objapplication.StatusBar.SetText("Validating...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim GetData() As String
                GetData = objform.Items.Item("txtdepdata").Specific.String.ToString.Split(",")
                Dim Flag As Boolean = False
                odbdsDetails = objform.DataSources.DBDataSources.Item("JDT1")
                For i As Integer = 0 To GetData.Length - 1
                    For RecRow As Integer = 0 To odbdsDetails.Size - 1
                        If odbdsDetails.GetValue("Ref1", RecRow) = GetData(i) Then
                            'odbdsDetails.SetValue("", RecRow + 1, "Y")
                            Flag = True
                            Matrix0.Columns.Item("27").Cells.Item(RecRow + 1).Specific.Checked = True
                        End If
                    Next
                    'For Row As Integer = 1 To Matrix0.VisualRowCount
                    '    If Matrix0.Columns.Item("15").Cells.Item(Row).Specific.String = GetData(i) Then
                    '        Matrix0.Columns.Item("27").Cells.Item(Row).Specific.Checked = True
                    '    End If
                    'Next
                Next
                If Flag = True Then
                    objaddon.objapplication.StatusBar.SetText("Data Validated...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    objaddon.objapplication.StatusBar.SetText("No Data found to auto select the transactions...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If

            Catch ex As Exception

            End Try


        End Sub

        Private Sub Form_LoadAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                objform = objaddon.objapplication.Forms.GetForm("60800", pVal.FormTypeCount)
                odbdsDetails = objform.DataSources.DBDataSources.Item("JDT1")
            Catch ex As Exception
            End Try

        End Sub

        Private Sub Button0_ClickBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Button0.ClickBefore
            Try
                'Dim oUserDs As SAPbouiCOM.UserDataSource
                'oUserDs = objform.DataSources.UserDataSources.Item("SYS_75")
                'Dim cc As String = oUserDs.ValueEx
                'oUserDs.ValueEx = "N"
                'Matrix0.Columns.Item("27").Cells.Item(1).Specific.Checked = False
                'Matrix0.FlushToDataSource()
                If objform.Items.Item("txtdepdata").Specific.String = "" Then BubbleEvent = False : Exit Sub
                For RecRow As Integer = 0 To odbdsDetails.Size - 1
                    If Matrix0.Columns.Item("27").Cells.Item(RecRow + 1).Specific.Checked = True Then
                        Matrix0.Columns.Item("27").Cells.Item(RecRow + 1).Specific.Checked = False
                    End If
                Next
            Catch ex As Exception

            End Try

        End Sub

    End Class
End Namespace
