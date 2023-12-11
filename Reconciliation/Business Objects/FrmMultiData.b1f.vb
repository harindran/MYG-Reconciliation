Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Reconciliation
    <FormAttribute("MULDATA", "Business Objects/FrmMultiData.b1f")>
    Friend Class FrmMultiData
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form

        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("btnok").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Button2 = CType(Me.GetItem("btnclear").Specific, SAPbouiCOM.Button)
            Me.Matrix0 = CType(Me.GetItem("mtxdata").Specific, SAPbouiCOM.Matrix)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MULDATA", 1)
                objform = objaddon.objapplication.Forms.ActiveForm
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                Matrix0.Columns.Item("Code").Cells.Item(1).Click()
                'objform.EnableMenu("1292", True)
                objform.EnableMenu("773", True)
                objform.EnableMenu("1281", False)
                objform.EnableMenu("1282", False)
                bModal = True
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try
        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix

#End Region

        Private Sub Button2_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try
                Matrix0.Clear()
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Matrix0_ValidateBefore(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As Boolean) Handles Matrix0.ValidateBefore
            Try
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Code", "#")
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button0_ClickAfter(sboObject As Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                Dim Code As String = ""
                For i As Integer = 1 To Matrix0.VisualRowCount
                    Code = Code + Matrix0.Columns.Item("Code").Cells.Item(i).Specific.String + ","
                Next
                Code = Code.Remove(Code.Length - 2)
                objMultiLoad.Items.Item("txtdepdata").Specific.String = ""
                objMultiLoad.Items.Item("txtdepdata").Specific.String = Code

                objMultiLoad = Nothing
                objform.Close()
            Catch ex As Exception

            End Try

        End Sub
    End Class
End Namespace
