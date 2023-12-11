Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework

Namespace Reconciliation
    <FormAttribute("120060803", "Business Objects/SysFrmInternalRecSelectionCriteria.b1f")>
    Friend Class SysFrmInternalRecSelectionCriteria
        Inherits SystemFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("120000001").Specific, SAPbouiCOM.Button)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub

#Region "Fields"

        Private WithEvents Button0 As SAPbouiCOM.Button

#End Region

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("120060803", 1)
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                'GetRecData = ""
                Dim objedit As SAPbouiCOM.EditText
                Dim Matrix0 As SAPbouiCOM.Matrix
                objedit = objform.Items.Item("540000095").Specific
                If objedit.Value = "" Then
                    RecFromDate = Date.Now.Date
                Else
                    RecFromDate = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                End If
                objedit = objform.Items.Item("540000097").Specific
                If objedit.Value = "" Then
                    RecToDate = Date.Now.Date
                Else
                    RecToDate = Date.ParseExact(objedit.Value, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                End If
                ''RecFromDate = DateAdd(DateInterval.Day, -1, DateTime.Now.Date)
                ''RecToDate = Date.Now.Date
                'GetRecData = GetData(DateAdd(DateInterval.Day, -1, DateTime.Now.Date), Date.Now.Date)
                Matrix0 = objform.Items.Item("10000085").Specific
                For i As Integer = 1 To Matrix0.VisualRowCount
                    If Matrix0.Columns.Item("10000003").Cells.Item(i).Specific.String <> "" Then
                        If i = 1 Then
                            CardCode = "'" + Matrix0.Columns.Item("10000003").Cells.Item(i).Specific.String + "'"
                        Else
                            CardCode += ",'" + Matrix0.Columns.Item("10000003").Cells.Item(i).Specific.String + "'"
                        End If
                    End If
                Next
            Catch ex As Exception

            End Try

        End Sub

        Private Function GetData(ByVal FromDate As Date, ByVal ToDate As Date) As String
            Try
                Dim Str As String
                If objaddon.HANA Then
                    Str = "select ""DueDate"",""AcctCode"",""Ref"",""Memo"",""Memo2"",""DebAmount"" from OBNK  where ""DueDate"" between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "'" ' order by ""DebAmount"""
                Else
                    Str = "select DueDate,AcctCode,Ref,Memo,Memo2,DebAmount from OBNK  where DueDate between '" & FromDate.ToString("yyyyMMdd") & "' and '" & ToDate.ToString("yyyyMMdd") & "' " 'order by DebAmount"
                End If
                Return Str
            Catch ex As Exception
                Return Nothing
            End Try
        End Function


    End Class
End Namespace
