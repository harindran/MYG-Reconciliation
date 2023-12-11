Namespace Reconciliation

    Public Class clsRightClickEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Dim ocombo As SAPbouiCOM.ComboBox
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    'Case "385"
                    '    ProcessExternal_RightClickEvent(eventInfo, BubbleEvent)
                    'Case "60800"
                    '    ReconciliationBankStatement_RightClickEvent(eventInfo, BubbleEvent)
                    'Case "MULDATA"
                    '    RightClickMenu_Delete("1280", "GTS")
                    Case "CBRS"
                        ExternalBankReconciliationConsolidated_RightClickEvent(eventInfo, BubbleEvent)
                End Select
            Catch ex As Exception
            End Try
        End Sub

        Private Sub RightClickMenu_Add(ByVal MainMenu As String, ByVal NewMenuID As String, ByVal NewMenuName As String, ByVal position As Integer)
            Dim omenus As SAPbouiCOM.Menus
            Dim omenuitem As SAPbouiCOM.MenuItem
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If Not omenuitem.SubMenus.Exists(NewMenuID) Then
                oCreationPackage.UniqueID = NewMenuID
                oCreationPackage.String = NewMenuName
                oCreationPackage.Position = position
                oCreationPackage.Enabled = True
                omenus = omenuitem.SubMenus
                omenus.AddEx(oCreationPackage)
            End If
        End Sub

        Private Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Dim omenuitem As SAPbouiCOM.MenuItem
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If omenuitem.SubMenus.Exists(NewMenuID) Then
                objaddon.objapplication.Menus.RemoveEx(NewMenuID)
            End If
        End Sub

        Private Sub ProcessExternal_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                objform = objaddon.objapplication.Forms.ActiveForm
                If eventInfo.BeforeAction Then
                    If objform.Items.Item("txtFName").Specific.String <> "" Then
                        RightClickMenu_Add("1280", "CLF", "Clear File", 0)
                    End If
                Else
                    RightClickMenu_Delete("1280", "CLF")
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub ReconciliationBankStatement_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                objform = objaddon.objapplication.Forms.ActiveForm
                If eventInfo.BeforeAction Then
                    If eventInfo.ItemUID <> "" Then Exit Sub
                    RightClickMenu_Add("1280", "GTS", "Load Deposit Data", 0)
                Else
                    RightClickMenu_Delete("1280", "GTS")
                End If
            Catch ex As Exception
            End Try
        End Sub

        Private Sub ExternalBankReconciliationConsolidated_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                objform = objaddon.objapplication.Forms.ActiveForm
                strsql = objform.DataSources.DBDataSources.Item("@AT_OBRS").GetValue("Canceled", 0)
                If eventInfo.BeforeAction Then
                    If eventInfo.ItemUID = "" Then
                        If objform.Items.Item("cstat").Specific.Selected.Value = "C" And objform.Items.Item("trecono").Specific.String <> "" And strsql = "N" Then
                            objform.EnableMenu("1284", True) 'Cancel
                        Else
                            objform.EnableMenu("1284", False) 'Cancel
                        End If
                    End If
                Else
                    objform.EnableMenu("1284", False) 'Cancel
                    objform.EnableMenu("1285", False) 'Restore
                End If
            Catch ex As Exception
            End Try
        End Sub


    End Class

End Namespace
