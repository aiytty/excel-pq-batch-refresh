Attribute VB_Name = "RefreshModule"
' ============================================================================
'  Excel Power Query 批量刷新宏
'  用于 SKILL: excel-pq-batch-refresh
'  适用场景: Power Automate Desktop 调用此宏刷新 Power Query 数据源
' ============================================================================

Option Explicit

Public Sub RefreshPowerQuery()
    On Error GoTo ErrorHandler

    ' 刷新所有数据连接（含 Power Query）
    ThisWorkbook.RefreshAll

    ' 等待所有异步查询完成
    Application.CalculateUntilAsyncQueriesDone

    ' 保存工作簿
    ThisWorkbook.Save

    Exit Sub

ErrorHandler:
    MsgBox "RefreshPowerQuery 失败: " & Err.Description & vbCrLf & _
           "Err.Number: " & Err.Number, vbCritical, "Excel PQ 刷新"
    Err.Clear
End Sub