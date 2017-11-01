Attribute VB_Name = "CustomTaskItemFactory"
Option Explicit

' カスタムタスクを取得する
Public Function GetInstance(ByRef Task As Outlook.TaskItem) As CustomTaskItem
    Dim CustomTask As CustomTaskItem
    Set CustomTask = New CustomTaskItem
    Set GetInstance = CustomTask.Init(Task)
End Function


' カスタムタスクを作成する
Public Function NewInstance() As CustomTaskItem
    Dim NewTask As Outlook.TaskItem
    Set NewTask = Application.CreateItem(olTaskItem)
    Set NewInstance = GetInstance(NewTask)
End Function

