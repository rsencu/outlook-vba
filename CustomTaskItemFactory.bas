Attribute VB_Name = "CustomTaskItemFactory"
Option Explicit

' �J�X�^���^�X�N���擾����
Public Function GetInstance(ByRef Task As Outlook.TaskItem) As CustomTaskItem
    Dim CustomTask As CustomTaskItem
    Set CustomTask = New CustomTaskItem
    Set GetInstance = CustomTask.Init(Task)
End Function


' �J�X�^���^�X�N���쐬����
Public Function NewInstance() As CustomTaskItem
    Dim NewTask As Outlook.TaskItem
    Set NewTask = Application.CreateItem(olTaskItem)
    Set NewInstance = GetInstance(NewTask)
End Function

