Attribute VB_Name = "TimeTracker"
Option Explicit

' 名前
'   タスクの作業時間を記録する
'
' 概要
'   タスクビューで選択しているタスクの作業時間を記録する。
'
'   タイムトラッキングの基本的な流れは以下のとおりになる。
'
'     タスクが終わっているか？ --[Yes]-> 進捗状況を完了する
'       ↓[No]
'     タスクの作業時間を記録する() をひたすら実行する
'
'
' 詳細説明
'   このマクロを実行することで、タスクビューで選択しているタスクに対し、
'   状態に応じて、開始日時・終了日時が設定される。
'   選択しているタスクが複数存在する場合、処理対象は1つ目のタスクのみとなる。
'
'   完了したタスクや開始日時・終了日時が設定されているタスクの場合、
'   元のタスクをコピーして、
'     1. コピーしたタスクに対し、進捗状況を完了し、
'     2. 元のタスクに対し、実績がリセットされ開始日時が設定される。
'
'
'   以下に使い方のイメージを示す。
'
'     状態   見積時間 開始日時 終了日時 作業時間
'     ----   -------- -------- -------- --------
'     未着手 4H
'
'     ↓
'     ↓ 10:00 作業の開始
'     ↓   タスクの作業時間を記録する() の実行
'     ↓
'
'     状態   見積時間 開始日時 終了日時 作業時間
'     ----   -------- -------- -------- --------
'     進行中 4H       10:00
'
'     ↓
'     ↓ 12:00 作業の一時中断
'     ↓   タスクの作業時間を記録する() の実行
'     ↓
'
'     状態   見積時間 開始日時 終了日時 作業時間
'     ----   -------- -------- -------- --------
'     進行中 4H       10:00    12:00    120
'
'     ↓
'     ↓ 13:00 作業の再開
'     ↓   タスクの作業時間を記録する() の実行
'     ↓
'
'     状態   見積時間 開始日時 終了日時 作業時間
'     ----   -------- -------- -------- --------
'     進行中 4H       13:00
'     完了            10:00    12:00    120
'
'     ↓
'     ↓ 16:00 作業の終了
'     ↓   タスクの作業時間を記録する() の実行
'     ↓
'
'     状態   見積時間 開始日時 終了日時 作業時間
'     ----   -------- -------- -------- --------
'     進行中 4H       13:00    16:00    180
'     完了            10:00    12:00    120
'
'     ↓
'     ↓   進捗状況を完了にする
'     ↓
'
'     状態   見積時間 開始日時 終了日時 作業時間
'     ----   -------- -------- -------- --------
'     完了   4H       13:00    16:00    180
'     完了            10:00    12:00    120
'
Public Sub タスクの作業時間を記録する()
    ' タスクビューで選択中のタスクを取得する
    Dim SelectedTasks As Outlook.Selection
    Set SelectedTasks = Application.ActiveExplorer.Selection
    
    Dim Task As Outlook.TaskItem
    Dim Item As Variant
    For Each Item In SelectedTasks
        If TypeName(Item) = "TaskItem" Then
            Set Task = Item
            Dim CustomTask As CustomTaskItem
            Set CustomTask = CustomTaskItemFactory.GetInstance(Task)
            
            Call CustomTask.RecordTime
            ' 1回で終了する
            Exit For
        End If
    Next
End Sub

' 名前
'   選択タスクの実績をリセットする
'
' 概要
'   タスクビューで選択しているタスクの開始日時・終了日時をリセットする。
'
'
' 詳細説明
'   このマクロを実行し、ダイアログにて [はい] を選択することで、
'   タスクビューで選択しているタスクに対し、開始日時・終了日時をリセットする。
'   状態は未着手にならない。
'
Public Sub 選択タスクの実績をリセットする()
    Dim Message As String
    Dim Response As Integer
    
    Message = "実績をリセットしますか？"
    Response = MsgBox(Message, vbYesNo)
    
    If Response = vbNo Then
        ' キャンセルされたので終了
        Exit Sub
    End If
    
    ' タスクビューで選択中のタスクを取得する
    Dim SelectedTasks As Outlook.Selection
    Set SelectedTasks = Application.ActiveExplorer.Selection
    
    Dim Item As Variant
    Dim Task As Outlook.TaskItem
    For Each Item In SelectedTasks
        If TypeName(Item) = "TaskItem" Then
            Set Task = Item
            Dim CustomTask As CustomTaskItem
            Set CustomTask = CustomTaskItemFactory.GetInstance(Task)
            
            Call CustomTask.ResetActualTime
        End If
    Next
End Sub

' 名前
'   選択タスクのプロジェクトを設定する
'
' 概要
'   タスクビューで選択しているタスクのプロジェクトを設定する。
'
'
' 詳細説明
'   このマクロを実行することで、
'   タスクビューで選択しているタスクに対し、ダイアログにて入力したプロジェクト名を設定する。
'
Public Sub 選択タスクのプロジェクトを設定する()
    Dim Prompt As String
    Dim Title As String
    Dim NewProject As String
    
    Prompt = "プロジェクト名を入力してください。"
    Title = "プロジェクトの設定"
    NewProject = InputBox(Prompt, Title)
    
    If Len(NewProject) = 0 Then
        ' キャンセルされたので終了
        Exit Sub
    End If
    
    Dim SelectedTasks As Outlook.Selection
    Set SelectedTasks = Application.ActiveExplorer.Selection
    
    Dim Item As Variant
    Dim Task As Outlook.TaskItem
    For Each Item In SelectedTasks
        If TypeName(Item) = "TaskItem" Then
            Set Task = Item
            Dim CustomTask As CustomTaskItem
            Set CustomTask = CustomTaskItemFactory.GetInstance(Task)
            
            CustomTask.Project = NewProject
        End If
    Next
End Sub



