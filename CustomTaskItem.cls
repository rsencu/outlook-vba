VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomTaskItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================
' 定数
' ==========================================================
' 日付の初期値(空の日付の値)
Const EMPTY_DATE_VALUE As Date = #1/1/4501#

' ユーザー定義フィールド[プロジェクト]の名前
Const USER_PROPERTY_NAME_PROJECT As String = "プロジェクト"
' ユーザー定義フィールド[開始日時]の名前
Const USER_PROPERTY_NAME_ACTUAL_START_TIME As String = "開始日時"
' ユーザー定義フィールド[終了日時]の名前
Const USER_PROPERTY_NAME_ACTUAL_END_TIME As String = "終了日時"
' ユーザー定義フィールド[見積時間]の名前
Const USER_PROPERTY_NAME_ESTIMATED_TIME As String = "見積時間"
' ユーザー定義フィールド[作業時間]の名前
Const USER_PROPERTY_NAME_ACTUAL_WORK_TIME As String = "作業時間"


' ==========================================================
' クラス変数
' ==========================================================
' タスク
Private MyTask As Outlook.TaskItem
' ユーザー定義フィールド[プロジェクト]
Private MyProject As Outlook.UserProperty
' ユーザー定義フィールド[開始日時]
Private MyActualStartTime As Outlook.UserProperty
' ユーザー定義フィールド[終了日時]
Private MyActualEndTime As Outlook.UserProperty
' ユーザー定義フィールド[見積時間]
Private MyEstimatedTime As Outlook.UserProperty
' ユーザー定義フィールド[作業時間]
Private MyActualWorkTime As Outlook.UserProperty


' ==========================================================
' メソッド
' ==========================================================
' 初期処理
Public Function Init(ByRef aTask As Outlook.TaskItem) As CustomTaskItem
    Set MyTask = aTask
    Debug.Print "CustomTaskItem#Init(): " & ToString()
    Set Init = Me
End Function

' 作業時間を記録する
Public Function RecordTime() As CustomTaskItem
    ' 完了                 → 作業時間の記録を再開する
    ' 未着手(開始日時が空) → 作業時間の記録を開始する
    ' 進行中(終了日時が空) → 作業時間の記録を停止する
    ' 上記以外             → 作業時間の記録を再開する
    
    If Status = olTaskComplete Then
        Set RecordTime = ReStartTime()
    ElseIf ActualStartTime = EMPTY_DATE_VALUE Then
        Set RecordTime = StartTime()
    ElseIf ActualEndTime = EMPTY_DATE_VALUE Then
        Set RecordTime = StopTime()
    Else
        Set RecordTime = ReStartTime()
    End If
End Function

' 作業時間の記録を開始する
Public Function StartTime() As CustomTaskItem
    ActualStartTime = Now()
    Status = olTaskInProgress
    
    Call MyTask.Save
    Set StartTime = Me
End Function

' 作業時間の記録を停止する
Public Function StopTime() As CustomTaskItem
    ActualEndTime = Now()
    Status = olTaskInProgress
    
    Call MyTask.Save
    Set StopTime = Me
End Function

' 作業時間の記録を再開する
Public Function ReStartTime() As CustomTaskItem
    ' ↓だと元のタスクが完了になってしまうのでダメ
    ' Set ReStartTime = Complete().Copy().NotStarted().StartTime()

    ' 見積時間もコピーすると後で集計するときに面倒なことになるので、見積時間はリセットする
    ' データは以下のようなイメージになる
    '
    ' 状態   見積時間 開始日時 終了日時
    ' ----   -------- -------- --------
    ' 作業中 1H       10:00    11:00
    '
    ' ↓
    ' ↓ 13:00 に再開
    ' ↓
    '
    ' 状態   見積時間 開始日時 終了日時
    ' ----   -------- -------- --------
    ' 作業中 1H       13:00             ← 元のタスク
    ' 完了            10:00    11:00    ← コピーしたタスク

    Dim NewCustomTask As CustomTaskItem
    Set NewCustomTask = Copy().Complete().ResetEstimatedTime()
    
    ' 元のタスクを未着手にして作業時間の記録を開始する
    Set ReStartTime = NotStarted().StartTime()
End Function

' タスクを完了状態にする
Public Function Complete() As CustomTaskItem
    If Status = olTaskComplete Then
        Set Complete = Me
        Exit Function
    End If
    
    If ActualStartTime = EMPTY_DATE_VALUE Then
        ActualStartTime = Now()
    End If
    If ActualEndTime = EMPTY_DATE_VALUE Then
        ActualEndTime = Now()
    End If
    Call MyTask.MarkComplete
    
    Call MyTask.Save
    Set Complete = Me
End Function

' タスクを未着手状態にする
Public Function NotStarted() As CustomTaskItem
    ActualStartTime = EMPTY_DATE_VALUE
    ActualEndTime = EMPTY_DATE_VALUE
    Status = olTaskNotStarted
    
    Call MyTask.Save
    Set NotStarted = Me
End Function

' 実績と見積をリセットする
Public Function Reset() As CustomTaskItem
    ActualStartTime = EMPTY_DATE_VALUE
    ActualEndTime = EMPTY_DATE_VALUE
    Status = olTaskNotStarted
    EstimatedTime = 0
    
    Call MyTask.Save
    Set Reset = Me
End Function

' 実績(作業時間の記録)をリセットする
Public Function ResetActualTime() As CustomTaskItem
    ActualStartTime = EMPTY_DATE_VALUE
    ActualEndTime = EMPTY_DATE_VALUE
    
    Call MyTask.Save
    Set ResetActualTime = Me
End Function

' 見積時間をリセットする
Public Function ResetEstimatedTime() As CustomTaskItem
    EstimatedTime = EMPTY_DATE_VALUE
    
    Call MyTask.Save
    Set ResetEstimatedTime = Me
End Function

' 分類項目をリセットする
Public Function ResetCategories() As CustomTaskItem
    Categories = ""
    
    Call MyTask.Save
    Set ResetCategories = Me
End Function

' タスクをコピーする
Public Function Copy() As CustomTaskItem
    Dim NewTask As Outlook.TaskItem
    Set NewTask = MyTask.Copy
    
    Set Copy = CustomTaskItemFactory.GetInstance(NewTask)
End Function

' タスクを移動する
Public Function Move(ByVal aFolderName As String) As CustomTaskItem
    Dim DestFolder As Outlook.Folder
    Set DestFolder = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderTasks).Folders(aFolderName)
    
    Call MyTask.Move(DestFolder)
    Set Move = Me
End Function

' タスクを保存する
Public Function Save() As CustomTaskItem
    Call MyTask.Save
    Set Save = Me
End Function

' インスタンスの内容を文字列で返却する(デバッグ用)
Public Function ToString() As String
    ToString = "Project=" & Project _
        & "," & "Subject=" & Subject _
        & "," & "Status=" & Status _
        & "," & "StartDate=" & StartDate _
        & "," & "DueDate=" & DueDate _
        & "," & "DateCompleted=" & DateCompleted _
        & "," & "ActualStartTime=" & ActualStartTime _
        & "," & "ActualEndTime=" & ActualEndTime _
        & "," & "EstimatedTime=" & EstimatedTime _
        & "," & "Categories=" & Categories
End Function


' ==========================================================
' プロパティ
' ==========================================================
' 件名
Public Property Get Subject() As String
    Subject = MyTask.Subject
End Property

Public Property Let Subject(ByVal aSubject As String)
    MyTask.Subject = aSubject
    Call MyTask.Save
End Property

' 開始日
Public Property Get StartDate() As Date
    StartDate = MyTask.StartDate
End Property

Public Property Let StartDate(ByVal aStartDate As Date)
    MyTask.StartDate = aStartDate
    Call MyTask.Save
End Property

' 期限
Public Property Get DueDate() As Date
    DueDate = MyTask.DueDate
End Property

Public Property Let DueDate(ByVal aDueDate As Date)
    MyTask.DueDate = aDueDate
    Call MyTask.Save
End Property

' 完了日
Public Property Get DateCompleted() As Date
    DateCompleted = MyTask.DateCompleted
End Property

Public Property Let DateCompleted(ByVal aDateCompleted As Date)
    MyTask.DateCompleted = aDateCompleted
    Call MyTask.Save
End Property

' 状態
Public Property Get Status() As OlTaskStatus
    Status = MyTask.Status
End Property

Public Property Let Status(ByVal aStatus As OlTaskStatus)
    MyTask.Status = aStatus
    Call MyTask.Save
End Property

' 分類項目
Public Property Get Categories() As String
    Categories = MyTask.Categories
End Property

Public Property Let Categories(ByVal aCategories As String)
    MyTask.Categories = aCategories
    Call MyTask.Save
End Property

' プロジェクト
Public Property Get Project() As String
    If Not MyProject Is Nothing Then
        Project = MyProject.Value
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyProject = MyTask.UserProperties.Find(USER_PROPERTY_NAME_PROJECT)
    If Not MyProject Is Nothing Then
        Project = MyProject.Value
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して初期化する
    Set MyProject = MyTask.UserProperties.Add(USER_PROPERTY_NAME_PROJECT, olText)
    MyProject.Value = ""
    Call MyTask.Save
    
    Project = MyProject.Value
End Property

Public Property Let Project(ByVal aProject As String)
    If Not MyProject Is Nothing Then
        MyProject.Value = aProject
        Call MyTask.Save
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyProject = MyTask.UserProperties.Find(USER_PROPERTY_NAME_PROJECT)
    If Not MyProject Is Nothing Then
        MyProject.Value = aProject
        Call MyTask.Save
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して引数を設定する
    Set MyProject = MyTask.UserProperties.Add(USER_PROPERTY_NAME_PROJECT, olText)
    MyProject.Value = aProject
    Call MyTask.Save
End Property

' 開始日時
Public Property Get ActualStartTime() As Date
    If Not MyActualStartTime Is Nothing Then
        ActualStartTime = MyActualStartTime.Value
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyActualStartTime = MyTask.UserProperties.Find(USER_PROPERTY_NAME_ACTUAL_START_TIME)
    If Not MyActualStartTime Is Nothing Then
        ActualStartTime = MyActualStartTime.Value
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して初期化する
    Set MyActualStartTime = MyTask.UserProperties.Add(USER_PROPERTY_NAME_ACTUAL_START_TIME, olDateTime)
    MyActualStartTime.Value = EMPTY_DATE_VALUE
    Call MyTask.Save

    ActualStartTime = MyActualStartTime.Value
End Property

Public Property Let ActualStartTime(ByVal anActualStartTime As Date)
    If Not MyActualStartTime Is Nothing Then
        MyActualStartTime.Value = anActualStartTime
        Call MyTask.Save
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyActualStartTime = MyTask.UserProperties.Find(USER_PROPERTY_NAME_ACTUAL_START_TIME)
    If Not MyActualStartTime Is Nothing Then
        MyActualStartTime.Value = anActualStartTime
        Call MyTask.Save
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して引数を設定する
    Set MyActualStartTime = MyTask.UserProperties.Add(USER_PROPERTY_NAME_ACTUAL_START_TIME, olDateTime)
    MyActualStartTime.Value = anActualStartTime
    Call MyTask.Save
End Property

' 終了日時
Public Property Get ActualEndTime() As Date
    If Not MyActualEndTime Is Nothing Then
        ActualEndTime = MyActualEndTime.Value
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyActualEndTime = MyTask.UserProperties.Find(USER_PROPERTY_NAME_ACTUAL_END_TIME)
    If Not MyActualEndTime Is Nothing Then
        ActualEndTime = MyActualEndTime.Value
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して初期化する
    Set MyActualEndTime = MyTask.UserProperties.Add(USER_PROPERTY_NAME_ACTUAL_END_TIME, olDateTime)
    MyActualEndTime.Value = EMPTY_DATE_VALUE
    Call MyTask.Save

    ActualEndTime = MyActualEndTime.Value
End Property

Public Property Let ActualEndTime(ByVal anActualEndTime As Date)
    If Not MyActualEndTime Is Nothing Then
        MyActualEndTime.Value = anActualEndTime
        Call MyTask.Save
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyActualEndTime = MyTask.UserProperties.Find(USER_PROPERTY_NAME_ACTUAL_END_TIME)
    If Not MyActualEndTime Is Nothing Then
        MyActualEndTime.Value = anActualEndTime
        Call MyTask.Save
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して引数を設定する
    Set MyActualEndTime = MyTask.UserProperties.Add(USER_PROPERTY_NAME_ACTUAL_END_TIME, olDateTime)
    MyActualEndTime.Value = anActualEndTime
    Call MyTask.Save
End Property

' 見積時間
Public Property Get EstimatedTime() As Long
    If Not MyEstimatedTime Is Nothing Then
        EstimatedTime = MyEstimatedTime.Value
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyEstimatedTime = MyTask.UserProperties.Find(USER_PROPERTY_NAME_ESTIMATED_TIME)
    If Not MyEstimatedTime Is Nothing Then
        EstimatedTime = MyEstimatedTime.Value
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して初期化する
    Set MyEstimatedTime = MyTask.UserProperties.Add(USER_PROPERTY_NAME_ESTIMATED_TIME, olDuration)
    MyEstimatedTime.Value = EMPTY_DATE_VALUE
    Call MyTask.Save

    EstimatedTime = MyEstimatedTime.Value
End Property

Public Property Let EstimatedTime(ByVal anEstimatedTime As Long)
    If Not MyEstimatedTime Is Nothing Then
        MyEstimatedTime.Value = anEstimatedTime
        Call MyTask.Save
        Exit Property
    End If
    
    ' ユーザー定義フィールドの取得
    Set MyEstimatedTime = MyTask.UserProperties.Find(USER_PROPERTY_NAME_ESTIMATED_TIME)
    If Not MyEstimatedTime Is Nothing Then
        MyEstimatedTime.Value = anEstimatedTime
        Call MyTask.Save
        Exit Property
    End If
    
    ' ここまで来るのはユーザ定義フィールドがない場合なので、追加して引数を設定する
    Set MyEstimatedTime = MyTask.UserProperties.Add(USER_PROPERTY_NAME_ESTIMATED_TIME, olDuration)
    MyEstimatedTime.Value = anEstimatedTime
    Call MyTask.Save
End Property
