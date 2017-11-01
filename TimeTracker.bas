Attribute VB_Name = "TimeTracker"
Option Explicit

' ���O
'   �^�X�N�̍�Ǝ��Ԃ��L�^����
'
' �T�v
'   �^�X�N�r���[�őI�����Ă���^�X�N�̍�Ǝ��Ԃ��L�^����B
'
'   �^�C���g���b�L���O�̊�{�I�ȗ���͈ȉ��̂Ƃ���ɂȂ�B
'
'     �^�X�N���I����Ă��邩�H --[Yes]-> �i���󋵂���������
'       ��[No]
'     �^�X�N�̍�Ǝ��Ԃ��L�^����() ���Ђ�������s����
'
'
' �ڍא���
'   ���̃}�N�������s���邱�ƂŁA�^�X�N�r���[�őI�����Ă���^�X�N�ɑ΂��A
'   ��Ԃɉ����āA�J�n�����E�I���������ݒ肳���B
'   �I�����Ă���^�X�N���������݂���ꍇ�A�����Ώۂ�1�ڂ̃^�X�N�݂̂ƂȂ�B
'
'   ���������^�X�N��J�n�����E�I���������ݒ肳��Ă���^�X�N�̏ꍇ�A
'   ���̃^�X�N���R�s�[���āA
'     1. �R�s�[�����^�X�N�ɑ΂��A�i���󋵂��������A
'     2. ���̃^�X�N�ɑ΂��A���т����Z�b�g����J�n�������ݒ肳���B
'
'
'   �ȉ��Ɏg�����̃C���[�W�������B
'
'     ���   ���ώ��� �J�n���� �I������ ��Ǝ���
'     ----   -------- -------- -------- --------
'     ������ 4H
'
'     ��
'     �� 10:00 ��Ƃ̊J�n
'     ��   �^�X�N�̍�Ǝ��Ԃ��L�^����() �̎��s
'     ��
'
'     ���   ���ώ��� �J�n���� �I������ ��Ǝ���
'     ----   -------- -------- -------- --------
'     �i�s�� 4H       10:00
'
'     ��
'     �� 12:00 ��Ƃ̈ꎞ���f
'     ��   �^�X�N�̍�Ǝ��Ԃ��L�^����() �̎��s
'     ��
'
'     ���   ���ώ��� �J�n���� �I������ ��Ǝ���
'     ----   -------- -------- -------- --------
'     �i�s�� 4H       10:00    12:00    120
'
'     ��
'     �� 13:00 ��Ƃ̍ĊJ
'     ��   �^�X�N�̍�Ǝ��Ԃ��L�^����() �̎��s
'     ��
'
'     ���   ���ώ��� �J�n���� �I������ ��Ǝ���
'     ----   -------- -------- -------- --------
'     �i�s�� 4H       13:00
'     ����            10:00    12:00    120
'
'     ��
'     �� 16:00 ��Ƃ̏I��
'     ��   �^�X�N�̍�Ǝ��Ԃ��L�^����() �̎��s
'     ��
'
'     ���   ���ώ��� �J�n���� �I������ ��Ǝ���
'     ----   -------- -------- -------- --------
'     �i�s�� 4H       13:00    16:00    180
'     ����            10:00    12:00    120
'
'     ��
'     ��   �i���󋵂������ɂ���
'     ��
'
'     ���   ���ώ��� �J�n���� �I������ ��Ǝ���
'     ----   -------- -------- -------- --------
'     ����   4H       13:00    16:00    180
'     ����            10:00    12:00    120
'
Public Sub �^�X�N�̍�Ǝ��Ԃ��L�^����()
    ' �^�X�N�r���[�őI�𒆂̃^�X�N���擾����
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
            ' 1��ŏI������
            Exit For
        End If
    Next
End Sub

' ���O
'   �I���^�X�N�̎��т����Z�b�g����
'
' �T�v
'   �^�X�N�r���[�őI�����Ă���^�X�N�̊J�n�����E�I�����������Z�b�g����B
'
'
' �ڍא���
'   ���̃}�N�������s���A�_�C�A���O�ɂ� [�͂�] ��I�����邱�ƂŁA
'   �^�X�N�r���[�őI�����Ă���^�X�N�ɑ΂��A�J�n�����E�I�����������Z�b�g����B
'   ��Ԃ͖�����ɂȂ�Ȃ��B
'
Public Sub �I���^�X�N�̎��т����Z�b�g����()
    Dim Message As String
    Dim Response As Integer
    
    Message = "���т����Z�b�g���܂����H"
    Response = MsgBox(Message, vbYesNo)
    
    If Response = vbNo Then
        ' �L�����Z�����ꂽ�̂ŏI��
        Exit Sub
    End If
    
    ' �^�X�N�r���[�őI�𒆂̃^�X�N���擾����
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

' ���O
'   �I���^�X�N�̃v���W�F�N�g��ݒ肷��
'
' �T�v
'   �^�X�N�r���[�őI�����Ă���^�X�N�̃v���W�F�N�g��ݒ肷��B
'
'
' �ڍא���
'   ���̃}�N�������s���邱�ƂŁA
'   �^�X�N�r���[�őI�����Ă���^�X�N�ɑ΂��A�_�C�A���O�ɂē��͂����v���W�F�N�g����ݒ肷��B
'
Public Sub �I���^�X�N�̃v���W�F�N�g��ݒ肷��()
    Dim Prompt As String
    Dim Title As String
    Dim NewProject As String
    
    Prompt = "�v���W�F�N�g������͂��Ă��������B"
    Title = "�v���W�F�N�g�̐ݒ�"
    NewProject = InputBox(Prompt, Title)
    
    If Len(NewProject) = 0 Then
        ' �L�����Z�����ꂽ�̂ŏI��
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



