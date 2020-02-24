Attribute VB_Name = "sample_Draft_View_Lock_UnLock"
'vba sample_Draft_View_Lock_UnLock ver0.0.2  using-'KCL0.0.12'  by Kantoku
'��è�ނȼ�Ă̑S�ޭ���ۯ��E��ۯ�

'{GP:21}
'{Caption:Lock_UnLock}
'{ControlTipText:��è�ނȼ�Ă̑S�ޭ���ۯ��E��ۯ����܂�}
'{BackColor:12648447}
Option Explicit

Sub CATMain()
    '�޷���Ă�����
    If Not CanExecute("DrawingDocument") Then Exit Sub
    
    Dim Views As DrawingViews
    Set Views = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
    
    If Views.Count < 3 Then Exit Sub
    
    Dim View As DrawingView
    Set View = Views.Item(3)
    
    Dim LockState As Boolean
    LockState = View.LockStatus
    
    Dim Msg As String
    If LockState Then
        Msg = "��ۯ�"
        LockState = False
    Else
        Msg = "ۯ�"
        LockState = True
    End If
    
    Dim i As Long
    For i = 3 To Views.Count
        Set View = Views.Item(i)
        View.LockStatus = LockState
    Next
    
    MsgBox "�S�Ẵr���[��" & Msg & "���܂���"
End Sub
