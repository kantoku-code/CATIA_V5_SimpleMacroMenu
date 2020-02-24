Attribute VB_Name = "sample_IsExists_LeftAxis"
'vba sample_IsExists_LeftAxis_ver0.0.1  using-'KCL0.0.12'  by Kantoku
'Part���̍�����W�n�̗L������

'{Gp:1}
'{Ep:LeftHand}
'{Caption:LeftHandAxis}
'{ControlTipText:Part���̍�����W�n�̗L������}
'{BackColor:33023}
Option Explicit

Sub LeftHand()
    '�޷���Ă�����
    If Not CanExecute("PartDocument") Then Exit Sub
    
    Dim Doc As PartDocument: Set Doc = CATIA.ActiveDocument
    Dim Axs As AxisSystems: Set Axs = Doc.Part.AxisSystems
    
    Dim Ax As AxisSystem
    Dim Msg As String: Msg = vbNullString
    For Each Ax In Axs
        If IsLeft(Ax) Then
            Msg = Msg & Ax.Name & vbNewLine
        End If
    Next
    
    If Msg = vbNullString Then
        MsgBox "����n���W�������݂��Ă��܂���"
    Else
        MsgBox "����n���W�������݂��Ă��܂�" & vbNewLine & Msg
    End If
End Sub

'����n���W������
'Ax As AxisSystem��NG
Private Function IsLeft(ByVal Ax As Variant) As Boolean
    '���޸��
    Dim VecX(2), VecY(2), VecZ(2)
    Ax.GetXAxis VecX
    Ax.GetYAxis VecY
    Ax.GetZAxis VecZ
    
    'X��/Y���̊O��
    Dim Outer(2) As Double
    Outer(0) = VecX(1) * VecY(2) - VecX(2) * VecY(1)
    Outer(1) = VecX(2) * VecY(0) - VecX(0) * VecY(2)
    Outer(2) = VecX(0) * VecY(1) - VecX(1) * VecY(0)
    
    ' ���߂��O�ς�Z���Ƃ̓��ς����߂�
    IsLeft = _
        VecZ(0) * Outer(0) + VecZ(1) * Outer(1) + VecZ(2) * Outer(2) < 0
End Function



