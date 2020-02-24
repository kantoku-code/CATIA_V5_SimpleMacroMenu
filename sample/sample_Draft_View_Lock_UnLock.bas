Attribute VB_Name = "sample_Draft_View_Lock_UnLock"
'vba sample_Draft_View_Lock_UnLock ver0.0.2  using-'KCL0.0.12'  by Kantoku
'ｱｸﾃｨﾌﾞなｼｰﾄの全ﾋﾞｭｰをﾛｯｸ・ｱﾝﾛｯｸ

'{GP:21}
'{Caption:Lock_UnLock}
'{ControlTipText:ｱｸﾃｨﾌﾞなｼｰﾄの全ﾋﾞｭｰをﾛｯｸ・ｱﾝﾛｯｸします}
'{BackColor:12648447}
Option Explicit

Sub CATMain()
    'ﾄﾞｷｭﾒﾝﾄのﾁｪｯｸ
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
        Msg = "ｱﾝﾛｯｸ"
        LockState = False
    Else
        Msg = "ﾛｯｸ"
        LockState = True
    End If
    
    Dim i As Long
    For i = 3 To Views.Count
        Set View = Views.Item(i)
        View.LockStatus = LockState
    Next
    
    MsgBox "全てのビューを" & Msg & "しました"
End Sub
