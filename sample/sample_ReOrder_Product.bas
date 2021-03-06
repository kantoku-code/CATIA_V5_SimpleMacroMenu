Attribute VB_Name = "sample_ReOrder_Product"
'vba sample_ReOrder_Product ver0.0.1  using-'KCL0.0.12'  by Kantoku
'ｲﾝｽﾀﾝｽ名でのｿｰﾄ順にTreeを並び替えます

'{GP:11}
'{Caption:ﾘｵｰﾀﾞｰ}
'{ControlTipText:ｲﾝｽﾀﾝｽ名でのｿｰﾄ順にTreeを並び替えます
'{BackColor:16744703}
'{FONTSIZE:10.5}

Option Explicit

Sub CATMain()
    'ﾄﾞｷｭﾒﾝﾄのﾁｪｯｸ
    If Not CanExecute("ProductDocument") Then Exit Sub
    
    'Doc取得
    Dim ProDoc As ProductDocument: Set ProDoc = CATIA.ActiveDocument
    Dim Pros As Products: Set Pros = ProDoc.Product.Products
    If Pros.Count < 2 Then Exit Sub
    
    'ｵﾌﾟｼｮﾝ変更
    Dim AssyMode As AsmConstraintSettingAtt
    Set AssyMode = CATIA.SettingControllers.Item("CATAsmConstraintSettingCtrl")
    Dim OriginalMode As CatAsmPasteComponentMode
    OriginalMode = AssyMode.PasteComponentMode
    
    'ｵﾌﾟｼｮﾝ切り替え
    AssyMode.PasteComponentMode = catPasteWithCstOnCopyAndCut
    
    'ｿｰﾄ済み名前ﾘｽﾄ
    Dim Names: Set Names = Get_SortedNames(Pros)
    
    'ｶｯﾄ
    Dim Sel As Selection: Set Sel = ProDoc.Selection
    Dim Itm As Variant
    
    CATIA.HSOSynchronized = False
    
    Sel.Clear
    For Each Itm In Names
        Sel.Add Pros.Item(Itm)
    Next
    Sel.Cut
    
    'ﾍﾟｰｽﾄ
    With Sel
        .Clear
        .Add Pros
        .Paste
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    
    'ｵﾌﾟｼｮﾝ戻し,UpDate
    AssyMode.PasteComponentMode = OriginalMode
    ProDoc.Product.Update
End Sub

'ｲﾝｽﾀﾝｽ名でｿｰﾄ済みの名前ﾘｽﾄ
Private Function Get_SortedNames(ByVal Pros As Products) As Object
    Dim Lst As Object
    Set Lst = KCL.InitLst()
    
    Dim Pro As Product
    For Each Pro In Pros
        Lst.Add Pro.Name
    Next
    
    Lst.Sort
    
    Set Get_SortedNames = Lst
End Function
