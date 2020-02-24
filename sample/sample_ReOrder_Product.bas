Attribute VB_Name = "sample_ReOrder_Product"
'vba sample_ReOrder_Product ver0.0.1  using-'KCL0.0.12'  by Kantoku
'²İ½Àİ½–¼‚Å‚Ì¿°Ä‡‚ÉTree‚ğ•À‚Ñ‘Ö‚¦‚Ü‚·

'{GP:11}
'{Caption:Øµ°ÀŞ°}
'{ControlTipText:²İ½Àİ½–¼‚Å‚Ì¿°Ä‡‚ÉTree‚ğ•À‚Ñ‘Ö‚¦‚Ü‚·
'{BackColor:16744703}
'{FONTSIZE:10.5}

Option Explicit

Sub CATMain()
    'ÄŞ·­ÒİÄ‚ÌÁª¯¸
    If Not CanExecute("ProductDocument") Then Exit Sub
    
    'Docæ“¾
    Dim ProDoc As ProductDocument: Set ProDoc = CATIA.ActiveDocument
    Dim Pros As Products: Set Pros = ProDoc.Product.Products
    If Pros.Count < 2 Then Exit Sub
    
    'µÌß¼®İ•ÏX
    Dim AssyMode As AsmConstraintSettingAtt
    Set AssyMode = CATIA.SettingControllers.Item("CATAsmConstraintSettingCtrl")
    Dim OriginalMode As CatAsmPasteComponentMode
    OriginalMode = AssyMode.PasteComponentMode
    
    'µÌß¼®İØ‚è‘Ö‚¦
    AssyMode.PasteComponentMode = catPasteWithCstOnCopyAndCut
    
    '¿°ÄÏ‚İ–¼‘OØ½Ä
    Dim Names: Set Names = Get_SortedNames(Pros)
    
    '¶¯Ä
    Dim Sel As Selection: Set Sel = ProDoc.Selection
    Dim Itm As Variant
    
    CATIA.HSOSynchronized = False
    
    Sel.Clear
    For Each Itm In Names
        Sel.Add Pros.Item(Itm)
    Next
    Sel.Cut
    
    'Íß°½Ä
    With Sel
        .Clear
        .Add Pros
        .Paste
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    
    'µÌß¼®İ–ß‚µ,UpDate
    AssyMode.PasteComponentMode = OriginalMode
    ProDoc.Product.Update
End Sub

'²İ½Àİ½–¼‚Å¿°ÄÏ‚İ‚Ì–¼‘OØ½Ä
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
