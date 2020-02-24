Attribute VB_Name = "sample_ReOrder_Product"
'vba sample_ReOrder_Product ver0.0.1  using-'KCL0.0.12'  by Kantoku
'�ݽ�ݽ���ł̿�ď���Tree����ёւ��܂�

'{GP:11}
'{Caption:ص��ް}
'{ControlTipText:�ݽ�ݽ���ł̿�ď���Tree����ёւ��܂�
'{BackColor:16744703}
'{FONTSIZE:10.5}

Option Explicit

Sub CATMain()
    '�޷���Ă�����
    If Not CanExecute("ProductDocument") Then Exit Sub
    
    'Doc�擾
    Dim ProDoc As ProductDocument: Set ProDoc = CATIA.ActiveDocument
    Dim Pros As Products: Set Pros = ProDoc.Product.Products
    If Pros.Count < 2 Then Exit Sub
    
    '��߼�ݕύX
    Dim AssyMode As AsmConstraintSettingAtt
    Set AssyMode = CATIA.SettingControllers.Item("CATAsmConstraintSettingCtrl")
    Dim OriginalMode As CatAsmPasteComponentMode
    OriginalMode = AssyMode.PasteComponentMode
    
    '��߼�ݐ؂�ւ�
    AssyMode.PasteComponentMode = catPasteWithCstOnCopyAndCut
    
    '��čςݖ��Oؽ�
    Dim Names: Set Names = Get_SortedNames(Pros)
    
    '���
    Dim Sel As Selection: Set Sel = ProDoc.Selection
    Dim Itm As Variant
    
    CATIA.HSOSynchronized = False
    
    Sel.Clear
    For Each Itm In Names
        Sel.Add Pros.Item(Itm)
    Next
    Sel.Cut
    
    '�߰��
    With Sel
        .Clear
        .Add Pros
        .Paste
        .Clear
    End With
    
    CATIA.HSOSynchronized = True
    
    '��߼�ݖ߂�,UpDate
    AssyMode.PasteComponentMode = OriginalMode
    ProDoc.Product.Update
End Sub

'�ݽ�ݽ���ſ�čς݂̖��Oؽ�
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
