Imports INFITF
Imports CATAssemblyTypeLib ' Include all
Imports KnowledgewareTypeLib
Imports PARTITF 'Include Hole
Imports MECMOD 'Include Bodies, Body, PartDocument, Part, Shapes
Imports ProductStructureTypeLib 'Include Product

Module CoreModule
    Public CATIA As Object
    Public mainDoc As INFITF.Document
    Private czyShim10 As Boolean = False
    Private czyShim5 As Boolean = False

    'Startowy Subroutine
    Public Sub Main()
        Dim iErr As Integer

        On Error Resume Next
        CATIA = GetObject(, "CATIA.Application")
        iErr = Err.Number
        If (iErr <> 0) Then
            MsgBox("There is no open CATIA Application")
            Exit Sub
        End If

        mainDoc = CATIA.ActiveDocument

        If Err.Number <> 0 Then
            MsgBox("There is no open any component in CATIA")
            Exit Sub
        End If
        On Error GoTo 0
        If TypeName(mainDoc) <> "ProductDocument" Then
            MsgBox("In CATIA Active window must be the Assembly (.CATProduct)")
            Exit Sub
        Else ' Jeżeli wszystko działa to wykonuje się to co poniżej
            Select Case MsgBox("Do you want start to color holes?", MsgBoxStyle.YesNo, "Tool to color Holes")
                Case MsgBoxResult.Yes
                    Exit Select
                Case MsgBoxResult.No
                    Exit Sub
            End Select
            Select Case MsgBox("Do you have shim 10mm?", MsgBoxStyle.YesNo, "Tool to insert threads")
                Case MsgBoxResult.Yes
                    czyShim10 = True
                    Exit Select
                Case MsgBoxResult.No
                    czyShim10 = False
                    Select Case MsgBox("Do you have shim 5mm?", MsgBoxStyle.YesNo, "Tool to insert threads")
                        Case MsgBoxResult.Yes
                            czyShim5 = True
                            Exit Select
                        Case MsgBoxResult.No
                            czyShim5 = False
                            Exit Select
                    End Select
            End Select
            'arrayWithFeatures()
            DetectHoles()
        End If
    End Sub
    'Loading All 

    Sub DetectHoles()
        Dim arrayHoles(1, 3) As String
        Dim oSelection2
        Dim dwaElementy As Integer
        Dim czyGwintowanyOtwor As CatHoleThreadingMode
        Dim dlugoscHole1 As Double
        Dim dlugoscHole2 As Double
        Dim dlugoscSruby As Double
        Dim arrayOfVariantOfBSTR1(0)

        Dim newMatrix(11)

        Dim product1 As Product
        product1 = mainDoc.Product

        Dim products1 As Products
        products1 = product1.Products

        Dim product2 As Product
        Dim products2 As Products

        oSelection2 = mainDoc.Selection
        dwaElementy = oSelection2.Count

        'Ładowanie arrayHoles
        For i = 0 To oSelection2.Count - 1
            czyGwintowanyOtwor = oSelection2.ITEM(i + 1).Value.ThreadingMode
            If czyGwintowanyOtwor = CatHoleThreadingMode.catThreadedHoleThreading Then
                'jeżeli hole jest z gwintem to:
                arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.HoleThreadDescription.Value ' tylko dla thread np M10
            Else
                arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.Diameter.Value 'np 10 
            End If

            'arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.Name 'np Hole.1
            'arrayHoles(i, 1) = oSelection2.Item(i + 1).LeafProduct.PartNumber 'np Konsola
            arrayHoles(i, 1) = oSelection2.Item(i + 1).LeafProduct.Parent.Parent.Name ' np BG_04.1
            arrayHoles(i, 2) = oSelection2.Item(i + 1).Value.BottomLimit.Dimension.Value 'np Hole.1
        Next

        'przypisujemy wartosc
        product2 = products1.Item(arrayHoles(0, 1))
        products2 = product2.Products

        'Liczymy

        Double.TryParse(arrayHoles(0, 2), dlugoscHole1)
        Double.TryParse(arrayHoles(1, 2), dlugoscHole2)

        If czyShim10 = True Then
            dlugoscSruby = dlugoscHole1 + dlugoscHole2 + 10
        ElseIf czyShim5 = True Then
            dlugoscSruby = dlugoscHole1 + dlugoscHole2 + 5
        Else
            dlugoscSruby = dlugoscHole1 + dlugoscHole2
        End If

        'Wstawiamy elementy  głowna lokalizacja elementów: T:\01\pp\lib\KUKA\VISSERIE
        If dlugoscSruby > 55 Then
            arrayOfVariantOfBSTR1(0) = "E:\Pliki 3D\SrubaM10x55.CATPart"
        ElseIf dlugoscSruby > 50 Then
            arrayOfVariantOfBSTR1(0) = "E:\Pliki 3D\SrubaM10x50.CATPart"
        End If

        products2.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

    End Sub

    Sub MoveAddedElements()
        Dim iMatrix(11)
        iMatrix(0) = 1.0
        iMatrix(1) = 0.0
        iMatrix(2) = 0.0
        iMatrix(3) = 0.0
        iMatrix(4) = 1.0
        iMatrix(5) = 0.0
        iMatrix(6) = 0.0
        iMatrix(7) = 0.0
        iMatrix(8) = 1.0
        iMatrix(9) = 100.0
        iMatrix(10) = 100.0
        iMatrix(11) = 100.0

        'Przesuwanie elementów
        'Dim sruba As Product
        'sruba = products2.Item(products2.Count)

        'sruba.Move.Apply(iMatrix)

        ' Dim proba As Product
        'proba = products2.Item(3)

        'Dim pozycjaCompasu As DNBASY.AsyMotionTarget

        'GetCompassPosition(iMatrix, DNBASY.AsyMotionTargetDataFormat.AsyMotionTarget3x4Matrix)

        'proba.Move.Apply(iMatrix)
    End Sub

    Sub arrayWithFeatures()
        Dim arrayPomocne(0, 2) As String
        Dim oSelection
        Dim ileHole As Integer
        Dim visPropertySet As VisPropertySet

        oSelection = mainDoc.Selection
        oSelection.Clear()

        visPropertySet = oSelection.VisProperties

        oSelection.Search("n:*Hole*,all")

        ileHole = oSelection.Count
        ReDim arrayPomocne(ileHole - 1, 2)

        For i = 0 To oSelection.Count - 1
            arrayPomocne(i, 0) = oSelection.Item(i + 1).Value.Name
            arrayPomocne(i, 1) = oSelection.Item(i + 1).LeafProduct.PartNumber
            arrayPomocne(i, 2) = oSelection.Item(i + 1).Value.Diameter.Value
        Next

        'Druga Petla
        For InxSel = 0 To ileHole - 1
            oSelection.Clear()

            Dim documents1 As Documents
            documents1 = CATIA.Documents

            Dim partDocument1 As PartDocument
            partDocument1 = documents1.Item(arrayPomocne(InxSel, 1) & ".CATPart")

            Dim part1 As Part
            part1 = partDocument1.Part

            Dim bodies1 As Bodies
            bodies1 = part1.Bodies

            Dim body1 As Body
            body1 = bodies1.Item("PartBody")

            Dim shapes1 As Shapes
            shapes1 = body1.Shapes

            Dim hole1 As Hole
            hole1 = shapes1.Item(arrayPomocne(InxSel, 0))

            oSelection.Add(hole1)
            If arrayPomocne(InxSel, 2) = 4 Then
                oSelection.VisProperties.SetRealColor(0, 133, 255, 0)
            ElseIf arrayPomocne(InxSel, 2) = 5 Then
                oSelection.VisProperties.SetRealColor(0, 133, 255, 0)
            ElseIf arrayPomocne(InxSel, 2) = 6 Then
                oSelection.VisProperties.SetRealColor(0, 133, 255, 0)
            ElseIf arrayPomocne(InxSel, 2) = 8 Then
                oSelection.VisProperties.SetRealColor(0, 133, 255, 0)
            ElseIf arrayPomocne(InxSel, 2) = 10 Then
                oSelection.VisProperties.SetRealColor(0, 133, 255, 0)
            ElseIf arrayPomocne(InxSel, 2) = 12 Then
                oSelection.VisProperties.SetRealColor(0, 133, 255, 0)
            ElseIf arrayPomocne(InxSel, 2) = 4.5 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            ElseIf arrayPomocne(InxSel, 2) = 4.5 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            ElseIf arrayPomocne(InxSel, 2) = 5.5 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            ElseIf arrayPomocne(InxSel, 2) = 6.6 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            ElseIf arrayPomocne(InxSel, 2) = 9 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            ElseIf arrayPomocne(InxSel, 2) = 11 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            ElseIf arrayPomocne(InxSel, 2) = 13.5 Then
                oSelection.VisProperties.SetRealColor(0, 175, 0, 0)
            Else
                oSelection.VisProperties.SetRealColor(230, 239, 20, 0)
            End If
        Next
        oSelection.Clear()

        MsgBox("Colored: " & ileHole & " elements.")

    End Sub

End Module