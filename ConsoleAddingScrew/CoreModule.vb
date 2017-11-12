Imports INFITF
Imports CATAssemblyTypeLib ' Include all
Imports KnowledgewareTypeLib
Imports PARTITF 'Include Hole
Imports MECMOD 'Include Bodies, Body, PartDocument, Part, Shapes
Imports ProductStructureTypeLib 'Include Product

Module CoreModule
    Public CATIA As Object
    Public mainDoc As INFITF.Document
    Private orShim10 As Boolean = False
    'Private orShim10 As Boolean = False
    Public libraryLocation As String = "T:\01\pp\lib\KUKA\VISSERIE"
    Public containingProduct As Product
    Public containingProducts As Products

    'Start Subroutine
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
        Else 'If everything works, then do the following
            Select Case MsgBox("You running program to insert threads/dowels. Do you have shim 10mm?", MsgBoxStyle.YesNo, "Tool to insert threads/dowels")
                Case MsgBoxResult.Yes
                    orShim10 = True
                    Exit Select
                Case MsgBoxResult.No
                    orShim10 = False
                    'Select Case MsgBox("Do you have shim 5mm?", MsgBoxStyle.YesNo, "Tool to insert threads/dowels")
                    '    Case MsgBoxResult.Yes
                    '        orShim10 = True
                    '        Exit Select
                    '    Case MsgBoxResult.No
                    '        orShim10 = False
                    '        Exit Select
                    'End Select
            End Select
            DetectHoles()
            'MoveAddedElements()
        End If
    End Sub
    'Loading All 
    Sub DetectHoles()
        Dim arrayHoles(1, 3) As String
        Dim oSelection2
        Dim twoElements, orDowel As Integer
        Dim orThreadingHole As CatHoleThreadingMode
        Dim lenghtHole1, lenghtHole2, lenghtElement As Double
        Dim arrayOfVariantOfBSTR1(0)
        Dim CompoObject As Composition
        CompoObject = New Composition

        Dim product1 As Product
        product1 = mainDoc.Product

        Dim products1 As Products
        products1 = product1.Products

        oSelection2 = mainDoc.Selection
        twoElements = oSelection2.Count

        'Loading arrayHoles
        For i = 0 To twoElements - 1
            orThreadingHole = oSelection2.Item(i + 1).Value.ThreadingMode
            If orThreadingHole = CatHoleThreadingMode.catThreadedHoleThreading Then
                'If the hole is a thread:
                arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.HoleThreadDescription.Value 'Only for thread e.g M10
            Else
                arrayHoles(i, 0) = oSelection2.Item(i + 1).Value.Diameter.Value 'e.g 10
                orDowel += 1
            End If
            arrayHoles(i, 1) = oSelection2.Item(i + 1).LeafProduct.Parent.Parent.Name 'e.g BG_04.1
            arrayHoles(i, 2) = oSelection2.Item(i + 1).Value.BottomLimit.Dimension.Value 'e.g What depth?
        Next

        'we assign a value
        containingProduct = products1.Item(arrayHoles(0, 1))
        containingProducts = containingProduct.Products

        'Counting
        Double.TryParse(arrayHoles(0, 2), lenghtHole1)
        Double.TryParse(arrayHoles(1, 2), lenghtHole2)

        If orShim10 = True Then
            lenghtElement = lenghtHole1 + lenghtHole2 + 10
        Else
            lenghtElement = lenghtHole1 + lenghtHole2
        End If

        'Which element we need to insert
        If orDowel = 2 Then 'Insert dowel
            CompoObject.searchForDowel(arrayHoles(0, 0), lenghtElement)
            'In original: \/
            'arrayOfVariantOfBSTR1(0) = libraryLocation & CompoObject.file
            'For tests in house: \/
            arrayOfVariantOfBSTR1(0) = "E:\Pliki 3D\SrubaM10x55.CATPart"
        Else 'Insert screw
            'Dal and i = 1 because first I choose a through hole !!!!!!!!!!!!!!!!!!!!
            CompoObject.searchForScrew(arrayHoles(1, 0), lenghtElement)
            'In original: \/
            'arrayOfVariantOfBSTR1(0) = libraryLocation & CompoObject.file
            'For tests in house: \/
            arrayOfVariantOfBSTR1(0) = "E:\Pliki 3D\SrubaM10x55.CATPart"
        End If

        ''For test on constraints
        'containingProducts.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

        'Dim constraints1 As Constraints
        'constraints1 = containingProduct.Connections("CATIAConstraints")

        ''Axis of screw
        'Dim reference1 As Reference
        'reference1 = containingProduct.CreateReferenceFromName("siema")

        ''Axis of
        'Dim reference2 As Reference
        'reference2 = containingProduct.CreateReferenceFromName("siema")

        'Dim constraint1 As Constraint
        'constraint1 = constraints1.AddBiEltCst(CatConstraintType.catCstTypeOn, reference1, reference2)

AnotherElement:
        oSelection2.clear()

        oSelection2.add(containingProducts)

        CATIA.StartCommand("Existing Component With Positioning")

        AppActivate("CATIA V5 - [" & mainDoc.Name & "]")

        Threading.Thread.Sleep(500)

        My.Computer.Keyboard.SendKeys(arrayOfVariantOfBSTR1(0), True)

        My.Computer.Keyboard.SendKeys("{ENTER}", True)

        'Add element
        'containingProducts.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

        oSelection2.clear()

        Select Case MsgBox("Another one?", MsgBoxStyle.YesNo, "Tool to insert threads/dowels")
            Case MsgBoxResult.Yes
                GoTo AnotherElement
                Exit Select
            Case MsgBoxResult.No
                Exit Sub
        End Select
    End Sub
    'For this moment is not neseserry
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

        containingProduct = containingProducts.Item(containingProducts.Count)

        CATIA.StartCommand("Existing Component With Positioning")

        'containingProduct.Move.Apply(iMatrix)

        'Dim pozycjaCompasu As DNBASY.AsyMotionTarget

        'pozycjaCompasu.GetCompassPosition(iMatrix, DNBASY.AsyMotionTargetDataFormat.AsyMotionTarget3x4Matrix)

        'proba.Move.Apply(iMatrix)
    End Sub

End Module