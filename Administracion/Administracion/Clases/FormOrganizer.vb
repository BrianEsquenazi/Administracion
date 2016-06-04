﻿Public Delegate Function QueryFunction(ByVal text As String)
Public Delegate Sub ShowMethod(ByVal selectedValue)

Public Class FormOrganizer

    Private form As Form
    Private maxHeight As Integer
    Private width As Integer
    Private queryFunctions As New List(Of Tuple(Of QueryFunction, String, ShowMethod))
    Private queryFunction As QueryFunction
    Private queryText As CustomTextBox
    Private queryList As CustomListBox
    Private selectionList As CustomListBox
    Private allControls As New List(Of CustomControl)
    Private firstColumnControls As New List(Of CustomControl)
    Private columnedControls As List(Of Object)
    Private compactedControls As List(Of Object)
    Private compactedFirstColumnControls As New List(Of CustomControl)
    Private annexedControls As New List(Of CustomControl)
    Private buttons As New List(Of CustomButton)
    Private buttonsTop As New List(Of CustomButton)
    Private buttonsBottom As New List(Of CustomButton)
    Private btnAdd As CustomButton
    Private btnAddClick As EventHandler
    Private btnDelete As CustomButton
    Private btnDeleteClick As EventHandler
    Private btnClean As CustomButton
    Private btnCleanClick As EventHandler
    Private btnQuery As CustomButton
    Private btnQueryClick As EventHandler
    Private btnList As CustomButton
    Private btnListClick As EventHandler
    Private btnClose As CustomButton
    Private btnCloseClick As EventHandler
    Private showMethodFunction As ShowMethod
    Private usingQueryText As Boolean = True
    Private organizingCompactedControls As Boolean = False
    Private isCompact = False

    Private buttonsWidth As Integer
    Private topMargin As Integer = 30
    Private leftMargin As Integer = 30
    Private rightMargin As Integer = 30
    Private bottomMargin As Integer = 30
    Private separation As Integer = 6
    Private simpleButtonHeight As Integer = 35
    Private listQueryHeight As Integer = 240
    Private separationBetweenControlsAndButtons As Integer = 45
    Private controlSeparation As Integer = 10
    Private charPixelSize As Double = 7.5
    Private compactHeightConstant As Integer = 0
    Dim timer As New Timer

    Public Sub New(ByVal someForm As Form, ByVal formWidth As Integer, ByVal formHeight As Integer)
        form = someForm
        maxHeight = formHeight
        width = formWidth
        buttonsWidth = width - leftMargin - rightMargin
    End Sub

    Public Sub addControls(ByVal ParamArray controlsColumn() As Object)
        addControls(columnedControls, firstColumnControls, controlsColumn)
    End Sub

    Public Sub addCompactedControls(ByVal ParamArray controlsColumn() As Object)
        addControls(compactedControls, compactedFirstColumnControls, controlsColumn)
    End Sub

    Private Sub addControls(ByRef objectsCollection As List(Of Object), ByVal firstColumnCollection As List(Of CustomControl), ByVal ParamArray controlsColumn() As Object)
        objectsCollection = controlsColumn.ToList
        For Each control In objectsCollection
            If control.GetType.GetInterface(GetType(CustomControl).Name) = GetType(CustomControl) Then
                firstColumnCollection.Add(control)
                allControls.Add(control)
            Else
                Dim listOfControls As List(Of Control) = castToListOfControl(DirectCast(control, Object()))
                firstColumnCollection.Add(listOfControls.First)
                listOfControls.ForEach(Sub(aControl) allControls.Add(DirectCast(aControl, CustomControl)))
            End If
        Next
    End Sub

    Public Sub addAnnexedControls(ByVal someControls As List(Of CustomControl))
        annexedControls = someControls
        annexedControls.OrderBy(Function(control) control.LabelAssociationKey)
    End Sub

    Private Function controlsHeight()
        Return firstColumnControls.Sum(Function(control As CustomControl) DirectCast(control, Control).Height) + (firstColumnControls.Count - 1) * controlSeparation
    End Function

    Private Function buttonsHeight()
        Return simpleButtonHeight * 2 + separation 'Las dos filas de botones + separation entre botones
    End Function

    Private Function formNormalHeight()
        Return topMargin + bottomMargin + controlsHeight() + separationBetweenControlsAndButtons + buttonsHeight() + compactHeightConstant
    End Function

    Private Function formNeededHeight()
        Return formNormalHeight() + compactedFirstColumnControls.Sum(Function(control As CustomControl) DirectCast(control, Control).Height) + (compactedFirstColumnControls.Count - 1) * controlSeparation - buttonsHeight()
    End Function

    Private Function formWithQueryControlsHeight()
        Dim height As Integer = formNormalHeight() + separation + queryList.Height
        If usingQueryText Then
            height += separation + queryTextHeight()
        End If
        Return height
    End Function

    Private Function formWithSelectionListHeight()
        Return formNormalHeight() + separation + selectionList.Height
    End Function

    Public Sub organize()
        CommonEventsHandler.setIndexTab(form)
        form.Height = Math.Min(maxHeight, formNeededHeight)
        form.Width = width

        Dim btnsTop As Integer = organizeControls() + separationBetweenControlsAndButtons
        organizeButtons(btnsTop)
        organizeQueryControllers(btnsTop + buttonsHeight())
    End Sub

    Public Sub compactOrganize()
        topMargin = 10
        bottomMargin = 10
        leftMargin = 20
        simpleButtonHeight = 25
        separationBetweenControlsAndButtons = 10
        controlSeparation = separation
        compactHeightConstant = 30
        buttonsWidth = buttonsWidth \ 2
        isCompact = True
        organize()
    End Sub

    Private Function organizeControls()
        Dim top As Integer = topMargin
        top = organizeControls(top, leftMargin, columnedControls)
        If isCompact Then
            organizingCompactedControls = True
            organizeCompactedControls(top + separationBetweenControlsAndButtons)
            organizingCompactedControls = False
        End If
        Return top
    End Function

    Private Sub organizeCompactedControls(ByVal top As Integer)
        organizeControls(top, leftMargin + buttonsWidth + separation, compactedControls)
    End Sub

    Private Function organizeControls(ByVal top As Integer, ByVal realLeftMargin As Integer, ByVal collectionOfControls As List(Of Object))
        'Setteo el top y el left de los controls y labels
        For Each castControl As Object In collectionOfControls
            If castControl.GetType.GetInterface(GetType(CustomControl).Name) = GetType(CustomControl) Then
                top = organizeSimpleControl(castControl, top, realLeftMargin)
            Else
                top = organizeColumnControl(castControl, top, realLeftMargin)
            End If
        Next
        Return top - topMargin
    End Function

    Private Function organizeSimpleControl(ByVal castControl As Control, ByVal top As Integer, ByVal realLeftMargin As Integer)
        'EXCLUSIVO CONTROLES SIN COLUMNAS
        Dim left As Integer = realLeftMargin
        Dim control As CustomControl = DirectCast(castControl, CustomControl)

        setLabelTopFor(control.LabelAssociationKey, top)
        setLabelLeftFor(control.LabelAssociationKey, realLeftMargin)

        left += separation + maxLabelWidth() 'Se cambia el left para el próximo control

        castControl.Top = top - 3
        castControl.Left = left

        Dim maxWidth As Integer = maxWidthForColumnControls(New List(Of Control) From {castControl})
        Dim widthPercentage As Double = widthPercentageForColumnControls(New List(Of Control) From {castControl})
        castControl.Width = maxWidthFor(control, 0, maxWidth, widthPercentage)

        left += castControl.Width + separation

        Dim annexedCustomControls = annexedControlsFor(control.LabelAssociationKey)
        For Each annexedCustomControl As CustomControl In annexedCustomControls
            Dim annexedControl As Control = DirectCast(annexedCustomControl, Control)
            annexedControl.Top = top - 3
            annexedControl.Left = left
            annexedControl.Width = maxWidthFor(annexedCustomControl, 1, maxWidth, widthPercentage)
            left += annexedControl.Width + separation
        Next

        Return top + castControl.Height + controlSeparation
    End Function

    Private Function maxWidthFor(ByVal control As CustomControl, ByVal columnNumber As Integer, ByVal maxWidth As Integer, ByVal widthPercentage As Double)
        If variableWidthFor(control, columnNumber) > maxWidth Then
            Return variableWidthFor(control, columnNumber) * widthPercentage
        Else
            Return variableWidthFor(control, columnNumber)
        End If
    End Function

    Private Function organizeColumnControl(ByVal castControl As Object, ByVal top As Integer, ByVal realLeftMargin As Integer)
        'EXCLUSIVO PARA CONTROLS ENCOLUMNADOS
        Dim left As Integer = realLeftMargin
        Dim controlList As List(Of Control) = castToListOfControl(castControl)
        Dim maxWidth As Integer = maxWidthForColumnControls(controlList)

        For Each specificControl In controlList
            Dim custControl As CustomControl = DirectCast(specificControl, CustomControl)
            Dim index As Integer = controlList.IndexOf(specificControl)

            setLabelTopFor(custControl.LabelAssociationKey, top)
            setLabelLeftFor(custControl.LabelAssociationKey, left)
            If left = realLeftMargin Then
                left += separation + maxLabelWidth() 'Se cambia el left para el próximo control
            Else
                left += separation + labelWidthFor(custControl.LabelAssociationKey)
            End If

            specificControl.Top = top - 3
            specificControl.Left = left
            specificControl.Width = Math.Min(variableWidthFor(custControl, index), maxWidth)

            left += specificControl.Width + separation 'Se cambia el left para el próximo control

            Dim annexedCustomControls = annexedControlsFor(custControl.LabelAssociationKey)
            For Each annexedCustomControl As CustomControl In annexedCustomControls
                Dim annexedControl As Control = DirectCast(annexedCustomControl, Control)
                annexedControl.Top = top - 3
                annexedControl.Left = left

                annexedControl.Width = Math.Min(variableWidthFor(custControl, index), maxWidth)
                left += annexedControl.Width + separation 'Se cambia el left para el próximo control
            Next
        Next
        Return top + controlList(0).Height + controlSeparation
    End Function

    Private Function maxWidthForColumnControls(ByVal controlList As List(Of Control)) 'TODO Mejorar para que devuelva el coeficiente de resta y no una división de cantidad
        Dim annexedControlsCount As Integer = controlList.Sum(Function(control) annexedControlsFor(DirectCast(control, CustomControl).LabelAssociationKey).Count)

        Dim availableSpace As Integer
        If organizingCompactedControls Then
            availableSpace = buttonsWidth - maxLabelWidth() + labelWidthFor(DirectCast(controlList(0), CustomControl).LabelAssociationKey) 'Ya que el maxLabelWidth suma el label más largo de la primer columna, se "resta" el primero
        Else
            availableSpace = form.Width - leftMargin - rightMargin - maxLabelWidth() + labelWidthFor(DirectCast(controlList(0), CustomControl).LabelAssociationKey) 'Ya que el maxLabelWidth suma el label más largo de la primer columna, se "resta" el primero
        End If
        availableSpace = availableSpace - (controlList.Count - 1) * 2 * separation - controlList.Sum(Function(control) labelWidthFor(DirectCast(control, CustomControl).LabelAssociationKey))
        availableSpace = availableSpace - annexedControlsCount * separation

        Dim avgMax As Integer = availableSpace \ (controlList.Count + annexedControlsCount)

        Dim smallerControls As List(Of Control) = controlList.FindAll(Function(control) Math.Min(variableWidthFor(DirectCast(control, CustomControl), controlList.IndexOf(control)), avgMax) < avgMax)
        Dim smallerControlsSize As Integer = smallerControls.Sum(Function(control) variableWidthFor(DirectCast(control, CustomControl), controlList.IndexOf(control)) + 2 * separation)

        Dim biggerControlsCount As Integer = controlList.Count - smallerControls.Count

        If biggerControlsCount <> 0 Then
            Return (availableSpace - smallerControlsSize) \ (controlList.Count - smallerControls.Count + annexedControlsCount)
        Else
            Return avgMax
        End If
    End Function

    Private Function widthPercentageForColumnControls(ByVal controlList As List(Of Control)) As Double 'TODO Mejorar para que devuelva el coeficiente de resta y no una división de cantidad
        Dim annexedControlsInRow As New List(Of CustomControl)
        controlList.ConvertAll(Function(control) annexedControlsFor(DirectCast(control, CustomControl).LabelAssociationKey)).ForEach(Sub(list) annexedControlsInRow.AddRange(list))
        Dim annexedControlsCount As Integer = annexedControlsInRow.Count

        Dim availableSpace As Integer
        If organizingCompactedControls Then
            availableSpace = buttonsWidth - maxLabelWidth() + labelWidthFor(DirectCast(controlList(0), CustomControl).LabelAssociationKey) 'Ya que el maxLabelWidth suma el label más largo de la primer columna, se "resta" el primero
        Else
            availableSpace = form.Width - leftMargin - rightMargin - maxLabelWidth() + labelWidthFor(DirectCast(controlList(0), CustomControl).LabelAssociationKey) 'Ya que el maxLabelWidth suma el label más largo de la primer columna, se "resta" el primero
        End If
        availableSpace = availableSpace - (controlList.Count - 1) * 2 * separation - controlList.Sum(Function(control) labelWidthFor(DirectCast(control, CustomControl).LabelAssociationKey))
        availableSpace = availableSpace - annexedControlsCount * separation

        Dim avgMax As Integer = availableSpace \ (controlList.Count + annexedControlsCount)

        Dim smallerControls As List(Of Control) = controlList.FindAll(Function(control) Math.Min(variableWidthFor(DirectCast(control, CustomControl), controlList.IndexOf(control)), avgMax) < avgMax)
        Dim smallerControlsSize As Integer = smallerControls.Sum(Function(control) variableWidthFor(DirectCast(control, CustomControl), controlList.IndexOf(control)) + 2 * separation)

        Return Math.Min((availableSpace - smallerControlsSize) / (controlList.Sum(Function(control) variableWidthFor(control)) + annexedControlsInRow.Sum(Function(control) variableWidthFor(DirectCast(control, Control)))), 1)
    End Function

    Private Function castToListOfControl(ByVal controls As Object) As List(Of Control)
        Dim list As New List(Of Control)
        For Each c In DirectCast(controls, Object())
            list.Add(DirectCast(c, Control))
        Next
        Return list
    End Function

    Private Function variableWidthFor(ByVal control As CustomControl)
        Return variableWidthFor(control, 0)
    End Function

    Private Function variableWidthFor(ByVal textBox As CustomTextBox, ByVal columnNumber As Integer)
        Return Math.Min(Math.Round(textBox.MaxLength * charPixelSize), variableWidthOfControl(textBox, columnNumber))
    End Function

    Private Function variableWidthFor(ByVal control As CustomControl, ByVal columnNumber As Integer)
        Try
            Return variableWidthFor(DirectCast(control, CustomTextBox), columnNumber)
        Catch e As InvalidCastException
            Return variableWidthOfControl(control, columnNumber)
        End Try
    End Function

    Private Function variableWidthOfControl(ByVal control As CustomControl, ByVal columnNumber As Integer)
        Dim labelWidth As Integer
        If columnNumber = 0 Then
            labelWidth = maxLabelWidth()
        Else
            labelWidth = labelWidthFor(control.LabelAssociationKey)
        End If

        If organizingCompactedControls Then
            Return buttonsWidth - labelWidth - separation
        Else
            Return width - leftMargin - rightMargin - labelWidth - separation
        End If
    End Function

    Private Function maxLabelWidth(ByVal collection As List(Of CustomControl))
        Return collection.ConvertAll(Function(control) labelWidthFor(control.LabelAssociationKey)).Max()
    End Function

    Private Function maxLabelWidth()
        If organizingCompactedControls Then
            Return maxLabelWidth(compactedFirstColumnControls)
        Else
            Return maxLabelWidth(firstColumnControls)
        End If
    End Function

    Private Sub organizeButtons(ByVal top As Integer)
        createButtons()

        Dim buttonWidth As Integer = (buttonsWidth - separation * 3) \ 3
        buttons.ForEach(Sub(button) button.Width = buttonWidth)
        buttons.ForEach(Sub(button) button.Height = simpleButtonHeight)

        Dim left As Integer = leftMargin
        For Each button As CustomButton In buttonsTop
            setButtonPosition(button, top, left)
            left += buttonWidth + separation
        Next

        left = leftMargin
        For Each button As CustomButton In buttonsBottom
            setButtonPosition(button, top + button.Height + separation, left)
            left += buttonWidth + separation
        Next
    End Sub

    Private Sub organizeQueryControllers(ByVal top As Integer)
        If usingQueryText Then
            queryText = New CustomTextBox
            queryText.Parent = form
            queryText.Name = "txtQuery"
            queryText.Width = buttonsWidth
            queryText.Top = top + separation
            top = queryText.Top
            queryText.Left = leftMargin
            queryText.Visible = False
            AddHandler queryText.KeyDown, AddressOf queryTextEnterPressed
        End If

        queryList = New CustomListBox
        queryList.Parent = form
        queryList.Name = "lstQuery"
        queryList.Width = buttonsWidth
        queryList.Height = Math.Min(listQueryHeight, maxHeight - formNormalHeight() - compactHeightConstant)
        queryList.Top = top + queryTextHeight() + separation
        queryList.Left = leftMargin
        queryList.Visible = False
        AddHandler queryList.DoubleClick, AddressOf listDoubleClickEventWithHide

        If queryFunctions.Count > 1 Then
            selectionList = New CustomListBox
            selectionList.Parent = form
            selectionList.Name = "lstQuery"
            selectionList.Width = buttonsWidth
            selectionList.Height = Math.Min(Math.Round(listQueryHeight / 2), maxHeight - formNormalHeight() - compactHeightConstant)
            selectionList.Top = top + separation
            selectionList.Left = leftMargin
            selectionList.Visible = False
            AddHandler selectionList.DoubleClick, AddressOf selectionDoubleClick
        End If
    End Sub

    Private Sub setButtonPosition(ByVal button As CustomButton, ByVal top As Integer, ByVal left As Integer)
        button.Top = top
        button.Left = left
    End Sub

    Private Sub createButtons()
        buttonsTop.Add(addButton)
        buttonsTop.Add(deleteButton)
        buttonsTop.Add(cleanButton)

        buttonsBottom.Add(queryButton)
        buttonsBottom.Add(listButton)
        buttonsBottom.Add(closeButton)

        buttons.AddRange(buttonsTop)
        buttons.AddRange(buttonsBottom)
    End Sub

    Private Function addButton()
        Dim btn As CustomButton
        If IsNothing(btnAdd) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnAdd"
            btn.Text = "Agregar"
        Else
            btn = btnAdd
        End If

        If Not IsNothing(btnAddClick) Then
            AddHandler btn.Click, AddressOf addClickWithClean
        End If

        Return btn
    End Function

    Private Function deleteButton()
        Dim btn As CustomButton
        If IsNothing(btnDelete) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnDelete"
            btn.Text = "Eliminar"
        Else
            btn = btnDelete
        End If

        If Not IsNothing(btnDeleteClick) Then
            AddHandler btn.Click, AddressOf deleteClickWithConfirmation
        End If

        Return btn
    End Function

    Private Function cleanButton()
        Dim btn As CustomButton
        If IsNothing(btnClean) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnClean"
            btn.Text = "Limpiar"
        Else
            btn = btnClean
        End If

        If Not IsNothing(btnCleanClick) Then
            AddHandler btn.Click, btnCleanClick
        End If

        Return btn
    End Function

    Private Function queryButton()
        Dim btn As CustomButton
        If IsNothing(btnQuery) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnQuery"
            btn.Text = "Consulta"
        Else
            btn = btnQuery
        End If

        If Not IsNothing(btnQueryClick) Then
            AddHandler btn.Click, btnQueryClick
        End If

        Return btn
    End Function

    Private Function listButton()
        Dim btn As CustomButton
        If IsNothing(btnList) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnList"
            btn.Text = "Listado"
        Else
            btn = btnList
        End If

        If Not IsNothing(btnListClick) Then
            AddHandler btn.Click, btnListClick
        End If

        Return btn
    End Function

    Private Function closeButton()
        Dim btn As CustomButton
        If IsNothing(btnClose) Then
            btn = New CustomButton()
            btn.Parent = form
            btn.Name = "btnClose"
            btn.Text = "Cerrar"
        Else
            btn = btnClose
        End If

        If Not IsNothing(btnCloseClick) Then
            AddHandler btn.Click, btnCloseClick
        End If

        Return btn
    End Function

    Private Function labelWidthFor(ByVal index As Integer)
        If IsNothing(labelFor(index)) Then
            Return 0
        Else
            Return labelFor(index).Width
        End If
    End Function

    Private Sub setLabelTopFor(ByVal index As Integer, ByVal top As Integer)
        If Not IsNothing(labelFor(index)) Then
            labelFor(index).Top = top
        End If
    End Sub

    Private Sub setLabelLeftFor(ByVal index As Integer, ByVal left As Integer)
        If Not IsNothing(labelFor(index)) Then
            labelFor(index).Left = left
        End If
    End Sub

    Private Function labelFor(ByVal index As Integer) As CustomLabel
        Return form.Controls.OfType(Of CustomLabel).ToList.Find(Function(label) label.ControlAssociationKey = index)
    End Function

    Private Function annexedControlsFor(ByVal index As Integer) As List(Of CustomControl)
        Dim list = annexedControls.FindAll(Function(control) control.LabelAssociationKey = index)
        If IsNothing(list) Then
            Return New List(Of CustomControl)
        Else
            Return list
        End If
    End Function

    Private Function validateForDelete()
        Return validateControls(False)
    End Function
    Private Function validateForAdd()
        Return validateControls(True)
    End Function
    Private Function validateControls(ByVal isAdd As Boolean)
        Dim firstControl As CustomTextBox = allControls.Find(Function(control) control.EnterIndex = 1) 'TODO HACERLO GENÉRICO PARA TODOS LOS CONTROLERS
        Dim validator As New Validator
        validator.validate(firstControl.Text, firstControl.Validator, labelFor(firstControl.LabelAssociationKey).Text)
        If isAdd Then
            Dim controlsToValidate As List(Of CustomTextBox) = allControls.OfType(Of CustomTextBox).ToList 'TODO HACERLO GENÉRICO PARA TODOS LOS CONTROLERS
            controlsToValidate.Remove(firstControl)
            controlsToValidate.ForEach(Sub(control) validator.validate(control.Text, control.Validator, labelFor(control.LabelAssociationKey).Text))
        End If
        Return validator.flush()
    End Function

    Public Sub addButton(ByVal button As CustomButton)
        btnAdd = button
    End Sub
    Public Sub deleteButton(ByVal button As CustomButton)
        btnDelete = button
    End Sub
    Public Sub cleanButton(ByVal button As CustomButton)
        btnClean = button
    End Sub
    Public Sub queryButton(ByVal button As CustomButton)
        btnQuery = button
    End Sub
    Public Sub listButton(ByVal button As CustomButton)
        btnList = button
    End Sub
    Public Sub closeButton(ByVal button As CustomButton)
        btnClose = button
    End Sub

    Public Sub setAddButtonClick(ByRef btnClick As EventHandler)
        btnAddClick = btnClick
    End Sub
    Public Sub setDeleteButtonClick(ByRef btnClick As EventHandler)
        btnDeleteClick = btnClick
    End Sub
    Public Sub setCleanButtonClick(ByRef btnClick As EventHandler)
        btnCleanClick = btnClick
    End Sub
    Public Sub setQueryButtonClick(ByRef btnClick As EventHandler)
        btnQueryClick = btnClick
    End Sub
    Public Sub setListButtonClick(ByRef btnClick As EventHandler)
        btnListClick = btnClick
    End Sub
    Public Sub setCloseButtonClick(ByRef btnClick As EventHandler)
        btnCloseClick = btnClick
    End Sub

    Public Sub setDefaultCleanButtonClick()
        btnCleanClick = AddressOf defaultCleanClick
    End Sub
    Public Sub addQueryFunction(ByVal listFunction As QueryFunction, ByVal name As String, ByVal showFunction As ShowMethod)
        queryFunctions.Add(Tuple.Create(listFunction, name, showFunction))
        btnQueryClick = AddressOf defaultQueryClick
    End Sub
    Public Sub dontUseQueryText()
        usingQueryText = False
    End Sub
    Public Sub setDefaultCloseButtonClick()
        btnCloseClick = AddressOf defaultCloseClick
    End Sub

    Private Sub listDoubleClickEventWithHide(ByVal sender As Object, ByVal e As EventArgs)
        showMethodFunction.Invoke(queryList.SelectedValue)
        queryFunction = Nothing
        hideQueryControls()
    End Sub
    Private Sub selectionDoubleClick(ByVal sender As Object, ByVal e As EventArgs)
        queryFunction = selectionList.SelectedValue.Item1
        showMethodFunction = selectionList.SelectedValue.Item3
        hideSelectionList()
        showQueryList()
    End Sub

    Private Sub addClickWithClean(ByVal sender As Object, ByVal e As EventArgs)
        If validateForAdd() Then
            btnAddClick.Invoke(sender, e)
            Cleanner.clean(form)
        End If
    End Sub
    Private Sub deleteClickWithConfirmation(ByVal sender As Object, ByVal e As EventArgs)
        If validateForDelete() Then
            If MsgBox("¿Desea eliminar el registro?", MsgBoxStyle.YesNo, "Eliminar") = vbYes Then
                btnDeleteClick.Invoke(sender, e)
                Cleanner.clean(form)
            End If
        End If
    End Sub

    Private Sub defaultCleanClick(ByVal sender As Object, ByVal e As EventArgs)
        Cleanner.clean(form)
    End Sub
    Private Sub defaultQueryClick(ByVal sender As Object, ByVal e As EventArgs)
        Select Case queryFunctions.Count
            Case 0
            Case 1
                queryFunction = queryFunctions.First().Item1
                showMethodFunction = queryFunctions.First().Item3
                showQueryList()
            Case Else
                If queryList.Visible Then
                    hideQueryControls()
                End If
                showSelectionList()
        End Select
    End Sub
    Private Sub showQueryList()
        queryTextEnterPressed(Nothing, Nothing)
        If usingQueryText Then
            queryText.Visible = True
            queryText.Focus()
        End If
        queryList.Visible = True
        form.Height = formWithQueryControlsHeight()
    End Sub
    Private Sub showSelectionList()
        selectionList.DataSource = queryFunctions
        selectionList.DisplayMember = "Item2"
        selectionList.Visible = True
        form.Height = formWithSelectionListHeight()
    End Sub

    Private Sub defaultCloseClick(ByVal sender As Object, ByVal e As EventArgs)
        form.Close()
    End Sub

    Private Sub queryTextEnterPressed(ByVal sender As Object, ByVal e As KeyEventArgs)
        If IsNothing(e) OrElse e.KeyValue = Keys.Enter Then
            Dim searchText As String = ""
            If usingQueryText Then
                searchText = queryText.Text
            End If
            queryList.DataSource = queryFunction.Invoke(searchText)
        End If
    End Sub

    Private Sub hideQueryControls()
        If usingQueryText Then
            queryText.Visible = False
            queryText.Text = ""
        End If
        queryList.Visible = False
        hideSelectionList()
        form.Height = Math.Max(formNormalHeight(), formNeededHeight())
    End Sub
    Private Sub hideSelectionList()
        If Not IsNothing(selectionList) Then
            selectionList.Visible = False
            form.Height = formNormalHeight()
        End If
    End Sub

    Private Function queryTextHeight() As Integer
        Dim heigth As Integer = 0
        If usingQueryText Then
            heigth = queryText.Height
        End If
        Return heigth
    End Function

End Class
