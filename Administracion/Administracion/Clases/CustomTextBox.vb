Public Enum ValidatorType
    None = 0
    NotEmpty
    Numeric
    Positive
    PositiveWithMax
    DateFormat
End Enum

Public Enum NumericType
    None = 0
    ShortType
    IntegerType
    LongType
    DoubleType
    FloatType
End Enum

Public Interface CustomControl
    Property EnterIndex() As Integer
    Property Cleanable() As Boolean
    Property LabelAssociationKey() As Integer
    Sub setHeight(ByVal value As Integer)
    Sub setWidth(ByVal value As Integer)
    Sub setTop(ByVal value As Integer)
    Sub setLeft(ByVal value As Integer)
End Interface

'TextBox
Public Class CustomTextBox
    Inherits TextBox
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False
    Private emptyPermitted As Boolean = True
    Private validatorConstant As ValidatorType = ValidatorType.None
    Private associationKey As Integer = -1

    Private Function getEnterIndex() As Integer
        Return customIndex
    End Function

    Private Sub setEnterIndex(ByVal anIndex As Integer)
        customIndex = anIndex
    End Sub

    Private Function getCleanStatus() As Integer
        Return cleanStatus
    End Function

    Private Sub setCleanStatus(ByVal status As Integer)
        cleanStatus = status
    End Sub

    Private Function getEmptyPermitted() As Boolean
        Return emptyPermitted
    End Function

    Private Sub setEmptyPermitted(ByVal value As Boolean)
        emptyPermitted = value
    End Sub

    Private Function getValidatorType() As Integer
        Return validatorConstant
    End Function

    Private Sub setValidatorType(ByVal type As Integer)
        validatorConstant = type
    End Sub

    Private Function getAssociationKey() As Integer
        Return associationKey
    End Function

    Private Sub setAssociationKey(ByVal key As Integer)
        associationKey = key
    End Sub

    Public Property EnterIndex() As Integer Implements CustomControl.EnterIndex
        Get
            Return CType(getEnterIndex(), Integer)
        End Get
        Set(ByVal value As Integer)
            setEnterIndex(value)
        End Set
    End Property

    Public Property Cleanable() As Boolean Implements CustomControl.Cleanable
        Get
            Return CType(getCleanStatus(), Boolean)
        End Get
        Set(ByVal value As Boolean)
            setCleanStatus(value)
        End Set
    End Property

    Public Property Empty() As Boolean
        Get
            Return CType(getEmptyPermitted(), Boolean)
        End Get
        Set(ByVal value As Boolean)
            setEmptyPermitted(value)
        End Set
    End Property

    Public Property Validator() As ValidatorType
        Get
            Return CType(getValidatorType(), ValidatorType)
        End Get
        Set(ByVal value As ValidatorType)
            setValidatorType(value)
        End Set
    End Property

    Public Property LabelAssociationKey() As Integer Implements CustomControl.LabelAssociationKey
        Get
            Return CType(getAssociationKey(), Integer)
        End Get
        Set(ByVal value As Integer)
            setAssociationKey(value)
        End Set
    End Property

    Public Sub setHeight(ByVal value As Integer) Implements CustomControl.setHeight
        Me.Height = value
    End Sub
    Public Sub setWidth(ByVal value As Integer) Implements CustomControl.setWidth
        Me.Width = value
    End Sub
    Public Sub setTop(ByVal value As Integer) Implements CustomControl.setTop
        Me.Top = value
    End Sub
    Public Sub setLeft(ByVal value As Integer) Implements CustomControl.setLeft
        Me.Left = value
    End Sub
End Class


'ComboBox
Public Class CustomComboBox
    Inherits ComboBox
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False
    Private emptyPermitted As Boolean
    Private validatorConstant As ValidatorType = ValidatorType.None
    Private associationKey As Integer = -1

    Private Function getEnterIndex() As Integer
        Return customIndex
    End Function

    Private Sub setEnterIndex(ByVal anIndex As Integer)
        customIndex = anIndex
    End Sub

    Private Function getCleanStatus() As Integer
        Return cleanStatus
    End Function

    Private Sub setCleanStatus(ByVal status As Integer)
        cleanStatus = status
    End Sub

    Private Function getEmptyPermitted() As Boolean
        Return emptyPermitted
    End Function

    Private Sub setEmptyPermitted(ByVal value As Boolean)
        emptyPermitted = value
    End Sub

    Private Function getValidatorType() As Integer
        Return validatorConstant
    End Function

    Private Sub setValidatorType(ByVal type As Integer)
        validatorConstant = type
    End Sub

    Private Function getAssociationKey() As Integer
        Return associationKey
    End Function

    Private Sub setAssociationKey(ByVal key As Integer)
        associationKey = key
    End Sub

    Public Property EnterIndex() As Integer Implements CustomControl.EnterIndex
        Get
            Return CType(getEnterIndex(), Integer)
        End Get
        Set(ByVal value As Integer)
            setEnterIndex(value)
        End Set
    End Property

    Public Property Cleanable() As Boolean Implements CustomControl.Cleanable
        Get
            Return CType(getCleanStatus(), Boolean)
        End Get
        Set(ByVal value As Boolean)
            setCleanStatus(value)
        End Set
    End Property

    Public Property Empty() As Boolean
        Get
            Return CType(getEmptyPermitted(), Boolean)
        End Get
        Set(ByVal value As Boolean)
            setEmptyPermitted(value)
        End Set
    End Property

    Public Property Validator() As ValidatorType
        Get
            Return CType(getValidatorType(), ValidatorType)
        End Get
        Set(ByVal value As ValidatorType)
            setValidatorType(value)
        End Set
    End Property

    Public Property LabelAssociationKey() As Integer Implements CustomControl.LabelAssociationKey
        Get
            Return CType(getAssociationKey(), Integer)
        End Get
        Set(ByVal value As Integer)
            setAssociationKey(value)
        End Set
    End Property

    Public Sub setHeight(ByVal value As Integer) Implements CustomControl.setHeight
        Height = value
    End Sub
    Public Sub setWidth(ByVal value As Integer) Implements CustomControl.setWidth
        Width = value
    End Sub
    Public Sub setTop(ByVal value As Integer) Implements CustomControl.setTop
        Top = value
    End Sub
    Public Sub setLeft(ByVal value As Integer) Implements CustomControl.setLeft
        Left = value
    End Sub
End Class

'ListBox
Public Class CustomListBox
    Inherits ListBox
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False
    Private associationKey As Integer = -1

    Private Function getEnterIndex() As Integer
        Return customIndex
    End Function

    Private Sub setEnterIndex(ByVal anIndex As Integer)
        customIndex = anIndex
    End Sub

    Private Function getCleanStatus() As Integer
        Return cleanStatus
    End Function

    Private Sub setCleanStatus(ByVal status As Integer)
        cleanStatus = status
    End Sub

    Private Function getAssociationKey() As Integer
        Return associationKey
    End Function

    Private Sub setAssociationKey(ByVal key As Integer)
        associationKey = key
    End Sub

    Public Property EnterIndex() As Integer Implements CustomControl.EnterIndex
        Get
            Return CType(getEnterIndex(), Integer)
        End Get
        Set(ByVal value As Integer)
            setEnterIndex(value)
        End Set
    End Property

    Public Property Cleanable() As Boolean Implements CustomControl.Cleanable
        Get
            Return CType(getCleanStatus(), Boolean)
        End Get
        Set(ByVal value As Boolean)
            setCleanStatus(value)
        End Set
    End Property

    Public Property LabelAssociationKey() As Integer Implements CustomControl.LabelAssociationKey
        Get
            Return CType(getAssociationKey(), Integer)
        End Get
        Set(ByVal value As Integer)
            setAssociationKey(value)
        End Set
    End Property

    Public Sub setHeight(ByVal value As Integer) Implements CustomControl.setHeight
        Height = value
    End Sub
    Public Sub setWidth(ByVal value As Integer) Implements CustomControl.setWidth
        Width = value
    End Sub
    Public Sub setTop(ByVal value As Integer) Implements CustomControl.setTop
        Top = value
    End Sub
    Public Sub setLeft(ByVal value As Integer) Implements CustomControl.setLeft
        Left = value
    End Sub
End Class

'Button
Public Class CustomButton
    Inherits Button
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False
    Private associationKey As Integer = -1

    Private Function getEnterIndex() As Integer
        Return customIndex
    End Function

    Private Sub setEnterIndex(ByVal anIndex As Integer)
        customIndex = anIndex
    End Sub

    Private Function getCleanStatus() As Integer
        Return cleanStatus
    End Function

    Private Sub setCleanStatus(ByVal status As Integer)
        cleanStatus = status
    End Sub

    Private Function getAssociationKey() As Integer
        Return associationKey
    End Function

    Private Sub setAssociationKey(ByVal key As Integer)
        associationKey = key
    End Sub

    Public Property EnterIndex() As Integer Implements CustomControl.EnterIndex
        Get
            Return CType(getEnterIndex(), Integer)
        End Get
        Set(ByVal value As Integer)
            setEnterIndex(value)
        End Set
    End Property

    Public Property Cleanable() As Boolean Implements CustomControl.Cleanable
        Get
            Return CType(getCleanStatus(), Boolean)
        End Get
        Set(ByVal value As Boolean)
            setCleanStatus(value)
        End Set
    End Property

    Public Property LabelAssociationKey() As Integer Implements CustomControl.LabelAssociationKey
        Get
            Return CType(getAssociationKey(), Integer)
        End Get
        Set(ByVal value As Integer)
            setAssociationKey(value)
        End Set
    End Property

    Public Sub setHeight(ByVal value As Integer) Implements CustomControl.setHeight
        Height = value
    End Sub
    Public Sub setWidth(ByVal value As Integer) Implements CustomControl.setWidth
        Width = value
    End Sub
    Public Sub setTop(ByVal value As Integer) Implements CustomControl.setTop
        Top = value
    End Sub
    Public Sub setLeft(ByVal value As Integer) Implements CustomControl.setLeft
        Left = value
    End Sub
End Class


'Label
Public Class CustomLabel
    Inherits Label

    Private associationKey As Integer = -1

    Private Function getAssociationKey() As Integer
        Return associationKey
    End Function

    Private Sub setAssociationKey(ByVal key As Integer)
        associationKey = key
    End Sub

    Public Property ControlAssociationKey() As Integer
        Get
            Return CType(getAssociationKey(), Integer)
        End Get
        Set(ByVal value As Integer)
            setAssociationKey(value)
        End Set
    End Property

    Public Sub setTop(ByVal value As Integer)
        Me.Top = value
    End Sub
    Public Sub setLeft(ByVal value As Integer)
        Me.Left = value
    End Sub
End Class