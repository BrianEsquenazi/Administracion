Public Enum ValidatorType
    None = 0
    NotEmpty
    OnlyNumbers
End Enum

Public Interface CustomControl
    Property EnterIndex() As Integer
    Property Cleanable() As Boolean
End Interface

'TextBox
Public Class CustomTextBox
    Inherits TextBox
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False
    Private validatorConstant As ValidatorType = ValidatorType.none

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

    Private Function getValidatorType() As Integer
        Return validatorConstant
    End Function

    Private Sub setValidatorType(ByVal type As Integer)
        validatorConstant = type
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

    Public Property Validator() As ValidatorType
        Get
            Return CType(getValidatorType(), ValidatorType)
        End Get
        Set(ByVal value As ValidatorType)
            setValidatorType(value)
        End Set
    End Property
End Class


'ComboBox
Public Class CustomComboBox
    Inherits ComboBox
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False
    Private validatorConstant As ValidatorType = ValidatorType.None

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

    Private Function getValidatorType() As Integer
        Return validatorConstant
    End Function

    Private Sub setValidatorType(ByVal type As Integer)
        validatorConstant = type
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

    Public Property Validator() As ValidatorType
        Get
            Return CType(getValidatorType(), ValidatorType)
        End Get
        Set(ByVal value As ValidatorType)
            setValidatorType(value)
        End Set
    End Property
End Class

'ListBox
Public Class CustomListBox
    Inherits ListBox
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False

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
End Class

'Button
Public Class CustomButton
    Inherits Button
    Implements CustomControl

    Private customIndex As Integer = -1
    Private cleanStatus As Boolean = False

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
End Class