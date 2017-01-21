Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class BudgetDetails
#Region "Variables"

    Protected mStartDate As String
    Protected mEndDate As String
    Protected mBudgetID As Long
    Protected mStatusID As Long
    Protected mCreatedBy As Long
    Protected mUpdatedBy As Long
    Protected mBudgetAmount As Single
    Protected mBudgetOvershoot As Single
    Protected mCreatedDate As String
    Protected mUpdatedOn As String
    Protected mBudgetName As String
    Protected mMsgFlg As String

    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long

#End Region

#Region "Properties"
    Public Property MsgFlg As String
        Get
            Return mMsgFlg
        End Get
        Set(value As String)
            mMsgFlg = value
        End Set
    End Property
    Public ReadOnly Property Database() As Database
        Get
            Return db
        End Get
    End Property

    Public ReadOnly Property OwnerType() As String
        Get
            Return Me.GetType.Name
        End Get
    End Property

    Public ReadOnly Property ConnectionName() As String
        Get
            Return mConnectionName
        End Get
    End Property

    Public Property StartDate() As String
        Get
            Return mStartDate
        End Get
        Set(ByVal value As String)
            mStartDate = value
        End Set
    End Property

    Public Property EndDate() As String
        Get
            Return mEndDate
        End Get
        Set(ByVal value As String)
            mEndDate = value
        End Set
    End Property

    Public Property BudgetID() As Long
        Get
            Return mBudgetID
        End Get
        Set(ByVal value As Long)
            mBudgetID = value
        End Set
    End Property

    Public Property StatusID() As Long
        Get
            Return mStatusID
        End Get
        Set(ByVal value As Long)
            mStatusID = value
        End Set
    End Property

    Public Property CreatedBy() As Long
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As Long)
            mCreatedBy = value
        End Set
    End Property

    Public Property UpdatedBy() As Long
        Get
            Return mUpdatedBy
        End Get
        Set(ByVal value As Long)
            mUpdatedBy = value
        End Set
    End Property

    Public Property BudgetAmount() As Single
        Get
            Return mBudgetAmount
        End Get
        Set(ByVal value As Single)
            mBudgetAmount = value
        End Set
    End Property

    Public Property BudgetOvershoot() As Single
        Get
            Return mBudgetOvershoot
        End Get
        Set(ByVal value As Single)
            mBudgetOvershoot = value
        End Set
    End Property

    Public Property CreatedDate() As String
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As String)
            mCreatedDate = value
        End Set
    End Property

    Public Property UpdatedOn() As String
        Get
            Return mUpdatedOn
        End Get
        Set(ByVal value As String)
            mUpdatedOn = value
        End Set
    End Property

    Public Property BudgetName() As String
        Get
            Return mBudgetName
        End Get
        Set(ByVal value As String)
            mBudgetName = value
        End Set
    End Property

#End Region

#Region "Methods"

#Region "Constructors"

    Public Sub New(ByVal ConnectionName As String, ByVal ObjectUserID As Long)

        mObjectUserID = ObjectUserID
        mConnectionName = ConnectionName
        db = New DatabaseProviderFactory().Create(ConnectionName)

    End Sub

#End Region

    Public Sub Clear()

        mStartDate =
        mEndDate =
        BudgetID = 0
        mStatusID = 0
        mCreatedBy = mObjectUserID
        mUpdatedBy = 0
        mBudgetAmount = 0
        mBudgetOvershoot = 0
        mCreatedDate = ""
        mUpdatedOn = ""
        mBudgetName = ""
        mMsgFlg = ""

    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mBudgetID)

    End Function

    Public Overridable Function Retrieve(ByVal BudgetID As Long) As Boolean

        Dim sql As String

        If BudgetID > 0 Then
            sql = "SELECT * FROM tblBudgetDetails WHERE BudgetID = " & BudgetID
        Else
            sql = "SELECT * FROM tblBudgetDetails WHERE BudgetID = " & mBudgetID
        End If

        Return Retrieve(sql)

    End Function

    Protected Overridable Function Retrieve(ByVal sql As String) As Boolean

        Try

            Dim dsRetrieve As DataSet = db.ExecuteDataSet(CommandType.Text, sql)

            If dsRetrieve IsNot Nothing AndAlso dsRetrieve.Tables.Count > 0 AndAlso dsRetrieve.Tables(0).Rows.Count > 0 Then

                LoadDataRecord(dsRetrieve.Tables(0).Rows(0))

                dsRetrieve = Nothing
                Return True

            Else

                mMsgFlg = "BudgetDetails not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetBudgetDetails() As System.Data.DataSet

        Return GetBudgetDetails(mBudgetID)

    End Function

    Public Overridable Function GetBudgetDetails(ByVal BudgetID As Long) As DataSet

        Dim sql As String

        If BudgetID > 0 Then
            sql = "SELECT * FROM tblBudgetDetails WHERE BudgetID = " & BudgetID
        Else
            sql = "SELECT * FROM tblBudgetDetails WHERE BudgetID = " & mBudgetID
        End If

        Return GetBudgetDetails(sql)

    End Function

    Protected Overridable Function GetBudgetDetails(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function
    Public Function UpdateBudgetDetails(ByVal str As String) As Boolean
        Try
            db.ExecuteNonQuery(CommandType.Text, str)
            Return True
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mStartDate = IIf(IsDBNull(.Item("StartDate")), Now.Date, .Item("StartDate"))
            mEndDate = IIf(IsDBNull(.Item("EndDate")), Now.Date, .Item("EndDate"))
            mBudgetID = IIf(IsDBNull(.Item("BudgetID")), 0, .Item("BudgetID"))
            mStatusID = IIf(IsDBNull(.Item("StatusID")), 0, .Item("StatusID"))
            mCreatedBy = IIf(IsDBNull(.Item("CreatedBy")), 0, .Item("CreatedBy"))
            mUpdatedBy = IIf(IsDBNull(.Item("UpdatedBy")), 0, .Item("UpdatedBy"))
            mBudgetAmount = IIf(IsDBNull(.Item("BudgetAmount")), 0, .Item("BudgetAmount"))
            mBudgetOvershoot = IIf(IsDBNull(.Item("BudgetOvershoot")), 0, .Item("BudgetOvershoot"))
            mCreatedDate = IIf(IsDBNull(.Item("CreatedDate")), "", .Item("CreatedDate"))
            mUpdatedOn = IIf(IsDBNull(.Item("UpdatedOn")), "", .Item("UpdatedOn"))
            mBudgetName = IIf(IsDBNull(.Item("BudgetName")), "", .Item("BudgetName"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@StartDate", DbType.Date, mStartDate)
        db.AddInParameter(cmd, "@EndDate", DbType.Date, mEndDate)
        db.AddInParameter(cmd, "@BudgetID", DbType.Int32, mBudgetID)
        db.AddInParameter(cmd, "@StatusID", DbType.Int32, mStatusID)
        db.AddInParameter(cmd, "@UpdatedBy", DbType.Int32, mObjectUserID)
        db.AddInParameter(cmd, "@BudgetAmount", DbType.Currency, mBudgetAmount)
        db.AddInParameter(cmd, "@BudgetOvershoot", DbType.Currency, mBudgetOvershoot)
        db.AddInParameter(cmd, "@UpdatedOn", DbType.String, mUpdatedOn)
        db.AddInParameter(cmd, "@BudgetName", DbType.String, mBudgetName)
        db.AddInParameter(cmd, "@CreatedBy", DbType.Int32, mObjectUserID)

    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_BudgetDetails")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mBudgetID = ds.Tables(0).Rows(0)(0)

            End If

            Return True

        Catch ex As Exception

            mMsgFlg = ex.Message
            Return False

        End Try

    End Function

#End Region

#Region "Delete"

    Public Overridable Function Delete() As Boolean

        'Return Delete("UPDATE tblBudgetDetails SET Deleted = 1 WHERE BudgetID = " & mBudgetID) 
        Return Delete("DELETE FROM tblBudgetDetails WHERE BudgetID = " & mBudgetID)

    End Function

    Protected Overridable Function Delete(ByVal DeleteSQL As String) As Boolean

        Try

            db.ExecuteNonQuery(CommandType.Text, DeleteSQL)
            Return True

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

#End Region

#End Region

End Class