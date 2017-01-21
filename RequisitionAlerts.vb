Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class RequisitionAlerts

#Region "Variables"

    Protected mRequestDate As String
    Protected mUpdateDate As String
    Protected mMsgID As Long
    Protected mRequestTypeID As Long
    Protected mRequestStatusID As Long
    Protected mCreatedBy As Long
    Protected mSMSNotificationStatus As Long
    Protected mRequestAmount As Single
    Protected mCreatedOn As String
    Protected mTargetAuthorizationNumber As String
    Protected mRequestOriginator As String
    Protected mRequestDescription As String
    Protected mMsgFlg As String

    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long

#End Region

#Region "Properties"

    Public Property MsgFlag As String
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

    Public Property RequestDate() As String
        Get
            Return mRequestDate
        End Get
        Set(ByVal value As String)
            mRequestDate = value
        End Set
    End Property

    Public Property UpdateDate() As String
        Get
            Return mUpdateDate
        End Get
        Set(ByVal value As String)
            mUpdateDate = value
        End Set
    End Property

    Public Property MsgID() As Long
        Get
            Return mMsgID
        End Get
        Set(ByVal value As Long)
            mMsgID = value
        End Set
    End Property

    Public Property RequestTypeID() As Long
        Get
            Return mRequestTypeID
        End Get
        Set(ByVal value As Long)
            mRequestTypeID = value
        End Set
    End Property

    Public Property RequestStatusID() As Long
        Get
            Return mRequestStatusID
        End Get
        Set(ByVal value As Long)
            mRequestStatusID = value
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

    Public Property SMSNotificationStatus() As Long
        Get
            Return mSMSNotificationStatus
        End Get
        Set(ByVal value As Long)
            mSMSNotificationStatus = value
        End Set
    End Property

    Public Property RequestAmount() As Single
        Get
            Return mRequestAmount
        End Get
        Set(ByVal value As Single)
            mRequestAmount = value
        End Set
    End Property

    Public Property CreatedOn() As String
        Get
            Return mCreatedOn
        End Get
        Set(ByVal value As String)
            mCreatedOn = value
        End Set
    End Property

    Public Property TargetAuthorizationNumber() As String
        Get
            Return mTargetAuthorizationNumber
        End Get
        Set(ByVal value As String)
            mTargetAuthorizationNumber = value
        End Set
    End Property

    Public Property RequestOriginator() As String
        Get
            Return mRequestOriginator
        End Get
        Set(ByVal value As String)
            mRequestOriginator = value
        End Set
    End Property

    Public Property RequestDescription() As String
        Get
            Return mRequestDescription
        End Get
        Set(ByVal value As String)
            mRequestDescription = value
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

        mRequestDate =
        mUpdateDate =
        MsgID = 0
        mRequestTypeID = 0
        mRequestStatusID = 0
        mCreatedBy = mObjectUserID
        mSMSNotificationStatus = 0
        mRequestAmount = 0
        mCreatedOn = ""
        mTargetAuthorizationNumber = ""
        mRequestOriginator = ""
        mRequestDescription = ""

    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mMsgID)

    End Function

    Public Overridable Function Retrieve(ByVal MsgID As Long) As Boolean

        Dim sql As String

        If MsgID > 0 Then
            sql = "SELECT * FROM tbl_RequisitionsAlerts WHERE MsgID = " & MsgID
        Else
            sql = "SELECT * FROM tbl_RequisitionsAlerts WHERE MsgID = " & mMsgID
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

                mMsgFlg = "RequisitionAlerts not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetRequisitionAlerts() As System.Data.DataSet

        Return GetRequisitionAlerts(mMsgID)

    End Function

    Public Overridable Function GetRequisitionAlerts(ByVal MsgID As Long) As DataSet

        Dim sql As String

        If MsgID > 0 Then
            sql = "SELECT * FROM tbl_RequisitionsAlerts WHERE MsgID = " & MsgID
        Else
            sql = "SELECT * FROM tbl_RequisitionsAlerts WHERE MsgID = " & mMsgID
        End If

        Return GetRequisitionAlerts(sql)

    End Function

    Protected Overridable Function GetRequisitionAlerts(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mRequestDate = IIf(IsDBNull(.Item("RequestDate")), Now.Date, .Item("RequestDate"))
            mUpdateDate = IIf(IsDBNull(.Item("UpdateDate")), Now.Date, .Item("UpdateDate"))
            mMsgID = IIf(IsDBNull(.Item("MsgID")), 0, .Item("MsgID"))
            mRequestTypeID = IIf(IsDBNull(.Item("RequestTypeID")), 0, .Item("RequestTypeID"))
            mRequestStatusID = IIf(IsDBNull(.Item("RequestStatusID")), 0, .Item("RequestStatusID"))
            mCreatedBy = IIf(IsDBNull(.Item("CreatedBy")), 0, .Item("CreatedBy"))
            mSMSNotificationStatus = IIf(IsDBNull(.Item("SMSNotificationStatus")), 0, .Item("SMSNotificationStatus"))
            mRequestAmount = IIf(IsDBNull(.Item("RequestAmount")), 0, .Item("RequestAmount"))
            mCreatedOn = IIf(IsDBNull(.Item("CreatedOn")), "", .Item("CreatedOn"))
            mTargetAuthorizationNumber = IIf(IsDBNull(.Item("TargetAuthorizationNumber")), "", .Item("TargetAuthorizationNumber"))
            mRequestOriginator = IIf(IsDBNull(.Item("RequestOriginator")), "", .Item("RequestOriginator"))
            mRequestDescription = IIf(IsDBNull(.Item("RequestDescription")), "", .Item("RequestDescription"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@RequestDate", DbType.Date, mRequestDate)
        db.AddInParameter(cmd, "@UpdateDate", DbType.Date, mUpdateDate)
        db.AddInParameter(cmd, "@MsgID", DbType.Int32, mMsgID)
        db.AddInParameter(cmd, "@RequestTypeID", DbType.Int32, mRequestTypeID)
        db.AddInParameter(cmd, "@RequestStatusID", DbType.Int32, mRequestStatusID)
        db.AddInParameter(cmd, "@SMSNotificationStatus", DbType.Int32, mSMSNotificationStatus)
        db.AddInParameter(cmd, "@RequestAmount", DbType.Currency, mRequestAmount)
        'db.AddInParameter(cmd, "@CreatedOn", DbType.String, mCreatedOn)
        db.AddInParameter(cmd, "@TargetAuthorizationNumber", DbType.String, mTargetAuthorizationNumber)
        db.AddInParameter(cmd, "@RequestOriginator", DbType.String, mRequestOriginator)
        db.AddInParameter(cmd, "@RequestDescription", DbType.String, mRequestDescription)
        db.AddInParameter(cmd, "@CreatedBy", DbType.Int32, mObjectUserID)

    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_RequisitionAlerts")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mMsgID = ds.Tables(0).Rows(0)(0)

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

        'Return Delete("UPDATE tbl_RequisitionsAlerts SET Deleted = 1 WHERE MsgID = " & mMsgID) 
        Return Delete("DELETE FROM tbl_RequisitionsAlerts WHERE MsgID = " & mMsgID)

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