Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class RequestDetails
#Region "Variables"

    Protected mRequestItemID As Long
    Protected mRequestHeaderID As Long
    Protected mStatusID As Long
    Protected mCreatedBy As Long
    Protected mUnitPrice As Single
    Protected mTotalAmount As Single
    Protected mApprovedAmount As Single
    Protected mCreatedOn As String
    Protected mQuantity As Decimal
    Protected mDescription As String
    Protected mMsgFlg As String
    Protected mSupplier As String
    Protected mOrderNo As String
    Protected mBranchName As String

    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long

#End Region

#Region "Properties"
    Public Property BranchName As String
        Get
            Return mBranchName
        End Get
        Set(value As String)
            mBranchName = value
        End Set
    End Property
    Public Property OrderNo As String
        Get
            Return mOrderNo
        End Get
        Set(value As String)
            mOrderNo = value
        End Set
    End Property
    Public Property Supplier As String
        Get
            Return mSupplier
        End Get
        Set(value As String)
            mSupplier = value
        End Set
    End Property
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

    Public Property RequestItemID() As Long
        Get
            Return mRequestItemID
        End Get
        Set(ByVal value As Long)
            mRequestItemID = value
        End Set
    End Property

    Public Property RequestHeaderID() As Long
        Get
            Return mRequestHeaderID
        End Get
        Set(ByVal value As Long)
            mRequestHeaderID = value
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

    Public Property UnitPrice() As Single
        Get
            Return mUnitPrice
        End Get
        Set(ByVal value As Single)
            mUnitPrice = value
        End Set
    End Property

    Public Property TotalAmount() As Single
        Get
            Return mTotalAmount
        End Get
        Set(ByVal value As Single)
            mTotalAmount = value
        End Set
    End Property

    Public Property ApprovedAmount() As Single
        Get
            Return mApprovedAmount
        End Get
        Set(ByVal value As Single)
            mApprovedAmount = value
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

    Public Property Quantity() As Decimal
        Get
            Return mQuantity
        End Get
        Set(ByVal value As Decimal)
            mQuantity = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
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

        RequestItemID = 0
        mRequestHeaderID = 0
        mStatusID = 0
        mCreatedBy = mObjectUserID
        mUnitPrice = 0
        mTotalAmount = 0
        mApprovedAmount = 0
        mCreatedOn = ""
        mQuantity = 0.0
        mDescription = ""
        mMsgFlg = ""
        mSupplier = ""
        mOrderNo = ""
        mBranchName = ""
    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mRequestItemID)

    End Function

    Public Overridable Function Retrieve(ByVal RequestItemID As Long) As Boolean

        Dim sql As String

        If RequestItemID > 0 Then
            sql = "SELECT * FROM tblRequestDetails WHERE RequestItemID = " & RequestItemID
        Else
            sql = "SELECT * FROM tblRequestDetails WHERE RequestItemID = " & mRequestItemID
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

                mMsgFlg = "RequestDetails not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetRequestDetails() As System.Data.DataSet

        Return GetRequestDetails(mRequestItemID)

    End Function

    Public Overridable Function GetRequestDetails(ByVal RequestItemID As Long) As DataSet

        Dim sql As String

        If RequestItemID > 0 Then
            sql = "SELECT * FROM tblRequestDetails WHERE RequestItemID = " & RequestItemID
        Else
            sql = "SELECT * FROM tblRequestDetails WHERE RequestItemID = " & mRequestItemID
        End If

        Return GetRequestDetails(sql)

    End Function

    Protected Overridable Function GetRequestDetails(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mRequestItemID = IIf(IsDBNull(.Item("RequestItemID")), 0, .Item("RequestItemID"))
            mRequestHeaderID = IIf(IsDBNull(.Item("RequestHeaderID")), 0, .Item("RequestHeaderID"))
            mStatusID = IIf(IsDBNull(.Item("StatusID")), 0, .Item("StatusID"))
            mCreatedBy = IIf(IsDBNull(.Item("CreatedBy")), 0, .Item("CreatedBy"))
            mUnitPrice = IIf(IsDBNull(.Item("UnitPrice")), 0, .Item("UnitPrice"))
            mTotalAmount = IIf(IsDBNull(.Item("TotalAmount")), 0, .Item("TotalAmount"))
            mApprovedAmount = IIf(IsDBNull(.Item("ApprovedAmount")), 0, .Item("ApprovedAmount"))
            mCreatedOn = IIf(IsDBNull(.Item("CreatedOn")), "", .Item("CreatedOn"))
            mQuantity = IIf(IsDBNull(.Item("Quantity")), 0.0, .Item("Quantity"))
            mDescription = IIf(IsDBNull(.Item("Description")), "", .Item("Description"))

            mSupplier = IIf(IsDBNull(.Item("Supplier")), "", .Item("Supplier"))
            mOrderNo = IIf(IsDBNull(.Item("OrderNo")), "", .Item("OrderNo"))
            mBranchName = IIf(IsDBNull(.Item("BranchName")), "", .Item("BranchName"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@RequestItemID", DbType.Int32, mRequestItemID)
        db.AddInParameter(cmd, "@RequestHeaderID", DbType.Int32, mRequestHeaderID)
        db.AddInParameter(cmd, "@StatusID", DbType.Int32, mStatusID)
        db.AddInParameter(cmd, "@UnitPrice", DbType.Currency, mUnitPrice)
        db.AddInParameter(cmd, "@TotalAmount", DbType.Currency, mTotalAmount)
        db.AddInParameter(cmd, "@ApprovedAmount", DbType.Currency, mApprovedAmount)
        db.AddInParameter(cmd, "@CreatedOn", DbType.String, mCreatedOn)
        db.AddInParameter(cmd, "@Quantity", DbType.Decimal, mQuantity)
        db.AddInParameter(cmd, "@Description", DbType.String, mDescription)

        db.AddInParameter(cmd, "@Supplier", DbType.String, mSupplier)
        db.AddInParameter(cmd, "@OrderNo", DbType.String, mOrderNo)
        db.AddInParameter(cmd, "@BranchName", DbType.String, mBranchName)

    End Sub


    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_RequestDetails")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mRequestItemID = ds.Tables(0).Rows(0)(0)

            End If

            Return True

        Catch ex As Exception

            mMsgFlg = ex.Message
            Return False

        End Try

    End Function

#End Region

#Region "Delete"

    Public Function UpdateRequestDetailLineStatus(ByVal ItemID As Long, ByVal StatusID As Long, ByVal UpdateType As Long, ByVal ApprovedBy As Long, ByVal Comments As String) As Boolean
        Try
            Dim str As String = ""
            If (UpdateType = 1) Then
                str = "update tblRequestDetails  set StatusID=" & StatusID & ", HODApprovedBy=" & ApprovedBy & ",HODApprovedOn=getDate(), HODCOmment='" & Comments & "' where RequestItemID = " & ItemID & ""
            ElseIf (UpdateType = 2) Then
                str = "update tblRequestDetails  set StatusID=" & StatusID & ", AccountsApprovedBy=" & ApprovedBy & ",AccountsApprovedOn=getDate(), FinanceCOmment='" & Comments & "' where RequestItemID = " & ItemID & ""
            ElseIf (UpdateType = 3) Then
                str = "update tblRequestDetails  set StatusID=" & StatusID & ", ExecutiveApprovedBy=" & ApprovedBy & ",ExecutiveApprovedOn=getDate(), ExecCOmment='" & Comments & "' where RequestItemID = " & ItemID & ""
            End If

            db.ExecuteNonQuery(CommandType.Text, Str)
            Return True
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Function UpdateRequestHeaderStatus(ByVal ItemID As Long, ByVal StatusID As Long) As Boolean
        Try
            Dim str As String = "update tblRequestsHeader  set StatusID=" & StatusID & " where RequestID = " & ItemID & ""
            db.ExecuteNonQuery(CommandType.Text, str)
            Return True
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Function DeleteDetailLine(ByVal RequestDetailLineID As Long) As Boolean
        Try
            Dim str As String = "delete from tblRequestDetails where RequestItemID=" & RequestDetailLineID & ""
            db.ExecuteNonQuery(CommandType.Text, str)
            Return True
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Function ValidateDetailLinesForOveralStatusUpdate(ByVal RequestHeaderID As Long, ByVal OldStatus As Long, ByVal NewStatus As Long) As Boolean
        Try
            Dim str As String = "select * from tblRequestDetails where RequestHeaderID = " & RequestHeaderID & ""
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, str)
            If (Not IsNothing(ds) AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) Then
                Dim str1 As String = "select * from tblRequestDetails where RequestHeaderID = " & RequestHeaderID & " and statusID=" & OldStatus & ""
                Dim ds1 As DataSet = db.ExecuteDataSet(CommandType.Text, str1)
                If (Not IsNothing(ds1) AndAlso ds1.Tables.Count > 0 AndAlso ds1.Tables(0).Rows.Count > 0) Then
                    If (ds1.Tables(0).Rows.Count = ds.Tables(0).Rows.Count) Then
                        Dim str2 As String = "update tbl"
                    End If
                End If
            End If
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Overridable Function Delete() As Boolean

        'Return Delete("UPDATE tblRequestDetails SET Deleted = 1 WHERE RequestItemID = " & mRequestItemID) 
        Return Delete("DELETE FROM tblRequestDetails WHERE RequestItemID = " & mRequestItemID)

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