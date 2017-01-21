Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class RequisitionsHeader
#Region "Variables"

    Protected mRequestDueDate As String
    Protected mRequestID As Long
    Protected mOrderTypeID As Long
    Protected mBranchID As Long
    Protected mRequestBy As Long
    Protected mStatusID As Long
    Protected mTotalRequestAmount As Single
    Protected mApprovedAmount As Single
    Protected mCreatedDate As String
    Protected mRequestDate As String
    Protected mOrderNumber As String
    Protected mDescription As String
    Protected mApprovedSupplier As String
    Protected mMsgFlg As String
    Protected mCompanyID As Long

    Protected mDisbursementType As String
    Protected mTransferType As String
    Protected mBank As String
    Protected mBranch As String
    Protected mAccount As String

    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long

#End Region

#Region "Properties"
    Public Property CompanyID As Long
        Get
            Return mCompanyID
        End Get
        Set(value As Long)
            mCompanyID = value
        End Set
    End Property
    Public Property DisbursementType As String
        Get
            Return mDisbursementType
        End Get
        Set(value As String)
            mDisbursementType = value
        End Set
    End Property
    Public Property TransferType As String
        Get
            Return mTransferType
        End Get
        Set(value As String)
            mTransferType = value
        End Set
    End Property
    Public Property Bank As String
        Get
            Return mBank
        End Get
        Set(value As String)
            mBank = value
        End Set
    End Property
    Public Property Branch As String
        Get
            Return mBranch
        End Get
        Set(value As String)
            mBranch = value
        End Set
    End Property
    Public Property Account As String
        Get
            Return mAccount
        End Get
        Set(value As String)
            mAccount = value
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

    Public Property RequestDueDate() As String
        Get
            Return mRequestDueDate
        End Get
        Set(ByVal value As String)
            mRequestDueDate = value
        End Set
    End Property

    Public Property RequestID() As Long
        Get
            Return mRequestID
        End Get
        Set(ByVal value As Long)
            mRequestID = value
        End Set
    End Property

    Public Property OrderTypeID() As Long
        Get
            Return mOrderTypeID
        End Get
        Set(ByVal value As Long)
            mOrderTypeID = value
        End Set
    End Property

    Public Property BranchID() As Long
        Get
            Return mBranchID
        End Get
        Set(ByVal value As Long)
            mBranchID = value
        End Set
    End Property

    Public Property RequestBy() As Long
        Get
            Return mRequestBy
        End Get
        Set(ByVal value As Long)
            mRequestBy = value
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

    Public Property TotalRequestAmount() As Single
        Get
            Return mTotalRequestAmount
        End Get
        Set(ByVal value As Single)
            mTotalRequestAmount = value
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

    Public Property CreatedDate() As String
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As String)
            mCreatedDate = value
        End Set
    End Property

    Public Property RequestDate() As String
        Get
            Return mRequestDate
        End Get
        Set(ByVal value As String)
            mRequestDate = value
        End Set
    End Property

    Public Property OrderNumber() As String
        Get
            Return mOrderNumber
        End Get
        Set(ByVal value As String)
            mOrderNumber = value
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

    Public Property ApprovedSupplier() As String
        Get
            Return mApprovedSupplier
        End Get
        Set(ByVal value As String)
            mApprovedSupplier = value
        End Set
    End Property

#End Region

#Region "Methods"

#Region "Constructors"

    Public Sub New(ByVal ConnectionName As String, ByVal ObjectUserID As Long)

        mObjectUserID = ObjectUserID
        mConnectionName = ConnectionName
        'db = DatabaseFactory.CreateDatabase(ConnectionName)
        db = New DatabaseProviderFactory().Create(ConnectionName)

    End Sub

#End Region

    Public Sub Clear()

        mRequestDueDate =
        RequestID = 0
        mOrderTypeID = 0
        mBranchID = 0
        mRequestBy = 0
        mStatusID = 0
        mTotalRequestAmount = 0
        mApprovedAmount = 0
        mCreatedDate = ""
        mRequestDate = ""
        mOrderNumber = ""
        mDescription = ""
        mApprovedSupplier = ""
        mMsgFlg = ""
        mCompanyID = 0

        mDisbursementType = ""
        mTransferType = ""
        mBank = ""
        mBranch = ""
        mAccount = ""

    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mRequestID)

    End Function

    Public Overridable Function Retrieve(ByVal RequestID As Long) As Boolean

        Dim sql As String

        If RequestID > 0 Then
            sql = "SELECT * FROM tblRequestsHeader WHERE RequestID = " & RequestID
        Else
            sql = "SELECT * FROM tblRequestsHeader WHERE RequestID = " & mRequestID
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

                mMsgFlg = "RequisitionsHeader not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetRequisitionsHeader() As System.Data.DataSet

        Return GetRequisitionsHeader(mRequestID)

    End Function

    Public Overridable Function GetRequisitionsHeader(ByVal RequestID As Long) As DataSet

        Dim sql As String

        If RequestID > 0 Then
            sql = "SELECT * FROM tblRequestsHeader WHERE RequestID = " & RequestID
        Else
            sql = "SELECT * FROM tblRequestsHeader WHERE RequestID = " & mRequestID
        End If

        Return GetRequisitionsHeader(sql)

    End Function
    'Public Function GetAllPendingRequisitions(ByVal )
    Public Function GetOrderDetails(ByVal HeaderID As Long) As DataSet
        Try
            Dim str As String = "select RequestItemID , RequestHeaderID ,Rd.Description,Rd.BranchName ,Quantity,cast(UnitPrice as numeric(10,2)) as UnitPrice,cast(TotalAmount as numeric(10,2)) as TotalAmount,Rd.ApprovedAmount,Rs.Description as Status,Rd.OrderNo,Rd.Supplier,Rd.HODComment from tblRequestDetails Rd inner join luRequestsStatus Rs on Rs.RequestStatusID = Rd.StatusID where Rd.RequestHeaderID=" & HeaderID & " and Rd.StatusID in (" & DatalookUp.RequisitionsStatus.Capturing & "," & DatalookUp.RequisitionsStatus.RejectedatHOD & ")"
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, str)
            If (Not IsNothing(ds) AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return Nothing
        End Try
    End Function
    Public Function GetOrderNumber() As String
        Try
            Dim OrderNumber As String = ""
            Dim str As String = "select * from tblRequestsHeader"
            Dim strCompanyCode As String = "select * from tblCompanyGroups where CompanyID = (select CompanyID from tblUsers where UserID =" & mObjectUserID & ")"
            Dim dsi As DataSet = db.ExecuteDataSet(CommandType.Text, strCompanyCode)
            Dim CompanyCode As String = ""
            If (Not IsNothing(dsi) AndAlso dsi.Tables.Count > 0 AndAlso dsi.Tables(0).Rows.Count > 0) Then
                CompanyCode = dsi.Tables(0).Rows(0)(4).ToString()
            Else
                CompanyCode = "REQ"
            End If
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, str)
            If (Not IsNothing(ds) AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) Then
                If (ds.Tables(0).Rows.Count <= 9) Then
                    OrderNumber = CompanyCode & "00000" & ds.Tables(0).Rows.Count + 1
                ElseIf (ds.Tables(0).Rows.Count >= 10 AndAlso ds.Tables(0).Rows.Count <= 99) Then
                    OrderNumber = CompanyCode & "0000" & ds.Tables(0).Rows.Count + 1
                ElseIf (ds.Tables(0).Rows.Count >= 100 AndAlso ds.Tables(0).Rows.Count <= 999) Then
                    OrderNumber = CompanyCode & "000" & ds.Tables(0).Rows.Count + 1
                ElseIf (ds.Tables(0).Rows.Count >= 1000 AndAlso ds.Tables(0).Rows.Count <= 9999) Then
                    OrderNumber = CompanyCode & "00" & ds.Tables(0).Rows.Count + 1
                ElseIf (ds.Tables(0).Rows.Count >= 10000 AndAlso ds.Tables(0).Rows.Count <= 99999) Then
                    OrderNumber = CompanyCode & "0" & ds.Tables(0).Rows.Count + 1
                Else
                    OrderNumber = CompanyCode & ds.Tables(0).Rows.Count + 1
                End If
            Else
                OrderNumber = CompanyCode & "00000" & ds.Tables(0).Rows.Count + 1
            End If
            Return OrderNumber
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return ""
        End Try
    End Function

    Protected Overridable Function GetRequisitionsHeader(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mRequestDueDate = IIf(IsDBNull(.Item("RequestDueDate")), Now.Date, .Item("RequestDueDate"))
            mRequestID = IIf(IsDBNull(.Item("RequestID")), 0, .Item("RequestID"))
            mOrderTypeID = IIf(IsDBNull(.Item("OrderTypeID")), 0, .Item("OrderTypeID"))
            mBranchID = IIf(IsDBNull(.Item("BranchID")), 0, .Item("BranchID"))
            mRequestBy = IIf(IsDBNull(.Item("RequestBy")), 0, .Item("RequestBy"))
            mStatusID = IIf(IsDBNull(.Item("StatusID")), 0, .Item("StatusID"))
            mTotalRequestAmount = IIf(IsDBNull(.Item("TotalRequestAmount")), 0, .Item("TotalRequestAmount"))
            mApprovedAmount = IIf(IsDBNull(.Item("ApprovedAmount")), 0, .Item("ApprovedAmount"))
            mCreatedDate = IIf(IsDBNull(.Item("CreatedDate")), "", .Item("CreatedDate"))
            mRequestDate = IIf(IsDBNull(.Item("RequestDate")), "", .Item("RequestDate"))
            mOrderNumber = IIf(IsDBNull(.Item("OrderNumber")), "", .Item("OrderNumber"))
            mDescription = IIf(IsDBNull(.Item("Description")), "", .Item("Description"))
            mApprovedSupplier = IIf(IsDBNull(.Item("ApprovedSupplier")), "", .Item("ApprovedSupplier"))

            mDisbursementType = IIf(IsDBNull(.Item("DisbursementType")), "", .Item("DisbursementType"))
            mTransferType = IIf(IsDBNull(.Item("TransferType")), "", .Item("TransferType"))
            mBank = IIf(IsDBNull(.Item("Bank")), "", .Item("Bank"))
            mBranch = IIf(IsDBNull(.Item("Branch")), "", .Item("Branch"))
            mAccount = IIf(IsDBNull(.Item("Account")), "", .Item("Account"))
            mCompanyID = IIf(IsDBNull(.Item("CompanyID")), 0, .Item("CompanyID"))


        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        'db.AddInParameter(cmd, "@RequestDueDate", DbType.Date, mRequestDueDate)
        db.AddInParameter(cmd, "@RequestID", DbType.Int32, mRequestID)
        db.AddInParameter(cmd, "@OrderTypeID", DbType.Int32, mOrderTypeID)
        db.AddInParameter(cmd, "@BranchID", DbType.Int32, mBranchID)
        db.AddInParameter(cmd, "@RequestBy", DbType.Int32, mRequestBy)
        db.AddInParameter(cmd, "@StatusID", DbType.Int32, mStatusID)
        db.AddInParameter(cmd, "@TotalRequestAmount", DbType.Currency, mTotalRequestAmount)
        db.AddInParameter(cmd, "@ApprovedAmount", DbType.Currency, mApprovedAmount)
        'db.AddInParameter(cmd, "@RequestDate", DbType.String, mRequestDate)
        db.AddInParameter(cmd, "@OrderNumber", DbType.String, mOrderNumber)
        db.AddInParameter(cmd, "@Description", DbType.String, mDescription)
        db.AddInParameter(cmd, "@ApprovedSupplier", DbType.String, mApprovedSupplier)

        db.AddInParameter(cmd, "@DisbursementType", DbType.String, mDisbursementType)
        db.AddInParameter(cmd, "@TransferType", DbType.String, mTransferType)
        db.AddInParameter(cmd, "@Bank", DbType.String, mBank)
        db.AddInParameter(cmd, "@Branch", DbType.String, mBranch)
        db.AddInParameter(cmd, "@Account", DbType.String, mAccount)

        db.AddInParameter(cmd, "@CompanyID", DbType.Int32, mCompanyID)


    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_RequisitionsHeader")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mRequestID = ds.Tables(0).Rows(0)(0)

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

        'Return Delete("UPDATE tblRequestsHeader SET Deleted = 1 WHERE RequestID = " & mRequestID) 
        Return Delete("DELETE FROM tblRequestsHeader WHERE RequestID = " & mRequestID)

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