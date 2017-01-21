Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class DatalookUp
    Protected mMsgflag As String
    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long
    Public Property MsgFlg As String
        Get
            Return mMsgflag
        End Get
        Set(value As String)
            mMsgflag = value
        End Set
    End Property
    Property mdtLkup As DataSet
    Public Property Dtlkup As DataSet
        Get
            Return mdtLkup
        End Get
        Set(value As DataSet)
            mdtLkup = value
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
    Public Sub New(ByVal Connectionname As String, ByVal UserID As Long)
        Try
            mObjectUserID = UserID
            mConnectionName = Connectionname
            db = New DatabaseProviderFactory().Create(Connectionname)
        Catch ex As Exception
            mMsgflag = ex.Message
        End Try

    End Sub
    Public Function getLuData(ByVal str As String) As DataSet
        Try
            If (Not IsNothing(getData(str))) Then
                Return getData(str)
            Else
                mMsgflag = "No data returned"
                Return Nothing
            End If

        Catch ex As Exception
            mMsgflag = ex.Message
            Return Nothing
        End Try
    End Function
    Protected Function getData(ByVal str As String) As DataSet
        Try
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, str)
            If (Not IsNothing(ds) AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            mMsgflag = ex.Message
            Return Nothing
        End Try
    End Function
    Public Enum RequisitionsStatus
        Capturing = 1
        AwaitingHODApproval = 2
        AwaitingFinanceApproval = 3
        AwaitingTreasuryApprocal = 4
        AwaitingExecutiveApproval = 5
        AwaitingDisbursement = 6
        DisbursedbyTransfer = 7
        CashDisbursement = 8
        Cancelledatcapturing = 9
        CancelledatHOD = 10
        CancelledatFinance = 11
        CancelledatTreasuray = 12
        CancelledatExecutive = 13
        RejectedatHOD = 14
        RejectedatFinance = 15
        RejectedatTreasuary = 16
        RejectedatExecutive = 17
        RejectedatDisbursement = 18
        HOLD = 19
        Active = 20
        Inactive = 21
        Deleted = 22

    End Enum
End Class
