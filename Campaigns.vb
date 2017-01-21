Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class Campaigns


#Region "Variables"

    Protected mMsgFlg As String
    Protected mMsgID As Long
    Protected mCreatedBy As Long
    Protected mListID As Long
    Protected mStatusID As Long
    Protected mCreatedOn As String
    Protected mmsgBody As String
    Protected mAttachementLocation As String
    Protected mSimulationAddress As String
    Protected mmsgType As String
    Protected mmsgSubject As String
    Protected mMsgFooter As String
    Protected mHtmlTemplate As String
    Protected mMsgFormat As Long

    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long

#End Region

#Region "Properties"

    Public Property MsgFormat As Long
        Get
            Return mMsgFormat
        End Get
        Set(value As Long)
            mMsgFormat = value
        End Set
    End Property
    Public Property HtmlTemplate As String
        Get
            Return mHtmlTemplate
        End Get
        Set(value As String)
            mHtmlTemplate = value
        End Set
    End Property
    Public Property MsgFooter As String
        Get
            Return mMsgFooter
        End Get
        Set(value As String)
            mMsgFooter = value
        End Set
    End Property
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

    Public Property MsgID() As Long
        Get
            Return mMsgID
        End Get
        Set(ByVal value As Long)
            mMsgID = value
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

    Public Property ListID() As Long
        Get
            Return mListID
        End Get
        Set(ByVal value As Long)
            mListID = value
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

    Public Property CreatedOn() As String
        Get
            Return mCreatedOn
        End Get
        Set(ByVal value As String)
            mCreatedOn = value
        End Set
    End Property

    Public Property msgBody() As String
        Get
            Return mmsgBody
        End Get
        Set(ByVal value As String)
            mmsgBody = value
        End Set
    End Property

    Public Property AttachementLocation() As String
        Get
            Return mAttachementLocation
        End Get
        Set(ByVal value As String)
            mAttachementLocation = value
        End Set
    End Property

    Public Property SimulationAddress() As String
        Get
            Return mSimulationAddress
        End Get
        Set(ByVal value As String)
            mSimulationAddress = value
        End Set
    End Property

    Public Property msgType() As String
        Get
            Return mmsgType
        End Get
        Set(ByVal value As String)
            mmsgType = value
        End Set
    End Property

    Public Property msgSubject() As String
        Get
            Return mmsgSubject
        End Get
        Set(ByVal value As String)
            mmsgSubject = value
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

        MsgID = 0
        mCreatedBy = mObjectUserID
        mListID = 0
        mStatusID = 0
        mCreatedOn = ""
        mmsgBody = ""
        mAttachementLocation = ""
        mSimulationAddress = ""
        mmsgType = ""
        mmsgSubject = ""
        mMsgFooter = 0
        mHtmlTemplate = ""
        mMsgFormat = 0

    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mMsgID)

    End Function

    Public Overridable Function Retrieve(ByVal MsgID As Long) As Boolean

        Dim sql As String

        If MsgID > 0 Then
            sql = "SELECT * FROM tbl_Campaigns WHERE MsgID = " & MsgID
        Else
            sql = "SELECT * FROM tbl_Campaigns WHERE MsgID = " & mMsgID
        End If

        Return Retrieve(sql)

    End Function

    Public Function getBuinessAffiliateContacts(ByVal MailingListID As Long) As DataSet
        Try
            Dim sql As String = "select * from tbl_BusinessAffiliatesContacts where BusinessAffiliateID in (select BusinessAffiliateID from tbl_MailingLists where MailingListID = " & MailingListID & ")"
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, sql)
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

    Protected Overridable Function Retrieve(ByVal sql As String) As Boolean

        Try

            Dim dsRetrieve As DataSet = db.ExecuteDataSet(CommandType.Text, sql)

            If dsRetrieve IsNot Nothing AndAlso dsRetrieve.Tables.Count > 0 AndAlso dsRetrieve.Tables(0).Rows.Count > 0 Then

                LoadDataRecord(dsRetrieve.Tables(0).Rows(0))

                dsRetrieve = Nothing
                Return True

            Else
                mMsgFlg = "Campaigns not found."


                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetCampaigns() As System.Data.DataSet

        Return GetCampaigns(mMsgID)

    End Function
    Public Function getCampaignDetails(ByVal ListID As Integer) As DataSet
        Try
            Dim sql As String = "select Ml.MailingListName, Ml.BusinessAffiliateID, Ml.CreatedDate, U.EmailAddress , U.firstname  from tbl_MailingLists Ml inner join tbl_Users U on U.userid = Ml.CreatedBy where Ml.MailingListID = " & ListID & ""
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, sql)
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
    Public Overridable Function GetCampaigns(ByVal MsgID As Long) As DataSet

        Dim sql As String

        If MsgID > 0 Then
            sql = "SELECT * FROM tbl_Campaigns WHERE MsgID = " & MsgID
        Else
            sql = "SELECT * FROM tbl_Campaigns WHERE MsgID = " & mMsgID
        End If

        Return GetCampaigns(sql)

    End Function

    Protected Overridable Function GetCampaigns(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mMsgID = IIf(IsDBNull(.Item("MsgID")), 0, .Item("MsgID"))
            mCreatedBy = IIf(IsDBNull(.Item("CreatedBy")), 0, .Item("CreatedBy"))
            mListID = IIf(IsDBNull(.Item("ListID")), 0, .Item("ListID"))
            mStatusID = IIf(IsDBNull(.Item("StatusID")), 0, .Item("StatusID"))
            mCreatedOn = IIf(IsDBNull(.Item("CreatedOn")), "", .Item("CreatedOn"))
            mmsgBody = IIf(IsDBNull(.Item("msgBody")), "", .Item("msgBody"))
            mAttachementLocation = IIf(IsDBNull(.Item("AttachementLocation")), "", .Item("AttachementLocation"))
            mSimulationAddress = IIf(IsDBNull(.Item("SimulationAddress")), "", .Item("SimulationAddress"))
            mmsgType = IIf(IsDBNull(.Item("msgType")), "", .Item("msgType"))
            mmsgSubject = IIf(IsDBNull(.Item("msgSubject")), "", .Item("msgSubject"))
            mMsgFooter = IIf(IsDBNull(.Item("msgFooter")), "", .Item("msgFooter"))
            mHtmlTemplate = IIf(IsDBNull(.Item("HtmlTemplate")), "", .Item("HtmlTemplate"))
            mMsgFormat = IIf(IsDBNull(.Item("msgFormat")), "", .Item("msgFormat"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@MsgID", DbType.Int32, mMsgID)
        db.AddInParameter(cmd, "@ListID", DbType.Int32, mListID)
        db.AddInParameter(cmd, "@StatusID", DbType.Int32, mStatusID)
        db.AddInParameter(cmd, "@msgBody", DbType.String, mmsgBody)
        db.AddInParameter(cmd, "@AttachementLocation", DbType.String, mAttachementLocation)
        db.AddInParameter(cmd, "@SimulationAddress", DbType.String, mSimulationAddress)
        db.AddInParameter(cmd, "@msgType", DbType.String, mmsgType)
        db.AddInParameter(cmd, "@msgSubject", DbType.String, mmsgSubject)
        db.AddInParameter(cmd, "@CreatedBy", DbType.Int32, mCreatedBy)
        db.AddInParameter(cmd, "@msgFooter", DbType.String, mMsgFooter)
        db.AddInParameter(cmd, "@HtmlTemplate", DbType.String, mHtmlTemplate)
        db.AddInParameter(cmd, "@CreatedOn", DbType.DateTime, Date.Now)
        db.AddInParameter(cmd, "@MsgFormat", DbType.Int32, mMsgFormat)

    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_Campaigns")

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

        'Return Delete("UPDATE tbl_Campaigns SET Deleted = 1 WHERE MsgID = " & mMsgID) 
        Return Delete("DELETE FROM tbl_Campaigns WHERE MsgID = " & mMsgID)

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


#Region "Authorization"
    '' Getting Authorization Lists Ready for sending out
    Public Function getPendingLists() As DataSet
        Try
            'Dim sql As String = "select distinct count(Ds.MsgAddress) as CampaignRecipient, C.listID,ML.MailingListName, C.CreatedOn , C.AttachementLocation as Attachment,C.SimulationAddress ,C.msgType ,C.msgSubject ,C.msgBody,C.msgFooter,C.HtmlTemplate  from tbl_Campaigns C inner join tbl_Users U on U.userid = C.CreatedBy inner join tbl_MailingLists ML on ML.MailingListID  = C.ListID left join tbl_MsgDistributionsSchedule DS on Ds.DistributionListID = C.ListID where C.StatusID=1 group by ML.MailingListName, C.CreatedOn , C.AttachementLocation,C.SimulationAddress ,C.msgType ,C.msgSubject ,C.msgBody,C.msgFooter,C.HtmlTemplate, C.ListID"
            Dim sql As String = "select distinct count(c.SimulationAddress) as CampaignRecipient, L.MailingListID as listID, L.MailingListName as MailingListName, U.EmailAddress as SimulationAddress, l.CreatedDate as CreatedOn , C.msgType,c.msgBody ,c.msgSubject ,c.msgBody ,c.msgFooter ,c.HtmlTemplate,c.AttachementLocation as Attachment   from tbl_MailingLists L inner join tbl_Campaigns c on c.ListID = l.MailingListID inner join tbl_Users U on U.userid = L.CreatedBy where c.StatusID = 3 group by L.MailingListID , L.MailingListName ,U.EmailAddress, l.CreatedDate ,C.msgType, c.msgBody ,c.msgSubject ,c.msgBody ,c.msgFooter ,c.HtmlTemplate,c.AttachementLocation"
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, sql)
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

#End Region


End Class