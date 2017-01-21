Imports Microsoft.Practices.EnterpriseLibrary.Data

Public Class AccessLogs


#Region "Variables"

    Protected mLogId As Long
    Protected mCreatedOn As String
    Protected mUsername As String
    Protected mMacAddress As String
    Protected mLANIpAddress As String
    Protected mPublicIpAddress As String
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

    Public Property LogId() As Long
        Get
            Return mLogId
        End Get
        Set(ByVal value As Long)
            mLogId = value
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

    Public Property Username() As String
        Get
            Return mUsername
        End Get
        Set(ByVal value As String)
            mUsername = value
        End Set
    End Property

    Public Property MacAddress() As String
        Get
            Return mMacAddress
        End Get
        Set(ByVal value As String)
            mMacAddress = value
        End Set
    End Property

    Public Property LANIpAddress() As String
        Get
            Return mLANIpAddress
        End Get
        Set(ByVal value As String)
            mLANIpAddress = value
        End Set
    End Property

    Public Property PublicIpAddress() As String
        Get
            Return mPublicIpAddress
        End Get
        Set(ByVal value As String)
            mPublicIpAddress = value
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

        LogId = 0
        mCreatedOn = ""
        mUsername = ""
        mMacAddress = ""
        mLANIpAddress = ""
        mPublicIpAddress = ""
        mMsgFlg = ""
    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mLogId)

    End Function

    Public Overridable Function Retrieve(ByVal LogId As Long) As Boolean

        Dim sql As String

        If LogId > 0 Then
            sql = "SELECT * FROM tblAccessLogs WHERE LogId = " & LogId
        Else
            sql = "SELECT * FROM tblAccessLogs WHERE LogId = " & mLogId
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

                mMsgFlg = "AccessLogs not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetAccessLogs() As System.Data.DataSet

        Return GetAccessLogs(mLogId)

    End Function

    Public Overridable Function GetAccessLogs(ByVal LogId As Long) As DataSet

        Dim sql As String

        If LogId > 0 Then
            sql = "SELECT * FROM tblAccessLogs WHERE LogId = " & LogId
        Else
            sql = "SELECT * FROM tblAccessLogs WHERE LogId = " & mLogId
        End If

        Return GetAccessLogs(sql)

    End Function

    Protected Overridable Function GetAccessLogs(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mLogId = IIf(IsDBNull(.Item("LogId")), 0, .Item("LogId"))
            mCreatedOn = IIf(IsDBNull(.Item("CreatedOn")), "", .Item("CreatedOn"))
            mUsername = IIf(IsDBNull(.Item("Username")), "", .Item("Username"))
            mMacAddress = IIf(IsDBNull(.Item("MacAddress")), "", .Item("MacAddress"))
            mLANIpAddress = IIf(IsDBNull(.Item("LANIpAddress")), "", .Item("LANIpAddress"))
            mPublicIpAddress = IIf(IsDBNull(.Item("PublicIpAddress")), "", .Item("PublicIpAddress"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@LogId", DbType.Int32, mLogId)
        'db.AddInParameter(cmd, "@CreatedOn", DbType.String, mCreatedOn)
        db.AddInParameter(cmd, "@Username", DbType.String, mUsername)
        db.AddInParameter(cmd, "@MacAddress", DbType.String, mMacAddress)
        db.AddInParameter(cmd, "@LANIpAddress", DbType.String, mLANIpAddress)
        db.AddInParameter(cmd, "@PublicIpAddress", DbType.String, mPublicIpAddress)

    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_AccessLogs")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mLogId = ds.Tables(0).Rows(0)(0)

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

        'Return Delete("UPDATE tblAccessLogs SET Deleted = 1 WHERE LogId = " & mLogId) 
        Return Delete("DELETE FROM tblAccessLogs WHERE LogId = " & mLogId)

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