Imports Microsoft.Practices.EnterpriseLibrary.Data
Public Class BranchAccountsMapping

#Region "Variables"

    Protected mBranchAccountMapID As Long
    Protected mBranchID As Long
    Protected mAccountID As Long
    Protected mCreatedBy As Long
    Protected mStatusID As Long
    Protected mCreatedOn As String
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

    Public Property BranchAccountMapID() As Long
        Get
            Return mBranchAccountMapID
        End Get
        Set(ByVal value As Long)
            mBranchAccountMapID = value
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

    Public Property AccountID() As Long
        Get
            Return mAccountID
        End Get
        Set(ByVal value As Long)
            mAccountID = value
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

        BranchAccountMapID = 0
        mBranchID = 0
        mAccountID = 0
        mCreatedBy = mObjectUserID
        mStatusID = 0
        mCreatedOn = ""
        mMsgFlg = ""
    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mBranchAccountMapID)

    End Function

    Public Overridable Function Retrieve(ByVal BranchAccountMapID As Long) As Boolean

        Dim sql As String

        If BranchAccountMapID > 0 Then
            sql = "SELECT * FROM tblBranchAccountsMapping WHERE BranchAccountMapID = " & BranchAccountMapID
        Else
            sql = "SELECT * FROM tblBranchAccountsMapping WHERE BranchAccountMapID = " & mBranchAccountMapID
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

                mMsgFlg = "BranchAccountsMapping not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetBranchAccountsMapping() As System.Data.DataSet

        Return GetBranchAccountsMapping(mBranchAccountMapID)

    End Function

    Public Overridable Function GetBranchAccountsMapping(ByVal BranchAccountMapID As Long) As DataSet

        Dim sql As String

        If BranchAccountMapID > 0 Then
            sql = "SELECT * FROM tblBranchAccountsMapping WHERE BranchAccountMapID = " & BranchAccountMapID
        Else
            sql = "SELECT * FROM tblBranchAccountsMapping WHERE BranchAccountMapID = " & mBranchAccountMapID
        End If

        Return GetBranchAccountsMapping(sql)

    End Function

    Protected Overridable Function GetBranchAccountsMapping(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mBranchAccountMapID = IIf(IsDBNull(.Item("BranchAccountMapID")), 0, .Item("BranchAccountMapID"))
            mBranchID = IIf(IsDBNull(.Item("BranchID")), 0, .Item("BranchID"))
            mAccountID = IIf(IsDBNull(.Item("AccountID")), 0, .Item("AccountID"))
            mCreatedBy = IIf(IsDBNull(.Item("CreatedBy")), 0, .Item("CreatedBy"))
            mStatusID = IIf(IsDBNull(.Item("StatusID")), 0, .Item("StatusID"))
            mCreatedOn = IIf(IsDBNull(.Item("CreatedOn")), "", .Item("CreatedOn"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@BranchAccountMapID", DbType.Int32, mBranchAccountMapID)
        db.AddInParameter(cmd, "@BranchID", DbType.Int32, mBranchID)
        db.AddInParameter(cmd, "@AccountID", DbType.Int32, mAccountID)
        db.AddInParameter(cmd, "@StatusID", DbType.Int32, mStatusID)
        db.AddInParameter(cmd, "@CreatedOn", DbType.String, mCreatedOn)
        db.AddInParameter(cmd, "@CreatedBy", DbType.Int32, mObjectUserID)

    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_BranchAccountsMapping")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mBranchAccountMapID = ds.Tables(0).Rows(0)(0)

            End If

            Return True

        Catch ex As Exception

            mMsgFlg = ex.Message
            Return False

        End Try

    End Function

#End Region

    Public Function ValidateAccountBeforeSave(ByVal BranchID As Long, ByVal AccountID As Long) As Boolean
        Try
            Dim str As String = "select * from tblBranchAccountsMapping where BranchID =" & BranchID & " and AccountID = " & AccountID & ""
            Dim obj As New DatalookUp(ConnectionName, mObjectUserID)
            With obj
                If (.getLuData(str) IsNot Nothing) Then
                    Return True
                Else
                    Return False
                End If
            End With
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
#Region "Delete"

    Public Overridable Function Delete(ByVal BranchAccountMapID As Long) As Boolean

        'Return Delete("UPDATE tblBranchAccountsMapping SET Deleted = 1 WHERE BranchAccountMapID = " & mBranchAccountMapID) 
        Return Delete("DELETE FROM tblBranchAccountsMapping WHERE BranchAccountMapID = " & BranchAccountMapID)

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
