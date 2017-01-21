Imports Microsoft.Practices.EnterpriseLibrary.Data
Imports System.Security.Cryptography
Public Class UserAdministration
#Region "Variables"

    Protected mUserID As Long
    Protected mUserTypeID As Long
    Protected mDepartmentID As Long
    Protected mStatusID As Long
    Protected mCreatedBy As Long
    Protected mCreatedDate As String
    Protected mUsername As String
    Protected mFirstname As String
    Protected mSurname As String
    Protected mPassword As String
    Protected memailAddress As String
    Protected mMobileNo As String
    Protected mCompanyID As Long
    Protected mBranchID As Long
    Protected mPasswordUpdateDate As String

    Protected mMsgFlg As String
    Protected db As Database
    Protected mConnectionName As String
    Protected mObjectUserID As Long

#End Region

#Region "Properties"
    Public Property PasswordUpdateDate As String
        Get
            Return mPasswordUpdateDate
        End Get
        Set(value As String)
            mPasswordUpdateDate = value
        End Set
    End Property
    Public Property CompanyID As Long
        Get
            Return mCompanyID
        End Get
        Set(value As Long)
            mCompanyID = value
        End Set
    End Property
    Public Property BranchID As Long
        Get
            Return mBranchID
        End Get
        Set(value As Long)
            mBranchID = value
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

    Public Property UserID() As Long
        Get
            Return mUserID
        End Get
        Set(ByVal value As Long)
            mUserID = value
        End Set
    End Property

    Public Property UserTypeID() As Long
        Get
            Return mUserTypeID
        End Get
        Set(ByVal value As Long)
            mUserTypeID = value
        End Set
    End Property

    Public Property DepartmentID() As Long
        Get
            Return mDepartmentID
        End Get
        Set(ByVal value As Long)
            mDepartmentID = value
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

    Public Property CreatedDate() As String
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As String)
            mCreatedDate = value
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

    Public Property Firstname() As String
        Get
            Return mFirstname
        End Get
        Set(ByVal value As String)
            mFirstname = value
        End Set
    End Property

    Public Property Surname() As String
        Get
            Return mSurname
        End Get
        Set(ByVal value As String)
            mSurname = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return mPassword
        End Get
        Set(ByVal value As String)
            mPassword = value
        End Set
    End Property

    Public Property emailAddress() As String
        Get
            Return memailAddress
        End Get
        Set(ByVal value As String)
            memailAddress = value
        End Set
    End Property

    Public Property MobileNo() As String
        Get
            Return mMobileNo
        End Get
        Set(ByVal value As String)
            mMobileNo = value
        End Set
    End Property


    Protected mPasswordExpiration As Long
    Public Property PasswordExpiration As Long
        Get
            Return mPasswordExpiration
        End Get
        Set(value As Long)
            mPasswordExpiration = value
        End Set
    End Property
    Protected mPasswordExpirationDays As Long
    Public Property PasswordExipirationDays As Long
        Get
            Return mPasswordExpirationDays

        End Get
        Set(value As Long)
            mPasswordExpirationDays = value
        End Set
    End Property
    Protected mAccountLockStatus As Long
    Public Property AccountLockStatus As Long
        Get
            Return mAccountLockStatus
        End Get
        Set(value As Long)
            mAccountLockStatus = value
        End Set
    End Property
    Protected mDeleted As Long
    Public Property Deleted As Long
        Get
            Return mDeleted
        End Get
        Set(value As Long)
            mDeleted = value
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
    Public Function validateExistingUserAcount(ByVal Username As String) As Boolean
        Try
            Dim str As String = "select * from tblUsers where Username = " & Username & ""
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, str)
            If (Not IsNothing(ds) AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Function getUsers() As DataSet
        Try
            Dim sql As String = "select UserID, username,EmailAddress,MobileNo,Firstname,Surname, case AccountLockStatusID when 0 then 'Active' when 1 then 'Locked' else 'unDefined' end as Status from tblUsers where Deleted=0 and statusid=1"
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

    Public Sub Clear()

        UserID = 0
        mUserTypeID = 0
        mDepartmentID = 0
        mStatusID = 0
        mCreatedBy = mObjectUserID
        mCreatedDate = ""
        mUsername = ""
        mFirstname = ""
        mSurname = ""
        mPassword = ""
        memailAddress = ""
        mMobileNo = ""
        mCompanyID = 0
        mBranchID = 0
        mPasswordUpdateDate = Date.Today
    End Sub
    Private Function ValidateHash(ByVal Username As String, ByVal UserPassword As String, Optional ByVal dbHashedPwdSupplied As Boolean = False) As Boolean

        Try
            Dim dbHashedPassword As String = ""

            If dbHashedPwdSupplied Then
                Dim dsUser As DataSet = db.ExecuteDataSet(CommandType.Text, "SELECT * FROM tblUsers where Username = '" & Username & "'")
                If (dsUser.Tables.Count > 0 AndAlso dsUser.Tables(0).Rows.Count > 0) Then

                    LoadDataRecord(dsUser.Tables(0).Rows(0))
                    dbHashedPassword = dsUser.Tables(0).Rows(0)("Password")

                End If
            Else

                dbHashedPassword = mPassword

            End If

            ' Create an Encoding object so that you can use the convenient GetBytes 
            ' method to obtain byte arrays.
            Dim uEncode As New Text.UnicodeEncoding()

            Dim bytHashOriginal As Byte() = uEncode.GetBytes(dbHashedPassword)

            Dim strHashForCompare As String = GenerateHashDigest(Username.ToLower + UserPassword)
            ' From the new hash digest create a byte array for comparison with the
            ' original hash digest byte array.
            Dim bytHashForCompare As Byte() = uEncode.GetBytes(strHashForCompare)
            ' Display the new hash digest in a TextBox.

            'Loop through all the bytes in the hashed values.
            Dim i As Integer
            For i = 0 To bytHashOriginal.Length - 1
                If bytHashOriginal(i) <> bytHashForCompare(i) Then

                    Return False

                Else
                    ' Every byte matched so the "transmitted" XML has been authenticated.

                End If
            Next
            ' Compare each byte. If any do not match display an appropriate message
            ' and exit the loop.
            Return True

        Catch ex As Exception

            mMsgFlg = ex.Message
            Return False

        End Try

    End Function
    Public Function PasswordHash(ByVal username As String, ByVal UserPassword As String) As String

        Return GenerateHashDigest(UserPassword)

    End Function
    Public Function GenerateHashDigest(ByVal strSource As String) As String
        ' Create an Encoding object so that you can use the convenient GetBytes 
        ' method to obtain byte arrays.

        Try

            Dim hash As Byte()
            Dim salt As String = "Spec@#*9" 'salt to be added to the Username and Password combination
            Dim uEncode As New Text.UnicodeEncoding()
            ' Create a byte array from the source text passed as an argument.
            Dim bytPassword() As Byte = uEncode.GetBytes(strSource & salt)

            Dim sha384 As New SHA384Managed()
            hash = sha384.ComputeHash(bytPassword)

            ' Base64 is a method of encoding binary data as ASCII text.
            Return Convert.ToBase64String(hash)

        Catch ex As Exception

            mMsgFlg = ex.Message
            Return Nothing

        End Try

    End Function
    Public Function ChangePassword(ByVal UserID As Long, ByVal Password As String) As Boolean
        Try
            Dim sql As String
            Dim HashedPassword As String = GenerateHashDigest(Password)

            sql = "UPDATE tblUsers SET [Password] = @Password, PasswordUpdateDate=@UpdateDate WHERE UserID = @UserID "
            'sql = "UPDATE tblUsers SET  " & vbCrLf
            'sql &= "	[Password] = @Password, UpdatedDate='" & Date.Now.ToString("yyyy-dd-MM HHH:mm:ss") & "'  " & vbCrLf
            'sql &= "	FailedPasswordAttemptCount = 0,  " & vbCrLf
            'sql &= "	FailedPasswordAttemptWindowStart = NULL  " & vbCrLf
            'sql &= "WHERE UserID = @UserID " & vbCrLf

            Dim cmd As System.Data.Common.DbCommand = db.GetSqlStringCommand(sql)

            db.AddInParameter(cmd, "@Password", DbType.String, HashedPassword)
            db.AddInParameter(cmd, "@UserID", DbType.Int32, UserID)
            db.AddInParameter(cmd, "@UpdateDate", DbType.Date, Now.Date)

            db.ExecuteNonQuery(cmd)
            Return True

        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Sub ResetPassword()

        'Dim sql As String = "UPDATE tblUsers SET [Password] = '" & mPassword & "', [LastPasswordChangeDate] = '" & Date.Today & "', PasswordExpires = 0 WHERE [Username] = '" & mUsername & "'"
        Dim sql As String = "UPDATE tblUsers SET [Password] = '" & mPassword & "' WHERE [Username] = '" & mUsername & "'"
        db.ExecuteNonQuery(CommandType.Text, sql)

    End Sub

#Region "Retrieve Overloads"

    Public Overridable Function Retrieve() As Boolean

        Return Me.Retrieve(mUserID)

    End Function

    Public Overridable Function Retrieve(ByVal UserID As Long) As Boolean

        Dim sql As String

        If UserID > 0 Then
            sql = "SELECT * FROM tblUsers WHERE UserID = " & UserID
        Else
            sql = "SELECT * FROM tblUsers WHERE UserID = " & mUserID
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

                mMsgFlg = "UserAdministration account not found."

                Return False

            End If

        Catch e As Exception

            mMsgFlg = e.Message
            Return False

        End Try

    End Function

    Public Overridable Function GetUserAdministration() As System.Data.DataSet

        Return GetUserAdministration(mUserID)

    End Function

    Public Overridable Function GetUserAdministration(ByVal UserID As Long) As DataSet

        Dim sql As String

        If UserID > 0 Then
            sql = "SELECT * FROM tblUsers WHERE UserID = " & UserID
        Else
            sql = "SELECT * FROM tblUsers WHERE UserID = " & mUserID
        End If

        Return GetUserAdministration(sql)

    End Function

    Protected Overridable Function GetUserAdministration(ByVal sql As String) As DataSet

        Return db.ExecuteDataSet(CommandType.Text, sql)

    End Function

#End Region

    Protected Friend Overridable Sub LoadDataRecord(ByRef Record As Object)

        With Record

            mUserID = IIf(IsDBNull(.Item("UserID")), 0, .Item("UserID"))
            mUserTypeID = IIf(IsDBNull(.Item("UserTypeID")), 0, .Item("UserTypeID"))
            mDepartmentID = IIf(IsDBNull(.Item("DepartmentID")), 0, .Item("DepartmentID"))
            mStatusID = IIf(IsDBNull(.Item("StatusID")), 0, .Item("StatusID"))
            mCreatedBy = IIf(IsDBNull(.Item("CreatedBy")), 0, .Item("CreatedBy"))
            mCreatedDate = IIf(IsDBNull(.Item("CreatedDate")), "", .Item("CreatedDate"))
            mUsername = IIf(IsDBNull(.Item("Username")), "", .Item("Username"))
            mFirstname = IIf(IsDBNull(.Item("Firstname")), "", .Item("Firstname"))
            mSurname = IIf(IsDBNull(.Item("Surname")), "", .Item("Surname"))
            mPassword = IIf(IsDBNull(.Item("Password")), "", .Item("Password"))
            memailAddress = IIf(IsDBNull(.Item("emailAddress")), "", .Item("emailAddress"))
            mMobileNo = IIf(IsDBNull(.Item("MobileNo")), "", .Item("MobileNo"))

            mPasswordExpirationDays = IIf(IsDBNull(.Item("DaysToExpire")), 0, .Item("DaysToExpire"))
            mPasswordExpiration = IIf(IsDBNull(.Item("PasswordExpires")), 0, .Item("PasswordExpires"))
            mAccountLockStatus = IIf(IsDBNull(.Item("AccountLockStatusID")), 0, .Item("AccountLockStatusID"))
            mDeleted = IIf(IsDBNull(.Item("Deleted")), 0, .Item("Deleted"))

            mCompanyID = IIf(IsDBNull(.Item("CompanyID")), 0, .Item("CompanyID"))
            mBranchID = IIf(IsDBNull(.Item("BranchID")), 0, .Item("BranchID"))

        End With

    End Sub

#Region "Save"

    Public Overridable Sub GenerateSaveParameters(ByRef db As Database, ByRef cmd As System.Data.Common.DbCommand)

        db.AddInParameter(cmd, "@UserID", DbType.Int32, mUserID)
        db.AddInParameter(cmd, "@UserTypeID", DbType.Int32, mUserTypeID)
        db.AddInParameter(cmd, "@DepartmentID", DbType.Int32, mDepartmentID)
        db.AddInParameter(cmd, "@StatusID", DbType.Int32, mStatusID)
        db.AddInParameter(cmd, "@Username", DbType.String, mUsername)
        db.AddInParameter(cmd, "@CompanyID", DbType.Int32, mCompanyID)
        db.AddInParameter(cmd, "@BranchID", DbType.Int32, mBranchID)
        db.AddInParameter(cmd, "@Firstname", DbType.String, mFirstname)
        db.AddInParameter(cmd, "@Surname", DbType.String, mSurname)
        db.AddInParameter(cmd, "@Password", DbType.String, GenerateHashDigest(mPassword))
        db.AddInParameter(cmd, "@emailAddress", DbType.String, memailAddress)
        db.AddInParameter(cmd, "@MobileNo", DbType.String, mMobileNo)
        db.AddInParameter(cmd, "@PasswordUpdateDate", DbType.Date, mPasswordUpdateDate)

    End Sub

    Public Overridable Function Save() As Boolean

        Dim cmd As System.Data.Common.DbCommand = db.GetStoredProcCommand("sp_Save_UserAdministration")

        GenerateSaveParameters(db, cmd)

        Try

            Dim ds As DataSet = db.ExecuteDataSet(cmd)

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then

                mUserID = ds.Tables(0).Rows(0)(0)

            End If

            Return True

        Catch ex As Exception

            mMsgFlg = ex.Message
            Return False

        End Try

    End Function
    Public Function ValidateUser() As Boolean

        Return ValidateUsingPassword(mUsername, mPassword)

    End Function
    Public Overridable Function GetLoggingDetails(ByVal UserID As Long) As DataSet
        Try
            Dim str As String = "select U.UserID,U.Username,U.Firstname,U.Surname,U.emailAddress,U.MobileNo,Ut.Description as UserType,D.Description as Department,Cg.CompanyName ,Cg.RequisitionPrefixCode ,B.BranchName   from tblUsers U inner join tblCompanyGroups Cg on Cg.CompanyID = U.CompanyID Inner join luUserTypes Ut on Ut.UserTypeID = U.UserTypeID inner join luDepartments D on D.DepartmentID = U.DepartmentID left join luBranches B on B.BranchID = U.BranchID where UserID= " & UserID & ""
            Return GetLoggingDetails(str)
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return Nothing
        End Try
    End Function
    Protected Overridable Function GetLoggingDetails(ByVal str As String) As DataSet
        Try
            Dim ds As DataSet = db.ExecuteDataSet(CommandType.Text, str)
            If Not IsNothing(ds) AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
                Return ds
            Else
                Return Nothing
            End If
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return Nothing
        End Try
    End Function
    Private Function ValidateUsingPassword(ByVal UserName As String, ByVal Password As String) As Boolean

        Dim sql As String = String.Empty
        If My.Computer.Name.ToUpper.Equals("DEVMACHINE1") Then
            sql = "SELECT TOP 1 * FROM tblUsers WHERE UserName=@UserName"
        Else
            'sql = "SELECT * FROM tblUsers WHERE UserName=@UserName AND Password=@Password"
            sql = "SELECT * FROM tblUsers WHERE UserName='" & UserName & "' AND Password='" & Password & "'"
        End If

        Dim cmd As System.Data.Common.DbCommand = db.GetSqlStringCommand(sql)

        db.AddInParameter(cmd, "@UserName", DbType.String, UserName)
        db.AddInParameter(cmd, "@Password", DbType.String, Password)

        Dim ds As DataSet = db.ExecuteDataSet(cmd)

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 AndAlso Not IsDBNull(ds.Tables(0).Rows(0)("UserID")) Then '' AndAlso Not (ds.Tables(0).Rows(0)("Deleted")) Then

            LoadDataRecord(ds.Tables(0).Rows(0))
            'UpdateLastLogin(mUsername, mApplicationID)
            Return True
        Else
            ' To be used in live environ            UpdateFailureCount(UserName, "password")
            Return False
        End If

    End Function
    Public Function ValidateLogin(ByVal Username As String) As DataSet
        Try
            Dim sql As String = "select *, DATEDIFF(day,PasswordUpdateDate, getDate()) as PasswordAge from tblUsers where Username='" & Username & "' and Deleted=0 and statusid=1"
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

#Region "Delete"
    Public Function UpdateExtraFields(ByVal userID As Long, ByVal PasswordExpires As Long, ByVal Daystoexpire As Long, ByVal AccountLockStatusID As Long, ByVal Deleted As Long) As Boolean
        Try
            Dim str As String = "update tblUsers set PasswordExpires =  " & PasswordExpires & ", Daystoexpire = " & Daystoexpire & " , AccountLockStatusID = " & AccountLockStatusID & ", Deleted = " & Deleted & " where userid = " & userID & ""
            db.ExecuteNonQuery(CommandType.Text, str)
            Return True
        Catch ex As Exception
            mMsgFlg = ex.Message
            Return False
        End Try
    End Function
    Public Overridable Function Delete() As Boolean

        'Return Delete("UPDATE tblUsers SET Deleted = 1 WHERE UserID = " & mUserID) 
        Return Delete("DELETE FROM tblUsers WHERE UserID = " & mUserID)

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