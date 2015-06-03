Namespace MSprintEx2014
    ''' <summary>
    ''' Summary description for TG.
    ''' </summary>
    ''' 
    <Serializable()> _
    Public Class TG
        Private gender As Integer()
        Private age As Integer()
        Private sec As Integer()
        Private hh As Integer()
        Private tgName As String
        Private qryString As String
        Private uniqueID As String

        Public Sub New()
            'constructor
            Me.gender = New Integer(1) {}
            Me.age = New Integer(6) {}
            Me.sec = New Integer(3) {}
            '  Me.hh = New Integer(1) {}
            Me.hh = New Integer(3) {}
            For i As Integer = 0 To Me.gender.GetUpperBound(0)
                Me.gender(i) = 0
            Next
            For i As Integer = 0 To Me.age.GetUpperBound(0)
                Me.age(i) = 0
            Next
            For i As Integer = 0 To Me.sec.GetUpperBound(0)
                Me.sec(i) = 0
            Next
            For i As Integer = 0 To Me.hh.GetUpperBound(0)
                Me.hh(i) = 0
            Next

            Me.tgName = ""
            Me.setUniqueID()
            Me.uniqueID = Me.getUniqueID()
            Me.qryString = ""
        End Sub
        ' end TG() constructor
        Public Function getTGName() As String
            Return Me.tgName
        End Function
        Public Sub setTGName(ByVal inpt As String)
            Me.tgName = inpt
        End Sub

        Public Function getUniqueID() As String
            Return Me.uniqueID
        End Function
        Public Function getGenderIds(ByVal i As Integer) As Integer
            Return Me.gender(i)
        End Function
        Public Sub setGenderIds(ByVal i As Integer)
            Me.gender(i) = 1
        End Sub
        Public Function getAgeIds(ByVal i As Integer) As Integer
            Return Me.age(i)
        End Function
        Public Sub setAgeIds(ByVal i As Integer)
            Me.age(i) = 1
        End Sub
        Public Function getSECIds(ByVal i As Integer) As Integer
            Return Me.sec(i)
        End Function
        Public Sub setSECIds(ByVal i As Integer)
            Me.sec(i) = 1
        End Sub
        Public Function getHHIds(ByVal i As Integer) As Integer
            Return Me.hh(i)
        End Function
        Public Sub setHHIds(ByVal i As Integer)
            Me.hh(i) = 1
        End Sub
        Public Sub setUniqueID()
            Dim idToSet As String = ""
            For i As Integer = 0 To Me.gender.GetUpperBound(0)
                idToSet += Me.gender(i).ToString()
            Next
            For i As Integer = 0 To Me.age.GetUpperBound(0)
                idToSet += Me.age(i).ToString()
            Next
            For i As Integer = 0 To Me.sec.GetUpperBound(0)
                idToSet += Me.sec(i).ToString()
            Next
            For i As Integer = 0 To Me.hh.GetUpperBound(0)
                idToSet += Me.hh(i).ToString()
            Next
            Me.uniqueID = idToSet
        End Sub

        Public Function getQueryString() As String
            Return Me.qryString
        End Function
        Public Sub setQueryString(ByVal fmPreviouslyBuiltTGQueryString As String)
            Me.qryString = fmPreviouslyBuiltTGQueryString
        End Sub
        Public Sub setUniqueID(ByVal fmPreviouslyBuiltTGUniqueID As String)
            Me.uniqueID = fmPreviouslyBuiltTGUniqueID
        End Sub

        Public Sub setQueryString()
            Dim qryStrToSet As String = ""
            Dim genderStr As String = Me.returnDimString("SEX_ID", Me.gender, Me.checkAllCellsChecked(Me.gender))
            Dim ageStr As String = Me.returnDimString("AGE_ID", Me.age, Me.checkAllCellsChecked(Me.age))
            Dim secStr As String = Me.returnDimString("SEC_ID", Me.sec, Me.checkAllCellsChecked(Me.sec))
            Dim hhStr As String = Me.returnDimString("HOUSE_HOLD_ID", Me.hh, Me.checkAllCellsChecked(Me.hh))
            qryStrToSet = genderStr & " and " & ageStr & " and " & secStr & " and  " & hhStr
            Me.qryString = qryStrToSet
        End Sub

        Private Function checkAllCellsChecked(ByVal paraArray As Integer()) As Boolean
            Dim blnAllChecked As Boolean = True
            ' go through the array values. one item not equal to zero means 
            ' all values are not checked
            For i As Integer = 0 To paraArray.GetUpperBound(0)
                If paraArray(i) = 0 Then
                    blnAllChecked = False
                End If
            Next
            Return blnAllChecked
        End Function

        Private Function returnDimString(ByVal [dim] As String, ByVal dimArray As Integer(), ByVal blnAllChecked As Boolean) As String
            Dim strToRet As String = ""
            If blnAllChecked Then
                strToRet = "0 = 0"
            Else
                ' else blnallchecked = false
                Dim dimUB As Integer = dimArray.GetUpperBound(0)
                Dim endCounter As Integer = 0
                Dim begCounter As Integer = 0
                For i As Integer = 0 To dimUB
                    If dimArray(i) = 1 Then
                        endCounter += 1
                    End If
                Next
                If endCounter = 1 Then
                    For jk As Integer = 0 To dimArray.GetUpperBound(0)
                        If dimArray(jk) = 1 Then
                            Dim valToIncl As Integer = jk + 1
                            strToRet += [dim] & " in ( " & valToIncl & " ) "
                        End If
                    Next
                Else
                    ' else nkp
                    For lm As Integer = 0 To dimArray.GetUpperBound(0)
                        'forab
                        If dimArray(lm) = 1 Then
                            'ifab
                            Dim valToIncl As Integer = lm + 1
                            If begCounter = 0 Then

                                strToRet += [dim] & " in ( " & valToIncl & " , "
                            Else
                                If begCounter < endCounter - 1 Then
                                    strToRet += valToIncl & " , "
                                Else
                                    strToRet += valToIncl & " ) "
                                End If
                            End If
                            begCounter += 1
                        End If
                        'end ifab and forab - the ifab being a single statement
                    Next
                    ' end else nkp
                End If
            End If
            'end else blnalchecked = false
            Return strToRet
        End Function

    End Class
    'end class
End Namespace
'end namespace