Class cDbCursor
	Private numCur
	Private lastMoveNext
	
	Public Default Function Init(connection, query)
		numCur = $DBCursorOpenSQL(connection, query)
		Set Init = Me
	End Function
	
	Public Function GetValue(strColumn)
		GetValue = $DBCursorGetValue(numCur, strColumn)
	End Function
	
	Public Function GetValues(columnNames)
		Dim ret, row
		ReDim ret(UBound(columnNames))
		
		For row=0 To UBound(columnNames) -1 
			ret(row) = GetValue(columnNames(row))
		Next
		
		GetValues = ret
	End Function
	
	Public Function ColumnInfo(numColumn, numTypeInfo)
		ColumnInfo = $DBCursorColumnInfo(numCur, numColumn, numTypeInfo)
	End Function
	
	Public Function ColumnName(numColumn)
		ColumnName = ColumnInfo(numColumn, 0)
	End Function
	
	Public Function ColumnType(numColumn)
		ColumnType = ColumnInfo(numColumn, 1)
	End Function
	
	Public Function ColumnCount()
		ColumnCount = $DBCursorColumnCount(numCur)
	End Function
	
	Function ColumnNames()
		Dim ret, row
		ReDim ret(ColumnCount()-1)
	
		For row=1 To ColumnCount()
			ret(row-1) = ColumnName(row)
		Next
	
		ColumnNames = ret
	End Function
	
	Public Function MoveNext()
		lastMoveNext = $DBCursorNext(numCur)

		If lastMoveNext => 0 Then
			MoveNext = True
		Else
			MoveNext = False
		End If
	End Function
	
	Public Function EOF()
		If lastMoveNext = 0 Then
			EOF = False
		Else
			EOF = True
		End If
	End Function
	
	Public Function MoveFirst()
		lastMoveNext = $DBCursorMoveTo(numCur, 0)
		End Function
	
	Public Function Count()
		Count = $DBCursorRowCount(numCur)
	End Function
	
	Public Sub CloseCnn()
		$DBCursorClose(numCur)
	End Sub
	
	Private Sub Class_Terminate
   	CloseCnn()
  	End Sub
End Class
