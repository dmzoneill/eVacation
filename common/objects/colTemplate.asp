<%
			 COLLECTION TEMPLATE ***
		Class mCol_name_
			'====================================
			'Use this template to create a
			'specialised Collection Object, it
			'inherits objCollection properties
			'and methods by explicit declaration.
			'====================================
			Private m_objCollection
	
			Private Sub Class_Initialize()
				Set m_objCollection = New mObjCollection
			End Sub
			
			Private Sub Class_Terminate()
				Set m_objCollection = Nothing
			End Sub
	
			Public Property Get Count
				Count = m_objCollection.Count
			End Property
	
			Public Default Property Get Item(ByVal loclngIndex)
				Set Item = m_objCollection(loclngIndex)
			End Property
			
			Public Sub Add(ByVal locObject)
				m_objCollection.Add(locObject)
			End Sub
		End Class
%>