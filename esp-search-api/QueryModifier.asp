<%
Class QueryModifier
	Private varField	
	Private varValue	
	
	'Constructor
	Private Sub class_initialize

	End Sub
	
	'Setter e Getters
	Public Property Get Field()
		Field = varField
	End Property
	Public Property Let Field(aux)
		varField = aux
	End Property
	
	Public Property Get Value()
		Value = varValue
	End Property
	Public Property Let Value(aux)
		varValue = aux
	End Property
	
End Class
%>