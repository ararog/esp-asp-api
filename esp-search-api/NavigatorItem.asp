<%
'Classe que representa a Tag NAVIGATIONELEMENT de um navegador(NAVIGATIONENTRY)
Class NavigatorItem
	Private varLabel 			'Atributo NAME da TAG NAVIGATIONELEMENT
	Private varValue 			'Atributo MODIFIER da TAG NAVIGATIONELEMENT
	Private varCount 			'Atributo COUNT da TAG NAVIGATIONELEMENT
	
	'Constructor
	Private Sub class_Initialize
	end Sub
	
	'Setters e Getters
	Public Property Get Label
		Label = varLabel
	End Property
	
	Public Property Let Label(paramLabel)
		varLabel = paramLabel
	End Property
	
	Public Property Get Value
		Value = varValue
	End Property
	
	Public Property Let Value(paramValue)
		varValue = paramValue
	End Property
	
	Public Property Get Count
		Count = varCount
	End Property
	
	Public Property Let Count(paramCount)
		varCount = paramCount
	End Property
	
End Class
%>