<%
Class QueryTransformation

	Private varName
	Private varAction
	Private varQuery
	Private varCustom
	Private varMessage
	Private varMessageId

	'Constructor
	Private Sub class_initialize

	End Sub

	Public Property Get Name()
		Name = varName
	End Property
	Public Property Let Name(aux)
		varName = aux
	End Property
	
	Public Property Get Action()
		Action = varAction
	End Property
	Public Property Let Action(aux)
		varAction = aux
	End Property
	
	Public Property Get Query()
		Query = varQuery
	End Property
	Public Property Let Query(aux)
		varQuery = aux
	End Property		

	Public Property Get Custom()
		Custom = varCustom
	End Property
	Public Property Let Custom(aux)
		varCustom = aux
	End Property		

	Public Property Get Message()
		Message = varMessage
	End Property
	Public Property Let Message(aux)
		varMessage = aux
	End Property		

	Public Property Get MessageId()
		MessageId = varMessageId
	End Property
	Public Property Let MessageId(aux)
		varMessageId = aux
	End Property		

End Class
%>