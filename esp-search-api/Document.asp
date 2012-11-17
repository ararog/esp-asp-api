<%
Class Document
	Private varFields
	
	'Construtor
	Private Sub class_Initialize
		Set varFields = Server.CreateObject("Scripting.Dictionary")
	End Sub
	
	Public Property Get Fields()
		Set Fields = varFields
	End Property
	
	Public Function addField( key , value )
		varFields.Add key , value
	End Function
	
End Class
%>