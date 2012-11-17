<!-- #include virtual="NavigatorItem.asp" -->
<%
Class Navigator
	Private varNavigationElement() 	'Lista de elementos Tag NAVIGATIONELEMENT do navegador
	Private varName 				
	Private varDisplayName 		
	
	'Constructor
	Private Sub class_Initialize
		reDim varNavigationElement(0)
		varName = ""
		varDisplayName = ""
	End Sub
	
	'Methods Public
	Public Property Get Name
		Name = varName
	End Property
	
	Public Property Let Name(nNavigator)
		varName = nNavigator
	End Property
	
	Public Property Get DisplayName
		DisplayName = varDisplayName
	End Property
	
	Public Property Let DisplayName(nNavigator)
		varDisplayName = nNavigator
	End Property
	
	Public Property Get NavigationElement
		NavigationElement = varNavigationElement
	End Property
	
	Public Property Let NavigationElement(paramNavigationElement)
		varNavigationElement = paramNavigationElement
	End Property
	
	Public Function addItem(paramNavigatorItem)
		Dim quantidade
		quantidade = Ubound(varNavigationElement)
		reDim Preserve varNavigationElement( quantidade + 1 )
		Set varNavigationElement( quantidade ) = paramNavigatorItem
	End Function
	
End Class
%>