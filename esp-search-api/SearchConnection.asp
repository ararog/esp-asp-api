<%
Class SearchConnection
	Private varHost
	Private varPort
	Private varParameters 
	Private varURL 		
	Private varPath 	
	
	'Constructor
	Private Sub class_Initialize
		varHost = "localhost" 
		varPort = "15100"
		varPath = "/cgi-bin/xsearch"
	End Sub
	
	'Setters e Getters
	Public Property Get Host()
		Host = varHost
	End Property
	Public Property Let Host(aux)
		varHost = aux
	End Property
	
	Public Property Get Port()
		Port = varPort
	End Property
	Public Property Let Port(aux)
		varPort = aux
	End Property
	
	Public Property Get Path()
		Path = varPath
	End Property
	Public Property Let Path(aux)
		varPath = aux
	End Property
	
	Public Function setParameters( aux )
		Set varParameters = aux
	End Function
	
	Public Property Get URL()
		URL = varURL
	End Property
	Public Property Let URL(aux)
		varURL = aux
	End Property

	'Functions
	Public Function DiscoverSearchProfile(paramSearchView)

		Dim objXMLHTTP
		Set objXMLHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objXMLHTTP.Open "GET" , "http://" & Host & ":" & Port & "/get?view=" & paramSearchView, false
		objXMLHTTP.Send
		
		Dim objXML
		Set objXML = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.3.0")
		objXML.validateOnParse = false
		objXML.resolveExternals = false
		objXML.preserveWhiteSpace = false
		objXML.async = false
		
		Dim xmlResponse
		xmlResponse = objXMLHTTP.responseText
		If (Len(xmlResponse) > 0) Then
			objXML.LoadXML( xmlResponse )

			Dim objSearchProfile
			Set objSearchProfile = new SearchProfile
			objSearchProfile.parseConfiguration(objXML)
		
			Path = "/cgi-bin/xml-" & objSearchProfile.ResultView

			Set DiscoverSearchProfile = objSearchProfile
		End If	
	End Function
	
	'Functions
	Public Function MontaURL()
		varURL = "http://" & Host & ":" & Port & Path
		
		If varParameters.Count > 0 Then 
			Dim quantidade
			quantidade = varParameters.Count
			
			Dim varKeys, varValues
			varKeys = varParameters.Keys
			varValues = varParameters.Items
			
			varURL = varURL & "?encoding=iso-8859-1"
			For i = 0 To ( quantidade - 1 )
				varURL = varURL & "&" & varKeys(i) & "=" & varValues(i)
			Next
		End If
	End Function
	
End Class
%>