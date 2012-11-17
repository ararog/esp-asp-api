<%
Class SearchProfile
	Private varResultView
	Private varRankProfile
	
	'Constructor
	Private Sub class_Initialize

	End Sub
	
	'Setters e Getters
	Public Property Get ResultView()
		ResultView = varResultView
	End Property
	Public Property Let ResultView(aux)
		varResultView = aux
	End Property
	
	Public Property Get RankProfile()
		RankProfile = varRankProfile
	End Property
	Public Property Let RankProfile(aux)
		varRankProfile = aux
	End Property
	
	'Functions
	Public Function ParseConfiguration(xmlDocument)

		Set resultSpecEntry = xmlDocument.selectSingleNode("//view/result-spec") 

		ResultView  = resultSpecEntry.Attributes.GetNamedItem("default-result-view").Text
		RankProfile = resultSpecEntry.Attributes.GetNamedItem("default-rank-profile").Text
	End Function
	
End Class
%>