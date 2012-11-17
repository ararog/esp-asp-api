<!-- #include virtual="Navigator.asp" -->
<!-- #include virtual="Document.asp" -->
<!-- #include virtual="QueryResult.asp" -->
<%
Class Query
	Private varParameters 		
	Private varConnection 		
	
	Private varQuery 				
	Private varHits 				
	Private varOffset 				
	Private varLemmatize 				
	Private varNavigation 				
							
							
	
	Private varNavigators 

	Private varLanguage 
	Private varSortby
	Private varSpellcheck 
	Private varSynomyns 	
	
	Private varInCache 
	
	Private varSearchViewName
	
	Private varModifiers()
	
	'Constructor
	Private Sub class_initialize
		Set varParameters = Server.CreateObject("Scripting.Dictionary")
		varHits = 10
		varOffset = 0
		varLanguage = "pt"
		varNavigation = True
		varLemmatize = False
		varSynomyns = False
		varInCache = False
		
		reDim varModifiers(0)
	End Sub
	
	'Setter e Getters
	Public Property Get Parameters()
		Parameters = varParameters
	End Property
	Public Property Let Parameters(aux)
		Set varParameters = aux
	End Property
	
	Public Property Get Connection()
		Set Connection = varConnection
	End Property
	Public Property Let Connection(auxCon)
		Set varConnection = auxCon
	End Property
	
	Public Property Get OffSet()
		OffSet = varOffSet
	End Property
	Public Property Let OffSet(aux)
		varOffSet = aux
	End Property
	
	Public Property Get Query()
		Query = varQuery
	End Property
	Public Property Let Query(aux)
		Set varQuery = aux
	End Property
	
	Public Property Get Hits()
		Hits = varHits
	End Property
	Public Property Let Hits(aux)
		varHits = aux
	End Property
	
	Public Property Get Lemmatize()
		Lemmatize = varLemmatize
	End Property
	Public Property Let Lemmatize(aux)
		varLemmatize = aux
	End Property
	
	Public Property Get Navigation()
		Navigation = varNavigation
	End Property
	Public Property Let Navigation(aux)
		varNavigation = aux
	End Property

	Public Property Get Navigators()
		Navigators = varNavigators
	End Property
	Public Property Let Navigators(aux)
		varNavigators = aux
	End Property
	
	Public Property Get Language()
		Language = varLanguage
	End Property
	Public Property Let Language(aux)
		varLanguage = aux
	End Property
	
	Public Property Get Sortby()
		Sortby = varSortby
	End Property
	Public Property Let Sortby(aux)
		varSortby = aux
	End Property
	
	Public Property Get Spellcheck()
		Spellcheck = varSpellcheck
	End Property
	Public Property Let Spellcheck(aux)
		Set varSpellcheck = aux
	End Property
	
	Public Property Get Synomyns()
		Synomyns = varSynomyns
	End Property
	Public Property Let Synomyns(aux)
		varSynomyns = aux
	End Property

	Public Property Get InCache()
		InCache = varInCache
	End Property
	Public Property Let InCache(aux)
		varInCache = aux
	End Property

	Public Property Get SearchViewName()
		SearchViewName = varSearchViewName
	End Property
	Public Property Let SearchViewName(aux)
		varSearchViewName = aux
	End Property

	Public Property Get Modifiers()
		Modifiers = varModifiers
	End Property
	Public Property Let Modifiers(aux)
		varModifiers = aux
	End Property
	
	'Functions
	
	Public Function addModifier( aux )
		Dim quantidadeNaLista
		quantidadeNaLista = Ubound( varModifiers )
		reDim Preserve varModifiers( quantidadeNaLista +1 )
		Set varModifiers( quantidadeNaLista ) = aux
	End Function	
	
	Private Function prepareQuery()
		If ( Len(varQuery) > 0 ) Then
			varParameters.add "query" , varQuery
		End If
		If ( Int(Trim(varHits)) > 0 ) Then
			varParameters.add "hits" , varHits
		End If
		If ( Int(Trim(varOffset)) > 0 ) Then
			varParameters.add "offset" , varOffset
		End If
		If ( Len(varSearchViewName) > 0 ) Then
			varParameters.add "view" , varSearchViewName
		End If
		If ( Len(varLanguage) > 0 ) Then
			varParameters.add "language" , varLanguage
		End If
		If ( Len(varSortby) > 0 ) Then
			varParameters.add "sortby" , varSortby
		End If
		If ( varInCache ) Then
			varParameters.add "qtf_teaser:view" , "hithighlight"
		End If
		If ( Len(varNavigators) > 0 ) Then
			varParameters.add "rpf_navigation:navigators" , varNavigators
		End If

		If ( varNavigation ) Then
			varParameters.add "rpf_navigation:enabled" , "true"
		Else	
			varParameters.add "rpf_navigation:enabled" , "false"
		End If
		If ( varLemmatize ) Then
			varParameters.add "qtf_lemmatize" , "true"
		Else	
			varParameters.add "qtf_lemmatize" , "false"
		End If
		If ( varSynomyns ) Then
			varParameters.add "qtf_querysynonyms" , "true"
		Else	
			varParameters.add "qtf_querysynonyms" , "false"
		End If
		
		If ( Len(varSpellcheck) > 0 ) Then
			If ( Trim(varSpellcheck) = "yes" OR Trim(varSpellcheck) = "1" ) Then
				varParameters.add "spell" , "1"
			End If
			If ( Trim(varSpellcheck) = "no" OR Trim(varSpellcheck) = "0" ) Then
				varParameters.add "spell" , "0"
			End If
			If ( Trim(varSpellcheck) = "Suggest" OR Trim(varSpellcheck) = "suggest" ) Then
				varParameters.add "spell" , "suggest"
			End If
		End If
				
		Dim filters
		filters = ""
		
		Dim modifiersCount
		modifiersCount = Ubound( varModifiers )		
		If ( modifiersCount > 0 ) Then
			For i = 0 To ( modifiersCount - 1 )
				Set modifier = varModifiers(i)
				value = modifier.Value
				If (InStrRev( value , " " ) > 0) Then
					value = chr(34) & modifier.Value & chr(34)
				End If
				
				If (IsNumeric(value)) Then
					filters = filters & "+" & modifier.Field & ":" & value 
				Else
					filters = filters & "+" & modifier.Field & ":^" & value & "$"
				End If
				
			Next
			varParameters.add "navigation" , filters
		End	If
	End Function
	
	Public Function execute(term, subportal)
		If ( Len(termo) > 0 ) Then
			varQuery = "string(""" & Replace(term, """", "\""") & """, mode=""simpleall"", annotation_class=""user"")"
		End If
		
		prepareQuery()
		varConnection.setParameters( varParameters )
		varConnection.MontaURL()
		
		Dim objQueryResult
		Set objQueryResult = new QueryResult
		
		Dim objXMLHTTP
		Set objXMLHTTP = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objXMLHTTP.Open "GET" , varConnection.URL , false
		objXMLHTTP.setRequestHeader "Content-Type", "text/html; charset=utf-8"
		objXMLHTTP.Send
		
		Dim objXML
		Set objXML = Server.CreateObject("Msxml2.FreeThreadedDOMDocument.3.0")
		objXML.validateOnParse = false
		objXML.resolveExternals = false
		objXML.preserveWhiteSpace = false
		objXML.async = false
		
		objXML.Load( objXMLHTTP.ResponseStream  )
		
		objQueryResult.parseServerResponse(objXML )
		
		Set execute = objQueryResult

	End Function
	
End Class
%>