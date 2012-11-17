<!-- #include virtual="QueryTransformation.asp" -->
<%
Class QueryResult
	Private varDocuments()
	Private varNavigators()
	Private varQueryTransformations()
	Private varFirsthit
	Private varLasthit
	Private varHits
	Private varTotalHits
	Private varTime

	'Constructor
	Private Sub class_Initialize
		reDim varDocuments(0)
		reDim varNavigators(0)
		reDim varQueryTransformations(0)
	End Sub
	
	'Setters and Getters
	Public Property Get Documents()
		Documents = varDocuments
	End Property
	
	Public Property Let Documents(aux)
		varDocuments = aux
	End Property
	
	Public Property Get Navigators()
		Navigators = varNavigators
	End Property
	
	Public Property Let Navigators(aux)
		varNavigators = aux
	End Property

	Public Property Get QueryTransformations()
		QueryTransformations = varQueryTransformations
	End Property
	
	Public Property Let QueryTransformations(aux)
		varQueryTransformations = aux
	End Property
	
	Public Property Get Firsthit
		Firsthit = varFirsthit
	End Property
	
	Public Property Let Firsthit(nFIRSTHIT)
		varFirsthit = nFIRSTHIT
	End Property
	
	Public Property Get Lasthit
		Lasthit = varLasthit
	End Property
	
	Public Property Let Lasthit(nLASTHIT)
		varLasthit = nLASTHIT
	End Property
	
	Public Property Get Hits
		Hits = varHits
	End Property
	
	Public Property Let Hits(nHITS)
		varHits = nHITS
	End Property
	
	Public Property Get TotalHits
		TotalHits = varTotalHits
	End Property
	
	Public Property Let TotalHits(nTOTALHITS)
		varTotalHits = nTOTALHITS
	End Property
	
	Public Property Get Time
		Time = varTime
	End Property
	
	Public Property Let Time(nTIME)
		varTime = nTIME
	End Property
	
	Private Function addDocument( aux )
		Dim quantidadeNaLista
		quantidadeNaLista = Ubound( varDocuments )
		reDim Preserve varDocuments( quantidadeNaLista +1 )
		Set varDocuments( quantidadeNaLista ) = aux
	End Function

	Private Function addNavigator( aux )
		Dim quantidadeNaLista
		quantidadeNaLista = Ubound( varNavigators )
		reDim Preserve varNavigators( quantidadeNaLista +1 )
		Set varNavigators( quantidadeNaLista ) = aux
	End Function

	Private Function addQueryTransformation( aux )
		Dim quantidadeNaLista
		quantidadeNaLista = Ubound( varQueryTransformations )
		reDim Preserve varQueryTransformations( quantidadeNaLista +1 )
		Set varQueryTransformations( quantidadeNaLista ) = aux
	End Function

	Public Function parseServerResponse( xmlDocument ) 

		Set queryTransformationsEntries = xmlDocument.getElementsByTagName("QUERYTRANSFORM")
		If ( queryTransformationsEntries.length > 0 ) Then
			For i = 0 To (queryTransformationsEntries.length - 1)
				Set queryTransformationEntry = queryTransformationsEntries.item(i)
				
				Set objQueryTransformation = new QueryTransformation
				objQueryTransformation.Name = queryTransformationEntry.Attributes.GetNamedItem("NAME").Text
				objQueryTransformation.Action = queryTransformationEntry.Attributes.GetNamedItem("ACTION").Text
				objQueryTransformation.Query = queryTransformationEntry.Attributes.GetNamedItem("QUERY").Text
				objQueryTransformation.Custom = queryTransformationEntry.Attributes.GetNamedItem("CUSTOM").Text
				objQueryTransformation.Message = queryTransformationEntry.Attributes.GetNamedItem("MESSAGE").Text
				objQueryTransformation.MessageId = queryTransformationEntry.Attributes.GetNamedItem("MESSAGEID").Text
				
				addQueryTransformation( objQueryTransformation )
			Next
		End If
	
		Set navigatorsEntries = xmlDocument.getElementsByTagName("NAVIGATIONENTRY")
		If ( navigatorsEntries.length > 0 ) Then
			For i = 0 To (navigatorsEntries.length - 1)
				Set navigatorEntry = navigatorsEntries.item(i)
				
				Set objNavigator = new Navigator
				objNavigator.Name = navigatorEntry.Attributes.GetNamedItem("NAME").Text
				objNavigator.DisplayName = navigatorEntry.Attributes.GetNamedItem("DISPLAYNAME").Text
				
				Set navigatorItems = navigatorEntry.getElementsByTagName("NAVIGATIONELEMENT")
				For j =0 To ( navigatorItems.length - 1 )
					Set objNavigatorItem = new NavigatorItem
					Set Itens = navigatorItems.item(j)
					
					objNavigatorItem.Label = Itens.Attributes.GetNamedItem("NAME").Text
					objNavigatorItem.Value = Itens.Attributes.GetNamedItem("MODIFIER").Text
					objNavigatorItem.Count = Itens.Attributes.GetNamedItem("COUNT").Text
					objNavigator.addItem( objNavigatorItem )
				Next
				addNavigator( objNavigator )
			Next
		End If
		
		Dim docCacheUrl
		
		Set documentsEntries = xmlDocument.getElementsByTagName("HIT")
		
		If ( documentsEntries.length > 0 ) Then
			For i = 0 To ( documentsEntries.length - 1 )
				Set documentEntry = documentsEntries.item(i)
				
				'Objeto Document
				Set objDocument = new Document
				
				Set fields = documentEntry.getElementsByTagName("FIELD")

				docCacheUrl = ""

				Set re = new RegExp
				re.Pattern = "doccacheurl=\""(.*)\"""
				re.IgnoreCase = true					
				
				For j = 0 To ( fields.length - 1 )
					Set field = fields.item(j)

					Set Matches = re.execute ( field.Text )
					For Each Match in Matches
						If Len(docCacheUrl) = 0 Then
							docCacheUrl = Match.SubMatches(0)
						End If
					Next
					
					fieldName = field.Attributes.GetNamedItem("NAME").Text
					If ( fieldName <> "viewsourceurl" ) Then
						If ( fieldName <> "body") Then
							objDocument.addField field.Attributes.GetNamedItem("NAME").Text , field.Text
						Else
							If Not field.HasChildNodes Then
								objDocument.addField field.Attributes.GetNamedItem("NAME").Text , field.Text
							Else
								Dim value
								For k = 0 To ( field.ChildNodes.length - 1 )
									Set node = field.ChildNodes.item(k)
									value = value & node.Text
									If (node.nodeName = "sep") Then
										value = value & "..."
									End If
								Next
								objDocument.addField field.Attributes.GetNamedItem("NAME").Text , value
								value = ""
							End If	
						End If							
					Else
						If (Len(docCacheUrl) > 0 And Len(field.Text) > 0) Then
							objDocument.addField field.Attributes.GetNamedItem("NAME").Text , docCacheUrl
						Else	
							objDocument.addField field.Attributes.GetNamedItem("NAME").Text , field.Text
						End If	
					End If
				Next
				
				'Adiciona o Document na lista de Documents no QueryResult
				addDocument( objDocument )
			Next
		End If
		
		Set resultSets = xmlDocument.getElementsByTagName("RESULTSET")
		If ( resultSets.length > 0 ) Then
			For i = 0 To ( resultSets.length - 1 )
				Set resultSet = resultSets.item(i)
				
				Me.Firsthit 	= resultSet.Attributes.GetNamedItem("FIRSTHIT").Text
				Me.Lasthit 		= resultSet.Attributes.GetNamedItem("LASTHIT").Text
				Me.Hits 		= resultSet.Attributes.GetNamedItem("HITS").Text
				Me.TotalHits 	= resultSet.Attributes.GetNamedItem("TOTALHITS").Text
				Me.Time 		= resultSet.Attributes.GetNamedItem("TIME").Text
				
			Next
		End If	
	
	End Function
	
End Class
%>