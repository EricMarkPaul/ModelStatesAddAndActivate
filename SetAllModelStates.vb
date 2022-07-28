''' <summary>
''' Program start
''' </summary>
Sub Main
	Dim doc As AssemblyDocument = ThisDoc.Document
	
	'Manage Support 1
	SetModelState(doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 1 Right"), Support1Distance)
	SetModelState(doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 1 Left"), Support1Distance)
	
	'Manage Support 2
	SetModelState(doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 2 Right"), Support2Distance)
	SetModelState(doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 2 Left"), Support2Distance)
	
	'Manage Support 3
	SetModelState(doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 3 Right"), Support3Distance)
	SetModelState(doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 3 Left"), Support3Distance)
	
	InventorVb.DocumentUpdate()
End Sub

''' <summary>
''' Set or create a model state for the given distance
''' Checks against the SupportDistance Parameter
''' </summary>
''' <param name="bracketOcc">ComponentOcurrence for the bracket</param>
''' <param name="distance">Used as both the model state name and "SupportDistance" parameter value</param>
Private Sub SetModelState(bracketOcc As ComponentOccurrence, distance As String)
	Dim stateToSet As ModelState

	Dim bracketCompDef As PartComponentDefinition = bracketOcc.Definition.Document.componentdefinition
	'This factory document is new!
	Dim bracketFacDoc As PartDocument = bracketCompDef.FactoryDocument
	
	'Check if the model state exists already by looking at names
	Dim bracketModelStates As ModelStates = bracketFacDoc.ComponentDefinition.ModelStates
	For Each modelState As ModelState In bracketModelStates
		If ModelState.name = distance Then
			bracketOcc.ActiveModelState = ModelState.name
			stateToSet = ModelState
			Exit For
		End If
	Next
	
	'If not model state was found, we'll create one as well as set the correct parameter value
	If stateToSet Is Nothing Then
		'Create a new model state
		stateToSet = bracketModelStates.Add(distance)
		'Activate the new model state before we update the user parameter as this is different per state
		stateToSet.Activate
		'Find the "SupportDistance" user parameter
		Dim foundParameter As UserParameter = (From param As UserParameter In bracketFacDoc.ComponentDefinition.Parameters.UserParameters
											   Where param.Name = "SupportDistance").FirstOrDefault
		foundParameter.Expression = distance
		
		'Finally set this new model state as active for the bracket occurrence we were given
		bracketOcc.ActiveModelState = stateToSet.Name
	End If
End Sub