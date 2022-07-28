'Inventor API
Dim doc As AssemblyDocument = ThisDoc.Document

Dim bracketOcc As ComponentOccurrence = doc.ComponentDefinition.Occurrences.ItemByName("Flexible Bracket 1 Right")

Dim bracketCompDef As PartComponentDefinition = bracketOcc.Definition.Document.componentdefinition

Dim bracketFacDoc As PartDocument = bracketCompDef.FactoryDocument

Dim bracketModelStates As ModelStates = bracketFacDoc.ComponentDefinition.ModelStates

For Each foundModelState As ModelState In bracketModelStates
	If foundModelState.Name = CStr(Support1Distance)
		bracketOcc.ActiveModelState = foundModelState.Name
		'Component.ActiveModelState(bracketOcc.Name) = CStr(Support1Distance)
		Exit For
	End If
Next

'iLogic Only
'Component.ActiveModelState("Flexible Bracket 1 Right") = CStr(Support1Distance)