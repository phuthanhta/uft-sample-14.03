'System config properties
Dim agrs
Set agrs = WScript.Arguments.Named
Dim test_path 'test path 
test_path = agrs.Item("run-path")
Dim result_path 'result path
result_path = agrs.Item("result-path")
Dim templates_path 'templates path 
templates_path = agrs.Item("templates-path")


'Define TestStep Class
Class TestStep
	Public description
	Public status
	Public order
End Class

'Define TestCase Class
Class TestCase
	Public test_suite
	Public parent_module
	Public execution_start_date
	Public execution_end_date
	Public name
	Public automation_content
	Public description
	Public testSteps
	Public status
	
	Public Sub Class_Initialize()
         Set testSteps = CreateObject("Scripting.Dictionary") 
     End Sub
End Class

Function writeXml(ByVal testCase)
	Dim xmlDoc, root, node
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	Set root = xmlDoc.createElement("TestCase")
	xmlDoc.appendChild root
	
	Set node = xmlDoc.createElement("test_suite")
	node.Text = testCase.test_suite
	root.appendChild node
	
	Set node = xmlDoc.createElement("parent_module")
	node.Text = testCase.parent_module 
	root.appendChild node
	
	Set node = xmlDoc.createElement("name")
	node.Text = testCase.name 
	root.appendChild node
	
	Set node = xmlDoc.createElement("execution_start_date")
	node.Text = testCase.execution_start_date 
	root.appendChild node
	
	Set node = xmlDoc.createElement("execution_end_date")
	node.Text = testCase.execution_end_date 
	root.appendChild node
	
	Set node = xmlDoc.createElement("description")
	node.Text = testCase.description 
	root.appendChild node
	
	Set node = xmlDoc.createElement("status")
	node.Text = testCase.status 
	root.appendChild node
	
	Set node = xmlDoc.createElement("automation_content")
	node.Text = testCase.automation_content 
	root.appendChild node
  
  Dim nodeSteps
  Set nodeSteps = xmlDoc.createElement("test_steps")
	root.appendChild nodeSteps
	
	For Each index In testCase.testSteps
		Dim tStep, tmpStep
		Set tmpStep = TestCase.testSteps(index)
		Set tStep = xmlDoc.createElement("test_step")
		nodeSteps.appendChild tStep
		
		Set node = xmlDoc.createElement("description")
		node.Text = tmpStep.description
		tStep.appendChild node
		
		Set node = xmlDoc.createElement("status")
		node.Text = tmpStep.status
		tStep.appendChild node	
		
		Set node = xmlDoc.createElement("order")
		node.Text = tmpStep.order
		tStep.appendChild node
	Next
	
	
	xmlDoc.save result_path & "/result-info.xml"
End Function

Dim App   'As Application
Set App = CreateObject("QuickTest.Application")

Dim tcTest

'launch App
App.Visible = False
App.Launch


'open test
App.Open test_path, False

Set tcTest = App.Test

tcTest.Settings.Run.IterationMode = "rngAll"

tcTest.Settings.Run.ObjectSyncTimeOut = 20000
tcTest.Settings.Run.DisableSmartIdentification = False
tcTest.Settings.Run.OnError = "NextStep"
tcTest.Settings.Web.BrowserNavigationTimeout = 60000

App.Options.Run.ViewResults = True
App.Options.Run.RunMode = "Fast"
App.Options.Run.ReportFormat = "RRV"


'read test actions
Dim tc
Set tc = New TestCase

tc.Name = tcTest.Name
tc.automation_content = test_path

Dim actions,action
Set actions = App.Test.Actions

'For i = 1 To actions.Count
'	Set action = actions(i)
'	Dim tStep
'	Set tStep = New TestStep	
'	tStep.description = action.Name
'	tc.testSteps.Add i, tStep	
'Next

'config run result
Dim qtResultsOpt
Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") 
qtResultsOpt.ResultsLocation = result_path

tcTest.Run qtResultsOpt

'collect result logs
tc.status = tcTest.LastRunResults.Status
Dim steps
Set steps = tc.testSteps

Dim xmlDoc, actNodes, summaryNode
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.load(result_path & "\Report\Results.xml")
Set actNodes = xmlDoc.selectNodes("Report/Doc/DIter/Action")


Set summaryNode = xmlDoc.selectSingleNode("Report/Doc/Summary")
tc.execution_start_date = summaryNode.getAttribute("sTime")
tc.execution_end_date = summaryNode.getAttribute("eTime")

For i = 0  To actNodes.length - 1
	Dim actNode, nodeargs, status, tmpStep, actionName, nameNode
	status = "PASSED"
	Set tmpStep = New TestStep
	Set actNode = actNodes(i)
	Set nodeargs = actNode.selectSingleNode("NodeArgs")
	Set nameNode = actNode.selectSingleNode("AName")
	actionName = nameNode.text
	status = nodeargs.getAttribute("status")
	tmpStep.description = actionName
	tmpStep.status = status
	tmpStep.order = i
	steps.Add i, tmpStep
Next

writeXml tc
Dim strErr
Function TransformXmlToHtml(ByVal inputXML, ByVal inputXSL, ByVal outputFile)
	sXMLLib = "Microsoft.XMLDOM"
	Set xmlDoc = CreateObject(sXMLLib)
	Set xslDoc = CreateObject(sXMLLib)
	
	xmlDoc.validateOnParse = True

	xslDoc.validateOnParse = True
	xmlDoc.async = False
	xslDoc.async = False
	
	xmlDoc.load inputXML

	If Err.Number <> 0 Then
		strErr = Err.Description & vbCrLf
		strErr = strErr & xmlDoc.parseError.reason & " line: " & xmlDoc.parseError.Line & " col: " & xmlDoc.parseError.linepos & " text: " & xmlDoc.parseError.srcText
		MsgBox strErr, vbCritical, "Error loading the Transform"
	End If

	xslDoc.load inputXSL

	If Err.Number <> 0 Then
		strErr = Err.Description & vbCrLf
		strErr = strErr & xslDoc.parseError.reason & " line: " & xslDoc.parseError.Line & " col: " & xslDoc.parseError.linepos & " text: " & xslDoc.parseError.srcText
		MsgBox strErr, vbCritical, "Error loading the Transform"
	End If

	outputText = xmlDoc.transformNode(xslDoc.documentElement)

	Set FSO = CreateObject("Scripting.FileSystemObject")

	Set outFile = FSO.CreateTextFile(outputFile,True)
	outFile.Write outputText
	outFile.Close


	Set outFile = Nothing
	Set FSO = Nothing
	Set xmlDoc = Nothing
	Set xslDoc = Nothing
	Set xmlResults = Nothing
End Function

Function CopyFiles(ByVal FiletoCopy,ByVal DestinationFolder)
   Dim fso
                Dim Filepath,WarFileLocation
                Set fso = CreateObject("Scripting.FileSystemObject")
                If  Right(DestinationFolder,1) <>"\"Then
                    DestinationFolder=DestinationFolder&"\"
                End If
    fso.CopyFile FiletoCopy,DestinationFolder,True
                FiletoCopy = Split(FiletoCopy,"\")

End Function

If Len(templates_path) <> 0 Then
	TransformXmlToHtml result_path & "\Report\Results.xml", templates_path & "\PDetails.xsl", result_path & "\Report\Results.html"
	CopyFiles templates_path & "\PResults.css", result_path & "\Report"
End If

App.Quit
Set App = Nothing