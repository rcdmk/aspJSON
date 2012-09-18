#JSON object class 2.0a - September, 17th - 2012
##By RCDMK - rcdmk@rcdmk.com

###Licence:
Creative Commons BY: http://creativecommons.org/licenses/by/3.0/
You are free to use, share, distribute or change this work, as long as you mantain a reference to the author in this file and, when applicable, in a readme, about or some form of credit screen or dialog on the application where this code is being used if the source of the app is not distributed.


###How to use:

<!-- languages: vbscript, vb -->

    ' instantiate the class
	Dim oJSON = New JSON
	
	' add properties
	oJSON.Add "prop1", "someString"
	oJSON.Add "prop2", 12.3
	oJSON.Add "prop3", Array(1, 2, "three")
	
	' change some values
	oJSON.Change "prop1", "someOtherString"
	oJSON.Change "prop4", "thisWillBeCreated" ' this property doen't exists and will be created automagically
	
	' get the values
	Response.Write oJSON.Value("prop1") & "<br>"
	Response.Write oJSON.Value("prop2") & "<br>"
	Response.Write oJSON.Value("prop3") & "<br>"
	Response.Write oJSON.Value("prop4") & "<br>"
	
	' get the JSON formatted output
	Dim jsonSting
	jsonString = oJSON.Serialize() ' this will contain the string representation of the JSON object
	
	oJSON.Write() ' this will write the output to the Response - equivalent to: Response.Write oJSON.Serialize()
	
	
	' load and parse some JSON formatted string
	jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""arrays"": [1, ""2"", 3.4, [5, 6, [7, 8]]], ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] } }" ' double quotes here because of the VBScript quote scaping
	
	oJSON.Parse jsonString
	
	oJSON.Write()
	