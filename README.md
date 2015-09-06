#JSON object class 2.2.2
##By RCDMK - rcdmk[at]hotmail[dot]com

###Licence:
MIT license: http://opensource.org/licenses/mit-license.php
The MIT License (MIT)
Copyright (c) 2012 RCDMK - rcdmk[at]hotmail[dot]com

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:  

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.  

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  

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
	Response.Write oJSON("prop3") & "<br>" ' default function is equivalent to `.Value(propName)`
	Response.Write oJSON("prop4") & "<br>"
	
	' get the JSON formatted output
	Dim jsonSting
	jsonString = oJSON.Serialize() ' this will contain the string representation of the JSON object
	
	oJSON.Write() ' this will write the output to the Response - equivalent to: Response.Write oJSON.Serialize()
	
	' load and parse some JSON formatted string
	jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""arrays"": [1, ""2"", 3.4, [5, 6, [7, 8]]], ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] } }" ' double quotes here because of the VBScript quote scaping
	
	oJSON.Parse(jsonString) ' set this to a variable if your string to load can be an Array, since the function returns the parsed object and arrays are parsed to JSONarray objects
	' if the string represents an object, the current object is returned so there is not need to set the return to a new variable
	
	oJSON.Write()
	
To load records from a database:
	
	' load records from an ADODB.Recordset
	dim cn, rs
	set cn = CreateObject("ADODB.Connection")
	cn.Open "yourConnectionStringGoesHere"
	
	set rs = cn.execute("SELECT id, nome, valor FROM pedidos ORDER BY id ASC")
	' this could also be:
	' set rs = CreateObject("ADODB.Recordset")
	' rs.Open "SELECT id, nome, valor FROM pedidos ORDER BY id ASC", cn	
	
	oJSON.LoadRecordset rs
	oJSONarray.LoadRecordset rs
	
	rs.Close
	cn.Close
	set rs = Nothing
	set cn = Nothing
	
	oJSON.Write() 		' outputs: {"data":[{"id":1,"nome":"nome 1","valor":10.99},{"id":2,"nome":"nome 2","valor":19.1}]}
	oJSONarray.Write() 	' outputs: [{"id":1,"nome":"nome 1","valor":10.99},{"id":2,"nome":"nome 2","valor":19.1}]
	
If you want to use arrays, I have something for you too

    ' instantiate the class
	Dim oJSONarray = New JSONarray
	
	' add something to the array
	oJSONarray.Push oJSON 	' Can be JSON objects, and even JSON arrays
	oJSONarray.Push 1.25 	' Can be numbers
	oJSONarray.Push "and strings too"
	
	' write to page
	oJSONarray.Write() ' Gess what? This does the same as the Write method from JSON object
	