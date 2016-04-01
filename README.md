#How to use:

var docx = new DocxGenerator(@"C:\Temp\MyTestTemplates\Hello Tag 1.docx");

Leaflet leaflet = new Leaflet();

leaflet.AddGlobal("Tag 1", "Henny");

for (int i = 0; i < 10; ++i)
{
	leaflet.AddLine("FirstList",
		new KeyValuePair<string, string>("fName", "Станислав"),
		new KeyValuePair<string, string>("sName", "Владимирович"),
		new KeyValuePair<string, string>("pName", "Березовиков")
		);
}

for (int i = 0; i < 50; ++i)
{
	leaflet.AddLine("List",
		new KeyValuePair<string, string>("Item1", (i + 1).ToString()),
		new KeyValuePair<string, string>("Item2", "1Текст первой калонки"),
		new KeyValuePair<string, string>("Item3", "1Текст 3-й калонки"),
		new KeyValuePair<string, string>("Item4", "1Текст 4-й калонки")
		);
}

docx.AddLeaflet(leaflet);
docx.AddLeaflet(leaflet);



docx.Generate(@"C:\Temp\MyTestTemplates\outFile");
