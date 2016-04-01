# Single leaflet #
## Prepare .docx for leaflets ##
Add text fields to your document: DEVELOPER -> Controls -> Rich Text Content Control.
Set control's tag value to name you will use later.

## Coding: ##
1) Create 'DocxGenerator':
<pre><code>var docx = new DocxGenerator(@"C:\Temp\MyTestTemplates\Hello Tag 1.docx");</code></pre>
2) Create 'Leaflet':
<pre><code>Leaflet leaflet = new Leaflet();
docx.AddLeaflet(leaflet);</code></pre>
3) Fill tags:
<pre><code>leaflet.AddGlobal("tag name", "value");</code></pre>
4) Save document:
<pre><code>docx.Generate(@"C:\Temp\MyTestTemplates\outFile");</code></pre>

# Multiple leaflets #
Do the same as for single leaflet, but create and add Leaflet several times. Example:
<pre><code>var docx = new DocxGenerator(@"C:\Temp\MyTestTemplates\Hello Tag 1.docx");

foreach (var leafletData in <<yourLeafletsSource>>)
{
  Leaflet leaflet = new Leaflet();
  leaflet.AddGlobal("Tag 1", "Henny");
  docx.AddLeaflet(leaflet);
}

docx.Generate(@"C:\Temp\MyTestTemplates\outFile");</code></pre>

# Registry #
## Prepare .docx for leaflets ##
1) Add table, add same tags to it.
2) Select line you are going to repeat for registry.
3) With line selected, click: DEVELOPER -> Controls -> Repeating Section Content Control.
4) Set tag value, for example 'MyList'.

## Coding: ##
1) Create 'DocxGenerator':
<pre><code>var docx = new DocxGenerator(@"C:\Temp\MyTestTemplates\Hello Tag 1.docx");</code></pre>
2) Create 'Leaflet':
<pre><code>Leaflet leaflet = new Leaflet();
docx.AddLeaflet(leaflet);</code></pre>
3) Fill tags for whole document:
<pre><code>leaflet.AddGlobal("tag name", "value");</code></pre>
4) Add row, where it is tagged as 'MyList':
<pre><code>leaflet.AddLine("MyList",
		new KeyValuePair<string, string>("Item1", (i + 1).ToString()),
		new KeyValuePair<string, string>("Item2", "Value 2"),
		new KeyValuePair<string, string>("Item3", "Value 3"),
		new KeyValuePair<string, string>("Item4", "Hello Registry!")
		);</code></pre>
5) Save document:
<pre><code>docx.Generate(@"C:\Temp\MyTestTemplates\outFile");</code></pre>
