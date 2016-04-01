using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

/*

How to use:

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

 */

namespace DocumentGenerator
{
	public class DocxGenerator
	{
		readonly List<Leaflet> _leaflets;
		readonly string _templatePath;

		readonly int _maxSize;
		Dictionary<string, byte[]> _files;
		bool _updateFields;

		public DocxGenerator(string templatePath, int maxSize = 5000)
		{
			_templatePath = templatePath;
			_maxSize = maxSize;
			_leaflets = new List<Leaflet>();
		}

		public void AddLeaflet(Dictionary<string, string> tags)
		{
			Leaflet leaflet = new Leaflet();
			leaflet.Globals = tags;
			AddLeaflet(leaflet);
		}

		public void AddLeaflet(Leaflet leaflet)
		{
			CheckLeafletData(leaflet);
			_leaflets.Add(leaflet);
		}

		private void CheckLeafletData(Leaflet leaflet)
		{
			if (!_leaflets.Any())
				return;

			const string errText = "All cards must have the same tags";

			var firstLeaflet = _leaflets[0];
			if (firstLeaflet.Globals.Count != leaflet.Globals.Count)
				throw new InvalidOperationException(errText);

			foreach (var newCardField in leaflet.Globals)
				if (!firstLeaflet.Globals.ContainsKey(newCardField.Key))
					throw new InvalidOperationException(errText);
		}

		public void GenerateFiles(string fileName)
		{
			_files = new Dictionary<string, byte[]>();

			int pratNum = 0;
			foreach (var leaflets in SplitLeaflets())
			{
				++pratNum;
				string postFix = string.Empty;
				if (_leaflets.Count > _maxSize)
					postFix = "_p" + pratNum.ToString();
				string fullName = fileName + postFix + ".docx";

				var buffer = File.ReadAllBytes(_templatePath);
				MemoryStream stream;
				using (stream = new MemoryStream())
				{
					stream.Write(buffer, 0, buffer.Count());
					Generate(stream, leaflets);
				}
				_files.Add(fullName, stream.ToArray());
			}
		}


		private List<List<Leaflet>> SplitLeaflets()
		{
			List<List<Leaflet>> parts = new List<List<Leaflet>>();
			List<Leaflet> part = null;

			for (int i = 0; i < _leaflets.Count; ++i)
			{
				var card = _leaflets[i];

				if (part == null || part.Count == _maxSize)
				{
					part = new List<Leaflet>();
					parts.Add(part);
				}

				part.Add(card);
			}

			return parts;
		}

		private void Generate(Stream stream, List<Leaflet> part)
		{
			using (var wordprocessingDocument = WordprocessingDocument.Open(stream, true))
			{
				var docPart = wordprocessingDocument.MainDocumentPart;
				// Assign a reference to the existing document body.
				Body body = docPart.Document.Body;

				List<OpenXmlElement> template = ReadTemplate(body.ChildElements);
				SectionProperties secProperties = ReadTemplateEnd(body.ChildElements);
				body.RemoveAllChildren();

				for (int i = 0; i < part.Count; ++i)
				{
					Leaflet leaflet = part[i];
					var newLeafletNodes = template.Select(element => element.CloneNode(true)).ToList();
					var tags = GetTags(newLeafletNodes);
					var controls = GetControls(tags);

					// prepare leaflet data
					ApplyLeafletToTemplate(leaflet, controls);

					// add leaflet
					foreach (OpenXmlElement element in newLeafletNodes)
						body.AppendChild(element);

					// insert separator
					OpenXmlElement p;

					if (i + 1 == part.Count)
						p = secProperties;
					else
						p = new Paragraph(new ParagraphProperties(secProperties.CloneNode(true)));

					body.AppendChild(p);
				}

				FixParams(body);
				UpdateFields(docPart.DocumentSettingsPart.Settings);
				UpdateParts(docPart, part);
			}
		}

		/// <summary>
		/// Update data in header and footer if only one leaflet exists.
		/// Otherwise do nothing
		/// </summary>
		private void UpdateParts(MainDocumentPart docPart, List<Leaflet> parts)
		{
			if (parts.Count != 1)
				return;
			Leaflet leaflet = parts[0];

			// update footers
			foreach (var footerPart in docPart.FooterParts)
			{
				var tags = GetTags(footerPart.RootElement.ToList());
				var controls = GetControls(tags);
				ApplyGlobalTags(leaflet, controls);
			}

			// update headers
			foreach (var headerPart in docPart.HeaderParts)
			{
				var tags = GetTags(headerPart.RootElement.ToList());
				var controls = GetControls(tags);
				ApplyGlobalTags(leaflet, controls);
			}
		}

		private void UpdateFields(Settings settings)
		{
			if (!_updateFields)
				return;

			// Create object to update fields on open
			UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
			updateFields.Val = new OnOffValue(true);

			// Insert object into settings part.
			settings.PrependChild(updateFields);
			settings.Save();
		}

		private static Dictionary<string, List<SdtElement>> GetTags(List<OpenXmlElement> elements)
		{
			List<SdtElement> sdtBlocks = elements.SelectMany(ee => ee.Descendants<SdtBlock>()).Cast<SdtElement>().ToList();
			sdtBlocks.AddRange(elements.SelectMany(ee => ee.Descendants<SdtRun>()));

			sdtBlocks.AddRange(elements.OfType<SdtBlock>());
			sdtBlocks.AddRange(elements.OfType<SdtRun>());

			var outTags = new Dictionary<string, List<SdtElement>>();

			foreach (var sdt in sdtBlocks)
			{
				SdtProperties sdtProperties = sdt.SdtProperties;
				if (sdtProperties == null)
					continue;

				Tag tag = sdtProperties.GetFirstChild<Tag>();
				if (tag == null)
					continue;

				OpenXmlCompositeElement content;

				if (sdt is SdtBlock)
					content = (sdt as SdtBlock).SdtContentBlock;
				else if (sdt is SdtRun)
					content = (sdt as SdtRun).SdtContentRun;
				else
					throw new InvalidOperationException("Unknown type of Sdt = '" + sdt.GetType().Name + "'");


				if (content == null || !content.Any())
					throw new NoNullAllowedException("Поле не может быть пустым");

				var first = content.First(x => !x.LocalName.StartsWith("bookmark"));
				content.RemoveAllChildren();
				content.AppendChild(first);

				sdtProperties.RemoveAllChildren<SdtPlaceholder>();

				var tagName = tag.Val;

				List<SdtElement> item;
				if (!outTags.TryGetValue(tagName, out item))
				{
					item = new List<SdtElement>();
					outTags.Add(tagName, item);
				}
				item.Add(sdt);
			}

			return outTags;
		}

		private static LeafletConstrols GetControls(Dictionary<string, List<SdtElement>> tags)
		{
			LeafletConstrols constrols = new LeafletConstrols();
			constrols.GlobalElements = new Dictionary<string, List<SdtElement>>();
			constrols.ListsElements = new Dictionary<string, SdtRow>();

			foreach (var tag in tags)
			{
				foreach (var element in tag.Value)
				{
					SdtRow listElement = GetRepeatedSection(element);

					if (listElement != null)
					{
						Tag tagNode = (Tag)listElement.SdtProperties.First(x => x is Tag);
						if (tagNode == null)
							throw new InvalidOperationException("Tag name not spesified");
						string listName = tagNode.Val;
						if (!constrols.ListsElements.ContainsKey(listName))
							constrols.ListsElements.Add(listName, listElement);
					}

					List<SdtElement> elements;
					if (!constrols.GlobalElements.TryGetValue(tag.Key, out elements))
					{
						elements = new List<SdtElement>();
						constrols.GlobalElements.Add(tag.Key, elements);
					}

					elements.Add(element);
				}
			}

			return constrols;
		}

		private static SdtRow GetRepeatedSection(OpenXmlElement element)
		{
			while (element != null)
			{
				SdtRow tmp = element as SdtRow;
				if (tmp != null)
				{
					var repeatedSection = tmp.SdtProperties.FirstOrDefault(x => x is DocumentFormat.OpenXml.Office2013.Word.SdtRepeatedSection);
					if (repeatedSection != null)
						return tmp;
				}

				element = element.Parent;
			}

			return null;
		}

		private static void FixParams(Body body)
		{
			var docPrs = body.Descendants<DocProperties>();

			uint id = 0;
			foreach (var docPr in docPrs)
				docPr.Id = ++id;
		}

		private static List<OpenXmlElement> ReadTemplate(OpenXmlElementList childElements)
		{
			List<OpenXmlElement> elements = new List<OpenXmlElement>();

			foreach (OpenXmlElement element in childElements.ToArray())     // ToArray needs because i delete elements from it, and it can break looping
			{
				// delete all 'proofErr'
				foreach (var proofErr in element.Descendants<ProofError>().ToArray())
					proofErr.Remove();

				// delete all '<w:placeholder>'
				foreach (var placeholder in element.Descendants<SdtPlaceholder>().ToArray())
					placeholder.Remove();

				if (childElements.Count != elements.Count + 1)
					elements.Add(element);
			}

			return elements;
		}

		private static SectionProperties ReadTemplateEnd(OpenXmlElementList childElements)
		{
			var element = childElements[childElements.Count - 1];

			SectionProperties secProperties = element as SectionProperties;
			if (secProperties == null)
				throw new InvalidOperationException("Invalid last tag name. Must be 'sectPr'.");

			secProperties.ClearAllAttributes();
			secProperties.AppendChild(new PageNumberType { Start = 1 });

			return secProperties;
		}

		private void ApplyLeafletToTemplate(Leaflet leaflet, LeafletConstrols controls)
		{
			ApplyGlobalTags(leaflet, controls);
			foreach (var list in leaflet.Lists)
				FillList(controls, list.Key, list.Value);
		}

		private void ApplyGlobalTags(Leaflet leaflet, LeafletConstrols constrols)
		{
			foreach (var tag in leaflet.Globals)
			{
				List<SdtElement> elements;
				if (!constrols.GlobalElements.TryGetValue(tag.Key, out elements))
					continue;

				foreach (var element in elements)
					ApplyTextToElement(element, tag.Value);
			}
		}

		private void FillList(LeafletConstrols constrols, string listName, List<KeyValuePair<string, string>[]> lines)
		{
			SdtRow controlList;
			if (!constrols.ListsElements.TryGetValue(listName, out controlList))
				return;

			// запомнить шаблон строки и убрать из документа
			var listContent = controlList.Parent;
			var rowTemplate = controlList.SdtContentRow.OfType<SdtRow>().First().SdtContentRow;

			foreach (var line in lines)
			{
				var newRow = rowTemplate.CloneNode(true);
				var tags = GetTags(newRow.ChildElements.ToList());

				// заполнить новую строку
				foreach (var item in line)
				{
					List<SdtElement> textTags;
					if (!tags.TryGetValue(item.Key, out textTags))
						continue;

					foreach (var textTag in textTags)
						ApplyTextToElement(textTag, item.Value);
				}

				// склонировать в документ
				var newRowElements = newRow.Elements().ToArray();
				newRow.RemoveAllChildren();
				foreach (var element in newRowElements)
					controlList.InsertBeforeSelf(element);
			}
			listContent.RemoveChild(controlList);
		}

		private void ApplyTextToElement(SdtElement element, string newText)
		{
			//element
			var texts = element.Descendants<TextType>().ToArray();

			TextType text = null;
			foreach (var t in texts)
			{
				if (text == null)
					text = t;
				else
					t.Remove();
			}

			if (text == null)
				throw new InvalidOperationException("Text field must have at least one Text element");

			if (text is FieldCode)
				_updateFields = true;

			var block = element.OfType<SdtContentBlock>().FirstOrDefault();
			if (block == null)
			{
				text.Text = newText;
			}
			else
			{
				Paragraph p = block.OfType<Paragraph>().First();

				string[] textParts = newText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
				for (int i = 0; i < textParts.Length; ++i)
				{
					text.Text = textParts[i];
					Paragraph newP = (Paragraph)p.CloneNode(true);
					p.InsertBeforeSelf(newP);
				}
				p.Remove();
			}



			if (element.Parent == null)
			{
				// element.InsertBeforeSelf(child); thows an exception
				// because internaly it use parrent property.
				// TODO: maybe there is a better way
				return;
			}

			OpenXmlCompositeElement content = element.OfType<SdtContentRun>().FirstOrDefault();
			if (content == null)
				content = element.OfType<SdtContentBlock>().First();
			var children = content.ChildElements.ToArray();
			content.RemoveAllChildren();
			foreach (var child in children)
				element.InsertBeforeSelf(child);
			element.Remove();
		}

		public Dictionary<string, byte[]> GetFilesAsBytes()
		{
			return _files;
		}

		public List<string> GetFilesAsPaths(string path)
		{
			var files = new List<string>(_files.Count);

			foreach (var file in _files)
			{
				string filePath = Path.Combine(path, file.Key);
				File.WriteAllBytes(filePath, file.Value);
				files.Add(filePath);
			}

			return files;
		}
	}
}
