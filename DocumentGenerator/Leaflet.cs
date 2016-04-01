using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentGenerator
{
	public class Leaflet
	{
		internal Dictionary<string, string> Globals;
		internal Dictionary<string, List<KeyValuePair<string, string>[]>> Lists;

		public Leaflet()
		{
			Globals = new Dictionary<string, string>();
			Lists = new Dictionary<string, List<KeyValuePair<string, string>[]>>();
		}

		public void AddGlobal(string tagName, string value)
		{
			Globals.Add(tagName, value);
		}

		public void AddLine(string listName, params KeyValuePair<string, string>[] items)
		{
			List<KeyValuePair<string, string>[]> lines;
			if (Lists.TryGetValue(listName, out lines))
			{
				var firstLine = lines[0];
				if (firstLine.Length != items.Length)
					throw new InvalidOperationException();  // ошибка если разное количество полей.

				for (int i = 0; i < firstLine.Length; ++i)
					if (string.CompareOrdinal(firstLine[i].Key, items[i].Key) != 0)
						throw new InvalidOperationException();      // ошибка если хоть одно поле отличается от другого
			}
			else
			{
				lines = new List<KeyValuePair<string, string>[]>();
				Lists.Add(listName, lines);
			}
			lines.Add(items);
		}
	}
}
