using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentGenerator
{
	class LeafletConstrols
	{
		/// <summary>
		/// содержит все теги не находящиеся в списке
		/// </summary>
		public Dictionary<string, List<SdtElement>> GlobalElements;
		/// <summary>
		/// все списки найденные в документе
		/// </summary>
		public Dictionary<string, SdtRow> ListsElements;
	}
}
