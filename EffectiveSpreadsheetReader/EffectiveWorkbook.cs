using System.Collections.Generic;
using System.IO;
using System.Linq;
using Pri.EffectiveSpreadsheet.Reader.Abstractions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.CompilerServices;

[assembly:InternalsVisibleTo("Tests")]
namespace Pri.EffectiveSpreadsheet.Reader
{
	public sealed partial class EffectiveWorkbook : IEffectiveWorkbook
	{
		private readonly Stream stream;
		private readonly SpreadsheetDocument document;

		internal EffectiveWorkbook(string fileName)
			: this(File.Open(fileName, FileMode.Open))
		{
		}

		private EffectiveWorkbook(Stream stream)
		{
			this.stream = stream;
			document = SpreadsheetDocument.Open(stream, stream.CanWrite);
		}

		public static EffectiveWorkbook Open(Stream stream)
		{
			return new EffectiveWorkbook(stream);
		}

		public IAssociativelyCollective<string, ISpreadsheet> Sheets
		{
			get
			{
				var workbookPart = document.WorkbookPart;
				IEnumerable<SpreadsheetPartAdapter> enumerable = workbookPart.WorksheetParts.OrderBy(w => w.Uri.ToString()).Select((worksheetPart, index) =>
				  {
					  var single = workbookPart.Workbook.Sheets.Cast<Sheet>()
						  .Single(sheet => sheet.Id == workbookPart.GetIdOfPart(worksheetPart));
					  return new SpreadsheetPartAdapter(worksheetPart, single.Name, (uint)index);
				  });
				return new PredicatingAssociativeCollection<string, ISpreadsheet>(enumerable,
					(worksheet, name) => worksheet.Name == name);
			}
		}

		public void Dispose()
		{
			document.Dispose();
			stream.Dispose();
		}
	}
}
