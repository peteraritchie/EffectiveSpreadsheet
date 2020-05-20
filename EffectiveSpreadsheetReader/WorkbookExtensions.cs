using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pri.EffectiveSpreadsheet.Reader
{
	internal static class WorkbookExtensions
	{
		//public static string GetDefinedName(this Cell cell)
		//{
		//	var x = cell.GetWorkbook().DefinedNames;

		//	return cell.Parent.LocalName;
		//}

		//public static Workbook GetWorkbook(this OpenXmlElement element)
		//{
		//	var worksheet = element.GetAncestor<Worksheet>();
		//	var workbookPart = worksheet.WorksheetPart.GetParentParts().First() as WorkbookPart;
		//	return workbookPart?.Workbook;
		//}

		//public static T GetAncestor<T>(this OpenXmlElement cell) where T : OpenXmlElement
		//{
		//	switch (cell.Parent)
		//	{
		//		case null:
		//			return null;
		//		case T ancestor:
		//			return ancestor;
		//	}

		//	return GetAncestor<T>(cell.Parent);
		//}

		public static string GetCellValue(this WorksheetPart wsPart, Cell theCell)
		{
			// If the cell represents an integer number, you are done. 
			// For dates, this code returns the serialized value that 
			// represents the date. The code handles strings and 
			// Booleans individually. For shared strings, the code 
			// looks up the corresponding value in the shared string 
			// table. For Booleans, the code converts the value into 
			// the words TRUE or FALSE.
			if (theCell.DataType != null &&
				(theCell.DataType == CellValues.SharedString || theCell.DataType == CellValues.Boolean))
			{
				var innerText = theCell.CellValue.InnerText;

				if (theCell.DataType.Value == CellValues.SharedString)
				{
					var wbPart = (WorkbookPart)wsPart.GetParentParts().First();

					// For shared strings, look up the value in the
					// shared strings table.
					var stringTable =
						wbPart.GetPartsOfType<SharedStringTablePart>()
							.FirstOrDefault();

					// If the shared string table is missing, something 
					// is wrong. Return the index that is in
					// the cell. Otherwise, look up the correct text in 
					// the table.
					if (stringTable == null) return innerText;

					// TODO: Exception instead of null?
					if (!int.TryParse(innerText, out var index)) return innerText;

					return stringTable.SharedStringTable
						.ElementAt(index).InnerText;
				}

				if (theCell.DataType.Value == CellValues.Boolean)
				{
					return innerText == "0" ? "FALSE" : "TRUE";
				}
			}

			return theCell.CellValue?.InnerText;
		}

		public static string GetColumn(this Cell cell)
		{
			return Regex.Match(GetLocation(cell), "^[A-Za-z]*").Value;
		}

		public static string GetRow(this Cell cell)
		{
			return Regex.Match(GetLocation(cell), "[0-9]*$").Value;
		}

		private static string GetLocation(Cell cell)
		{
			return cell.CellReference.Value;
		}
	}
}
