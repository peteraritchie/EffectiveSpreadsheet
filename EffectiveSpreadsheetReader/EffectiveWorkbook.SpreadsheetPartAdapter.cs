using System.Collections.Generic;
using System.Linq;
using Pri.EffectiveSpreadsheet.Reader.Abstractions;
//using Pri.EffectiveSpreadsheet.Reader.Framework;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pri.EffectiveSpreadsheet.Reader
{
	public sealed partial class EffectiveWorkbook
	{
		private class SpreadsheetPartAdapter : IIndexedSpreadsheet
		{
			private readonly WorksheetPart worksheetPart;
			private readonly SheetData sheetData;
			private readonly IEnumerable<Cell> cells;

			public SpreadsheetPartAdapter(WorksheetPart worksheetPart, string name, uint index)
			{
				this.worksheetPart = worksheetPart;// ?? throw new ArgumentNullException(nameof(worksheetPart));
				Name = name;
				Index = index;
#if DEBUG
				sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
				cells = worksheetPart.Worksheet.Descendants<Cell>();
				if(cells.Any())
				{
					var cellText = worksheetPart.GetCellValue(sheetData.Elements<Row>().First().Elements<Cell>().First());
				}
#endif
			}

			public string Name { get; }
			public uint Index { get; }

			//public override string ToString()
			//{
			//	return Name;
			//}

			public IAssociativelyCollective<uint, IIndexedRow> Rows =>
				new PredicatingAssociativeCollection<uint, IIndexedRow>(
					sheetData.Elements<Row>().Select((row, index) => new RowAdapter(worksheetPart, row, row.RowIndex.Value)),
					(row, index) => row.Index == index);

			//public IAssociativelyCollective<uint, IIndexedColumn> Columns
			//{
			//	get
			//	{
			//		IEnumerable<Columns> columns = worksheetPart.Worksheet.Elements<Columns>();
			//		if (columns.Any())
			//		{
			//			return new PredicateAssociativelyCollective<uint, IIndexedColumn>(
			//			columns
			//			.FirstOrDefault()
			//			?.Select((column, index)
			//				=> new ColumnAdapter(worksheetPart, (Column)column, (uint)index)),
			//				(column, index) => column.Index == index);
			//		}
			//		return new PredicateAssociativelyCollective<uint, IIndexedColumn>(
			//			Enumerable.Empty<IIndexedColumn>(),
			//				(column, index) => column.Index == index);
			//	}
			//}

			public IAssociativelyCollective<string, ICell> Cells =>
				new PredicatingAssociativeCollection<string, ICell>(
					cells.Select(cell => new CellAdapter(worksheetPart, cell)),
					(cell, index) => cell.Location == index);
		}
	}
}
