using Pri.EffectiveSpreadsheet.Reader.Abstractions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pri.EffectiveSpreadsheet.Reader
{
	public sealed partial class EffectiveWorkbook
	{
		private class CellAdapter : ICell
		{
			private readonly WorksheetPart worksheetPart;
			private readonly Cell cell;

			public CellAdapter(WorksheetPart worksheetPart, Cell cell)
			{
				this.worksheetPart = worksheetPart;// ?? throw new ArgumentNullException(nameof(worksheetPart));
				this.cell = cell;// ?? throw new ArgumentNullException(nameof(cell));
			}

			public string ValueText => worksheetPart.GetCellValue(cell);
			public string Location => cell.CellReference.Value;

			public string ColumnIndex => cell.GetColumn();

			public string RowIndex => cell.GetRow();
		}
	}
}
