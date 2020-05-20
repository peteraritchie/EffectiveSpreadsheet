namespace Pri.EffectiveSpreadsheet.Reader
{
	public sealed partial class EffectiveWorkbook
	{
		//private class ColumnAdapter : IIndexedColumn
		//{
		//	private readonly WorksheetPart worksheetPart;
		//	private readonly Column column;

		//	public ColumnAdapter(WorksheetPart worksheetPart, Column column, uint index)
		//	{
		//		this.worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
		//		this.column = column ?? throw new ArgumentNullException(nameof(column));
		//		Index = index;
		//	}

		//	public ICell this[string index] => Cells[index];

		//	public /*readonly*/ IAssociativelyCollective<string, ICell> Cells =>
		//		new PredicateAssociativelyCollective<string, ICell>(GetCellAdapters(),
		//			(cell, index) => cell.Location == index);

		//	public IEnumerator<ICell> GetEnumerator()
		//	{
		//		return GetCellAdapters().GetEnumerator();
		//	}
		//	private IEnumerable<CellAdapter> GetCellAdapters()
		//	{
		//		return worksheetPart.Worksheet.Descendants<Cell>()
		//			.Where(e => e.GetColumn() == GetName())
		//			.Select(cell => new CellAdapter(worksheetPart, cell));
		//	}

		//	IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
		//	private string GetName()
		//	{
		//		uint div = column.Min.Value;
		//		string colLetter = String.Empty;
		//		uint mod = 0;

		//		while (div > 0)
		//		{
		//			mod = (div - 1) % 26;
		//			colLetter = (char)(65 + mod) + colLetter;
		//			div = (uint)((div - mod) / 26);
		//		}
		//		return colLetter;
		//	}

		//	public uint Index { get; }
		//}
	}
}
