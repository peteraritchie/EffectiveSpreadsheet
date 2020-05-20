namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface ICell
	{
		string ValueText { get; }

		string Location { get; }
		// reference
		// style
		// formula
		// comments
		string ColumnIndex { get; }
		string RowIndex { get; }
	}
}