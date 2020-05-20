namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface ISpreadsheet
	{
		IAssociativelyCollective<uint, IIndexedRow> Rows { get; }
		//IAssociativelyCollective<uint, IIndexedColumn> Columns { get; }
		IAssociativelyCollective<string, ICell> Cells { get; }
		string Name { get; }
	}
}