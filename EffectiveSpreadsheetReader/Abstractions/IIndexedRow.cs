namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IIndexedRow : IRow, IIndexed<uint>
	{
		IAssociativelyCollective<string, ICell> Cells { get; }
	}
}