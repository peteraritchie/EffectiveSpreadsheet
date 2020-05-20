
namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IEffectiveWorkbook : System.IDisposable
	{
		IAssociativelyCollective<string, ISpreadsheet> Sheets { get; }
	}
}