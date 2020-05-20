namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IIndexed<out T>
	{
		T Index { get; }
	}
}
