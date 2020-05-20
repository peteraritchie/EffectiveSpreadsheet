namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IMutablyAssociativelyCollective<in TKey, TValue> : IAssociativelyCollective<TKey, TValue>
	{
		new TValue this[TKey index] { get; set; }
		void Add(TKey key, TValue value);
		TValue Remove(TKey key);
	}
}
