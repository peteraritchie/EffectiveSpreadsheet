using System.Collections.Generic;

namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IAssociativelyCollective<in TKey1, in TKey2, out TValue> : IEnumerable<TValue>
	{
		TValue this[TKey1 index] { get; }
		TValue this[TKey2 index] { get; }
		// extension IReadOnlyDictionary<TKey, TValue> ToReadOnlyDictionary();
	}

	public interface IAssociativelyCollective<in TKey, out TValue> : IEnumerable<TValue>
	{
		TValue this[TKey index] { get; }
		// extension IReadOnlyDictionary<TKey, TValue> ToReadOnlyDictionary();
	}
}
