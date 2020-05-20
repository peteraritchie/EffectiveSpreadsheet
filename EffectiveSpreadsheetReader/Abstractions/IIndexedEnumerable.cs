using System.Collections.Generic;

namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IIndexedEnumerable<out TValue> : IEnumerable<TValue>
	{
		TValue this[int index] { get; }
	}
}
