using Pri.EffectiveSpreadsheet.Reader.Abstractions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Pri.EffectiveSpreadsheet.Reader
{
	public class PredicatingAssociativeCollection<TKey, TValue> : WrappedEnumerable<TValue>,
		IAssociativelyCollective<TKey, TValue>
	{
		private readonly Func<TValue, TKey, bool> predicate;

		public PredicatingAssociativeCollection(IEnumerable<TValue> enumerable, Func<TValue, TKey, bool> predicate)
			: base(enumerable)
		{
			this.predicate = predicate ?? throw new ArgumentNullException(nameof(predicate));
		}

		public TValue this[TKey index] => enumerable.SingleOrDefault(e => predicate(e, index));
	}

	//public class PredicatedAssociativeCollection<TKey1, TKey2, TValue> : WrappedEnumerable<TValue>,
	//	IAssociativelyCollective<TKey1, TKey2, TValue>
	//{
	//	private readonly Func<TValue, TKey1, bool> predicate1;
	//	private readonly Func<TValue, TKey2, bool> predicate2;

	//	public PredicatedAssociativeCollection(IEnumerable<TValue> enumerable, Func<TValue, TKey1, bool> predicate1,
	//		Func<TValue, TKey2, bool> predicate2)
	//		: base(enumerable)
	//	{
	//		this.predicate1 = predicate1 ?? throw new ArgumentNullException(nameof(predicate1));
	//		this.predicate2 = predicate2 ?? throw new ArgumentNullException(nameof(predicate2));
	//	}

	//	public TValue this[TKey1 index] => enumerable.SingleOrDefault(e => predicate1(e, index));
	//	public TValue this[TKey2 index] => enumerable.SingleOrDefault(e => predicate2(e, index));
	//}
}