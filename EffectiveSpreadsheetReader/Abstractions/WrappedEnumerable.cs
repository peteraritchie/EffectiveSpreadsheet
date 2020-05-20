using System;
using System.Collections;
using System.Collections.Generic;

namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public abstract class WrappedEnumerable<TValue> : IEnumerable<TValue>
	{
		protected readonly IEnumerable<TValue> enumerable;

		protected WrappedEnumerable(IEnumerable<TValue> enumerable)
		{
			this.enumerable = enumerable ?? throw new ArgumentNullException(nameof(enumerable));
		}

		public IEnumerator<TValue> GetEnumerator()
		{
			return enumerable.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return ((IEnumerable) enumerable).GetEnumerator();
		}
	}
}
