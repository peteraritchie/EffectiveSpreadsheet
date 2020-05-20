using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Pri.EffectiveSpreadsheet.Reader.Abstractions;
//using Pri.EffectiveSpreadsheet.Reader.Framework;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pri.EffectiveSpreadsheet.Reader
{
	public sealed partial class EffectiveWorkbook
	{
		private class RowAdapter : IIndexedRow
		{
			private readonly WorksheetPart worksheetPart;
			private readonly Row row;

			public RowAdapter(WorksheetPart worksheetPart, Row row, uint index)
			{
				this.worksheetPart = worksheetPart;// ?? throw new ArgumentNullException(nameof(worksheetPart));
				this.row = row;// ?? throw new ArgumentNullException(nameof(row));
				Index = index;
				//Debug.Assert(row.RowIndex.Value == index + 1);
				var cells = row.Descendants<Cell>();
			}

			public /*readonly*/ IAssociativelyCollective<string, ICell> Cells =>
				new PredicatingAssociativeCollection<string, ICell>(GetCellAdapters(),
					(cell, index) => cell.Location == index);

			private IEnumerable<CellAdapter> GetCellAdapters()
			{
				return row.Descendants<Cell>().Select(cell => new CellAdapter(worksheetPart, cell));
			}

			public ICell this[string index] => Cells[index];

			public IEnumerator<ICell> GetEnumerator()
			{
				return GetCellAdapters().GetEnumerator();
			}

			IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

			public uint Index { get; }
		}
	}
}
