using DocumentFormat.OpenXml.Math;
using Pri.EffectiveSpreadsheet.Reader;
using Pri.EffectiveSpreadsheet.Reader.Abstractions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace Tests
{
	/// <summary>
	/// TODO: generate some XML via the SDK to reach some of the corner cases
	/// like boolean cell values.
	/// </summary>
	[Collection("Worksheet Open")]
	public class OpenTests
	{

		[Fact]
		public void TestOpen()
		{
			const string filename =
				//@"C:\local\documents\OneDrive - Knex\MortgageDataPointsStructuresEventsAndMilestones.xlsx";
				@"worksheet.xlsx";
			using (var workbook = new EffectiveWorkbook(filename))
			{
				Assert.NotNull(workbook);
				Assert.NotNull(workbook.Sheets);
			}
		}

		[Fact]
		public void Test1()
		{
			const string sheetName = "Data Dictionary";

			const string filename =
			//@"C:\local\documents\OneDrive - Knex\MortgageDataPointsStructuresEventsAndMilestones.xlsx";
			@"worksheet.xlsx";

			using (var stream = File.Open(filename, FileMode.Open, FileAccess.Read))
			{
				using (var effectiveBook = EffectiveWorkbook.Open(stream))
				{
					var sheets = effectiveBook.Sheets;
					Assert.NotNull(sheets);
					Assert.True(sheets.Count() > 1);

					var dataDictionarySheet = sheets[sheetName];
					Assert.NotNull(dataDictionarySheet);
					Assert.Equal(3u, ((IIndexedSpreadsheet)dataDictionarySheet).Index);

					var firstSheet = sheets.ElementAt(0);
					Assert.NotNull(firstSheet);
					Assert.Equal("Sheet1", firstSheet.Name);

					var topLeftCell = dataDictionarySheet.Cells["A1"];
					//Assert.NotNull(topLeftCell);
					//Assert.Equal("Unique ID", topLeftCell.ValueText);

					foreach (var sheet in sheets)
					{
						var rows = sheet.Rows;
						Assert.NotNull(rows);

						if (rows.Any())
						{
							var firstRow = rows.First();

							Assert.NotNull(firstRow);
							Assert.NotEmpty(firstRow);
						}

						foreach (var row in rows)
						{
							foreach (var cell in row)
							{
								Assert.NotNull(cell);

								string message = $"{sheet.Name} {cell.ValueText}";
								//Debug.WriteLine(message);
							}
						}
					}
				}
			}
		}
	}

	[Collection("Worksheet Open")]
	public class UnitTest1 : IClassFixture<WorksheetFixture>
	{
		readonly WorksheetFixture worksheetFixture;

		[Fact]
		public void GetValues()
		{
			List<string> values = new List<string>();
			var sheet = worksheetFixture.Workbook.Sheets["Data Dictionary"];
			Assert.Equal(6, sheet.Rows.Count());
			foreach (var row in sheet.Rows)
			{
				foreach (var cell in row.Cells)
				{
					values.Add(cell.ValueText);
				}
			}
			Assert.Equal(4 * 4 + 2, values.Count);
		}

		public UnitTest1(WorksheetFixture worksheetFixture)
		{
			this.worksheetFixture = worksheetFixture;
		}

		//[Fact]
		//public void NoColums()
		//{
		//	EffectiveWorkbook workbook = worksheetFixture.Workbook;
		//	var sheet = workbook.Sheets["Sheet4"];
		//	var columns = sheet.Rows;
		//	var enumerable = (IEnumerable)columns;
		//}

		[Fact]
		public void Values()
		{
			var sheet = worksheetFixture.Workbook.Sheets["Sheet4"];
			foreach (var cell in sheet.Cells)
			{ }
		}

		[Fact]
		public void ExplicitEnumerable()
		{
			// TODO: a spreadsheet may not contain columns despite having rows
			// assumption: the inverse is also true
			var sheet = worksheetFixture.Workbook.Sheets["Data Dictionary"];
			var rows = sheet.Rows;
			var enumerable = (IEnumerable)rows;

			Assert.NotNull(enumerable);
			Assert.NotNull(enumerable.GetEnumerator());
		}

		//[Fact]
		//public void Test4()
		//{
		//	// TODO: a spreadsheet may not contain columns despite having rows
		//	// assumption: the inverse is also true
		//	var sheet = worksheetFixture.Workbook.Sheets["Data Dictionary"];
		//	var columns = sheet.Columns;
		//	var column = columns.SingleOrDefault(e => e.Any() && e.First()?.ValueText == "Unique ID");

		//	Assert.NotNull(column);
		//}

		[Fact]
		public void RowIndex()
		{
			Assert.Equal("1", worksheetFixture.Workbook.Sheets["Data Dictionary"].Rows.First().First().RowIndex);
		}

		[Fact]
		public void ColumnIndex()
		{
			Assert.Equal("A", worksheetFixture.Workbook.Sheets["Data Dictionary"].Rows.First().First().ColumnIndex);
		}

		[Fact]
		public void EnumerateCells()
		{
			foreach (var row in worksheetFixture.Workbook.Sheets["Data Dictionary"].Rows)
			{
				foreach (ICell cell in row)
				{
					Assert.NotNull(cell);
					Assert.NotNull(cell.ValueText);
				}
			}
		}

		[Fact]
		public void Test3()
		{
			var cell = worksheetFixture.Workbook.Sheets["Data Dictionary"].Rows[1].Cells["A1"];

			Assert.NotNull(cell);
			Assert.Equal("A1", cell.Location);
			Assert.Equal("Unique ID", cell.ValueText);
		}

		[Fact]
		public void Test2()
		{
			foreach (var sheet in worksheetFixture.Workbook.Sheets)
			{
				_ = sheet;
			}

			var dataDictionarySheet = worksheetFixture.Workbook.Sheets["Data Dictionary"];
			Assert.NotNull(dataDictionarySheet); // unique id, form section name
			const string v = "Unique ID";

			//var t1 = dataDictionarySheet.Columns.Select(e => e.FirstOrDefault()?.ValueText);
			//var col = dataDictionarySheet.Columns.FirstOrDefault(e => e.FirstOrDefault()?.ValueText == v);

			IIndexedRow row = dataDictionarySheet.Rows[1];
			var cell = row["A1"];
			cell = row.Cells["A1"];
			Assert.NotNull(cell);

			foreach (var r in dataDictionarySheet.Rows)
			{
			}

		}
	}
	public sealed class WorksheetFixture : IDisposable
	{
		private static FileStream stream;
		private Lazy<EffectiveWorkbook> lazy = new Lazy<EffectiveWorkbook>(() =>
		{
			const string filename =
				//@"C:\local\documents\OneDrive - Knex\MortgageDataPointsStructuresEventsAndMilestones.xlsx";
				@"worksheet.xlsx";
			stream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
			return EffectiveWorkbook.Open(stream);
		});

		public EffectiveWorkbook Workbook => lazy?.Value;

		public void Dispose()
		{
			EffectiveWorkbook w;
			var s = stream;
			if (s != null)
			{
				w = Workbook;
				lazy = null;
				stream = null;
				using (s)
				using (w)
				{ /*intentionally blank*/}
			}
		}
	}
	public class WrappedEnumerable
	{
		private class SUT : WrappedEnumerable<int>
		{
			public SUT() : base(new List<int> {1,2,3,4,5 })
			{
			}
		}
		private class NullSUT : WrappedEnumerable<string>
		{
			public NullSUT() : base(null)
			{
			}
		}

		[Fact]
		public void ValueType()
		{
			var sut = new SUT();
			foreach (int i in sut)
			{
				Assert.True(i > 0);
				Assert.True(i < 6);
			}
		}
		[Fact]
		public void NullCTorEnumerableThrows()
		{
			Assert.Throws<ArgumentNullException>(() => new NullSUT());
		}
		[Fact]
		public void PredicatingAssociativeCollectionWithNullEnumerableThrows()
		{
			Assert.Throws<ArgumentNullException>(() 
				=> new PredicatingAssociativeCollection<string,string>(
					default, (string o,string t)=>t==o));
		}
		[Fact]
		public void PredicatingAssociativeCollectionWithNullPredicateThrows()
		{
			Assert.Throws<ArgumentNullException>(() 
				=> new PredicatingAssociativeCollection<string,string>(
					new[] { "a", "b" }, default));
		}
	}
}
