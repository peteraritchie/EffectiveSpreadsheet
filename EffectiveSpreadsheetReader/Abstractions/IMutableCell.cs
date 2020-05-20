namespace Pri.EffectiveSpreadsheet.Reader.Abstractions
{
	public interface IMutableCell : ICell
	{
		new string ValueText { get; set; }
	}
}