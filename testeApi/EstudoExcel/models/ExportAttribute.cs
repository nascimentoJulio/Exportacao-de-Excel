namespace EstudoExcel.models
{
  public class ExportAttribute : System.Attribute
  {
    public string ColumnName { get; set; }
    public bool DiscardColumn { get; set; }
  }
}