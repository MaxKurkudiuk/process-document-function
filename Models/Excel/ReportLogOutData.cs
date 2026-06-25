namespace ProcessDocumentFunction.Models.Excel;

public class ReportLogOutData
{
  public string SheetName { get; set; } = null!;
  public object[,] RawData { get; set; } = null!;
  public int HeaderIndex { get; set; } = -1;
  public List<string> Headers { get; set; } = null!;
  public List<string> MissingHeaders { get; set; } = null!;
  public Dictionary<string, int> HeaderColumnsDictionary { get; set; } = null!;

  public int GetColumnNumberByName(string name) 
  => HeaderColumnsDictionary.FirstOrDefault(x => x.Key.Equals(name, StringComparison.CurrentCultureIgnoreCase)).Value;

  public bool IsValid { get; set; } = true;
  public bool IsStats { get; set; } = false;
}
