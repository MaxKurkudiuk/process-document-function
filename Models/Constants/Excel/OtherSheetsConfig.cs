namespace ProcessDocumentFunction.Models.Constants.Excel;

public static class OtherSheetsConfig
{
  public static string[] IncorrectWorkTitles { get; set; } = [.. new string[]
  {
      "Vacation",
      "Illness",
      "Day off",
      "Sick leave"
  }.Select(x => x.ToLower().Replace(" ", ""))];
}
