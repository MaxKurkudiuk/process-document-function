namespace ProcessDocumentFunction.Models.Excel;

public class VacationData(int rowIdx, string title, string date, string userName, string timeSpent)
{
  public int RowIdx { get; set; } = rowIdx;
  public string Title { get; set; } = title;
  public string Date { private get; set; } = date;
  public string UserName { get; set; } = userName;
  public string TimeSpent { get; set; } = timeSpent;
  public DateTime GetDate() => DateTime.Parse(Date);
  public float GetTimeSpent()
  {
    if (int.TryParse(TimeSpent, out var time))
      return time;
    return 0;
  }
}
