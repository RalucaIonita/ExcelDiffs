namespace ExcelDiffs.Models
{
    public class CustomError
    {
        public string Title { get; set; }
        public string Body { get; set; }

        public CustomError(string title, string body)
        {
            Title = title;
            Body = body;
        }

    }
}
