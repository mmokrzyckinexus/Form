namespace Form.Models
{
    public class HomeIndexViewModel
    {
        public AzureUserViewModel AzureUser { get; set; }
        public List<SectionQuestionModel> AvailableQuestions { get; set; } = new();
        public List<QuestionInstance> Questions { get; set; } = new();
    }
}
