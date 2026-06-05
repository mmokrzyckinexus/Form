namespace Form.Models
{
    public class HomeIndexViewModel
    {
        public AzureUserViewModel AzureUser { get; set; }
        public List<QuestionDefinitionModel> AvailableQuestions { get; set; } = new();
        public List<QuestionInstance> Questions { get; set; } = new();
    }
}
