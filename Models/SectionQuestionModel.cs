namespace Form.Models
{
    public class SectionQuestionModel
    {
        public string Section { get; set; }
        public List<QuestionDefinitionModel> Questions { get; set; }
    }
}
