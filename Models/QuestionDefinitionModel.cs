namespace Form.Models
{
    public enum QuestionType
    {
        Text,
        SingleSelect,
        MultiSelect,
        Checkbox
    }
    public class QuestionDefinitionModel
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Label { get; set; } = "";
        public string Target { get; set; } = "";
        public QuestionType Type { get; set; }
        public List<string>? Questions { get; set; }
    }

    public class QuestionInstance
    {
        public string QuestionId { get; set; }
        public string Label { get; set; }
        public string Target { get; set; }
        public QuestionType Type { get; set; }

        public string? Value { get; set; }
        public List<string>? Values { get; set; }

    }
}
