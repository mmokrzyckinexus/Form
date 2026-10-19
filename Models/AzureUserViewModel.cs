using Microsoft.Graph.Models;

namespace Form.Models
{
    public class AzureUserViewModel
    {
        public required User User { get; set; }
        public required DirectoryObject Manager { get; set; }
        public IList<User>? DirectReports { get; set; }
        public IList<DirectoryObject>? MembersOf { get; set; }
    }
}
