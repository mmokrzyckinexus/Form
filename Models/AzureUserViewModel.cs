using Microsoft.Graph.Models;

namespace Form.Models
{
    public class AzureUserViewModel
    {
        public User User { get; set; }
        public DirectoryObject Manager { get; set; }
        public IList<User> DirectReports { get; set; }
        public IList<DirectoryObject> MembersOf { get; set; }
    }
}
