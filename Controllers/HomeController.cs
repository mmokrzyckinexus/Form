using System.Diagnostics;
using Form.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Drives;
using Microsoft.Graph.Me;
//using static Microsoft.Graph.Me.MeRequestBuilder;
using Microsoft.Kiota.Abstractions;
using Microsoft.Identity.Web;
using System.Threading.Tasks;

namespace Form.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly GraphServiceClient _graphClient;

        public HomeController(ILogger<HomeController> logger, GraphServiceClient graphClient)
        {
            _logger = logger;
            _graphClient = graphClient;
        }

        [AuthorizeForScopes(ScopeKeySection = "Graph:Scopes")]
        public async Task<IActionResult> Index()
        {
            var user = await _graphClient.Me.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new[]
                {
                    "displayName",
                    "givenName",
                    "mail",
                    "userPrincipalName",
                    "surname"
                };
            });

            DirectoryObject manager = null;
            try
            {
                manager = await _graphClient.Me.Manager.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new[]
                    { "id",
                      "displayName",
                      "userPrincipalName",
                      "givenName",
                      "surname",
                      "mail"
                    };
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching manager information: manager.");
            }

            var directReportsList = new List<User>();
            var directReport = await _graphClient.Me.DirectReports.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new[]
                {
                    "id",
                    "displayName",
                    "userPrincipalName"
                };
            });

            if (directReport?.Value != null)
            {
                foreach (var report in directReport.Value)
                {
                    if (report is User userReport) directReportsList.Add(userReport);
                }
            }

            var viewModel = new AzureUserViewModel
            {
                User = user,
                Manager = manager,
                DirectReports = directReportsList
            };

            return View(viewModel);
        }

        /*private static void ConfigurationRequest(RequestConfiguration<MeRequestBuilderGetRequestConfiguration> config)
        {
            config.QueryParameters.Select = new[] { "displayName", "givenName", "mail", "userPrincipalName" };
        }*/

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public class QuestionAnswer
        {
            public string Label { get; set; }
            public string Target { get; set; }
            public string Value { get; set; }
        }

        public class PreviewRequest
        {
            public string Subject { get; set; }
            public string DirectReportId { get; set; }
            public List<QuestionAnswer> Questions { get; set; } = new();
            public string Filename { get; set; }
        }

        public class PreviewViewModel
        {
            public string Subject { get; set; }
            public string Name { get; set; }
            public string Email { get; set; }
            public string Bio { get; set; }
        }

        [HttpPost]
        [Route("Home/Preview")]
        public async Task<IActionResult> Preview([FromBody] PreviewRequest request)
        {
            string baseLineFirst = string.Empty, baseLineLast = string.Empty, baseLineEmail = string.Empty;

            if (request?.Subject == "direct" && !string.IsNullOrEmpty(request.DirectReportId))
            {
                try
                {
                    var directReport = await _graphClient.Users[request.DirectReportId].GetAsync(config =>
                    {
                        config.QueryParameters.Select = new[]
                        {
                            "displayName",
                            "givenName",
                            "mail",
                            "userPrincipalName",
                            "surname"
                        };
                    });

                    if (directReport != null)
                    {
                        baseLineFirst = directReport.GivenName;
                        baseLineLast = directReport.Surname;
                        baseLineEmail = directReport.UserPrincipalName ?? directReport.Mail;
                    }
                }

                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error fetching direct report information for ID: {DirectReportId}", request.DirectReportId);
                }
            }

            if (string.IsNullOrEmpty(baseLineFirst) && string.IsNullOrEmpty(baseLineLast) && string.IsNullOrEmpty(baseLineEmail))
            {
                var currentUser = await _graphClient.Me.GetAsync(config =>
                {
                    config.QueryParameters.Select = new[]
                    {
                        "displayName",
                        "givenName",
                        "mail",
                        "userPrincipalName",
                        "surname"
                    };
                });

                baseLineFirst = !string.IsNullOrEmpty(baseLineFirst) ? baseLineFirst : currentUser.GivenName;
                baseLineLast = !string.IsNullOrEmpty(baseLineLast) ? baseLineLast : currentUser.Surname;
                baseLineEmail = !string.IsNullOrEmpty(baseLineEmail) ? baseLineEmail : (currentUser.UserPrincipalName ?? currentUser.Mail);
            }
        }
    }
}
