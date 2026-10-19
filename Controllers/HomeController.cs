using System.Diagnostics;
using Form.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Drives;
using Microsoft.Graph.Me;
using static Microsoft.Graph.Me.MeRequestBuilder;
using Microsoft.Kiota.Abstractions;
using Microsoft.Identity.Web;
using System.Threading.Tasks;
using Microsoft.Graph.DeviceManagement.DeviceConfigurations.Item.GetOmaSettingPlainTextValueWithSecretReferenceValueId;
using System.Text;
using Microsoft.Data.SqlClient;
using System.Text.Json;
using System.IO;
using System.Text.Json.Serialization;
using Microsoft.Graph.Models.Security;

namespace Form.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly GraphServiceClient _graphClient;
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _env;

        readonly List<SectionQuestionModel> AvailableQuestions = new();

        /*readonly List<SectionQuestionModel> AvailableQuestions = new()
        {
            new()
            {
                Section = "Lincolnshire Regulated Enterprise System Access",
                Questions = new List<QuestionDefinitionModel>
                {
                    new()
                    {
                        Label = "DicksonOne Environmentail Monitoring System",
                        Type = QuestionType.Checkbox,
                        Questions = new() { "Administrator", "Read Only", "Manage" }
                    },
                    new()
                    {
                        Label = "NEXUS-FMS02 (Oxygen Headspace): Lighthouse",
                        Type = QuestionType.SingleSelect,
                        Questions = new() { "Administrator", "QA/QC User", "R&D User" }
                    },
                    new()
                    {
                        Label = "LINCLAB-FTIR01 (Fourier-Transform Infrared Spectroscopy (FTIR)): ThermoFischer",
                        Type = QuestionType.SingleSelect,
                        Questions = new() { "Administrator", "User" }
                    }
                }
            },
            new()
            {
                Section = "Email Distribution Lists and Shared Drive Access - Check All That Apply:",
                Questions = new List<QuestionDefinitionModel>
                {
                    new()
                    {
                        Label = "Email Distribution Lists:",
                        Type = QuestionType.Text
                    },
                    new()
                    {
                        Label = "Shared Drives - Check All That Apply:",
                        Type = QuestionType.Checkbox,
                        Questions = new() { "H:", "H: (Human Resources)", "P: (Public Lab)" }
                    }
                }
            }
        };*/

        public HomeController(ILogger<HomeController> logger, GraphServiceClient graphClient, IConfiguration configuration, IWebHostEnvironment env)
        {
            _logger = logger;
            _graphClient = graphClient;
            _configuration = configuration;
            _env = env;

            try
            {
                var configured = _configuration.GetSection("AvailableQuestions").Get<List<SectionQuestionModel>>();
                if (configured != null && configured.Count > 0)
                {
                    AvailableQuestions.AddRange(configured);
                    _logger.LogDebug("Loaded {Count} sections from configuration.", configured.Count);
                }
                else
                {
                    _logger.LogWarning("No AvailableQuestions found in configuration (appsettings.json).");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to load AvailableQuestions from configuration.");
            }

            /*try
            {
                var candidate = new[]
                {
                    Path.Combine(_env.ContentRootPath, "questions.json")
                }.Where(p => !string.IsNullOrEmpty(p)).Distinct();

                string found = null;
                foreach (var path in candidate)
                {
                    _logger.LogInformation("Checking for questions.json at: {Path}", path);
                    if (System.IO.File.Exists(path))
                    {
                        found = path;
                        break;
                    }
                }
                if (found == null)
                {
                    _logger.LogError("questions.json not found in any expected location.");
                    throw new FileNotFoundException("questions.json not found.");
                }

                //var dataFile = Path.Combine(_env.ContentRootPath, "Data", "questions.json");
                if (System.IO.File.Exists(found))
                {
                    var jsonData = System.IO.File.ReadAllText(found);
                    _logger.LogInformation("questions.json size: {Length}", jsonData?.Length ?? 0);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    };

                    options.Converters.Add(new JsonStringEnumConverter());
                    List<SectionQuestionModel>? list = null;
                    try
                    {
                        list = JsonSerializer.Deserialize<List<SectionQuestionModel>>(jsonData, options);
                    }
                    catch (JsonException jsonEx)
                    {
                        _logger.LogError(jsonEx, "JSON deserialization error for questions.json.");
                        throw;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to load questions from JSON file.");
            }*/
        }


        [AuthorizeForScopes(ScopeKeySection = "Graph:Scopes")]
        public async Task<IActionResult> Index()
        {
            try
            {
                var connectionString = _configuration.GetConnectionString("DefaultConnection");
                using var connection = new SqlConnection(connectionString);
                connection.Open();
                ViewBag.DatabaseStatus = "Database connection successful.";
            }
            catch
            {
                _logger.LogError("Database connection failed.");
                ViewBag.DatabaseStatus = "Database connection failed.";
            }

            var user = await _graphClient.Me.GetAsync(requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Select = new[]
                {
                    "displayName",
                    "givenName",
                    "mail",
                    "userPrincipalName",
                    "surname",
                    "jobTitle",
                    "department",
                    "officelocation"
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

            var directReportsList = new List<Microsoft.Graph.Models.User>();
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
                    if (report is Microsoft.Graph.Models.User userReport) directReportsList.Add(userReport);
                }
            }

            /*var viewModel = new AzureUserViewModel
            {
                User = user,
                Manager = manager,
                DirectReports = directReportsList
            };*/

            var model = new HomeIndexViewModel
            {
                AzureUser = new AzureUserViewModel
                {
                    User = user,
                    Manager = manager,
                    DirectReports = directReportsList
                },
                AvailableQuestions = AvailableQuestions
            };

            return View(model);
        }

        [HttpGet]
        public async Task<IActionResult> UserDetails(string id)
        {
            try
            {
                var user = await _graphClient.Users[id].GetAsync(requestConfiguration =>
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
                    manager = await _graphClient.Users[id].Manager.GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Select = new[]
                        {
                            "id",
                            "displayName",
                            "userPrincipalName",
                            "givenName",
                            "surname",
                            "mail"
                        };
                    });

                }
                catch { }

                string managerName = null, managerEmail = null;
                if (manager is Microsoft.Graph.Models.User managerUser)
                {
                    managerName = managerUser.DisplayName;
                    managerEmail = managerUser.UserPrincipalName ?? managerUser.Mail;
                }
                else if (manager is OrgContact oc)
                {
                    managerName = oc.DisplayName;
                    managerEmail = oc.Mail;
                }

                return Ok(new
                {
                    userDisplayName = user.DisplayName,
                    userEmail = user.UserPrincipalName ?? user.Mail,
                    managerName = managerName,
                    managerEmail = managerEmail
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error fetching user information.");
                return StatusCode(500, new { Error = "An error occurred while fetching the user information." });
            }
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

        /*public IActionResult DatabaseStatus()
        {
            try
            {
                var connectionString = _configuration.GetConnectionString("DefaultConnection");
                using var connection = new SqlConnection(connectionString);
                connection.Open();

                return Ok(new { isConnected = true, Message = "Database connection successful." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Database connection failed.");
                return Ok(new { isConnected = false, Message = "Database connection failed." });
            }
        }*/
        /*public class QuestionAnswer
        {
            public string Label { get; set; }
            public string Target { get; set; }
            public string Value { get; set; }
            public List<string>? Values { get; set; }
        }

        public class PreviewRequest
        {
            public string Subject { get; set; }
            public string DirectReportId { get; set; }
            public List<QuestionAnswer> Questions { get; set; } = new();
            public string Filename { get; set; }
        }*/

        public class PreviewViewModel
        {
            public string Subject { get; set; }
            public string Name { get; set; }
            public string Email { get; set; }
            public string Bio { get; set; }
        }

        //[HttpPost]
        //[Route("Home/Preview")]
        /*public async Task<IActionResult> Preview([FromBody] PreviewRequest request)
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

            var finalFirst = baseLineFirst;
            var finalLast = baseLineLast;
            var finalEmail = baseLineEmail;
            var finalBio = string.Empty;

            if (request.Questions != null)
            {
                foreach (var question in request.Questions)
                {
                    if (string.IsNullOrEmpty(question.Value)) continue;
                    switch (question.Target)
                    {
                        case "firstName":
                            finalFirst = question.Value;
                            break;
                        case "lastName":
                            finalLast = question.Value;
                            break;
                        case "email":
                            finalEmail = question.Value;
                            break;
                        case "bio":
                            finalBio = question.Value;
                            break;
                    }
                }
            }

            var viewModel = new PreviewViewModel
            {
                Name = $"{finalFirst} {finalLast}".Trim(),
                Email = finalEmail,
                Bio = finalBio,
                Subject = request.Subject
            };

            return PartialView("_PreviewPartial", viewModel);
        }

        [HttpPost]
        [Route("Home/Save")]
        public async Task<IActionResult> Save([FromBody] PreviewRequest request)
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

            var finalFirst = baseLineFirst;
            var finalLast = baseLineLast;
            var finalEmail = baseLineEmail;
            var finalBio = string.Empty;

            if (request.Questions != null)
            {
                foreach (var question in request.Questions)
                {
                    if (string.IsNullOrEmpty(question.Value)) continue;
                    switch (question.Target)
                    {
                        case "firstName":
                            finalFirst = question.Value;
                            break;
                        case "lastName":
                            finalLast = question.Value;
                            break;
                        case "email":
                            finalEmail = question.Value;
                            break;
                        case "bio":
                            finalBio = question.Value;
                            break;
                    }
                }
            }

            var docHtml = new StringBuilder();
            docHtml.Append("<!doctype html><html><head><meta charset=\"utf-8\"><title>Form Save</title></head><body>");
            docHtml.Append($"<h2>{System.Net.WebUtility.HtmlEncode($"{finalFirst} {finalLast}".Trim())}</h2>");
            docHtml.Append($"<p><strong>Email:</strong> {System.Net.WebUtility.HtmlEncode(finalEmail ?? "")}</p>");
            docHtml.Append("<h3>Summary</h3>");
            docHtml.Append($"<div>{System.Net.WebUtility.HtmlEncode(finalBio ?? "").ReplaceLineEndings("\\n\", \"<br/>")}</div>");
            docHtml.Append("</body></html>");

            var bytes = Encoding.UTF8.GetBytes(docHtml.ToString());

            using var memoryStream = new MemoryStream(bytes);

            var fileName = string.IsNullOrEmpty(request.Filename) ? $"FormData_{DateTime.UtcNow:yyyyMMddHHmmss}.html" : $"{request.Filename}.html";
            var oneDrivePath = $"Desktop/{fileName}";

            try
            {
                memoryStream.Position = 0;
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
                var item = await _graphClient.Users[user.UserPrincipalName].Drive.GetAsync();
                var drive = await _graphClient.Drives[item.Id].Root.ItemWithPath(oneDrivePath).Content.PutAsync(memoryStream);

                return Ok(new { Success = true, Message = $"File '{fileName}' saved to OneDrive successfully.", FileId = item?.Id });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving file to OneDrive for user. Filename: {FileName}", fileName);
                return StatusCode(500, new { Success = false, Error = "An error occurred while saving the file to OneDrive." });
            }
        }*/
    }
}
