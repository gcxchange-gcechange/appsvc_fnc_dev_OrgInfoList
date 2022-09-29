using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;

namespace appsvc_fnc_dev_OrgInfoList
{
    public static class domain
    {
        [FunctionName("Domain")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.System, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            IConfiguration config = new ConfigurationBuilder()

           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            string siteid = config["SiteId"];
            string listid = config["ListDomain"];

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("expand", "fields(select=Legal_x0020_Title,RG_x0020_Code,GoCDomain)")
            };
            
            try
            {
                IListItemsCollectionPage  items = await graphAPIAuth.Sites[siteid].Lists[listid].Items
                .Request(queryOptions)
                .GetAsync();
                return new OkObjectResult(items);
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error when try get all domain {ex}");
                return new BadRequestObjectResult(ex);
            }
        }
    }
}
