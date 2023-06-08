using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;

namespace appsvc_fnc_dev_OrgInfoList
{
    public static class department
    {
        [FunctionName("Department")]
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
            string listid = config["ListDepartment"];

            try
            {
                List<ListItem> itemList = new List<ListItem>();

                var items = await graphAPIAuth.Sites[siteid].Lists[listid].Items.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=Legal_x0020_Title,Appellation_x0020_l_x00e9_gale,Abbr_x002e_,Abr_x00e9_v_x002e_,RG_x0020_Code)" };
                });
                itemList.AddRange(items.Value);

                while (items.OdataNextLink != null)
                {
                    var nextPageRequestInformation = new RequestInformation
                    {
                        HttpMethod = Method.GET,
                        UrlTemplate = items.OdataNextLink
                    };

                    items = await graphAPIAuth.RequestAdapter.SendAsync(nextPageRequestInformation, (parseNode) => new ListItemCollectionResponse());
                    itemList.AddRange(items.Value);
                }

                return new OkObjectResult(itemList);
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error when try get department {ex}");
                return new BadRequestObjectResult(ex);

            }
        }
    }
}
