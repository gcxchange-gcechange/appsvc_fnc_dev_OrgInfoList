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
using Newtonsoft.Json;

namespace appsvc_fnc_dev_OrgInfoList
{
    public static class department
    {
        [FunctionName("Department")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
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

            //var queryOptions = new List<QueryOption>
            //{
            //    //new QueryOption("expand", "fields(select=Legal_x0020_Title,Appellation_x0020_l_x00e9_gale,Abbr_x002e_,Abr_x00e9_v_x002e_,RG_x0020_Code)"),
            //    new QueryOption("expand", "fields(select=Title)"),
            //};

            try
            {
                //List<ListItem> itemList = new List<ListItem>();


                // ffc90682-ea8a-4aff-9d28-eb5f49a2e458
                // ffc90682-ea8a-4aff-9d28-eb5f49a2e458

                //siteid = "ee706188-1a6b-4bff-be3b-1a04127809d2";

                //IListItemsCollectionPage items = await graphAPIAuth.Sites[siteid].Lists[listid].Items
                //.Request(queryOptions)
                //.GetAsync();

                var result = await graphAPIAuth.Sites[siteid].Lists[listid].Items.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=Legal_x0020_Title,Appellation_x0020_l_x00e9_gale,Abbr_x002e_,Abr_x00e9_v_x002e_,RG_x0020_Code)" };
                });

                //for (int i = 0; i < items.Count -1; i++)
                //{

                //   var item = items[i];
                //   log.LogInformation($"item = {item}");
                //}

                return new OkObjectResult(result);

                //itemList.AddRange(items.CurrentPage);

                //while (items.NextPageRequest != null)
                //{
                //    items = await items.NextPageRequest.GetAsync();
                //    itemList.AddRange(items.CurrentPage);
                //}

                //return new OkObjectResult(itemList);
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error when try get department {ex}");
                return new BadRequestObjectResult(ex);

            }
        }
    }
}
