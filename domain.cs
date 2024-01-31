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
using Newtonsoft.Json;

namespace appsvc_fnc_dev_OrgInfoList
{
    public static class domain
    {
        private class Domain
        {
            public string Legal_x0020_Title { get; set; }
            public string Appellation_x0020_l_x00e9_gale { get; set; }
            public string Abbr_x002e_ { get; set; }
            public string Abr_x00e9_v_x002e_ { get; set; }
            public string RG_x0020_Code { get; set; }
            public string GoCDomain { get; set; }

            public Domain(string TitleEn, string TitleFr, string abbreviationEn, string abbreviationFr, string RGCode, string _GoCDomain)
            {
                Legal_x0020_Title = TitleEn;
                Appellation_x0020_l_x00e9_gale = TitleFr;
                Abbr_x002e_ = abbreviationEn;
                Abr_x00e9_v_x002e_ = abbreviationFr;
                RG_x0020_Code = RGCode;
                GoCDomain = _GoCDomain;
            }
        }

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

            try
            {
                List<ListItem> itemList = new List<ListItem>();
                List<Domain> domains = new();

                var items = await graphAPIAuth.Sites[siteid].Lists[listid].Items.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=Legal_x0020_Title,Appellation_x0020_l_x00e9_gale,Abbr_x002e_,Abr_x00e9_v_x002e_,RG_x0020_Code,GoCDomain)" };
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

                foreach (var item in itemList)
                {
                    domains.Add(new Domain(GetValue(item, "Legal_x0020_Title"), GetValue(item, "Appellation_x0020_l_x00e9_gale"), GetValue(item, "Abbr_x002e_"), GetValue(item, "Abr_x00e9_v_x002e_"), GetValue(item, "RG_x0020_Code"), GetValue(item, "GoCDomain")));
                }

                return new OkObjectResult(JsonConvert.SerializeObject(domains));
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error when try get all domain {ex}");
                return new BadRequestObjectResult(ex);
            }
        }

        private static string GetValue(ListItem item, string fieldName)
        {
            return item.Fields.AdditionalData.ContainsKey(fieldName) ? item.Fields.AdditionalData[fieldName].ToString() : "";
        }
    }
}