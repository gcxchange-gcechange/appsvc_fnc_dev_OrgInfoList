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
using ListItem = Microsoft.Graph.Models.ListItem;
using Newtonsoft.Json;

namespace appsvc_fnc_dev_OrgInfoList
{
    public static class department
    {
        private class Department
        {
            public string Legal_x0020_Title { get; set; }
            public string Appellation_x0020_l_x00e9_gale { get; set; }
            public string Abbr_x002e_ { get; set; }
            public string Abr_x00e9_v_x002e_ { get; set; }
            public string RG_x0020_Code { get; set; }
            public string B2B { get; set; }
            public string ProBList { get; set; }

            public Department(string TitleEn, string TitleFr, string abbreviationEn, string abbreviationFr, string RGCode, string _B2B, string _ProBList)
            {
                Legal_x0020_Title = TitleEn;
                Appellation_x0020_l_x00e9_gale = TitleFr;
                Abbr_x002e_ = abbreviationEn;
                Abr_x00e9_v_x002e_ = abbreviationFr;
                RG_x0020_Code = RGCode;
                B2B = _B2B;
                ProBList = _ProBList;
            }
        }

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
                List<Department> departments = new();

                var items = await graphAPIAuth.Sites[siteid].Lists[listid].Items.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Expand = new string[] { "fields($select=Legal_x0020_Title,Appellation_x0020_l_x00e9_gale,Abbr_x002e_,Abr_x00e9_v_x002e_,RG_x0020_Code,B2B,ProBList)" };
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
                    departments.Add(new Department(GetValue(item, "Legal_x0020_Title"), GetValue(item, "Appellation_x0020_l_x00e9_gale"), GetValue(item, "Abbr_x002e_"), GetValue(item, "Abr_x00e9_v_x002e_"), GetValue(item, "RG_x0020_Code"), GetValue(item, "B2B"), GetValue(item, "ProBList")));
                }

                return new OkObjectResult(JsonConvert.SerializeObject(departments));
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error when try get department {ex}");
                return new BadRequestObjectResult(ex);

            }
        }

        private static string GetValue(ListItem item, string fieldName)
        {
            return item.Fields.AdditionalData.ContainsKey(fieldName) ? item.Fields.AdditionalData[fieldName].ToString() : "";
        }
    }
}