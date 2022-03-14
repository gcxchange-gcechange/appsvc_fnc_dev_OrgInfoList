using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using System;
using Microsoft.Graph;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Http;

namespace appsvc_fnc_dev_OrgInfoList
{

    public interface IGraphClientWrapper
    {
        Task<object> Department(String siteid, String listid);
    }

    public class GraphClientMock : IGraphClientWrapper
    {
        private readonly String _result;

        public GraphClientMock(String result)
        {
            _result = result;
        }

        public async Task<object> Department(String siteid, String listid)
        {
            var mockResult = Task<object>.Run(() => { return _result; });
            return await mockResult;
        }
    }

    public static class Department
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

            string siteid = config["SiteId"];
            string listid = config["ListDepartment"];

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            try
            {
                await getDepartment(new GraphClientWrapper(graphAPIAuth), siteid, listid, log);
            }
            catch (ServiceException e)
            {
                return new BadRequestObjectResult("E1BadRequest");
            }

            return new OkObjectResult("Finished");
        }

        private class GraphClientWrapper : IGraphClientWrapper
        {
            private readonly GraphServiceClient _graphClient;

            public GraphClientWrapper(GraphServiceClient graphClient)
            {
                _graphClient = graphClient;
            }

            public async Task<object> Department(String siteid, String listid)
            {
                var queryOptions = new List<QueryOption>()
            {
                new QueryOption("expand", "fields(select=Legal_x0020_Title,Appellation_x0020_l_x00e9_gale,Abbr_x002e_,Abr_x00e9_v_x002e_,RG_x0020_Code)")
            };
                return await _graphClient.Sites[siteid].Lists[listid].Items
                .Request(queryOptions)
                .GetAsync();
            }
        }

        public static Task<object> getDepartment(IGraphClientWrapper graphClient, string siteid, string listid, ILogger Log)
        {
           
            
            var result = graphClient.Department(siteid, listid);

            if (result.Result != null)
            {
                Error err = new Error();
                err.Message = (String)result.Result;
                throw new ServiceException(err);
            }

            return result;
        }
    }
}