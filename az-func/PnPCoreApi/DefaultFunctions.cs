using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using System.Net.Http;
using System.Collections.Generic;
using PnP.Core.Services;
using PnP.Core.Auth;
using System.Security;
using System.Linq;
using PnP.Core.QueryModel;

namespace PnPCoreApi
{
    public class DefaultFunctions
    {
        private AppInfo _appInfo;
        private readonly IPnPContextFactory _pnpContextFactory;

        public DefaultFunctions(AppInfo appInfo, IPnPContextFactory pnpContextFactory)
        {
            _appInfo = appInfo;
            _pnpContextFactory = pnpContextFactory;
        }

        [FunctionName("GetLists")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "GetLists")] HttpRequestMessage request, ILogger log)
        {
            var query = request.RequestUri.ParseQueryString();
            var siteUrl = query["siteUrl"];
            var tenantId = query["tenantId"];
            var clientSecret = new SecureString();

            foreach (char c in _appInfo.ClientSecret) clientSecret.AppendChar(c);

            var onBehalfAuthProvider = new OnBehalfOfAuthenticationProvider(_appInfo.ClientId, tenantId, clientSecret, () => request.Headers.Authorization.Parameter);
            var results = new List<ListData>();
            using (var pnpContext = await _pnpContextFactory.CreateAsync(new System.Uri(siteUrl), onBehalfAuthProvider))
            {
                await pnpContext.Web.Lists.LoadAsync(l => l.Id, l => l.Title);

                foreach (var list in pnpContext.Web.Lists.AsRequested())
                {
                    results.Add(new ListData { Title = list.Title });
                }
            }

            return new JsonResult(results);
        }
    }
}
