using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Entities;

namespace SPApplyGAConsole
{
    class Program
    {


        static void Main(string[] args)
        {

            string ga_script = ConfigurationManager.AppSettings["GAScript"].ToString();
            string scriptSrc = string.Format(ga_script, ConfigurationManager.AppSettings["GA_ID"]);

            Console.WriteLine(string.Format("Starting site script update on host {0} at {1} {2}.", Environment.MachineName.ToString(), DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString()));

            Uri tenantAdminUri = new Uri(ConfigurationManager.AppSettings["TenantAdminUrl"]);
            string tenantRealm = TokenHelper.GetRealmFromTargetUrl(tenantAdminUri);
            var token = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, tenantAdminUri.Authority, tenantRealm).AccessToken;
            using (var adminContext = TokenHelper.GetClientContextWithAccessToken(tenantAdminUri.ToString(), token))
            {
                Tenant tenant = new Tenant(adminContext);
                IList<SiteEntity> siteCollections = tenant.GetSiteCollections().Where(s => s.Url.Contains(ConfigurationManager.AppSettings["SitesFilter"])).ToList<SiteEntity>();
                foreach (SiteEntity site in siteCollections)
                {
                    bool siteExists = tenant.SiteExists(site.Url);
                    if (siteExists)
                    {
                        Uri siteUri = new Uri(site.Url);
                        var accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, TokenHelper.GetRealmFromTargetUrl(siteUri)).AccessToken;
                        using (ClientContext ctx = TokenHelper.GetClientContextWithAccessToken(site.Url.ToString(), accessToken))
                        {
                            var ctxSite = ctx.Site;
                            ctx.Load(ctxSite);
                            //Remove existing
                            ctxSite.DeleteJsLink(ConfigurationManager.AppSettings["ScriptKey"]);
                            ctx.ExecuteQuery();
                            //Add updated script
                            ctxSite.AddJsBlock(ConfigurationManager.AppSettings["ScriptKey"], scriptSrc);
                            ctx.ExecuteQuery();
                        }
                    }
                }
            }
            Console.WriteLine(string.Format("Completed site script update on host {0} at {1} {2}.", Environment.MachineName.ToString(), DateTime.Now.ToShortDateString(), DateTime.Now.ToLongTimeString()));

#if DEBUG
            Console.ReadLine();
#endif
        }
    }
}
