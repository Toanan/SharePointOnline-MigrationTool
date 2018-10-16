using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Windows;


namespace SharePointOnline_MigrationTool
{
    class SPOLogic
    {

        public SPOLogic(SharePointOnlineCredentials Credentials, string Url)
        {
            this.Credentials = Credentials;
            this.Url = Url;

        }

        #region Props
        public string Url { get; set; }
        public SharePointOnlineCredentials Credentials { get; set; }
        #endregion

        // Method - Returns tenantSiteProps
        public SPOSitePropertiesEnumerable getTenantProp()
        { 
            try
            {
                ClientContext ctx = new ClientContext(Url);
                ctx.Credentials = Credentials;
                Tenant tenant = new Tenant(ctx);
                SPOSitePropertiesEnumerable prop = tenant.GetSitePropertiesFromSharePoint("0", true);
                ctx.Load(prop);
                ctx.ExecuteQuery();
                return prop;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return null;
        }// End Method

        // Method - Returns webProps
        public Web getWebProps(string Url, string CredName)
        {
            // Creating ClientContext and passing Credentials from CredentialManagement
            using (ClientContext ctx = new ClientContext(Url))
            {
                try
                {
                    ctx.Credentials = Credentials;
                    //Retrieving Web.Title and Web.SiteUsers
                    var web = ctx.Web;
                    ctx.Load(web, w => w.SiteUsers, w => w.Title, w => w.WebTemplate, w => w.Configuration);
                    ctx.ExecuteQuery();
                    return web;
                }
                catch(System.Exception ex)
                {
                    MessageBox.Show(ex.Message + " double check credentials" );
                }

                return null;
            }


        }// End Method

        // Method - Returns web.Lists
        public IEnumerable<List> getWebLists(string Url, SharePointOnlineCredentials Credentials)
        {
            // Using Clientcontext to avoid memory usage with no ctx.dispose()
            using (ClientContext ctx = new ClientContext(Url))
            {
                try
                {
                    ctx.Credentials = Credentials;

                    ListCollection lists = ctx.Web.Lists;

                    ctx.Load(ctx.Web.Lists);
                    ctx.ExecuteQuery();

                    return ctx.Web.Lists;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                return null;
            }
        }// End Method

    }

}