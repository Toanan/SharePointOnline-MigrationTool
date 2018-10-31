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

        #region MyRegion

        #endregion
        // Method - Returns tenantSiteProps
        public SPOSitePropertiesEnumerable getTenantProp()
        { 
            using (ClientContext ctx = new ClientContext(Url))
            {
                try
                {
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
            }
            
        }// End Method

        // Method - Returns webProps - NOT USED YET !
        public Web getWebProps(string Url)
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
        public IEnumerable<List> getWebLists()
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

        // Method - Migrate <=2mb file 
        public void migrateLightFile(string sourcePath, string targetLib)
        {
    
            // We set the fileName from sourcePath
            string fileName = sourcePath.Substring(sourcePath.LastIndexOf("\\") + 1);

            using (ClientContext ctx = new ClientContext(Url))
            {
                ctx.Credentials = Credentials;

                // We create the FileInfo to migrate
                FileCreationInformation fileInfo = new FileCreationInformation();
                fileInfo.Url = fileName;
                fileInfo.Overwrite = true;
                fileInfo.Content = System.IO.File.ReadAllBytes(sourcePath);

                Web web = ctx.Web;

                // We set the target library and folder
                List lib = web.Lists.GetByTitle(targetLib);
                Folder folder = lib.RootFolder.Folders.GetByUrl(string.Format("{0}/folder/", targetLib));

                // We push the file to SPO
                File file = folder.Files.Add(fileInfo);
                //File file = lib.RootFolder.Files.Add(fileInfo);
                ctx.ExecuteQuery();

                file.ListItemAllFields["Created"] = "2015-07-03";
                //file.ListItemAllFields["Modified"] = "2018-01-01";
                //file.ListItemAllFields["Author"] = "bob@toanan.onmicrosoft.com";
                //file.ListItemAllFields["Editor"] = "bobo@toanan.onmicrosoft.com";
                file.ListItemAllFields.Update();
                ctx.ExecuteQuery();
            }
        } // End Method
    }

}