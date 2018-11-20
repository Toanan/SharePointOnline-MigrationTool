using OfficeDevPnP.Core.Utilities;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Windows;


namespace SharePointOnline_MigrationTool
{

    /// <summary>
    /// Handle all the logic to interact with SPO
    /// </summary>
    class SPOLogic
    {
        #region Ctor
        /// <summary>
        /// Ctor - Provide Crendentials and admin site Url as props
        /// </summary>
        /// <param name="Credentials"></param>
        /// <param name="Url"></param>
        public SPOLogic(SharePointOnlineCredentials Credentials, string Url)
        {
            this.Credentials = Credentials;
            this.Url = Url;

        }
        #endregion

        #region Props

        public string Url { get; set; }

        public SharePointOnlineCredentials Credentials { get; set; }

        #endregion

        /// <summary>
        /// Return SPOSites tenant wide
        /// </summary>
        /// <returns></returns>
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
                    new SigningScreen().Show();
                }
                return null;
            }
            
        }

        /// <summary>
        /// Return SPOSite Title, Users, Webtemplate and Configuration
        /// </summary>
        /// <param name="Url">SPOSite Url</param>
        /// <returns>Web (filtered, see description)</returns>
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


        }

        /// <summary>
        /// Return List from a SPOSite
        /// </summary>
        /// <returns>Lists in SPOSite</returns>
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
        }

        /// <summary>
        /// Copy less than 2 MB file from a local sourcePath to a SPO Library
        /// </summary>
        /// <param name="sourcePath">Local source Path (folder only)</param>
        /// <param name="targetLib">Library Name</param>
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
        }

        /// <summary>
        /// Return files from an SPOLibrary
        /// </summary>
        /// <param name="targetLib">Library Name</param>
        /// <returns>listitems in the library</returns>
        public ListItemCollection getLibraryFile(string targetLib)
        {
            // using ClientContext
            using (ClientContext ctx = new ClientContext(Url))
            {
                ctx.Credentials = Credentials;

                // We target the list
                List list = ctx.Web.Lists.GetByTitle(targetLib);

                // We get the items from that list (max 10 000)
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View Scope =\"RecursiveAll\"></View>";
                ListItemCollection collListItem = list.GetItems(camlQuery);

                // Load and execute
                ctx.Load(collListItem);
                    /*, items => items.Include(
                         item => item.Id,
                         item => item.DisplayName,
                         item => item.HasUniqueRoleAssignments,
                         item => item.FieldValuesAsHtml,
                         item => item.RoleAssignments));
                    */
                ctx.ExecuteQueryRetry();

                return collListItem;
            }
        }

        public List getListName(string targetLib)
        {
            // using ClientContext
            using (ClientContext ctx = new ClientContext(Url))
            {
                ctx.Credentials = Credentials;

                var url = ctx.Url;
                var list = ctx.Web.Lists.GetByTitle(targetLib);
                ctx.Load(list, w => w.RootFolder.Name);
                ctx.ExecuteQueryRetry();

                return list;
            }
        }
    }

}