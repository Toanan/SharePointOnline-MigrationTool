using CredentialManagement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace SharePointOnline_MigrationTool
{
    /// <summary>
    /// Logique d'interaction pour MainMenu.xaml
    /// </summary>
    public partial class MainMenu : Window
    {
        public MainMenu(string Url, SharePointOnlineCredentials credential)
        {
            InitializeComponent();
            this.credential = credential;
            this.tenantUrl = Url;
        }

        #region Props
        public string tenantUrl { get; set; }

        public SharePointOnlineCredentials credential { get; set; }
        #endregion

        // Method - Window.loaded -Load Tenant sites TreeView
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            // Call the SPOLogic object
            SPOLogic sp = new SPOLogic(credential, tenantUrl);
            // Ask for Sites and loop 
            SPOSitePropertiesEnumerable Tenant = sp.getTenantProp();
            foreach (var site in Tenant)
            {
                var item = new TreeViewItem
                {
                    Header = site.Url,
                    Tag = site.Url,
                };
                // Adding dumy item.items for expand icon to show
                item.Items.Add(null);
                // Listen out for item being expanded
                item.Expanded += Folder_Expanded;
                SiteView.Items.Add(item);
            }     
        }// End Method

        // Method - TreeViewItem.Expand Listener - Call for Site Lists
        private void Folder_Expanded(object sender, RoutedEventArgs e)
        {

            //We declare the sender TreeViewItem
            var item = (TreeViewItem)sender;

            // If the item only contains the dumy data
            if (item.Items.Count != 1 || item.Items == null)
                return;
            //Clear dummy item
            item.Items.Clear();

            // Get Site library
            var SitePath = (string)item.Tag;

            // We populate TreeViewItems using Threading
            Task.Factory.StartNew(() =>
            {
                // Call the SPOLogic object and pass the item.Url
                var sp = new SPOLogic(credential, SitePath);
                // We call for this site Lists and filter hidden Lists
                IEnumerable<Microsoft.SharePoint.Client.List> lists = sp.getWebLists().Where(l => !l.Hidden);

                item.Dispatcher.Invoke(() =>
                {
                    // We push TreeViewIems from lists
                    foreach (var list in lists)
                    {
                        var subitem = new TreeViewItem
                        {
                            Header = list.Title + " (" + list.ItemCount + ") - " + list.BaseTemplate.ToString(),
                            Tag = list.BaseTemplate.ToString(),
                        };
                        item.Items.Add(subitem);
                    }
                });
            });// End Task        
        }// End Method

        // Method - Migrate.onClick - Copy Files from source to target library -TEST !
        private void Migrate_Click(object sender, RoutedEventArgs e)
        {
            //We set up source and target strings
            string source = @"c:\tmp\test.txt"; //TBSource.Text;
            string target = TBTarget.Text;

            // Call the SPOLogic object
            SPOLogic sp = new SPOLogic(credential, "https://toanan.sharepoint.com/sites/demo");
            try
            {
                //Try to copy the file and give success message
                sp.migrateLightFile(source, target);
                MessageBox.Show(string.Format("The file {0} has been migrated", source));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Method get List Items.onClick() - retrieve list items
        private void BtnGetListItems_Click(object sender, RoutedEventArgs e)
        {
            var selectedlib = SiteView.SelectedItem as TreeViewItem;
            string libfull = selectedlib.Header.ToString();
            string lib = libfull.Split('(')[0].Trim();
            var Parent = selectedlib.Parent as TreeViewItem;
            string siteUrl = Parent.Header.ToString();

            SPOLogic spol = new SPOLogic(credential, siteUrl);

            ListItemCollection listItems = spol.getLibraryFile(lib);

            foreach (ListItem listItem in listItems)
            {
                TBOut.Text += string.Format("{0} - {1}{2}", listItem["Title"], listItem["Modified"], Environment.NewLine);
            }


        }// End Method
    }
}
