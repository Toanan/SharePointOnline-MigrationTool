using CredentialManagement;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.SharePoint.Client;
using System.Net;
using System.Security;
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
            // Using CC, we call SPOLogic method to return Tenant Sites
            using (ClientContext ctx = new ClientContext(tenantUrl))
            {
                // Call SPOLogic
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
            }  
        }// End Method

        // Method - TreeViewItem.Expand Listener - Call for Site Lists
        private void Folder_Expanded(object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)sender;

            // If the item only contains the dumy data
            if (item.Items.Count != 1 || item.Items == null)
                return;
            //Clear dummy item
            item.Items.Clear();

            // Get Site library
            var SitePath = (string)item.Tag;

            Task.Factory.StartNew(() =>
            {
                // Call for the expended site Web
                var sp = new SPOLogic(credential, SitePath);
                // Filter on not hidden file
                IEnumerable<Microsoft.SharePoint.Client.List> lists = sp.getWebLists(SitePath, credential).Where(l => !l.Hidden);

                item.Dispatcher.Invoke(() =>
                {
                    // Creating TreeeViewIems from lists
                    foreach (var list in lists)
                    {
                        var subitem = new TreeViewItem
                        {
                            Header = list.Title + " (" + list.ItemCount + ") - " + list.BaseTemplate.ToString(),
                            Tag = list.BaseTemplate.ToString(),
                        };
                        item.Items.Add(subitem);
                    }
                });// End Dispatch
            });// End Task        
        }// End Method
    }
}
