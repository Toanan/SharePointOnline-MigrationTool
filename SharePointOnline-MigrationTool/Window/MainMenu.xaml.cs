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
using OfficeDevPnP.Core.Utilities;

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

        // Property : SPO Admin Url
        public string tenantUrl { get; set; }

        // Property : SharePointOnlineCredentials
        public SharePointOnlineCredentials credential { get; set; }

        // Method - Window.loaded
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            using (ClientContext ctx = new ClientContext(tenantUrl))
            {
                ctx.Credentials = credential;
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                try { 
                    ctx.ExecuteQuery();
                    MessageBox.Show(web.Title);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }  
            }
        }
    }
}
