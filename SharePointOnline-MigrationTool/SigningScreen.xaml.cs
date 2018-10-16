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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SharePointOnline_MigrationTool
{
    /// <summary>
    /// Login Screen - Handle Authentication method to the targeted Tenant
    /// </summary>
    public partial class SigningScreen : Window
    {
        public SigningScreen()
        {
            InitializeComponent();
        }

        // Method - Connect.onClick() 
        private void Connect_Click(object sender, RoutedEventArgs e)
        {
            string tenantUrl = "https://" + TBTenant.Text + "-admin.sharepoint.com";
            // We prepare the Credential object
            Credential Cred;

            // If Tenant is not filled, we give a message and return
            if (string.IsNullOrEmpty(TBTenant.Text))
            {
                MessageBox.Show("Please Provide a Tenant Url");
                return;
            }

            // If SaveCred is checked, we run the SaveCredentials scenario
            if (CBSaveCred.IsChecked == true)
            {
                if (string.IsNullOrEmpty(TBUserName.Text) || PBPassword.SecurePassword.Length == 0)
                {
                    MessageBox.Show("Please provide a User Name and Password");
                    return;
                }
                // We save the filled credentials to the Credential Manager
                Cred = SPCredentials.SaveCredentials(tenantUrl, TBUserName.Text, PBPassword.SecurePassword);

                MessageBox.Show("Saving credentials and using storedCredentials");

                // We hide this window, show the menu and pass the credentials
                //this.Hide();
                //x.show(Cred)
                return;
            }

            // If UserName is not filled, we run the login scenario
            if (string.IsNullOrEmpty(TBUserName.Text))
            {
                // We try to load the credential stored for the targeter Tenant
                Cred = SPCredentials.GetStoredCredentials(tenantUrl);
                // We hide this window, show the menu and pass the credentials
                // [TODO] 

                MessageBox.Show("using storedCredentials");

                //this.Hide();
                // x.show(Cred)
                return;
            }

            Cred = SPCredentials.GetCredentials(TBUserName.Text, PBPassword.SecurePassword);

            MessageBox.Show("using Credentials not saving");

            //this.Hide();
            // x.show(Cred)

            return;
        }
    }
}
