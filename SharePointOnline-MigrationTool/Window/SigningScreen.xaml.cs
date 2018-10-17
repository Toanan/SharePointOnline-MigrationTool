using System.Windows;
using Microsoft.SharePoint.Client;

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
            // We Set the Tenant Url
            string tenantUrl = "https://" + TBTenant.Text + "-admin.sharepoint.com";
            // We prepare the Credential object
            SharePointOnlineCredentials Cred;

            // If Tenant is not filled, we give a message and return
            if (string.IsNullOrEmpty(TBTenant.Text))
            {
                MessageBox.Show("Please Provide a Tenant Url");
                return;
            }

            // If SaveCred is checked, we run the SaveCredentials scenario
            if (CBSaveCred.IsChecked == true)
            {
                // We verify the requested field
                if (string.IsNullOrEmpty(TBUserName.Text) || PBPassword.SecurePassword.Length == 0)
                {
                    MessageBox.Show("Please provide a User Name and Password");
                    return;
                }
                
                // We save the credentials to the Credential Manager
                Cred = SPCredentials.SaveCredentials(tenantUrl, TBUserName.Text, PBPassword.SecurePassword);

                // Handling error on Credentials Creation
                if (Cred != null)
                {
                    this.Hide();
                    new MainMenu(tenantUrl, Cred).Show();
                }
                // We hide this window, show the menu and pass the credentials and Tenant Url
                
                return;
            }

            // If UserName is not filled, we run the login scenario
            if (string.IsNullOrEmpty(TBUserName.Text))
            {
                // We try to load the credential stored for the targeter Tenant
                Cred = SPCredentials.GetStoredCredentials(tenantUrl);
                
                // Handling Error on Credentials not set
                if (Cred != null)
                {
                    // We hide this window, show the menu and pass the credentials and Tenant Url
                    this.Hide();
                    new MainMenu(tenantUrl, Cred).Show();
                }
                return;
            }

            // Else we Create the Credential Object
            Cred = SPCredentials.GetCredentials(TBUserName.Text, PBPassword.SecurePassword);
            
            // We hide this window, show the menu and pass the credentials and Tenant Url
            this.Hide();
            new MainMenu(tenantUrl, Cred).Show();
            return;
        }
    }
}
