using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core;
using System.Security;
using CredentialManagement;
using System.Windows;

namespace SharePointOnline_MigrationTool
{
    /// <summary>
    /// Credential Logic - Get or Save credentials
    /// </summary>
    class SPCredentials
    {
        // Method - Return Stored Credentials
        public static Credential GetStoredCredentials(string Url)
        {
            // We create the Credential to load
            Credential cred = new Credential
            {
                Target = Url
            };

            cred.Load();
            // If the UserName returned is null, we return null and prompt for Registration
            if (string.IsNullOrEmpty(cred.Username))
            {
                MessageBox.Show("No credential found for this tenant, please register this tenant or login");
                return null;
            }
            // Else we return the Credentials loaded
            return cred;
        } // End Method

        // Method - Save Credentials in the credential Manager
        public static Credential SaveCredentials(string Url, string UserName, SecureString SecurePassWord)
        {
            // We create the Credentials to Save
            Credential Cred = new Credential
            {
                Target = Url, // Credential target = Site URL
                Username = UserName,
                SecurePassword = SecurePassWord,
                PersistanceType = PersistanceType.LocalComputer,
                Type = CredentialType.Generic,
            };
            // We push it the the Credential Manager and return it
            Cred.Save();
            return Cred;            
        } // End Method

        // Method - Return Credentials
        public static Credential GetCredentials(string UserName, SecureString Password)
        {
            // We create the Credentials Object
            Credential cred = new Credential
            {
                Username = UserName,
                SecurePassword = Password
            };
            // And return it
            return cred;
        } // End Method
    }
}
