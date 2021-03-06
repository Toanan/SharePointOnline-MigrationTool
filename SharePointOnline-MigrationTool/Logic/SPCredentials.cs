﻿using OfficeDevPnP.Core.Utilities;
using System.Security;
using Microsoft.SharePoint.Client;
using CredentialManagement;
using System.Windows;
using System;

namespace SharePointOnline_MigrationTool
{
    /// <summary>
    /// Credential Logic - Handle Credentials for the app
    /// Returns SharePointOnlineCredentials
    /// </summary>
    class SPCredentials
    {

        // Method - Save Credentials in the credential Manager
        public static SharePointOnlineCredentials SaveCredentials(string Url, string UserName, SecureString SecurePassWord)
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

            SharePointOnlineCredentials cred = CredentialManager.GetSharePointOnlineCredential(Url);

            return cred;
        } // End Method

        // Method - Return Stored Credentials
        public static SharePointOnlineCredentials GetStoredCredentials(string Url)
        {
            try
            {
                SharePointOnlineCredentials cred = CredentialManager.GetSharePointOnlineCredential(Url);
                return cred;
            }
            catch (System.ArgumentNullException)
            {
                MessageBox.Show("Please register this tenant before trying to login");
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }  
        } // End Method

        // Method - Return Credentials
        public static SharePointOnlineCredentials GetCredentials(string UserName, SecureString Password)
        {
            // We create the Credentials Object
            SharePointOnlineCredentials cred = new SharePointOnlineCredentials(UserName, Password);

            // And return it
            return cred;
        } // End Method

    }
}
