using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.IO;
using System.Text;

namespace SharePointOnline_MigrationTool
{
    /// <summary>
    /// Logique d'interaction pour MainMenu.xaml
    /// </summary>
    public partial class MainMenu : Window
    {
        #region Ctor
        public MainMenu(string Url, SharePointOnlineCredentials credential)
        {
            InitializeComponent();
            this.credential = credential;
            this.tenantUrl = Url;
        }
        #endregion

        #region Props
        public string tenantUrl { get; set; }

        public SharePointOnlineCredentials credential { get; set; }
        #endregion

        #region eventHandler

        /// <summary>
        /// Popullate treeview with SPOSite
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        }

        /// <summary>
        /// Expand SPOSite in the treeview to show non hidden lists
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        }

        /// <summary>
        /// Migrate single file (2mb max) from local directory to library - test purpose 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Retrive the items from the selected list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnGetListItems_Click(object sender, RoutedEventArgs e)
        {
            //We prompt for a folder path and retrieve related files
            string sourcePath = prompSourcePath();

            //We retrieve the sub dirinfos
            List<DirectoryInfo> sourceFolders = getSourceFolders(sourcePath);

            //We create the files fileinfo object
            List<FileInfo> files = new List<FileInfo>();

            // Start a task to loop on directories and retrieve file info
            Task.Factory.StartNew(() =>
            {
                //And loop inside all dir to retrieve the files fileinfo
                foreach (DirectoryInfo directory in sourceFolders)
                {
                    List<FileInfo> Currentfiles = getSourceFiles(directory.FullName);
                    foreach (FileInfo fi in Currentfiles)
                    {
                        files.Add(fi);
                    }
                }
                
                //We create the path 
                DateTime now = DateTime.Now;
                var date = now.ToString("yyyy-MM-dd-HH-mm-ss");
                string csvFileName = "Getfile";
                var appPath = AppDomain.CurrentDomain.BaseDirectory;
                var csvfilePath = $"{appPath}{csvFileName}{date}.csv" ;
                //We create the stringbuilder
                var csv = new StringBuilder();
                var header = "Filepath,FileName,LastAccessTime";
                csv.AppendLine(header);

                //Retrive fileinfo and write on csv file
                foreach (FileInfo file in files)
                {
                    var filePath = file.FullName;
                    var fileLastAccess = file.LastAccessTime;
                    var fileName = file.Name;
                    var newLine = string.Format("{0},{1},{2}", filePath, fileName, fileLastAccess);
                    csv.AppendLine(newLine);
                }

                System.IO.File.WriteAllText(csvfilePath, csv.ToString());
                MessageBox.Show("Task done");
            });

            /*
            TBOut.Text += sourceFolders;
            TBOut.Text += Environment.NewLine;
            TBOut.Text += files;

            //We retrieve the selected list from the Treeview and related SPOSite
            string[] selection = getSelectedTreeview();

            //We instanciate the SPOLogic class
            SPOLogic spol = new SPOLogic(credential, selection[1]);

            //We retrieve listitems from the selected library
            ListItemCollection listItems = spol.getLibraryFile(selection[0]);



            //We loop the listitems to show on TBOut
            foreach (ListItem listItem in listItems)
            {
                TBOut.Text += string.Format("{0} - {1}{2}{3}", listItem.FieldValues["FileLeafRef"], listItem.FieldValues["Modified"], listItem.FieldValues["FileRef"], Environment.NewLine);
            }*/
        }
        #endregion

        #region Functions

        /// <summary>
        /// Prompt user for a local folder path
        /// </summary>
        /// <returns></returns>
        private string prompSourcePath()
        {

            //We prompt user for directory selection
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            dialog.Multiselect = false;
            CommonFileDialogResult result = dialog.ShowDialog();
            string selectedPath = dialog.FileName;

            return selectedPath;
        }

        /// <summary>
        /// Retrive the selected list and related SPOSite url in an array 0=>Lib , 1 => SiteUrl
        /// </summary>
        private string[] getSelectedTreeview()
        {
            //We exctract the selected library name
            var selectedlib = SiteView.SelectedItem as TreeViewItem;
            string libfull = selectedlib.Header.ToString();
            string lib = libfull.Split('(')[0].Trim();

            //We extract the SPOSite related (Treeviewitem parent)
            var Parent = selectedlib.Parent as TreeViewItem;
            string siteUrl = Parent.Header.ToString();

            //We create the result array containing Library and site url
            string[] selected = { lib, siteUrl };

            return selected;
        }

        /// <summary>
        /// Retrive items from local directory
        /// </summary>
        /// <param name="url"></param>
        private List<DirectoryInfo> getSourceFolders(string path)
        {
            // TODO ADD the root directory !!
            string[] Folders = Directory.GetDirectories(path, "*.*", SearchOption.AllDirectories);
            //We create the list to put all directories
            List<DirectoryInfo> folders = new List<DirectoryInfo>();
            //We create the source rootFolder DirInfo and add it to the top of the list
            DirectoryInfo rootFolder = new DirectoryInfo(path);
            folders.Add(rootFolder);

            //We loop to populate directory info from directory path
            foreach (string folder in Folders)
            {   
                DirectoryInfo di = new DirectoryInfo(folder);
                folders.Add(di);
            }

            return folders;
        }

        /// <summary>
        /// Retrive items from local directory
        /// </summary>
        /// <param name="url"></param>
        private List<FileInfo> getSourceFiles(string path)
        {
            //We retrive file path from the directory path
            string[] Files = Directory.GetFiles(path, "*.*", SearchOption.TopDirectoryOnly);
            //We create the list to store files info
            List<FileInfo> files = new List<FileInfo>();

            //We loop to populate fileinfo from file path
            foreach (string File in Files)
            {
                FileInfo fi = new FileInfo(File);
                files.Add(fi);
            }

            return files;
        }
        #endregion
    }
}
