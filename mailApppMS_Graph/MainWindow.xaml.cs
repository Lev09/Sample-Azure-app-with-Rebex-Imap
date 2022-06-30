using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Identity.Client;
using Rebex;
using Rebex.Net;

namespace mailApppMS_Graph
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        //Set the API Endpoint to Graph 'me' endpoint
        string graphAPIEndpoint = "https://graph.microsoft.com/v1.0/me";

        //Set the scope for API call to user.read
        string[] scopes = new string[] {
            "user.read",
            "profile", // needed to retrieve the user name, which is required for Office365's IMAP authentication
            "email", // not required, but may be useful
            "openid", // required by the 'profile' and 'email' scopes
            "offline_access", // specify this scope to make it possible to refresh the access token when it expires (after one hour)
            "IMAP.AccessAsUser.All",
            "POP.AccessAsUser.All",
            "SMTP.Send"

        };


        public MainWindow()
        {
            InitializeComponent();
            Rebex.Licensing.Key = "==AZeCo71q3Wuc7QtHpIUoGG9FHCFBm1RirXjMVFkUNzGo==";
        }

        /// <summary>
        /// Call AcquireToken - to acquire a token requiring user to sign-in
        /// </summary>
        private async void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationResult authResult = null;
            var app = App.PublicClientApp;
            ResultText.Text = string.Empty;
            TokenInfoText.Text = string.Empty;

            var accounts = await app.GetAccountsAsync();
            var firstAccount = accounts.FirstOrDefault();

            try
            {
                authResult = await app.AcquireTokenSilent(scopes, firstAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException ex)
            {
                // A MsalUiRequiredException happened on AcquireTokenSilent.
                // This indicates you need to call AcquireTokenInteractive to acquire a token
                System.Diagnostics.Debug.WriteLine($"MsalUiRequiredException: {ex.Message}");

                try
                {
                    authResult = await app.AcquireTokenInteractive(scopes)
                        .WithAccount(accounts.FirstOrDefault())
                        .WithPrompt(Prompt.SelectAccount)
                        .ExecuteAsync();
                }
                catch (MsalException msalex)
                {
                    ResultText.Text = $"Error Acquiring Token:{System.Environment.NewLine}{msalex}";
                }
            }
            catch (Exception ex)
            {
                ResultText.Text = $"Error Acquiring Token Silently:{System.Environment.NewLine}{ex}";
                return;
            }

            if (authResult != null)
            {
                ResultText.Text = await GetHttpContentWithToken(graphAPIEndpoint, authResult.AccessToken);
                DisplayBasicTokenInfo(authResult);
                this.SignOutButton.Visibility = Visibility.Visible;

                await GetMessageListAsync(authResult.Account.Username, authResult.AccessToken);
            }
        }

        /// <summary>
        /// Display basic information contained in the token
        /// </summary>
        private void DisplayBasicTokenInfo(AuthenticationResult authResult)
        {
            TokenInfoText.Text = "";
            if (authResult != null)
            {
                TokenInfoText.Text += $"Username: {authResult.Account.Username}" + Environment.NewLine;
                TokenInfoText.Text += $"Token Expires: {authResult.ExpiresOn.ToLocalTime()}" + Environment.NewLine;
            }
        }

        /// <summary>
        /// Sign out the current user
        /// </summary>
        private async void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            var accounts = await App.PublicClientApp.GetAccountsAsync();

            if (accounts.Any())
            {
                try
                {
                    await App.PublicClientApp.RemoveAsync(accounts.FirstOrDefault());
                    this.ResultText.Text = "User has signed-out";
                    this.CallGraphButton.Visibility = Visibility.Visible;
                    this.SignOutButton.Visibility = Visibility.Collapsed;
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }

        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <returns>String containing the results of the GET operation</returns>
        public async Task<string> GetHttpContentWithToken(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }

        private async Task GetMessageListAsync(string userName, string accessToken)
        {
            try
            {
                // connect using Rebex IMAP and retrieve list of messages
                using (var client = new Imap())
                {
                    // communication logging (enable if needed)
                    client.LogWriter = new FileLogWriter("imap-oauth.log");

                    // connect to the server
                    Console.WriteLine("Connecting to IMAP...");
                    await client.ConnectAsync("outlook.office365.com", Imap.DefaultImplicitSslPort, SslMode.Implicit);


                    // NOTE: This is no longer needed in Rebex Secure Mail R5.7 or higher
                    Console.WriteLine("Authenticating to IMAP...");

                    // prepare (wrap) the authentication token for IMAP, POP3, or SMTP
                    //string userName = _credentials.UserName;
                    //string accessToken = _credentials.AccessToken;
                    //string pattern = string.Format("user={0}{1}auth=Bearer {2}{1}{1}", userName, '\x1', accessToken);
                    //string token = Convert.ToBase64String(Encoding.ASCII.GetBytes(pattern));

                    // authenticate using the wrapped access token
                    //await client.LoginAsync(token, ImapAuthentication.OAuth20);


                    // authenticate using the OAuth 2.0 access token
                    await client.LoginAsync(userName, accessToken, ImapAuthentication.OAuth20);

                    // list recent messages in the 'Inbox' folder

                    Console.WriteLine("Listing folder contents...");
                    await client.SelectFolderAsync("Inbox", readOnly: true);

                    int messageCount = client.CurrentFolder.TotalMessageCount;
                    var messageSet = new ImapMessageSet();
                    messageSet.AddRange(Math.Max(1, messageCount - 50), messageCount);

                    var list = await client.GetMessageListAsync(messageSet, ImapListFields.Envelope);
                    list.Sort(new ImapMessageInfoComparer(ImapMessageInfoComparerType.SequenceNumber, Rebex.SortingOrder.Descending));

                    foreach (ImapMessageInfo item in list)
                    {
                        Console.WriteLine($"{item.Date.LocalTime:yyyy-MM-dd} {item.From}: {item.Subject}");
                    }
                }

                Console.WriteLine("Finished successfully!");

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
