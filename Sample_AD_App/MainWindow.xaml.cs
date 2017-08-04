using System;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Globalization;
using System.Configuration;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Media.Imaging;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Data.OData;
using System.Data.Services.Client;
using System.Windows.Media;

namespace Sample_AD_App
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string aadInstance = ConfigurationManager.AppSettings["ida:AADInstance"];
        private static string tenant = ConfigurationManager.AppSettings["ida:Tenant"];
        private static string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static string resourceId = ConfigurationManager.AppSettings["todo:ResourceId"];
        private static string graphRequestCom = ConfigurationManager.AppSettings["ida:GraphRequest"];
        private static AuthenticationContext authContext = null;
        private static Uri redirectUri = new Uri(ConfigurationManager.AppSettings["ida:RedirectUri"]);

        private static string authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);

        public MainWindow()
        {
            InitializeComponent();
            authContext = new AuthenticationContext(authority, new FileCache());
            authContext.TokenCache.Clear();
            ClearCookies();
            btnSignOut.Visibility = Visibility.Collapsed;
            btnSignIn.Visibility = Visibility.Visible;
        }

        private static async Task<AADUser> GetUserInfo(string tenantId, string userId, string token)
        {
            string graphRequest = graphRequestCom.Replace("{tenantId}", tenantId)
                  .Replace("{userId}", userId);
            HttpClient client = new HttpClient();

            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            var response = await client.GetAsync(new Uri(graphRequest));

            string content = await response.Content.ReadAsStringAsync();
            var user = JsonConvert.DeserializeObject<AADUser>(content);
            return user;
        }

        private void ClearCookies()
        {
            const int INTERNET_OPTION_END_BROWSER_SESSION = 42;
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0);
        }

        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);

        private async void btnSignIn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                authContext = new AuthenticationContext(authority);
                AuthenticationResult result = null;
                result = await authContext.AcquireTokenAsync(resourceId, clientId, redirectUri, new PlatformParameters(PromptBehavior.Always));

                string token = result?.AccessToken;
                if (token == null)
                {
                    MessageBox.Show("If the error continues, please contact your administrator." + "Sorry, an error occurred while signing you in.");
                }
                if (token != null)
                {
                    btnSignIn.Visibility = Visibility.Collapsed;
                    btnSignOut.Visibility = Visibility.Visible;
                    var user = await GetUserInfo(result.TenantId, result.UserInfo.UniqueId, token);
                    lblOutput.Text = "";
                    lblOutput.Text += $"Display name: {user.displayName}" + Environment.NewLine;
                    lblOutput.Text += $"UPN: {user.userPrincipalName}" + Environment.NewLine;

                    if(user.thumbnailPhoto != null)
                    {
                        var bytes = user.thumbnailPhoto;

                        using (var ms = new MemoryStream(bytes))
                        {
                            var imageSource = new BitmapImage();
                            imageSource.BeginInit();
                            imageSource.StreamSource = ms;
                            imageSource.EndInit();

                            imgUser.Source = imageSource;
                        }
                    }

                    //var servicePoint = new Uri("https://graph.windows.net");
                    //var serviceRoot = new Uri(servicePoint, "massl.onmicrosoft.com"); //e.g. xxx.onmicrosoft.com
                    //const string clientId = "2fa773a1-e0af-43ec-94d6-a0bc99303fe9";
                    //const string secretKey = "";// ClientID and SecretKey are defined when you register application with Azure AD
                    //var authContext1 = new AuthenticationContext("https://login.windows.net/massl.onmicrosoft.com/oauth2/token");
                    //var credential = new ClientCredential(clientId, secretKey);

                    //ActiveDirectoryClient directoryClient = new ActiveDirectoryClient(serviceRoot, async () =>
                    //{
                    //    var result1 = await authContext1.AcquireTokenAsync("https://graph.windows.net/", credential);
                    //    return result1.AccessToken;
                    //});

                    //var user1 = await directoryClient.Users.Where(x => x.UserPrincipalName == user.userPrincipalName).ExecuteSingleAsync();
                    //DataServiceStreamResponse photo = await user1.ThumbnailPhoto.DownloadAsync();
                    //using (MemoryStream s = new MemoryStream())
                    //{
                    //    photo.Stream.CopyTo(s);
                    //    var encodedImage = Convert.ToBase64String(s.ToArray());
                    //    byte[] imgStr = new byte[encodedImage.Length];
                    //    for (int i = 0; i < encodedImage.Length; i++)
                    //    {
                    //        imgStr[i] = (byte)encodedImage[i];
                    //    }
                    //    imgUser.Source = ByteImageConverter.ByteToImage(imgStr);
                    //}
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception." + ex.Message);
            }
        }

        private void btnSignOut_Click(object sender, RoutedEventArgs e)
        {
            authContext.TokenCache.Clear();
            ClearCookies();
            btnSignOut.Visibility = Visibility.Collapsed;
            btnSignIn.Visibility = Visibility.Visible;
            lblOutput.Text = "";
        }
    }

    public class AADUser
    {
        public string displayName { get; set; }
        public string userPrincipalName { get; set; }
        public byte[] thumbnailPhoto { get; set; }
    }

    public class ByteImageConverter
    {
        public static ImageSource ByteToImage(byte[] imageData)
        {
            BitmapImage biImg = new BitmapImage();
            MemoryStream ms = new MemoryStream(imageData);
            biImg.BeginInit();
            biImg.StreamSource = ms;
            biImg.EndInit();

            ImageSource imgSrc = biImg as ImageSource;

            return imgSrc;
        }
    }
}