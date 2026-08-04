using System;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace AdrkhaTypograph
{
    /// <summary>
    /// يتحقق من توفر تحديثات عبر GitHub Releases API ويحمّل المثبت الجديد عند الحاجة.
    /// </summary>
    internal class UpdateChecker
    {
        private const string GitHubOwner = "adrkha";
        private const string GitHubRepo  = "AdrkhaTypograph";
        private const string ApiUrl      = "https://api.github.com/repos/" + GitHubOwner + "/" + GitHubRepo + "/releases/latest";
        private const string UserAgent   = "AdrkhaTypograph-AutoUpdater";

        public bool   UpdateAvailable { get; private set; }
        public string LatestVersion   { get; private set; }
        public string CurrentVersion  { get; private set; }
        public string DownloadUrl     { get; private set; }
        public string ReleaseNotes    { get; private set; }

        public event Action<UpdateChecker> UpdateDetected;
        public event Action<int>           DownloadProgressChanged;
        public event Action<string>        DownloadCompleted;

        public UpdateChecker()
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version;
            CurrentVersion = $"{v.Major}.{v.Minor}.{v.Build}";
        }

        public async Task CheckAsync()
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                    client.Timeout = TimeSpan.FromSeconds(8);
                    string json = await client.GetStringAsync(ApiUrl);

                    var tagMatch = Regex.Match(json, "\"tag_name\"\\s*:\\s*\"v?([\\d\\.]+)\"");
                    if (!tagMatch.Success) return;
                    LatestVersion = tagMatch.Groups[1].Value;

                    if (IsNewer(LatestVersion, CurrentVersion))
                    {
                        UpdateAvailable = true;
                        var urlMatch = Regex.Match(json, "\"browser_download_url\"\\s*:\\s*\"([^\"]+\\.exe)\"");
                        if (urlMatch.Success) DownloadUrl = urlMatch.Groups[1].Value;
                        var notesMatch = Regex.Match(json, "\"body\"\\s*:\\s*\"(.*?)\"", RegexOptions.Singleline);
                        if (notesMatch.Success)
                            ReleaseNotes = notesMatch.Groups[1].Value.Replace("\\n", "\n").Replace("\\r", "");
                        UpdateDetected?.Invoke(this);
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine("[UpdateChecker] " + ex.Message); }
        }

        public async Task DownloadUpdateAsync()
        {
            if (string.IsNullOrEmpty(DownloadUrl)) return;
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), $"AdrkhaTypograph_{LatestVersion}_Setup.exe");
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", UserAgent);
                    var response = await client.GetAsync(DownloadUrl, HttpCompletionOption.ResponseHeadersRead);
                    response.EnsureSuccessStatusCode();
                    long totalBytes = response.Content.Headers.ContentLength ?? -1;
                    long downloaded = 0;
                    using (var stream     = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
                    {
                        var buffer = new byte[8192];
                        int read;
                        while ((read = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, read);
                            downloaded += read;
                            if (totalBytes > 0)
                                DownloadProgressChanged?.Invoke((int)(downloaded * 100 / totalBytes));
                        }
                    }
                }
                DownloadCompleted?.Invoke(tempPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[UpdateChecker] Download failed: " + ex.Message);
                DownloadCompleted?.Invoke(null);
            }
        }

        public static void LaunchInstaller(string installerPath)
        {
            if (!File.Exists(installerPath)) return;
            Process.Start(new ProcessStartInfo(installerPath) { UseShellExecute = true });
        }

        private static bool IsNewer(string latestVer, string currentVer)
        {
            try { return new Version(latestVer) > new Version(currentVer); }
            catch { return false; }
        }
    }
}
