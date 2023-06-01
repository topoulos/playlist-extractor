using System.Diagnostics;
using System.Xml;
using Google.Apis.Services;
using Google.Apis.YouTube.v3;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PlaylistExtractor;

string? apiKey;
string? playlistId;
string? outputFilename;

var builder = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

IConfiguration configuration = builder.Build();

if (configuration["Defaults:GoogleApiKey"] != null)
{
    apiKey = configuration["Defaults:GoogleApiKey"];
}
else
{
    Console.WriteLine("Paste Google API Key: ");
    apiKey = Console.ReadLine();
    if (string.IsNullOrEmpty(apiKey))
    {
        Console.WriteLine("No API Key provided. Exiting...");
        return;
    }
}

if (configuration["Defaults:PlaylistId"] != null)
{
    playlistId = configuration["Defaults:PlaylistId"];
}
else
{
    Console.WriteLine("Paste Playlist ID: ");
    playlistId = Console.ReadLine();
    if (string.IsNullOrEmpty(playlistId))
    {
        Console.WriteLine("No Playlist ID provided. Exiting...");
        return;
    }
}

if (configuration["Defaults:OutputFilename"] != null)
{
    outputFilename = configuration["Defaults:OutputFilename"];
}
else
{
    Console.WriteLine("Enter Output filename:");
    outputFilename = Console.ReadLine();
}


if (string.IsNullOrEmpty(outputFilename))
    outputFilename = "PlayListDuration.xlsx";

if (!outputFilename.EndsWith(".xlsx"))
{
    outputFilename += ".xlsx";
}

Console.WriteLine($"Closing Excel file {outputFilename} if open...");

// Close Excel file "PlayListDuration.xlsx" if it is open
foreach (var process in Process.GetProcessesByName("EXCEL"))
{
    if (!process.MainWindowTitle.Contains(outputFilename)) continue;
    process.Kill();
}

var youtubeService = new YouTubeService(new BaseClientService.Initializer
{
    ApiKey = apiKey,
    ApplicationName = "PlaylistExtractor"
});

var playlistRequest = youtubeService.PlaylistItems.List("snippet");
playlistRequest.PlaylistId = playlistId;
playlistRequest.MaxResults = 150;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using var package = new ExcelPackage();
var worksheet = package.Workbook.Worksheets.Add("Youtube Data");

worksheet.Cells[1, 1].Value = "Thumbnail";
worksheet.Cells[1, 2].Value = "Title";
worksheet.Cells[1, 3].Value = "URL";
worksheet.Cells[1, 4].Value = "Duration";
worksheet.Cells[1, 5].Value = "Total Duration: ";

worksheet.Row(1).Style.Font.Bold = true;
worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
worksheet.Row(1).Style.Font.Size = 16;

var row = 2;

var httpClient = new HttpClient();

var nextPageToken = "";

double totalDuration = 0;

while (nextPageToken != null)
{
    playlistRequest.PageToken = nextPageToken;

    var playlistResponse = await playlistRequest.ExecuteAsync();
    
    Console.WriteLine("Found {0} videos for page token {1}", playlistResponse.Items.Count, nextPageToken);
    
    foreach (var playlistItem in playlistResponse.Items)
    {
        Console.WriteLine("Processing video {0}", playlistItem.Snippet.Title);
        
        var videoRequest = youtubeService.Videos.List("contentDetails");
        videoRequest.Id = playlistItem.Snippet.ResourceId.VideoId;

        var videoResponse = await videoRequest.ExecuteAsync();

        var duration = XmlConvert.ToTimeSpan(videoResponse.Items[0].ContentDetails.Duration).TotalSeconds;
        var url = $"https://www.youtube.com/watch?v={playlistItem.Snippet.ResourceId.VideoId}";

        // Download thumbnail image
        var thumbnailUrl = playlistItem.Snippet.Thumbnails.Default__.Url;
        var thumbnailResponse = await httpClient.GetStreamAsync(thumbnailUrl);

        // Add image to worksheet
        using var memoryStream = new MemoryStream();

        await thumbnailResponse.CopyToAsync(memoryStream);
        memoryStream.Position = 0; // Reset position

        var picture = worksheet.Drawings.AddPicture($"Thumbnail{row}", memoryStream);
        picture.SetPosition(row - 1, 10, 0, 10);
        picture.SetSize(100, 100);

        // Adjust row height and column width to fit image
        worksheet.Row(row).Height = 100;
        worksheet.Column(1).Width = Utility.ConvertEmuToPixel(picture.Size.Width) / 6; // Adjust as needed

        worksheet.Cells[row, 2].Value = playlistItem.Snippet.Title;
        worksheet.Cells[row, 3].Value = url;
        worksheet.Cells[row, 3].Hyperlink = new Uri(url);

        worksheet.Cells[row, 4].Value = Utility.TimeStringFromSeconds(duration);

        worksheet.Column(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        worksheet.Column(3).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        worksheet.Column(4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        totalDuration += duration;
        row++;
    }

    nextPageToken = playlistResponse.NextPageToken;
}


worksheet.Column(2).AutoFit();
worksheet.Column(3).AutoFit();
worksheet.Column(4).AutoFit();

worksheet.Cells[1, 6].Value = Utility.TimeStringFromSeconds(totalDuration);
worksheet.Cells[1, 6].Style.Font.Bold = true;
worksheet.Cells[1, 6].Style.Font.Size = 16;

worksheet.Column(5).AutoFit();

worksheet.Column(6).AutoFit();

// Delete Excel file if it already exists
if (File.Exists(outputFilename)) File.Delete(outputFilename);

// Save Excel file
package.SaveAs(new FileInfo(outputFilename));

// Open the Excel file
Process.Start(new ProcessStartInfo(outputFilename) { UseShellExecute = true });