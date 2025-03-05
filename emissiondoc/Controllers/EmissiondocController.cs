using Microsoft.AspNetCore.Mvc;
using Azure.Storage.Blobs;
using System.Text.Json;
using System.Text;
using System.Threading.Tasks;
using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;
using Xceed.Document.NET;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Emissions.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EmissionController : Controller
    {
        private readonly string connectionString = "DefaultEndpointsProtocol=https;AccountName=gowrii;AccountKey=0AoLb/gcjZunKKPPd6QGM+zhjDlukd5rzHZO61NZaNoyuIrfPZUHwCpRlbtytn1KwYMnvr1GHnPD+AStA/jZug==;EndpointSuffix=core.windows.net";
        private readonly string containerName = "report";

        private const string API_KEY = "72cf8fe946ad54c62f814a3b58a1f1d4";
        private const string CITY = "Guntur";

        [HttpGet("{activity}")]
        public async Task<IActionResult> GetEmissions(string activity, [FromQuery] double usageValue)
        {
            if (usageValue <= 0)
                return BadRequest(new { error = "Usage value must be greater than zero." });

            double temperature = await GetTemperatureAsync(CITY);
            string season = GetSeason(DateTime.Now);

            var emissions = CalculateEmissions(activity.ToLower(), usageValue);

            var result = new
            {
                Date = DateTime.Now.ToString("yyyy-MM-dd"),
                Temperature = $"{temperature} °C",
                Season = season,
                Activity = activity,
                Scope = emissions.Scope,
                Emissions = $"{Math.Round(emissions.Value, 2)} kg CO₂/day"
            };

            string fileName = $"{activity}_{DateTime.Now:yyyyMMddHHmmss}.docx";
            byte[] wordData = GenerateWordDocument(result);

            bool uploadSuccess = await UploadToBlobStorage(wordData, fileName);
            if (!uploadSuccess)
                return StatusCode(500, new { error = "Failed to upload data to Azure Blob Storage." });

            return Ok(new
            {
                message = "Emissions report generated and uploaded successfully.",
                fileName
            });
        }

        private async Task<double> GetTemperatureAsync(string city)
        {
            string apiUrl = $"https://api.openweathermap.org/data/2.5/weather?q={CITY}&appid={API_KEY}&units=metric";
            using HttpClient client = new HttpClient();
            var response = await client.GetAsync(apiUrl);
            if (!response.IsSuccessStatusCode)
                return -1;

            var jsonResponse = await response.Content.ReadAsStringAsync();
            using JsonDocument doc = JsonDocument.Parse(jsonResponse);
            return doc.RootElement.GetProperty("main").GetProperty("temp").GetDouble();
        }

        private string GetSeason(DateTime date)
        {
            return date.Month switch
            {
                12 or 1 or 2 => "Winter",
                3 or 4 or 5 => "Spring",
                6 or 7 or 8 => "Summer",
                9 or 10 or 11 => "Autumn",
                _ => "Unknown"
            };
        }

        private (string Scope, double Value) CalculateEmissions(string activity, double usageValue)
        {
            var emissionFactors = new Dictionary<string, (string Scope, double Factor)>
            {
                { "car", ("Scope 1 (Direct Emissions - Fuel Combustion)", 2.3) },
                { "boiler", ("Scope 1 (Direct Emissions - Gas Combustion)", 2.2) },
                { "ac", ("Scope 2 (Indirect Emissions - Electricity Usage)", 0.4) },
                { "electricity", ("Scope 2 (Indirect Emissions - Electricity Grid)", 0.4) },
                { "air travel", ("Scope 3 (Indirect Emissions - Business Travel)", 0.15) },
                { "purchased goods", ("Scope 3 (Indirect Emissions - Supply Chain)", 0.5) }
            };

            if (emissionFactors.TryGetValue(activity, out var factorData))
                return (factorData.Scope, usageValue * factorData.Factor);

            return ("Unknown Scope", 0);
        }

        private byte[] GenerateWordDocument(object result)
        {
            using MemoryStream stream = new MemoryStream();
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                var mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                body.AppendChild(new Paragraph(new Run(new Text("🌍 Emission Report"))) { ParagraphProperties = new ParagraphProperties(new Bold()) });
                body.AppendChild(new Paragraph(new Run(new Text($"📅 Date: {result.GetType().GetProperty("Date").GetValue(result)}"))));
                body.AppendChild(new Paragraph(new Run(new Text($"🌡 Temperature: {result.GetType().GetProperty("Temperature").GetValue(result)}"))));
                body.AppendChild(new Paragraph(new Run(new Text($"🍂 Season: {result.GetType().GetProperty("Season").GetValue(result)}"))));
                body.AppendChild(new Paragraph(new Run(new Text($"🚀 Activity: {result.GetType().GetProperty("Activity").GetValue(result)}"))));
                body.AppendChild(new Paragraph(new Run(new Text($"📌 Scope: {result.GetType().GetProperty("Scope").GetValue(result)}"))));
                body.AppendChild(new Paragraph(new Run(new Text($"💨 Emissions: {result.GetType().GetProperty("Emissions").GetValue(result)}"))));

                wordDocument.Save();
            }
            return stream.ToArray();
        }

        private async Task<bool> UploadToBlobStorage(byte[] fileData, string fileName)
        {
            try
            {
                BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);

                await containerClient.CreateIfNotExistsAsync();  // ✅ No public access required

                BlobClient blobClient = containerClient.GetBlobClient(fileName);
                using var stream = new MemoryStream(fileData);
                await blobClient.UploadAsync(stream, overwrite: true);

                Console.WriteLine($"✅ Successfully uploaded: {fileName}");
                return true;
            }
            catch (Azure.RequestFailedException ex)
            {
                Console.WriteLine($"❌ Azure Request Error: {ex.ErrorCode} - {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ General Error: {ex.Message}");
                return false;
            }
        }
    }
}
