

namespace CFRG.DocGenFunctionApp
{
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.Http;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Extensions.Logging;
    using Newtonsoft.Json;
    using System.Collections.Generic;
    using Aspose.Words;
    using System.Data;
    using System.Net.Http;
    using System;
    using Microsoft.Extensions.Configuration;

    public static class DocGenFunction
    {
        [FunctionName("processDocuments")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log, ExecutionContext context)
        {
            IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(context.FunctionAppDirectory)
            .AddJsonFile("local.settings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

            string apiKey = config["auth-basic-key"];

            if(string.IsNullOrWhiteSpace(apiKey))
            {
                throw new Exception("Function app is not properly setup. Please contact support.");
            }

            if (!string.Equals(req.Headers["auth-key"],  apiKey))
            {
                return new UnauthorizedResult();
            }

            log.LogInformation("C# HTTP trigger function processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var documents = JsonConvert.DeserializeObject<Dictionary<string, string>>(requestBody);


            Document mergedDocs = null;
            foreach (var document in documents)
            {
                Document processedDoc = new Document(document.Value);

                //// Fetch document data.
                string documentJsonData = string.Empty;
                using (var httpClient = new HttpClient())
                {
                    documentJsonData = await httpClient.GetStringAsync(document.Key);
                }

                // string documentJsonData = GetDummyData();

                // Create dataset.
                DataSet documentDataset = JsonConvert.DeserializeObject<DataSet>(documentJsonData);

                processedDoc.MailMerge.TrimWhitespaces = false;
                processedDoc.MailMerge.ExecuteWithRegions(documentDataset);
                if (mergedDocs == null)
                {
                    mergedDocs = processedDoc;
                }
                else
                {
                    mergedDocs.AppendDocument(processedDoc, ImportFormatMode.UseDestinationStyles);
                }
            }

            string fileName = GetNewDocumentName("Final_Doc_");

            MemoryStream memoryStream = new MemoryStream();
            mergedDocs.Save(memoryStream, SaveFormat.Docx);

            // Return the document as a MemoryStreamResult
            var fileContentResult = new FileContentResult(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = fileName
            };

            return fileContentResult;
        }

        public static string GetNewDocumentName(string prefix)
        {
            return $"{prefix}_{DateTime.Now.ToString("ddMMyyyyHHmmssfff")}_{Guid.NewGuid().ToString().Replace("-", "")}.docx";
        }
    }
}
