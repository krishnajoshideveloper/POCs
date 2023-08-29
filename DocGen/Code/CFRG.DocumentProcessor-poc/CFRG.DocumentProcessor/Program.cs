using Aspose.Words;
using Newtonsoft.Json;
using System.Data;

try
{

    // Target folder where document final documents need to save. 
    string targetFolder = "temp-docs";

    var documents = new List<(string docUrl, string docDataUrl)>();

    documents.Add((docUrl: "https://raw.githubusercontent.com/krishnajoshideveloper/POCs/main/DocGen/Data/sample-invoice.docx"
        , docDataUrl: "https://raw.githubusercontent.com/krishnajoshideveloper/POCs/main/DocGen/Data/sample-invoice-data-order-1.json"));

    documents.Add((docUrl: "https://raw.githubusercontent.com/krishnajoshideveloper/POCs/main/DocGen/Data/sample-invoice.docx"
        , docDataUrl: "https://raw.githubusercontent.com/krishnajoshideveloper/POCs/main/DocGen/Data/sample-invoice-data-order-2.json"));

    Document mergedDocs = null;
    foreach (var document in documents)
    {
        Document processedDoc = new Document(document.docUrl);

        // Fetch document data.
        string documentJsonData = string.Empty;
        using (var httpClient = new HttpClient())
        {
            documentJsonData = await httpClient.GetStringAsync(document.docDataUrl);
        }        

        // Create dataset.
        DataSet documentDataset = JsonConvert.DeserializeObject<DataSet>(documentJsonData);

        processedDoc.MailMerge.TrimWhitespaces = false;
        processedDoc.MailMerge.ExecuteWithRegions(documentDataset);
        if( mergedDocs == null)
        {
            mergedDocs = processedDoc;
        }
        else
        {
            mergedDocs.AppendDocument(processedDoc, ImportFormatMode.UseDestinationStyles);
        }
    }

    string targetFilePath = Path.Combine(targetFolder, $"Final_Doc_{DateTime.Now.ToString("ddMMyyyyHHmmssfff")}.docx");

    mergedDocs.Save(targetFilePath);
}
catch (Exception ex)
{
    Console.WriteLine(ex.ToString());
}