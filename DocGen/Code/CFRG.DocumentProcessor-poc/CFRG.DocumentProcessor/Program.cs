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
        //string documentJsonData = string.Empty;
        //using (var httpClient = new HttpClient())
        //{
        //    documentJsonData = await httpClient.GetStringAsync(document.docDataUrl);
        //}

        string documentJsonData = GetDummyData();

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

static string GetDummyData()
{
    return @"{
  ""Order"": [
    {
      ""Number"": ""23"",
      ""Address"": ""Nelson Street"",
      ""Suburb"": ""Howick"",
      ""City"": ""Auckland"",
      ""Phonenumber"": ""543 1234"",
      ""Date"": ""03/01/2010"",
      ""Total"": ""14.00"",
      ""Order_Id"": 0
    }
  ],
  ""Item"": [
    {
      ""Name"": ""BBQ Chicken Pizza"",
      ""Price"": ""6.00"",
      ""Quantity"": ""1"",
      ""ItemTotal"": ""6.00"",
      ""Order_Id"": 0
    },
    {
      ""Name"": ""1.5 Litre Coke"",
      ""Price"": ""4.00"",
      ""Quantity"": ""2"",
      ""ItemTotal"": ""8.00"",
      ""Order_Id"": 0
    }
  ]
}
";
}