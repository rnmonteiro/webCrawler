using System.Reflection.Metadata;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using HtmlAgilityPack;
using IronXL;

await StartCrawlerAsync();

static async Task StartCrawlerAsync()
{
    var url = "http://www.fundsexplorer.com.br/ranking";
    var httpClient = new HttpClient();
    httpClient
        .DefaultRequestHeaders
        .UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/104.0.5112.81 Safari/537.36 Edg/104.0.1293.54");
    var html = await httpClient.GetStringAsync(url);

    var doc = new HtmlDocument();

    doc.LoadHtml(html);

    createFile(doc);

}

static void createFile(HtmlDocument doc)
{
    var workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    var sheet = workbook.CreateWorkSheet("Result Sheet");

    var i = 2;

    foreach (var row in doc.DocumentNode
                        .SelectNodes("//*[@id='table-ranking']//tr"))
    {
        var header = row.SelectNodes("th");
        if (header != null)
        {
            sheet["A1"].Value = header[0].InnerText;
            sheet["B1"].Value = header[1].InnerText;
            sheet["C1"].Value = header[2].InnerText;
            sheet["D1"].Value = header[3].InnerText;
            sheet["E1"].Value = header[4].InnerText;
            sheet["F1"].Value = header[5].InnerText;
            sheet["G1"].Value = header[6].InnerText;
            sheet["H1"].Value = header[7].InnerText;
            sheet["I1"].Value = header[8].InnerText;
            sheet["J1"].Value = header[9].InnerText;
            sheet["K1"].Value = header[10].InnerText;
            sheet["L1"].Value = header[11].InnerText;
            sheet["M1"].Value = header[12].InnerText;
            sheet["N1"].Value = header[13].InnerText;
            sheet["O1"].Value = header[14].InnerText;
            sheet["P1"].Value = header[15].InnerText;
            sheet["Q1"].Value = header[16].InnerText;
            sheet["R1"].Value = header[17].InnerText;
            sheet["S1"].Value = header[18].InnerText;
            sheet["T1"].Value = header[19].InnerText;
            sheet["U1"].Value = header[20].InnerText;
            sheet["V1"].Value = header[21].InnerText;
            sheet["W1"].Value = header[22].InnerText;
            sheet["X1"].Value = header[23].InnerText;
            sheet["Y1"].Value = header[24].InnerText;
            sheet["Z1"].Value = header[25].InnerText;
        }

        var nodes = row.SelectNodes("td");

        if (nodes != null)
        {
            sheet["A" + (i)].Value = nodes[0].InnerText;
            sheet["B" + (i)].Value = nodes[1].InnerText;
            sheet["C" + (i)].Value = nodes[2].InnerText;
            sheet["D" + (i)].Value = nodes[3].InnerText;
            sheet["E" + (i)].Value = nodes[4].InnerText;
            sheet["F" + (i)].Value = nodes[5].InnerText;
            sheet["G" + (i)].Value = nodes[6].InnerText;
            sheet["H" + (i)].Value = nodes[7].InnerText;
            sheet["I" + (i)].Value = nodes[8].InnerText;
            sheet["J" + (i)].Value = nodes[9].InnerText;
            sheet["K" + (i)].Value = nodes[10].InnerText;
            sheet["L" + (i)].Value = nodes[11].InnerText;
            sheet["M" + (i)].Value = nodes[12].InnerText;
            sheet["N" + (i)].Value = nodes[13].InnerText;
            sheet["O" + (i)].Value = nodes[14].InnerText;
            sheet["P" + (i)].Value = nodes[15].InnerText;
            sheet["Q" + (i)].Value = nodes[16].InnerText;
            sheet["R" + (i)].Value = nodes[17].InnerText;
            sheet["S" + (i)].Value = nodes[18].InnerText;
            sheet["T" + (i)].Value = nodes[19].InnerText;
            sheet["U" + (i)].Value = nodes[20].InnerText;
            sheet["V" + (i)].Value = nodes[21].InnerText;
            sheet["W" + (i)].Value = nodes[22].InnerText;
            sheet["X" + (i)].Value = nodes[23].InnerText;
            sheet["Y" + (i)].Value = nodes[24].InnerText;
            sheet["Z" + (i)].Value = nodes[25].InnerText;

            i++;
        }
    }

    workbook.SaveAs("fiis.xlsx");
}