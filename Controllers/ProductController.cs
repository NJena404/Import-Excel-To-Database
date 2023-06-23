using System.IO;
using System.Web.Mvc;
using MongoDB.Driver;
using OfficeOpenXml;

public class ProductController : Controller
{
    private const string ConnectionString = "mongodb://localhost:27017";
    private const string DatabaseName = "ExcelData";
    private const string CollectionName = "Exceldatas";

    public ActionResult Upload()
    {
        return View();
    }

    [HttpPost]
    public ActionResult Upload(HttpPostedFileBase file)
    {
        if (file != null && file.ContentLength > 0)
        {
            using (var package = new ExcelPackage(file.InputStream))
            {
                var worksheet = package.Workbook.Worksheets[1]; 
                var rowCount = worksheet.Dimension.Rows;
                var dataList = new List<ExcelData>();

                for (var row = 2; row <= rowCount; row++) // Start from row 2 to skip header
                {
                    dataList.Add(new ExcelData
                    {
                        Id = worksheet.Cells[row, 1].Value?.ToString(),
	        device_id = worksheet.Cells[row, 2].Value?.ToString(),
                        power_type = worksheet.Cells[row, 3].Value?.ToString(),
                        description= worksheet.Cells[row, 4].Value?.ToString(),
                        event_date= worksheet.Cells[row, 5].Value?.ToString(),
	        added_on= worksheet.Cells[row, 6].Value?.ToString()
                    });
                }

               
                var client = new MongoClient(ConnectionString);
                var database = client.GetDatabase(DatabaseName);
                var collection = database.GetCollection<ExcelData>(CollectionName);
                collection.InsertMany(dataList);
            }
        }

        return RedirectToAction("Upload");
    }
}
