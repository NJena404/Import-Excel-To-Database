
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            try
            {
                using (var package = new ExcelPackage(file.InputStream))
                {
                    var worksheet = package.Workbook.Worksheets[1]; 

                    var rowCount = worksheet.Dimension.Rows;
                    var dataList = new List<ExcelData>();

                    for (var row = 2; row <= rowCount; row++) 
                    {
		        var Id = worksheet.Cells[row, 1].Value?.ToString();
                        var deviceId = worksheet.Cells[row, 2].Value?.ToString();
                        var powerType = worksheet.Cells[row, 3].Value?.ToString();
                        var description = worksheet.Cells[row, 4].Value?.ToString();
                        var eventDate = DateTime.Parse(worksheet.Cells[row, 5].Value?.ToString());
                        var addedOn = DateTime.Parse(worksheet.Cells[row, 6].Value?.ToString());

                        if (string.IsNullOrWhiteSpace(deviceId) || string.IsNullOrWhiteSpace(powerType) ||
                            string.IsNullOrWhiteSpace(description))
                        {
                            continue;
                        }

                        dataList.Add(new ExcelData
                        {
			    Id = Id,
                            DeviceId = deviceId,
                            PowerType = powerType,
                            Description = description,
                            EventDate = eventDate,
                            AddedOn = addedOn
                        });
                    }

                    var client = new MongoClient(ConnectionString);
                    var database = client.GetDatabase(DatabaseName);
                    var collection = database.GetCollection<ExcelData>(CollectionName);

	        #Data Entry to DB
                    var uniqueDataList = dataList.Distinct().ToList();
                    collection.InsertMany(uniqueDataList);

                    // Data Analytical 
                    var uptime = CalculateAverageUptime(uniqueDataList);
                    var downtime = CalculateAverageDowntime(uniqueDataList);
                    var peakUsageTimes = FindPeakUsageTimes(uniqueDataList);

                    return RedirectToAction("Upload");
                }
            }
            catch (Exception ex)
            {
                return View("Error");
            }
        }

        return RedirectToAction("Upload");
    }

    // Calculate average uptime
    private TimeSpan CalculateAverageUptime(List<ExcelData> dataList)
    {
        var totalUptime = TimeSpan.Zero;
        var uptimeCount = 0;

        foreach (var data in dataList)
        {
            if (data.PowerType == "On")
            {
                totalUptime += data.EventDate - data.AddedOn;
                uptimeCount++;
            }
        }

        return uptimeCount ;
    }

    // Calculate average downtime
    private TimeSpan CalculateAverageDowntime(List<ExcelData> dataList)
    {
        var totalDowntime = TimeSpan.Zero;
        var downtimeCount = 0;

        foreach (var data in dataList)
        {
            if (data.PowerType == "Off")
            {
                totalDowntime += data.EventDate - data.AddedOn;
                downtimeCount++;
            }
        }

        return downtimeCount ;
    }

    // Peak usage times
    
    private List<DateTime> FindPeakUsageTimes(List<ExcelData> dataList)
    {
        var peakTimes = new List<DateTime>();

        var groupedData = dataList.GroupBy(x => x.EventDate.Hour)
                                  .OrderByDescending(g => g.Count())
                                  .FirstOrDefault();

        if (groupedData != null)
        {
            var maxCount = groupedData.Count();

            foreach (var data in groupedData)
            {
                if (groupedData.Count(x => x.EventDate == data.EventDate) == maxCount)
                {
                    peakTimes.Add(data.EventDate);
                }
            }
        }

        return peakTimes;
    }
}
