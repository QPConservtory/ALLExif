using LevDan.Exif;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using ClosedXML.Excel;
using System.Data;
using System.Diagnostics;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AllEXIF
{
    class Program
    {
        const string NO_ARGS = "Please enter arguments for source path and results path. Use AllEXIF.exe -h for help";
        const string HELP = @"Usage is: AllEXIF.exe -s c:\pictures folder\ -d c:\report output path -k c:\kml output path  (folders in double quotes if there are spaces in the names)";

        static void Main(string[] args)
        {
            DataTable dtGPSData = new DataTable();

            //DataColumn[] dcs;

            //dtGPSData.Columns.AddRange(dcs);

            double latitude = 0;
            double longitude = 0;
            string altitude = string.Empty;
            string make = string.Empty;
            string model = string.Empty;
            string modifyDateTime = string.Empty;
            string gpsTimeStamp = string.Empty;
            string dateTimeOriginal = string.Empty;
            string dateTimeDigitized = string.Empty;
            string gpsDateStamp = string.Empty;
            List<string> headers = new List<string>();
            List<string> data = new List<string>();
            string exiftoolPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "exiftool.exe");
            List<string> results = new List<string>();

            List<string> photos = new List<string>();

            if (args.Length == 0)
            {
                Console.WriteLine(NO_ARGS);
            }
            else
            {
                if (args[0].ToString() == "-h")
                {
                    Console.WriteLine(HELP);
                }

                if (args[0].ToString() == "-s" && args[2].ToString() == "-d" && args[4].ToString() == "-k")
                {
                    DirectoryInfo source = new DirectoryInfo(args[1].ToString());
                    DirectoryInfo excelReportDestination = new DirectoryInfo(args[3].ToString());
                    DirectoryInfo kmlReportDestination = new DirectoryInfo(args[5].ToString());
                    int numrecs = 0;
                    foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
                    {
                        Console.WriteLine(DateTime.Now.ToLongTimeString());
                        Console.WriteLine("Processing: " + path);
                        numrecs++;
                        try
                        {
                            ExifGPSLatLonTagCollection exif = new ExifGPSLatLonTagCollection(path.FullName);
                            var proc = new Process
                            {
                                StartInfo = new ProcessStartInfo
                                {
                                    FileName = exiftoolPath,
                                    Arguments = string.Format("-g2 {0}", path.FullName),
                                    UseShellExecute = false,
                                    RedirectStandardOutput = true,
                                    CreateNoWindow = true
                                }
                            };

                            proc.Start();

                            //get all the output before processing
                            while (!proc.StandardOutput.EndOfStream)
                            {
                                string retVal = proc.StandardOutput.ReadLine();

                                //check for : in the stream, that means there is data
                                if (retVal.Contains(":"))
                                {
                                    results.Add(retVal);
                                }
                            }
                            //string metadata = string.Empty;
                            //store the data and haeders for each file
                            foreach (string result in results)
                            {
                                //there are lots of colons in the data, get rid of the separator between
                                //the header and data
                                StringBuilder builder = new StringBuilder(result);
                                builder.Replace(": ", "|");

                                //split by |, header is 0, data is 1
                                string[] splitRetVal = builder.ToString().Split('|');

                                string header = splitRetVal[0].ToString();
                                string metadata = splitRetVal[1].ToString();

                                data.Add(metadata);

                                int num = 1;
                                //check if item[0] is in header list
                                if (!headers.Contains(header))
                                {
                                    headers.Add(header);
                                }
                                else
                                {
                                    headers.Add(header + num.ToString());
                                    num++;
                                }
                            }

                            for (int i = 0; i < headers.Count; i++)
                            {
                                dtGPSData.Columns.Add(headers[i], typeof(string));
                            }

                            //instantiate row
                            DataRow row = dtGPSData.NewRow();

                            //fill row
                            for (int i = 0; i < headers.Count; i++)
                            {
                                row[i] = data[i].ToString();
                            }

                            //add row
                            dtGPSData.Rows.Add(row);

                            data.Clear();

                            foreach (DataRow dr in dtGPSData.Rows)
                            {
                                foreach (Cell c in dr.ItemArray)
                                {
                                    Console.WriteLine("Column value: " + c.CellValue.ToString());
                                }
                            }

                            /*if (latitude > 0 && longitude > 0)
                            {
                                dtGPSData.Rows.Add(Path.GetFileName(path.FullName).ToString(), latitude.ToString(),
                                    longitude.ToString(), altitude.ToString(), make.ToString(), model.ToString(),
                                    modifyDateTime.ToString(), dateTimeOriginal.ToString(), dateTimeDigitized.ToString(),
                                    gpsDateStamp.ToString(), gpsTimeStamp.ToString());

                                photos.Add(longitude + "," + latitude + "," + altitude + "," + modifyDateTime + "," +
                                    Path.GetFileName(path.FullName) + "," + make + "," + model);
                            }*/

                            latitude = 0;
                            longitude = 0;
                            altitude = string.Empty;
                            modifyDateTime = string.Empty;
                            make = string.Empty;
                            model = string.Empty;

                        }
                        catch
                        { }
                    }
                    /*
                    if (photos.Count > 0)
                    {
                        string path3 = Path.Combine(kmlReportDestination.FullName, "EXIF Report.kml");
                        KML.Create(photos, path3);
                    }*/

                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        string path = Path.Combine(excelReportDestination.FullName, "EXIF Report.xlsx");
                        wb.Worksheets.Add(dtGPSData, "GPS");
                        wb.SaveAs(path);
                    }
                    Console.WriteLine(DateTime.Now.ToLongTimeString());
                    Console.WriteLine("Num photos: " + numrecs); 
                }
                else
                {
                    Console.WriteLine(HELP);
                }
            }
        }
        /// <summary>
        /// Creates a Point and Placemark and prints the resultant KML on to the console.
        /// </summary>
        public static class KML
        {
            public static void Create(List<string> points, string path)
            {
                //photos.Add(longitude + "," + latitude + "," + altitude + "," + dateTime + "," + Path.GetFileName(path.FullName) + "," + make + "," + model);
                StringBuilder sb = new StringBuilder();

                using (XmlWriter writer = XmlWriter.Create(path))
                {
                    int i = 1;
                    writer.WriteStartElement("Document");
                    writer.WriteElementString("name", "photos.xml");
                    writer.WriteElementString("open", "1");

                    writer.WriteStartElement("Style");
                    writer.WriteStartElement("LabelStyle");
                    writer.WriteElementString("color", "ff0000cc");
                    writer.WriteEndElement();//LabelStyle
                    writer.WriteEndElement();//Style

                    foreach (string p in points)
                    {
                        string[] splitUp = p.Split(',');
                        sb.Append("Latitude: " + splitUp[0].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Longitude: " + splitUp[1].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Altitude: " + splitUp[2].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Date Time: " + splitUp[3].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("File Name: " + splitUp[4].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Make: " + splitUp[5].ToString());
                        sb.Append(Environment.NewLine);
                        sb.Append("Model: " + splitUp[6].ToString());

                        writer.WriteStartElement("Placemark");
                        writer.WriteElementString("description", sb.ToString());
                        writer.WriteElementString("name", splitUp[4].ToString());
                        writer.WriteStartElement("Point");
                        writer.WriteElementString("coordinates", splitUp[0].ToString() + "," + splitUp[1].ToString());
                        writer.WriteEndElement();//Point
                        writer.WriteEndElement();//Placemark
                        i++;

                        sb.Clear();
                    }

                    writer.WriteEndElement();//Document
                    writer.Flush();
                }
            }
        }
    }
}
