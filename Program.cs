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

namespace AllEXIF
{
    class Program
    {
        const string NO_ARGS = "Please enter arguments for source path and results path. Use AllEXIF.exe -h for help";
        const string HELP = @"Usage is: AllEXIF.exe -s c:\pictures folder\ -d c:\report output path -k c:\kml output path  (folders in double quotes if there are spaces in the names)";
        const string PROCESSING = "Processing photos from {0} to {1}";

        static void Main(string[] args)
        {
            DataTable dtGPSData = new DataTable();

            dtGPSData.Columns.AddRange(new DataColumn[11]
            {
                new DataColumn("Filename", typeof(string)),
                new DataColumn("Latitude", typeof(string)),
                new DataColumn("Altitude",typeof(string)),
                new DataColumn("Longitude", typeof(string)),
                new DataColumn("Make", typeof(string)),
                new DataColumn("Model",typeof(string)),
                new DataColumn("ModifyDateTime", typeof(string)),
                new DataColumn("GPSTimeAtomicClock", typeof(string)),
                new DataColumn("DateTimeOriginal",typeof(string)),
                new DataColumn("DateTimeDigitized", typeof(string)),
                new DataColumn("GPSDateStamp", typeof(string))
            });

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

                    foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
                    {
                        try
                        {
                            ExifGPSLatLonTagCollection exif = new ExifGPSLatLonTagCollection(path.FullName);
                            var proc = new Process
                            {
                                StartInfo = new ProcessStartInfo
                                {
                                    FileName = @"C:\Windows\exiftool.exe",
                                    Arguments = string.Format(@"-g2 {0}", path),
                                    UseShellExecute = false,
                                    RedirectStandardOutput = true,
                                    CreateNoWindow = true
                                }
                            };

                            proc.Start();
                            while (!proc.StandardOutput.EndOfStream)
                            {
                                //listBoxTopResults.Items.Add(proc.StandardOutput.ReadLine().Replace("  ", ""));
                            }

                            if (exif.Count() >= 3)
                            {
                                foreach (ExifTag tag in exif)
                                {
                                    string latRef = string.Empty;
                                    string lonRef = string.Empty;

                                    foreach (ExifTag tag2 in exif)
                                    {
                                        switch (tag2.FieldName)
                                        {
                                            case "GPSLatitudeRef":
                                                {
                                                    latRef = tag2.Value;
                                                    break;
                                                }
                                            case "GPSLongitudeRef":
                                                {
                                                    lonRef = tag2.Value;
                                                    break;
                                                }
                                        }
                                    }
                                    switch (tag.FieldName)
                                    {
                                        case "GPSLatitude":
                                            {
                                                if (!string.IsNullOrEmpty(latRef))
                                                {
                                                    latitude = Utilities.GPS.GetLatLonFromDMS(latRef.Substring(0, 1) + tag.Value);
                                                }
                                                latitude = Utilities.GPS.GetLatLonFromDMS(tag.Value);
                                                break;
                                            }
                                        case "GPSLongitude":
                                            {
                                                if (!string.IsNullOrEmpty(lonRef))
                                                {
                                                    longitude = Utilities.GPS.GetLatLonFromDMS(lonRef.Substring(0, 1) + tag.Value);
                                                }
                                                longitude = Utilities.GPS.GetLatLonFromDMS(tag.Value);
                                                break;
                                            }
                                        case "GPSAltitude":
                                            {
                                                altitude = tag.Value;
                                                break;
                                            }
                                        case "DateTime":
                                            {
                                                modifyDateTime = tag.Value;
                                                break;
                                            }
                                        case "Make":
                                            {
                                                make = tag.Value;
                                                break;
                                            }
                                        case "Model":
                                            {
                                                model = tag.Value;
                                                break;
                                            }
                                        case "DateTimeOriginal":
                                            {
                                                dateTimeOriginal = tag.Value;
                                                break;
                                            }
                                        case "DateTimeDigitized":
                                            {
                                                dateTimeDigitized = tag.Value;
                                                break;
                                            }
                                        case "GPSDateStamp":
                                            {
                                                gpsDateStamp = tag.Value;
                                                break;
                                            }
                                        case "GPSTimeStamp":
                                            {
                                                gpsTimeStamp = tag.Value;
                                                break;
                                            }
                                    }
                                }

                                if (latitude > 0 && longitude > 0)
                                {
                                    dtGPSData.Rows.Add(Path.GetFileName(path.FullName).ToString(), latitude.ToString(),
                                        longitude.ToString(), altitude.ToString(), make.ToString(), model.ToString(),
                                        modifyDateTime.ToString(), dateTimeOriginal.ToString(), dateTimeDigitized.ToString(),
                                        gpsDateStamp.ToString(), gpsTimeStamp.ToString());

                                    photos.Add(longitude + "," + latitude + "," + altitude + "," + modifyDateTime + "," +
                                        Path.GetFileName(path.FullName) + "," + make + "," + model);
                                }

                                latitude = 0;
                                longitude = 0;
                                altitude = string.Empty;
                                modifyDateTime = string.Empty;
                                make = string.Empty;
                                model = string.Empty;
                            }
                        }
                        catch { }
                    }

                    if (photos.Count > 0)
                    {
                        string path3 = Path.Combine(kmlReportDestination.FullName, "EXIF Report.kml");
                        KML.Create(photos, path3);
                    }

                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        string path = Path.Combine(excelReportDestination.FullName, "EXIF Report.xlsx");
                        wb.Worksheets.Add(dtGPSData, "GPS");
                        wb.SaveAs(path);
                    }
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
