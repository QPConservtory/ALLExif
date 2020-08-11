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
using DocumentFormat.OpenXml.Wordprocessing;
using Utilities;

namespace AllEXIF
{
    class Program
    {
        const string NO_ARGS = "Please enter arguments for source path and results path. Use AllExif.exe -h for help";
        const string HELP = @"Usage is: AllExif.exe -s c:\pictures folder\ -d c:\report output path -k c:\kml output path  (folders in double quotes if there are spaces in the names)";
        const string PROCESSING = "Processing photos from {0} to {1}";

        static void Main(string[] args)
        {
            DataTable dtGPSData = new DataTable();

            dtGPSData.Columns.AddRange(new DataColumn[24]
            {
                new DataColumn("Filename", typeof(string)),
                new DataColumn("ImageWidth", typeof(string)),
                new DataColumn("ImageHeight", typeof(string)),
                new DataColumn("FileCreationDate", typeof(string)),
                new DataColumn("FileModifyDate", typeof(string)),
                new DataColumn("FileOriginalDate", typeof(string)),
                new DataColumn("FileCreateDate", typeof(string)),
                new DataColumn("Latitude", typeof(string)),
                new DataColumn("Longitude", typeof(string)),
                new DataColumn("Altitude", typeof(string)),
                new DataColumn("CameraMake", typeof(string)),
                new DataColumn("CameraModel", typeof(string)),
                new DataColumn("GPSTimeStamp", typeof(string)),
                new DataColumn("GPSDateStamp", typeof(string)),
                new DataColumn("GPSSpeedRef", typeof(string)),
                new DataColumn("GPSSpeed", typeof(string)),
                new DataColumn("GPSImgDirectionRef", typeof(string)),
                new DataColumn("GPSImgDirection", typeof(string)),
                new DataColumn("GPSDestBearingRef", typeof(string)),
                new DataColumn("GPSDestBearing", typeof(string)),
                new DataColumn("GPSHorizontalPositioningError", typeof(string)),
                new DataColumn("GPSPosition", typeof(string)),
                new DataColumn("SerialNumber", typeof(string)),
                new DataColumn("FileAccessDate", typeof(string))
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
            string imageWidth = string.Empty;
            string gpsDateStamp = string.Empty;
            string imageHeight = string.Empty;
            string gpsSpeedRef = string.Empty;
            string gpsSpeed = string.Empty;
            string gpsImgDirectionRef = string.Empty;
            string gpsImgDirection = string.Empty;
            string gpsDestBearingRef = string.Empty;
            string gpsDestBearing = string.Empty;
            string gpsPosition = string.Empty;
            string serialNumber = string.Empty;
            string fileAccessDate = string.Empty;
            string gpsHPositioningError = string.Empty;

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
                    Console.WriteLine(string.Format("Starting processing {0} Files at {1}: ", source.GetFiles().Length, DateTime.Now));

                    foreach (FileInfo path in source.GetFiles("*", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file : " + path.FullName);
                        serialNumber = GetSerialNumber(path.FullName);
                        //gpsHPositioningError = GetGPSHPositionalError(path.FullName);
                        //fileAccessDate = GetFileAccessDate(path.FullName);

                        try
                        {
                            AllExifTagCollection exif = new AllExifTagCollection(path.FullName);

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
                                        case "ImageWidth":
                                            {
                                                imageWidth = tag.Value;
                                                break;
                                            }
                                        case "ImageHeight":
                                            {
                                                imageHeight = tag.Value;
                                                break;
                                            }
                                        case "GPSSpeedRef":
                                            {
                                                gpsSpeedRef = tag.Value;
                                                break;
                                            }
                                        case "GPSSpeed":
                                            {
                                                gpsSpeed = tag.Value;
                                                break;
                                            }
                                        case "GPSImgDirectionRef":
                                            {
                                                gpsImgDirectionRef = tag.Value;
                                                break;
                                            }
                                        case "GPSImgDirection":
                                            {
                                                gpsImgDirection = tag.Value;
                                                break;
                                            }
                                        case "GPSDestBearingRef":
                                            {
                                                gpsDestBearingRef = tag.Value;
                                                break;
                                            }
                                        case "GPSDestBearing":
                                            {
                                                gpsDestBearing = tag.Value;
                                                break;
                                            }
                                        case "GPSPosition":
                                            {
                                                gpsPosition = tag.Value;
                                                break;
                                            }
                                        case "SerialNumber":
                                            {
                                                serialNumber = tag.Value;
                                                break;
                                            }
                                        case "FileAccessDate":
                                            {
                                                fileAccessDate = tag.Value;
                                                break;
                                            }
                                        case "GPSHPositioningError":
                                            {
                                                gpsHPositioningError = tag.Value;
                                                break;
                                            }
                                    }
                                }

                                if (gpsPosition == string.Empty)
                                {
                                    gpsPosition = latitude + "," + longitude;
                                }

                                dtGPSData.Rows.Add(
                                    Path.GetFileName(path.FullName),
                                    imageWidth,
                                    imageHeight,
                                    dateTimeOriginal,
                                    modifyDateTime,
                                    dateTimeOriginal,
                                    dateTimeOriginal,
                                    latitude,
                                    longitude,
                                    altitude,
                                    make,
                                    model,
                                    gpsTimeStamp,
                                    gpsDateStamp,
                                    gpsSpeedRef,
                                    gpsSpeed,
                                    gpsImgDirectionRef,
                                    gpsImgDirection,
                                    gpsDestBearingRef,
                                    gpsDestBearing,
                                    gpsHPositioningError,
                                    gpsPosition,
                                    serialNumber,
                                    fileAccessDate);

                                if (latitude > 0 && longitude > 0)
                                {
                                    photos.Add(
                                        longitude + "," +
                                        latitude + "," +
                                        altitude + "," +
                                        modifyDateTime + "," +
                                        Path.GetFileName(path.FullName) + "," +
                                        make + "," +
                                        model);
                                }

                                latitude = 0;
                                longitude = 0;
                                altitude = string.Empty;
                                make = string.Empty;
                                model = string.Empty;
                                modifyDateTime = string.Empty;
                                gpsTimeStamp = string.Empty;
                                dateTimeOriginal = string.Empty;
                                dateTimeDigitized = string.Empty;
                                imageWidth = string.Empty;
                                gpsDateStamp = string.Empty;
                                imageHeight = string.Empty;
                                gpsSpeedRef = string.Empty;
                                gpsSpeed = string.Empty;
                                gpsImgDirectionRef = string.Empty;
                                gpsImgDirection = string.Empty;
                                gpsDestBearingRef = string.Empty;
                                gpsDestBearing = string.Empty;
                                gpsPosition = string.Empty;
                                serialNumber = string.Empty;
                                fileAccessDate = string.Empty;
                            }
                        }
                        catch { }
                    }

                    if (photos.Count > 0)
                    {
                        string path3 = Path.Combine(kmlReportDestination.FullName, "Photo GPS Report.kml");
                        KML.Create(photos, path3);
                    }

                    using (XLWorkbook wb = new XLWorkbook())
                    {
                        string path = Path.Combine(excelReportDestination.FullName, "Photo GPS Report.xlsx");
                        wb.Worksheets.Add(dtGPSData, "Data");
                        wb.SaveAs(path);
                    }
                    Console.WriteLine("Ending processing: " + DateTime.Now);
                }
                else
                {
                    Console.WriteLine(HELP);
                }
            }
        }
        public static string GetSerialNumber(string path)
        {
            string serialNumber = "Not in EXIF";


            string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            //This will strip just the working path name:
            string strWorkPath = Path.Combine(Path.GetDirectoryName(strExeFilePath), "exiftool.exe");

            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = strWorkPath,
                    Arguments = string.Format(@"-SerialNumber {0}", path),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            proc.Start();
            serialNumber = (proc.StandardOutput.ReadLine() != string.Empty) ? proc.StandardOutput.ReadLine() : "Not in EXIF";

            return serialNumber;
        }
        public static string GetGPSHPositionalError(string path)
        {
            string gpsHPositionalError = "Not in EXIF";

            string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            //This will strip just the working path name:
            string strWorkPath = Path.Combine(Path.GetDirectoryName(strExeFilePath), "exiftool.exe");

            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = strWorkPath,
                    Arguments = string.Format(@"-GPSHPositioningError {0}", path),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            proc.Start();
            gpsHPositionalError = (proc.StandardOutput.ReadLine() != string.Empty) ? proc.StandardOutput.ReadLine() : "Not in EXIF";

            return gpsHPositionalError;
        }
        public static string GetFileAccessDate(string path)
        {
            string fileAccessDate = "Not in EXIF";

            string strExeFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            //This will strip just the working path name:
            string strWorkPath = Path.Combine(Path.GetDirectoryName(strExeFilePath), "exiftool.exe");

            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = strWorkPath,
                    Arguments = string.Format(@"-FileAccessDate {0}", path),
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            proc.Start();
            fileAccessDate = (proc.StandardOutput.ReadLine() != string.Empty) ? proc.StandardOutput.ReadLine() : "Not in EXIF";

            return fileAccessDate;
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

