using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.IO;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;

namespace ExportXLS
{
    class Program
    {
        private static Settings settings = new Settings();

        static void Main(string[] args)
        {
            try
            {
                XDocument xml;
                List<XElement> root;
                DateTime day;
                string dateFormat = "yyyyMMdd";
                if (args.Length > 0)
                {
                    if (args.Contains("/xml"))
                    {
                        if (args.Length == 3)
                            day = DateTime.ParseExact(args[2], dateFormat, CultureInfo.InvariantCulture.DateTimeFormat);
                        else
                            day = DateTime.Now.Date.AddDays(-1);
                        xml = XDocument.Load(args[1]);
                        root = new List<XElement>(xml.Descendants("job"));
                        Work(root, day);
                        Logger.Log(string.Format("OK: {0} за {1}", args[1], day.ToString("dd.MM.yyyy")));
                    }
                    else if (args.Contains("/?"))
                    {
                        ShowHelp();
                    }
                    else
                    {
                        day = DateTime.ParseExact(args[1], dateFormat, CultureInfo.InvariantCulture.DateTimeFormat);
                        xml = XDocument.Load("ExportTasks.xml");
                        var query = from element in xml.Descendants("job")
                                    where element.Attribute("inn").Value == args[0]
                                    select element;
                        Work(new List<XElement>(query), DateTime.ParseExact(args[1],
                            dateFormat, CultureInfo.InvariantCulture.DateTimeFormat));
                        Logger.Log(string.Format("OK: {0} за {1}", args[0], day.ToString("dd.MM.yyyy")));
                    }
                }
                else
                    ShowHelp();
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message + (ex.InnerException == null ? "" : ": " + ex.InnerException.Message));
                Console.WriteLine();
                Console.WriteLine("Ошибка! Подробности смотри в Export.log");
            }            
        }

        private static void ShowHelp()
        {
            Console.WriteLine("ExportXLS - программа для экспорта данных из АИИС КУЭ Piramida2000");
            Console.WriteLine("Использование:");
            Console.WriteLine("1) ExportXLS /xml файл_описания_задачи.xml date");
            Console.WriteLine("2) ExportXLS inn date");
            Console.WriteLine("inn - Код смежника (элемент sender.inn в макете 80020)");
            Console.WriteLine("date - Дата результатов измерения (элемент datetime.day в макете 80020)");
            Console.WriteLine();
            Console.WriteLine("ExportXLS /? - этот текст");
            Console.WriteLine();
            Console.WriteLine("Для завершения нажмите любую кнопку");
            Console.Read();
        }

        private static void Work(List<XElement> xml, DateTime day)
        {
            foreach (XElement job in xml)
            {
                Console.Write(job.Attribute("description").Value + ": в процессе...");
                string fileName = job.Descendants("FileNamePrefix").First().Value +
                                  " " + day.ToString("yyyy-MM-dd") + ".xlsx";
                int sensorCount = job.Descendants("sensor").Count();
                string[] deviceCodes = new string[sensorCount]; // This array must be the same size as sensorCodes
                string[] sensorCodes = new string[sensorCount];
                string workingFolder = job.Descendants("workingdirectory").First().Value;
                string archiveFolder = job.Descendants("archive").First().Value;
                int i = 0;
                foreach (XElement device in job.Descendants("device"))
                {
                    string deviceCode = device.Attribute("code").Value;
                    foreach (XElement sensor in device.Descendants("sensor"))
                    {
                        deviceCodes[i] = deviceCode;
                        sensorCodes[i] = sensor.Attribute("code").Value;
                        i++;
                    } // end of foreach (XElement sensor in device.Descendants("sensor"))
                } // end of foreach (XElement device in job.Descendants("device"))
                ToXLS12(deviceCodes, sensorCodes, day, Path.Combine(workingFolder, fileName));
                string[] addresses = job.Descendants("email").Select(el => el.Value).ToArray();
                StringBuilder contacts=new StringBuilder();
                foreach (XElement cont in job.Descendants("contacts"))
                {
                    contacts.Append(cont.Value);
                    contacts.Append(Environment.NewLine);
                }
                SendMail(job.Attribute("description").Value + " за " + day.ToString("dd-MM-yyyy"),
                    Path.Combine(workingFolder, fileName), string.Join(",", addresses),
                    job.Descendants("from").First().Value, contacts.ToString());
                ToArchive(fileName, workingFolder, archiveFolder);
                Console.WriteLine();
            } // end of foreach (XElement job in xml)
        }

        private static void ToArchive(string fileName, string workingFolder, string archiveFolder)
        {
            if (!Directory.Exists(workingFolder))
                throw new Exception("Неверно задана рабочая папка: " + workingFolder);
            if(!Directory.Exists(archiveFolder))
                throw new Exception("Неверно задана архивная папка: " + archiveFolder);
            if (!File.Exists(Path.Combine(workingFolder, fileName)))
                throw new Exception("Ошибка доступа к файлу: " + fileName);
            try
            {
                File.Move(Path.Combine(workingFolder, fileName),
                          Path.Combine(archiveFolder,
                                       Path.GetFileNameWithoutExtension(fileName) +
                                       "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") +
                                       Path.GetExtension(fileName)));
            }
            catch (Exception ex)
            {
                throw new Exception("Не удалось переместить файл в архив: " + fileName, ex);
            }
        }

        private static void SendMail(string jobTitle, string fileName, string addresses,
                                     string addressFrom, string contacts)
        {
            try
            {
                MailMessage msg = new MailMessage(addressFrom, addresses, jobTitle, contacts);
                msg.Attachments.Add(new Attachment(fileName));
                msg.Body = msg.Body + Environment.NewLine + Environment.NewLine +
                    "Сообщение сформировано автоматически: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                SmtpClient smtp = new SmtpClient(settings.SMTPServer, settings.SMTPPort);
                if (settings.UseSSL == 1)
                    smtp.EnableSsl = true;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential(settings.SMTPUserName, settings.SMTPPassword);
                smtp.Send(msg);
                smtp.Dispose();
                msg.Dispose();
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка отправки сообщения " + fileName, ex);
            }
        }

        private static void ToXLS12_broken(string[] devices, string[] sensors, DateTime day, string fileName)
        {
            DataProvider d = new DataProvider(settings.Server, settings.Database, settings.UserName, settings.Password);
            Excel.Range c;
            Excel.Application xls;
            Excel.Workbook wb;
            int percent;
            int firstRow = 4;
            int totalSensors = sensors.Length;
            int totalRows;
            int totalData;
            int completed = 0;
            double firstHalf = 0; // The first halfhour of the two forming hour value as their average
            Dictionary<string, string> sensorInfo;
            string deviceCode;
            string sensorCode;
            double halfhour;
            xls = new Excel.Application();
            xls.SheetsInNewWorkbook = 2;
            wb = xls.Workbooks.Add();
            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            Excel.Worksheet ws2 = (Excel.Worksheet)wb.Worksheets[2];
            ws1.Name = "Получасовки";
            ws2.Name = "Часовки";
            #region Prepare headers of halfhours worksheet
            c = (Excel.Range)(ws1.Cells[1, 1]);
            c.Value = "Дата:";
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            c = (Excel.Range)(ws1.Cells[1, 2]);
            c.Value = day.ToShortDateString();
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;      
            c = (Excel.Range)(ws1.Cells[firstRow - 1, 1]);
            c.Value = "Дата";
            c.ColumnWidth = 12;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            c = (Excel.Range)(ws1.Cells[firstRow - 1, 2]);
            c.Value = "Время";
            c.ColumnWidth = 13;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            #endregion

            #region Prepare headers of hours worksheet
            c = (Excel.Range)(ws2.Cells[1, 1]);
            c.Value = "Дата";
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;            
            c = (Excel.Range)(ws2.Cells[1, 2]);
            c.Value = day.ToShortDateString();
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            c = (Excel.Range)(ws2.Cells[firstRow - 1, 1]);
            c.Value = "Дата";
            c.ColumnWidth = 12;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            c = (Excel.Range)(ws2.Cells[firstRow - 1, 2]);
            c.Value = "Время";
            c.ColumnWidth = 13;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            #endregion

            DateTime currentDate = day;

            #region Write dates and times into two leftmost columns of halfhours sheet
            int currentRow = 0;
            int currentColumn = 3;
            totalRows = 48;
            string[,] leftColumns = new string[totalRows, 2];
            while (currentDate < day.AddDays(1))
            {
                leftColumns[currentRow, 0] =
                    currentDate.Date.ToShortDateString();
                leftColumns[currentRow, 1] =
                    string.Format("{0:00}:{1:00} - {2:00}:{3:00}",
                                  currentDate.TimeOfDay.Hours,
                                  currentDate.TimeOfDay.Minutes,
                                  currentDate.AddMinutes(30).TimeOfDay.Hours,
                                  currentDate.AddMinutes(30).TimeOfDay.Minutes);
                currentDate = currentDate.AddMinutes(30);
                currentRow++;
            }
            c = (Excel.Range)ws1.Cells[firstRow, 1];
            c = c.Resize[totalRows, 2];
            c.Value = leftColumns;
            #endregion

            #region Write dates and times into two leftmost columns of hours sheet
            currentDate = day;
            currentRow = 0;
            totalRows = 24;
            leftColumns = new string[totalRows, 2];
            while (currentDate < day.AddDays(1))
            {
                leftColumns[currentRow, 0] =
                    currentDate.Date.ToShortDateString();
                leftColumns[currentRow, 1] =
                    string.Format("{0:00}:00 - {1:00}:00",
                                   currentDate.TimeOfDay.Hours,
                                   currentDate.AddHours(1).TimeOfDay.Hours);
                currentDate = currentDate.AddHours(1);
                currentRow++;
            }
            c = (Excel.Range)ws2.Cells[firstRow, 1];
            c = c.Resize[totalRows, 2];
            c.Value = leftColumns;
            #endregion

            totalRows = 48;
            totalData = totalRows * totalSensors;
            currentColumn = 3;
            string deviceName, sensorName;
            int cursorPosition = Console.CursorLeft - 2;
            for(int i=0;i<sensors.Length;i++)
            {                
                deviceCode = devices[i];
                sensorCode = sensors[i];
                currentDate = day.AddMinutes(30);
                currentRow = firstRow;
                #region Write devices' and sensors' names into first two rows
                // halfhours column headers
                c = (Excel.Range)(ws1.Cells[1, currentColumn]);
                c.ColumnWidth = 18;
                deviceName = d.GetDeviceName(deviceCode);
                c.Value = deviceName;
                sensorName = d.GetSensorName(deviceCode, sensorCode);
                ws1.Cells[2, currentColumn] = sensorName;
                c = (Excel.Range)(ws1.Cells[firstRow - 1, currentColumn]);
                c.FormulaR1C1 = string.Format("=SUM(R[1]C:R[{0}]C)/2", totalRows);
                c.Font.Bold = true;
                c.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                c.Interior.Color = Excel.XlRgbColor.rgbGrey;
                // hours column headers
                c = (Excel.Range)(ws2.Cells[1, currentColumn]);
                c.ColumnWidth = 18;
                c.Value = deviceName;
                ws2.Cells[2, currentColumn] = sensorName;
                c = (Excel.Range)(ws2.Cells[firstRow - 1, currentColumn]);
                c.FormulaR1C1 = string.Format("=SUM(R[1]C:R[{0}]C)", totalRows / 2);
                c.Font.Bold = true;
                c.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                c.Interior.Color = Excel.XlRgbColor.rgbGrey;
                #endregion
                while (currentDate <= day.AddDays(1))
                {
                    halfhour = d.GetSingleHalfhour(deviceCode, sensorCode, currentDate);
                    if (halfhour < 0)
                        ws1.Cells[currentRow, currentColumn] = "";
                    else
                        ws1.Cells[currentRow, currentColumn] = halfhour;
                    if ((currentRow - firstRow) % 2 == 0)
                    {
                        firstHalf = (halfhour < 0) ? 0 : halfhour;
                    }
                    else
                    {
                        c = (Excel.Range)ws2.Cells[(currentRow - firstRow) / 2 + firstRow, currentColumn];

                        c.Value = (firstHalf + ((halfhour < 0) ? 0 : halfhour)) / 2;
                        c.NumberFormat = "#,##0.00";
                        firstHalf = 0;
                    }
                    currentDate = currentDate.AddMinutes(30);
                    currentRow++;
                    completed++;
                    percent = 100 * completed / totalData;
                    Console.CursorLeft = cursorPosition;
                    Console.Write(percent.ToString() + "%");
                }
                currentColumn++;
            }            
            ws1.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ws2.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            c = (Excel.Range)ws2.Cells[firstRow, 3];
            ws2.Activate();
            c.Select();
            Excel.Windows xlsWindows = wb.Windows;
            Excel.Window xlsWindow = xlsWindows[1];
            xlsWindow.FreezePanes = true;
            c = (Excel.Range)ws1.Cells[firstRow, 3];
            ws1.Activate();
            c.Select();
            //xlsWindow = xlsWindows[1];
            xlsWindow.FreezePanes = true;
            c = (Excel.Range)ws1.Cells[firstRow - 1, 3];
            c = c.Resize[1, totalSensors];
            c.NumberFormat = "#,##0";
            c = (Excel.Range)ws2.Cells[firstRow - 1, 3];
            c = c.Resize[1, totalSensors];
            c.NumberFormat = "#,##0";
            xlsWindow.Activate();
            if (File.Exists(fileName))
                File.Delete(fileName);
            wb.SaveAs(fileName);
            wb.Close();
            xls.Quit();
            releaseObject(ws1);
            releaseObject(ws2);
            releaseObject(wb);
            releaseObject(xls);
        }

        private static void ToXLS101(string[] devices, string[] sensors, DateTime day, string fileName)
        {
            DataProvider d = new DataProvider(settings.Server, settings.Database, settings.UserName, settings.Password);
            Excel.Range c;
            Excel.Application xls;
            Excel.Workbook wb;
            int percent;
            int firstRow = 4;
            int totalSensors = sensors.Length;
            int totalRows;
            int totalData;
            int completed = 0;
            string deviceCode;
            string sensorCode;
            double value;
            xls = new Excel.Application();
            xls.SheetsInNewWorkbook = 1;
            wb = xls.Workbooks.Add();
            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            ws1.Name = "Показания";
            #region Prepare headers
            c = (Excel.Range)(ws1.Cells[1, 1]);
            c.Value = "Дата:";
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            c = (Excel.Range)(ws1.Cells[1, 2]);
            c.Value = day.ToShortDateString();
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            c = (Excel.Range)(ws1.Cells[firstRow - 1, 1]);
            c.Value = "Дата";
            c.ColumnWidth = 12;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            c = (Excel.Range)(ws1.Cells[firstRow - 1, 2]);
            c.Value = "Время";
            c.ColumnWidth = 13;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            #endregion

            DateTime currentDate = day;

            #region Write dates and times into two leftmost coumns
            int currentRow = 0;
            int currentColumn = 3;
            totalRows = 48;
            string[,] leftColumns = new string[totalRows, 2];
            while (currentDate < day.AddDays(1))
            {
                leftColumns[currentRow, 0] =
                    currentDate.Date.ToShortDateString();
                leftColumns[currentRow, 1] =
                    string.Format("{0:00}:{1:00} - {2:00}:{3:00}",
                                  currentDate.TimeOfDay.Hours,
                                  currentDate.TimeOfDay.Minutes,
                                  currentDate.AddMinutes(30).TimeOfDay.Hours,
                                  currentDate.AddMinutes(30).TimeOfDay.Minutes);
                currentDate = currentDate.AddMinutes(30);
                currentRow++;
            }
            c = (Excel.Range)ws1.Cells[firstRow, 1];
            c = c.Resize[totalRows, 2];
            c.Value = leftColumns;
            #endregion

            totalRows = 48;
            totalData = totalRows * totalSensors;
            currentColumn = 3;
            string deviceName, sensorName;
            int cursorPosition = Console.CursorLeft - 2;
            for (int i = 0; i < sensors.Length; i++)
            {
                deviceCode = devices[i];
                sensorCode = sensors[i];
                currentDate = day.AddMinutes(30);
                currentRow = firstRow;
                #region Write devices' and sensors' names into first two rows
                // halfhours column headers
                c = (Excel.Range)(ws1.Cells[1, currentColumn]);
                c.ColumnWidth = 18;
                deviceName = d.GetDeviceName(deviceCode);
                c.Value = deviceName;
                sensorName = d.GetSensorName(deviceCode, sensorCode);
                ws1.Cells[2, currentColumn] = sensorName;
                c = (Excel.Range)(ws1.Cells[firstRow - 1, currentColumn]);
                c.FormulaR1C1 = string.Format("=SUM(R[1]C:R[{0}]C)/2", totalRows);
                c.Font.Bold = true;
                c.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                c.Interior.Color = Excel.XlRgbColor.rgbGrey;
                #endregion
                while (currentDate <= day.AddDays(1))
                {
                    //value = d.GetSingleHalfhour(deviceCode, sensorCode, currentDate);
                    value = 1;
                    if (value < 0)
                        ws1.Cells[currentRow, currentColumn] = "";
                    else
                        ws1.Cells[currentRow, currentColumn] = value;
                    currentDate = currentDate.AddMinutes(30);
                    currentRow++;
                    completed++;
                    percent = 100 * completed / totalData;
                    Console.CursorLeft = cursorPosition;
                    Console.Write(percent.ToString() + "%");
                }
                currentColumn++;
            }
            ws1.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ws1.Activate();
            c.Select();
            Excel.Windows xlsWindows = wb.Windows;
            Excel.Window xlsWindow = xlsWindows[1];
            //xlsWindow.FreezePanes = true;
            c = (Excel.Range)ws1.Cells[firstRow, 3];
            ws1.Activate();
            c.Select();
            //xlsWindow = xlsWindows[1];
            xlsWindow.FreezePanes = true;
            c = (Excel.Range)ws1.Cells[firstRow - 1, 3];
            c = c.Resize[1, totalSensors];
            c.NumberFormat = "#,##0";
            xlsWindow.Activate();
            if (File.Exists(fileName))
                File.Delete(fileName);
            wb.SaveAs(fileName);
            wb.Close();
            xls.Quit();
            releaseObject(ws1);
            releaseObject(wb);
            releaseObject(xls);
        }

        private static void ToXLS12(string[] devices, string[] sensors, DateTime day, string fileName)
        {
            DataProvider d = new DataProvider(settings.Server, settings.Database, settings.UserName, settings.Password);
            Excel.Range c;
            Excel.Application xls;
            Excel.Workbook wb;
            int firstRow = 4;
            int totalSensors = sensors.Length;
            int totalRows;
            int totalData;
            string deviceCode;
            string sensorCode;
            double value;
            xls = new Excel.Application();
            xls.SheetsInNewWorkbook = 1;
            wb = xls.Workbooks.Add();
            Excel.Worksheet ws1 = (Excel.Worksheet)wb.Worksheets[1];
            ws1.Name = "Показания";
            #region Prepare headers
            c = (Excel.Range)(ws1.Cells[1, 1]);
            c.Value = "Дата:";
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            c = (Excel.Range)(ws1.Cells[1, 2]);
            c.Value = day.ToShortDateString();
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            c = (Excel.Range)(ws1.Cells[firstRow - 1, 1]);
            c.Value = "Дата";
            c.ColumnWidth = 12;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            c = (Excel.Range)(ws1.Cells[firstRow - 1, 2]);
            c.Value = "Время";
            c.ColumnWidth = 13;
            c.Interior.Color = Excel.XlRgbColor.rgbGrey;
            #endregion

            totalRows = 48;
            totalData = totalRows * totalSensors;
            string deviceName, sensorName;
            DateTime currentDate = day;

            #region Write dates and times into two leftmost coumns
            int currentRow = 0;
            int currentColumn = 3;
            totalRows = 48;
            string[,] leftColumns = new string[totalRows, 2];
            while (currentDate < day.AddDays(1))
            {
                leftColumns[currentRow, 0] =
                    currentDate.Date.ToShortDateString();
                leftColumns[currentRow, 1] =
                    string.Format("{0:00}:{1:00} - {2:00}:{3:00}",
                                  currentDate.TimeOfDay.Hours,
                                  currentDate.TimeOfDay.Minutes,
                                  currentDate.AddMinutes(30).TimeOfDay.Hours,
                                  currentDate.AddMinutes(30).TimeOfDay.Minutes);
                currentDate = currentDate.AddMinutes(30);
                currentRow++;
            }
            c = (Excel.Range)ws1.Cells[firstRow, 1];
            c = c.Resize[totalRows, 2];
            c.Value = leftColumns;
            #endregion


            #region Write channel names in first
            for (int i = 0; i < sensors.Length; i++)
            {
                deviceCode = devices[i];
                sensorCode = sensors[i];
                // halfhours column headers
                c = (Excel.Range)(ws1.Cells[1, currentColumn]);
                c.ColumnWidth = 18;
                deviceName = d.GetDeviceName(deviceCode);
                c.Value = deviceName;
                sensorName = d.GetSensorName(deviceCode, sensorCode);
                ws1.Cells[2, currentColumn] = sensorName;
                c = (Excel.Range)(ws1.Cells[firstRow - 1, currentColumn]);
                c.FormulaR1C1 = string.Format("=SUM(R[1]C:R[{0}]C)/2", totalRows);
                c.Font.Bold = true;
                c.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                c.Interior.Color = Excel.XlRgbColor.rgbGrey;
                currentColumn++;
            }
            #endregion
            int cursorPosition = Console.CursorLeft - 2;
            System.Data.DataTable halfhours = d.GetAllHalfhours(devices, sensors, day.AddMinutes(30), day.AddDays(1));
            string[,] allValues = new string[totalRows, totalSensors];
            currentColumn = 0;
            currentRow = 0;
            foreach (System.Data.DataRow row in halfhours.Rows)
            {
                for(int col=0;col<totalSensors;col++)
                {
                    allValues[currentRow, col] = row[col + 1].ToString().Replace(',', '.');
                }
                currentRow++;
            }
            c = (Excel.Range)ws1.Cells[firstRow, 3];
            c = c.Resize[totalRows, totalSensors];
            c.Value = allValues;
            ws1.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ws1.Activate();
            c.Select();
            Excel.Windows xlsWindows = wb.Windows;
            Excel.Window xlsWindow = xlsWindows[1];
            //xlsWindow.FreezePanes = true;
            c = (Excel.Range)ws1.Cells[firstRow, 3];
            ws1.Activate();
            c.Select();
            //xlsWindow = xlsWindows[1];
            xlsWindow.FreezePanes = true;
            c = (Excel.Range)ws1.Cells[firstRow - 1, 3];
            c = c.Resize[1, totalSensors];
            c.NumberFormat = "#,##0";
            xlsWindow.Activate();
            if (File.Exists(fileName))
                File.Delete(fileName);
            wb.SaveAs(fileName);
            wb.Close();
            xls.Quit();
            releaseObject(ws1);
            releaseObject(wb);
            releaseObject(xls);
            Console.WriteLine();
            Console.WriteLine("OK");
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Failed to release object" + obj.ToString(), ex);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static int GetRowNumber(DateTime dtStart, DateTime dtCurrent, bool halves, int offset)
        {
            int result = 0;
            TimeSpan dif = (dtCurrent - dtStart);
            result = (int)dif.TotalMinutes;
            if (halves)
                result = result / 30;
            else
                result = result / 60;
            result += (offset - 1);
            return result;
        }


        /*
         1) XMLimport
После обработки xml проверяет, есть ли ИНН в списке экспорта. Если есть, запускает

"c:\ExportXLS\ExportXLS.exe inn day"

где аргументы командной строки inn и day взяты из только что обработанной xml.

2) ExportXLS
Два режима работы. Первый = с ключом /xml
Тогда в качестве аргумента командной строки передаётся имя xml-файла, в котором приведены все нужные сведения для данной конкретной 
задачи: коды устройств и каналов, адреса электронной почты.
Второй режим - без ключа.
Тогда должно быть два аргумента в командной строке: ИНН и дата.
В папке с программой лежит большая xml-ка со всеми задачами. Из неё выбираем все задачи с переданным ИНН (их может быть больше 1) 
и начинаем выполнять по очереди. В каждой задаче те же сведения, что и в отдельных xml-ках в первом режиме.
Еще нужно добавить ключ /?

XML должны выглядеть так:

<?xml version="1.0"?>
<Export>
  <job inn="1111111111"	name="СКЖД" description="отправка Т-304 в ВЭС">
    <email>cdsir@ves.stavre.elektra.ru</email>
    <email>staskue@stavre.elektra.ru</email>
    <codes>
      <device code="4464">
        <sensor code="191"/>
        <sensor code="192"/>
        <sensor code="193"/>
        <sensor code="194"/>
      </device>
      <device code="4444">
        <sensor code="1"/>
        <sensor code="2"/>
      </device>
    </codes>
  </job>
</Export>
         */

    }
}
