using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TAR解碼
{
    class Program
    {
        public static void Main(string[] args)
        {
            string allSapPath = @"C:\\Users\\Kenneth\\Downloads\\今天比完\\" + DateTime.Now.ToString("yyyyMMdd") + "AllSap.csv";
            var dateList = GetDateList(DateTime.Now,0);

            foreach (var date in dateList)
            {
                string dailyDataPath = @"C:\\Users\\Kenneth\\Downloads\\今天比完\\" + date + " - 統一超商.csv";
                string resultDirectory = @"C:\Users\Kenneth\Downloads\今天比完\" + date + "Result\\";
                string resultPath = resultDirectory + date + " - 結果.csv";

                //get AllSap
                var todayAllSapRecords = new List<AllSap>();
                todayAllSapRecords = GetAllSaps(allSapPath);
                //get DailyChangeRecord
                var dailyChangeRecords = new List<DailyChangeRecord>();
                
                    dailyChangeRecords = GetDailyChangeRecords(dailyDataPath);

                    ////get每日異動檔(統一超商.csv)
                    //var persons = new List<Person>();
                    //persons = GetPersonFromDaliyChangeRecords(dailyDataPath);

                    //Compare Result
                    CompareResultAndExport(resultPath, resultDirectory, todayAllSapRecords, dailyChangeRecords);
                    //write Result
                    //StreamWriterMetod(resultPath, persons, allSapRecords);
                
            }

        }

        public static void StreamWriterMetod(string destinationPath, List<Person> persons, List<AllSap> allSapRecords)
        {
            if (Directory.Exists(destinationPath))
                Directory.Delete(destinationPath);

            try
            {
                using (StreamWriter swWriter = new StreamWriter(destinationPath, true, Encoding.UTF8))
                {
                    swWriter.WriteLine($"帳號," +
                                $"姓名," +
                                $"公司," +
                                $"職稱," +
                                $"部門," +
                                $"薪呈");
                    //依全建檔取得每日異動檔
                    foreach (var person in persons)
                    {
                        var personDataInAllSap = allSapRecords.Where(item => item.PersonId == person.PersonId && item.PersonName == person.PersonName).FirstOrDefault();
                        if (personDataInAllSap != null)
                            swWriter.WriteLine($"{"P" + personDataInAllSap.PersonId.PadLeft(8, '0')}," +
                                $"{personDataInAllSap.PersonName.Trim()}," +
                                $"{personDataInAllSap.Company.Trim()}," +
                                $"{personDataInAllSap.Title.Trim()}," +
                                $"{personDataInAllSap.Department.Trim()}," +
                                $"{personDataInAllSap.SalaryCode.Trim()}");
                        else
                            swWriter.WriteLine($"{"P" + person.PersonId.PadLeft(8, '0')}," +
                               $"{person.PersonName.Trim()}," +
                               $"{string.Empty}," +
                               $"{string.Empty}," +
                               $"{person.Department.Trim()}," +
                               $"{string.Empty}");
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static void CompareResultAndExport(string destinationPath, List<DailyChangeRecord> todayADRecords, List<DailyChangeRecord> dailyChangeRecords)
        {
            if (Directory.Exists(destinationPath))
                Directory.Delete(destinationPath);

            try
            {
                using (StreamWriter swWriter = new StreamWriter(destinationPath, true, Encoding.UTF8))
                {
                    swWriter.WriteLine($"帳號," +
                                $"姓名," +
                                $"狀態," +
                                $"公司," +
                                $"職稱," +
                                $"部門," +
                                $"部門," +
                                $"Remark");
                    //依全建檔取得每日異動檔
                    foreach (var dailyChangeRecord in dailyChangeRecords)
                    {
                        //var check = new DailyChangeRecord();
                        //var test = todayADRecords.Where(item => item.PersonId == "P10312710" && dailyChangeRecord.PersonId == "P10312710").FirstOrDefault();
                        //if (test != null)
                        //    check = test;

                        var todayADRecord = todayADRecords.Where(item => item.PersonId == dailyChangeRecord.PersonId).FirstOrDefault();
                        if (todayADRecord != null)
                        {
                            var differences = "";
                            if (dailyChangeRecord.SalaryCode.Substring(0, 1) != todayADRecord.SalaryCode)
                                differences += string.Format("AD薪呈為:{0};", todayADRecord.SalaryCode);
                            if (dailyChangeRecord.Title != todayADRecord.Title)
                                differences += string.Format("AD職稱為:{0};", todayADRecord.Title);
                            if (dailyChangeRecord.Department != todayADRecord.Department)
                                differences += string.Format("AD部門為:{0};", todayADRecord.Department);
                            if (dailyChangeRecord.PersonName != todayADRecord.PersonName)
                                differences += string.Format("AD名稱為:{0};", todayADRecord.PersonName);
                            differences = string.IsNullOrEmpty(differences) ? "沒異動" : differences;
                            swWriter.WriteLine($"{ dailyChangeRecord.PersonId}," +
                                $"{dailyChangeRecord.PersonName.Trim()}," +
                                $"{dailyChangeRecord.Status.Trim()}," +
                                $"{dailyChangeRecord.Company.Trim()}," +
                                $"{dailyChangeRecord.Title.Trim()}," +
                                $"{dailyChangeRecord.Department.Trim()}," +
                                $"{dailyChangeRecord.SalaryCode.Trim()}," +
                                $"{differences.Trim()}");
                        }
                        else
                            swWriter.WriteLine($"{ dailyChangeRecord.PersonId}," +
                               $"{dailyChangeRecord.PersonName.Trim()}," +
                               $"{dailyChangeRecord.Status.Trim()}," +
                               $"{dailyChangeRecord.Company.Trim()}," +
                               $"{dailyChangeRecord.Title.Trim()}," +
                               $"{dailyChangeRecord.Department.Trim()}," +
                               $"{dailyChangeRecord.SalaryCode.Trim()}," +
                               $"不存在於AD");
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static void CompareResultAndExport(string destinationPath, string destinationDirectory, List<AllSap> todayAllSapRecords, List<DailyChangeRecord> dailyChangeRecords)
        {
            if (Directory.Exists(destinationDirectory))
                Directory.Delete(destinationDirectory, true);

            Directory.CreateDirectory(destinationDirectory);

            try
            {
                using (StreamWriter swWriter = new StreamWriter(destinationPath, true, Encoding.UTF8))
                {
                    swWriter.WriteLine($"員編," +
                                $"姓名," +
                                $"狀態," +
                                $"公司," +
                                $"職稱," +
                                $"部門," +
                                $"職員群組," +
                                $"地區," +
                                $"薪層," +
                                $"到職日期," +
                                $"離職日期," +
                                $"更新日," +
                                $"Remark");
                    //依全建檔取得每日異動檔
                    foreach (var dailyChangeRecord in dailyChangeRecords)
                    {
                        var DaPer = dailyChangeRecord.PersonId.PadLeft(8, '0');
                        var todayADRecord = todayAllSapRecords.Where(item => item.PersonId == dailyChangeRecord.PersonId.PadLeft(8, '0')).FirstOrDefault();
                        if (todayADRecord != null)
                        {
                            var differences = "";
                            if (dailyChangeRecord.SalaryCode != todayADRecord.SalaryCode)
                                differences += string.Format("AllSap薪呈為:{0};", todayADRecord.SalaryCode);
                            if (dailyChangeRecord.Title != todayADRecord.Title)
                                differences += string.Format("AllSap職稱為:{0};", todayADRecord.Title);
                            if (dailyChangeRecord.Department != todayADRecord.Department)
                                differences += string.Format("AllSap部門為:{0};", todayADRecord.Department);
                            if (dailyChangeRecord.PersonName != todayADRecord.PersonName)
                                differences += string.Format("AllSap姓名為:{0};", todayADRecord.PersonName);

                            differences = string.IsNullOrEmpty(differences) ? "沒異動" : differences;
                            swWriter.WriteLine($"{ dailyChangeRecord.PersonId}," +
                                $"{dailyChangeRecord.PersonName.Trim()}," +
                                $"{dailyChangeRecord.Status.Trim()}," +
                                $"{dailyChangeRecord.Company.Trim()}," +
                                $"{dailyChangeRecord.Title.Trim()}," +
                                $"{dailyChangeRecord.Department.Trim()}," +
                                $"{dailyChangeRecord.EmployeeGroup.Trim()}," +
                                $"{dailyChangeRecord.Zone.Trim()}," +
                                $"{dailyChangeRecord.SalaryCode.Trim()}," +
                                $"{dailyChangeRecord.ArrivalDate.Trim()}," +
                                $"{dailyChangeRecord.ResignDate.Trim()}," +
                                $"{dailyChangeRecord.UpdateDate.Trim()}," +
                                $"{differences.Trim()}");
                        }
                        else
                            swWriter.WriteLine($"{ dailyChangeRecord.PersonId}," +
                               $"{dailyChangeRecord.PersonName.Trim()}," +
                               $"{dailyChangeRecord.Status.Trim()}," +
                               $"{dailyChangeRecord.Company.Trim()}," +
                               $"{dailyChangeRecord.Title.Trim()}," +
                               $"{dailyChangeRecord.Department.Trim()}," +
                               $"{string.Empty}," +
                               $"{string.Empty}," +
                               $"{dailyChangeRecord.SalaryCode.Trim()}," +
                               $"{dailyChangeRecord.ArrivalDate.Trim()}," +
                               $"{dailyChangeRecord.ResignDate.Trim()}," +
                               $"{dailyChangeRecord.UpdateDate.Trim()}," +
                               $"不存在於AllSap");
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static List<AllSap> GetAllSaps(string sourcePath)
        {
            //讀AllSap
            var allSapRecords = new List<AllSap>();
            var allSaplineCount = 0;
            string rougthAllSapData = "";
            using (StreamReader allsapReader = new StreamReader(sourcePath, Encoding.UTF8))
            {
                while (!allsapReader.EndOfStream)
                {
                    rougthAllSapData = allsapReader.ReadLine();
                    var allSapRecord = rougthAllSapData.Split('\t');
                    if (allSaplineCount >= 1)
                        allSapRecords.Add(
                            AllSap.createInstnace(
                                allSapRecord[0],
                                allSapRecord[1],
                                allSapRecord[2],
                                allSapRecord[3],
                                allSapRecord[4],
                                allSapRecord[5],
                                allSapRecord[6],
                                allSapRecord[7]
                                ));
                    allSaplineCount++;
                }
                allsapReader.Close();
            }
            return allSapRecords;
        }

        public static List<DailyChangeRecord> GetDailyChangeRecords(string sourcePath)
        {
            //讀AllSap
            var dailyChangeRecords = new List<DailyChangeRecord>();
            var dailyChangeRecordCount = 0;
            string rougthAllSapData = "";
            using (StreamReader allsapReader = new StreamReader(sourcePath, Encoding.UTF8))
            {
                while (!allsapReader.EndOfStream)
                {
                    rougthAllSapData = allsapReader.ReadLine();
                    var allSapRecord = rougthAllSapData.Split('\t');
                    if (dailyChangeRecordCount >= 3)
                        dailyChangeRecords.Add(
                            DailyChangeRecord.createInstnace(
                                allSapRecord[0],
                                allSapRecord[1],
                                allSapRecord[2],
                                allSapRecord[3],
                                allSapRecord[4],
                                allSapRecord[5],
                                allSapRecord[6],
                                allSapRecord[7],
                                allSapRecord[8],
                                allSapRecord[9],
                                allSapRecord[10],
                                allSapRecord[11]
                                ));
                    dailyChangeRecordCount++;
                }
                allsapReader.Close();
            }
            return dailyChangeRecords;
        }

        public static List<Person> GetPersonFromDaliyChangeRecords(string sourcePath)
        {
            var personFromDaliylineCount = 0;
            string rougthPersonFromDaliyChangeRecord = "";
            var personFromDaliyChangeRecords = new List<Person>();
            using (StreamReader dailyDataReader = new StreamReader(sourcePath, Encoding.UTF8))
            {
                while (!dailyDataReader.EndOfStream)
                {
                    rougthPersonFromDaliyChangeRecord = dailyDataReader.ReadLine();
                    var personFromDaliyChangeRecord = rougthPersonFromDaliyChangeRecord.Split('\t');
                    if (personFromDaliylineCount >= 3)
                        personFromDaliyChangeRecords.Add(Person.createInstnace(
                            personFromDaliyChangeRecord[0],
                            personFromDaliyChangeRecord[1],
                            personFromDaliyChangeRecord[6]));
                    personFromDaliylineCount++;
                }
                dailyDataReader.Close();
            }
            return personFromDaliyChangeRecords;
        }

        public static List<string> GetDateList(DateTime targetDate, int countDayBefore = 0)
        {
            var dateList = new List<string>();
            for (int i = countDayBefore; i <= 0; i++)
                dateList.Add(targetDate.AddDays(i).ToString("yyyyMMdd"));
            return dateList;
        }

        public class Person
        {
            public string PersonId { get; set; }
            public string PersonName { get; set; }
            public string Department { get; set; }

            public static Person createInstnace(string PersonId, string PersonName, string Department)
            {
                return new Person
                {
                    PersonId = PersonId,
                    PersonName = PersonName,
                    Department = Department
                };
            }
        }

        public class DailyChangeRecord
        {
            public string PersonId { get; set; }
            public string PersonName { get; set; }
            public string Status { get; set; }
            public string Company { get; set; }
            public string Title { get; set; }
            public string Department { get; set; }
            public string EmployeeGroup { get; set; }
            public string Zone { get; set; }
            public string SalaryCode { get; set; }
            public string ArrivalDate { get; set; }
            public string ResignDate { get; set; }
            public string UpdateDate { get; set; }
            public string Remark { get; set; }

            public static DailyChangeRecord createInstnace(
                string PersonId, string PersonName, string Status,
                string Company, string Title, string Department,
                string EmployeeGroup, string Zone, string SalaryCode,
                string ArrivalDate, string ResignDate, string UpdateDate)
            {
                return new DailyChangeRecord
                {
                    PersonId = PersonId,
                    PersonName = PersonName,
                    Status = Status,
                    Company = Company,
                    Title = Title,
                    Department = Department,
                    EmployeeGroup = EmployeeGroup,
                    Zone = Zone,
                    SalaryCode = SalaryCode,
                    ArrivalDate = ArrivalDate,
                    ResignDate = ResignDate,
                    UpdateDate = UpdateDate
                };
            }
        }

        public class AllSap
        {
            public string PersonId { get; set; }
            public string PersonName { get; set; }
            public string Company { get; set; }
            public string Title { get; set; }
            public string Department { get; set; }
            public string SalaryType { get; set; }
            public string Area { get; set; }
            public string SalaryCode { get; set; }

            public static AllSap createInstnace(string PersonId, string PersonName, string Company, string Title, string Department, string SalaryType, string Area, string SalaryCode)
            {
                return new AllSap
                {
                    PersonId = PersonId,
                    PersonName = PersonName,
                    Company = Company,
                    Title = Title,
                    Department = Department,
                    SalaryType = SalaryType,
                    Area = Area,
                    SalaryCode = SalaryCode,
                };
            }
        }
    }
}

