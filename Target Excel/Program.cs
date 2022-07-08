using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Target_Excel
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //create file location for excel files
            var file = new FileInfo(@"D:\test\Reference.xlsx"); //LOCATION OF REFERENCE EXCEL SHEET
            var file2 = new FileInfo(@"D:\test\Print.xlsx"); //LOCATION OF NEW FILLED EXCEL SHEET

            var people = GetSetupData();

            await SaveExcelFile(people, file, file2);
        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file, FileInfo file2)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                ExcelWorkbook wb = excelPackage.Workbook;
                ExcelWorksheet excelWorksheet = wb.Worksheets.First();
                var range = excelWorksheet.Cells["A3"].LoadFromCollection(people, false);
                range.AutoFitColumns();
                DeleteIfExists(file2);
                await excelPackage.SaveAsAsync(file2);
            }
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if(file.Exists)
            {
                file.Delete();
            }
        }
        private static List<PersonModel> GetSetupData()
        {
            //Console ReadLine WriteLine
            Console.WriteLine("Team Member Name:");
            string TM = Console.ReadLine();
            Console.WriteLine("Box Cutter #:");
            string BC = Console.ReadLine();
            Console.WriteLine("Walkie #:");
            string W = Console.ReadLine();
            Console.WriteLine("Your entering shift:");
            string BCTO = Console.ReadLine();
            string WTO = BCTO;
            string MYTO = WTO;
            string PTO = BCTO;
            Console.WriteLine("Your ending shift:");
            string BCTI = Console.ReadLine();
            string WTI = BCTI;
            string MYTI = BCTI;
            string PTI = BCTI;
            Console.WriteLine("My Device PDA #:");
            string MYD = Console.ReadLine();
            Console.WriteLine("Printer #:");
            string P = Console.ReadLine();

            List<PersonModel> output = new()
            {
                new()
                {
                    TeamMember = TM,
                    BoxCutter = BC,
                    BoxTimeOut = BCTO,
                    BoxTimeIn = BCTI,
                    Walkie = W,
                    WalkieTimeOut = WTO,
                    WalkieTimeIn = WTI,
                    MyDevicePda = MYD,
                    MyDeviceTimeOut = MYTO,
                    TEST = WTO,
                    MyDeviceTimeIn = MYTI,
                    PrinterNum = P,
                    PrinterTimeOut = PTO,
                    PrinterTimeIn = PTI,

                }
            };

            Console.WriteLine("Would you like to continue?");
            string answr = Console.ReadLine();
            if (answr == "Y")
            {
                while (answr == "Y")
                {
                    Console.Clear();
                    //Console ReadLine WriteLine
                    Console.WriteLine("Team Member Name:");
                    TM = Console.ReadLine();
                    Console.WriteLine("Box Cutter #:");
                    BC = Console.ReadLine();
                    Console.WriteLine("Walkie #:");
                    W = Console.ReadLine();
                    Console.WriteLine("Your entering shift:");
                    BCTO = Console.ReadLine();
                    WTO = BCTO;
                    MYTO = BCTO;
                    PTO = BCTO;
                    Console.WriteLine("Your ending shift:");
                    BCTI = Console.ReadLine();
                    WTI = BCTI;
                    MYTI = BCTI;
                    PTI = BCTI;
                    Console.WriteLine("My Device PDA #:");
                    MYD = Console.ReadLine();
                    Console.WriteLine("Printer #:");
                    P = Console.ReadLine();
                    output.Add(new PersonModel
                    {
                        TeamMember = TM,
                        BoxCutter = BC,
                        BoxTimeOut = BCTO,
                        BoxTimeIn = BCTI,
                        Walkie = W,
                        WalkieTimeOut = WTO,
                        WalkieTimeIn = WTI,
                        MyDevicePda = MYD,
                        MyDeviceTimeOut = MYTO,
                        TEST = MYTO,
                        MyDeviceTimeIn = MYTI,
                        PrinterNum = P,
                        PrinterTimeOut = PTO,
                        PrinterTimeIn = PTI,

                    });
                    Console.WriteLine("Would you like to continue?");
                    answr = Console.ReadLine();
                    
                }
            }
            return output;
        }
    }
}
