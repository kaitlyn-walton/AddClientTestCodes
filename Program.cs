using AddClientTestCodes.DTO;
using AddClientTestCodes.Helper;
using AddClientTestCodes.Repository;

namespace AddClientTestCodes
{
    public class Program
    {
        public static void Main()
        {
            ExcelHelper excelHelper = new ExcelHelper();
            ClientTestCodeRepository clientTestCodeRepository = new ClientTestCodeRepository();

            while (true)
            {
                Console.WriteLine("Type 'exit' at anytime to stop the program...");

                var hospital = clientTestCodeRepository.GetHospital();
                if (hospital == "")
                {
                    return;
                }
                else
                {
                    Console.WriteLine("Hospital Code: " + hospital + "\n");

                    var testCode = clientTestCodeRepository.GetTestCode();
                    if (testCode == 0)
                    {
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Test code: " + testCode + "\n");

                        var customClientTestCode = hospital + "-" + testCode;

                        List<ExcelSheetDTO>? records = excelHelper.GetExcelFile(customClientTestCode);

                        if (records == null)
                        {
                            return;
                        }
                        else
                        {
                            foreach (ExcelSheetDTO record in records)
                            {
                                var update = clientTestCodeRepository.FindTestCodes(record.IncompleteTestCode);

                                record.CustomTestCode = update;

                                Console.WriteLine("\n" + "New Client Test Code: " + record.CustomTestCode);

                                var insert = clientTestCodeRepository.InsertClientTestCode(hospital, testCode, record);
                                Console.WriteLine("New Id:" + insert + "\n");
                            }
                        }
                    }
                }
            }
        }
    }
}