using System.Data;
using AddClientTestCodes.DTO;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace AddClientTestCodes.Repository
{
    public class ClientTestCodeRepository
    {
        private IConfiguration _config;
        private string? _connectionString;
        private IDbConnection _dbConnection;
        public ClientTestCodeRepository()
        {
            _config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            _connectionString = _config.GetConnectionString("DefaultConnection");
            _dbConnection = new SqlConnection(_connectionString);
        }

        //check to make sure hospital exists in our records
        public string GetHospital()
        {
            while (true)
            {
                try
                {
                    Console.Write("\n" + "Enter a hospital code: ");
                    var hospitalCode = Console.ReadLine();

                    if (hospitalCode?.ToLower() == "exit")
                    {
                        return "";
                    }
                    else
                    {
                        string checkIfExists = @"SELECT HospitalCode FROM dbo.Hospital WHERE HospitalCode = '" + hospitalCode + "'";
                        var code = _dbConnection.QuerySingle<string>(checkIfExists);


                        return code.ToUpper();
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Hospital code does not exist");
                    continue;
                }

            }
        }

        //check to make sure test code exists in our records
        public int GetTestCode()
        {
            while (true)
            {
                try
                {
                    Console.Write("Enter a test code: ");
                    var testCode = Console.ReadLine();

                    if (testCode?.ToLower() == "exit")
                    {
                        return 0;
                    }
                    else
                    {
                        var convertCode = Convert.ToInt32(testCode);

                        string checkIfExists = @"SELECT TestCode FROM dbo.AllTestCodes WHERE TestCode = '" + convertCode + "'";
                        var code = _dbConnection.QuerySingle<int>(checkIfExists);

                        return code;
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Invalid code.");
                    continue;
                }
            }
        }

        //find any duplicate client test codes and then increment our client test code by 1
        public string FindTestCodes(string custom)
        {
            string findDuplicateTestCodes = @"SELECT ClientTestCode FROM dbo.ClientTestCodes WHERE ClientTestCode LIKE '%" + custom + "%'";

            List<string>? duplicateCodes = _dbConnection.Query<string>(findDuplicateTestCodes).ToList();

            int newVersion;

            if (duplicateCodes.Count > 0)
            {
                var versions = new int[duplicateCodes.Count];

                for (int i = 0; i < duplicateCodes.Count; i++)
                {
                    var split = duplicateCodes[i].Split("-");
                    var num = Convert.ToInt32(split[3]);
                    versions[i] = num;
                }
                var lastVersion = versions.Max();
                newVersion = lastVersion + 1;
            }
            else
            {
                newVersion = 1;
            }

            var newTestCode = custom + newVersion;

            return newTestCode;
        }

        //create new client test code
        public int InsertClientTestCode(string hospital, int testCode, ExcelSheetDTO record)
        {

            string addTestCode = @"INSERT INTO dbo.ClientTestCodes
                                                    (HospitalCode,
                                                     TestCode,
                                                     ClientTestCode,
                                                     Name,
                                                     CreatedOnDateTime)
                                            OUTPUT Inserted.Id
                                            VALUES ('" + hospital + "','"
                                                        + testCode + "','"
                                                        + record.CustomTestCode + "','"
                                                        + record.PanelName + "','"
                                                        + DateTime.Now + "')";

            var newId = _dbConnection.QuerySingle<int>(addTestCode);

            InsertClientTestCodeGenes(newId, record.GenesInPanel);

            return newId;
        }

        //insert client test code genes based on the new Id that is returned
        public void InsertClientTestCodeGenes(int newId, List<string>? genes)
        {
            int recordsAdded = 0;

            if (genes != null)
            {
                foreach (string aGene in genes)
                {
                    string addGenes = @"INSERT INTO dbo.ClientTestCodeGenes 
                                                        (ClientTestCodeId,
                                                        Gene)
                                                VALUES ('" + newId + "','" + aGene + "')";

                    int addSingleGene = _dbConnection.Execute(addGenes);

                    recordsAdded++;
                }
            }
            Console.WriteLine("# of genes added: " + recordsAdded);
        }
    }
}