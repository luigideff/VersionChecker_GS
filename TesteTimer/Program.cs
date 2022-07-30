using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Timers;
using System.Configuration;
using System.Diagnostics;

namespace Main
{
    class Program
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        //  private const string SpreadsheetId = "1Wz3oQ19JQTBDAlhJLBVZBpqFcYOffpGp3iLPpnmAync";
        private static string SpreadsheetId = ConfigurationManager.AppSettings["SpreadsheetID"];
        private static string InitialColumn = ConfigurationManager.AppSettings["InitialColumn"];
        private static string FinalColumn = ConfigurationManager.AppSettings["FinalColumn"];
        private const string GoogleCredentialsFileName = "gscred.json";
        private const string ReadRange = "A:B";

       
        public static async Task Main(string[] args)
        {

            Console.WriteLine("Atualizando planilha de versão dos clientes...");            

            await AtualizarPlanilha();
            //SetTimer();
            Console.WriteLine("Planilha atualizada");
            //Console.ReadLine();
        }

        static async Task LerPlanilha()
        {

            var serviceValues = GetSheetsService().Spreadsheets.Values;
            Console.Clear();
            await ReadAsync(serviceValues);
            await WriteAsync(serviceValues);
            Console.WriteLine("--------FIM DA PLANILHA---------");
        }

        static async Task AtualizarPlanilha()
        {

            var serviceValues = GetSheetsService().Spreadsheets.Values;
            await WriteAsync(serviceValues);
        }



        private static SheetsService GetSheetsService()
        {
            using (var stream =
                new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)

                };
                return new SheetsService(serviceInitializer);
            }
        }

        private static async Task ReadAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values;

            if (values == null || !values.Any())
            {
                Console.WriteLine("No data found.");
                return;
            }

            //var header = string.Join(" ", values.First().Select(r => r.ToString()));
            //Console.WriteLine($"Header: {header}");

            foreach (var row in values)
            {
                var res = string.Join(" - ", row.Select(r => r.ToString()));
                Console.WriteLine(res);
            }
        }



        private static async Task WriteAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            List<string> paths = File.ReadAllLines("paths.txt").ToList();

            int line = 2;
                foreach (string path in paths)
            {
                string[] vect = path.Split(';');
                if (GetFileVersion(vect[1]) != "O caminho especificado não aponta para um arquivo válido.")
                {

                    string WriteRange = $"{InitialColumn}{line}:{FinalColumn}{line}";
                    var valueRange = new ValueRange
                    {
                        Values = new List<IList<object>> { new List<object>
            {
                vect[0],
                GetFileVersion(vect[1]),
                GetFileModDate(vect[1])
            } }
                    };
                    var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    var response = await update.ExecuteAsync();

                }
                line++;

            }


        }



        public static string GetFileVersion(string filepath)
        {
            try
            {
                FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(filepath);
                string FileVersion = Convert.ToString(fileVersionInfo.FileVersion);
                return FileVersion;
            }

            catch (IOException)
            {

                return "O caminho especificado não aponta para um arquivo válido.";
            }
            catch (ArgumentException)
            {
                return String.Empty;
            }

        }

        public static string GetFileModDate(string fileName)
        {
            try
            {
                string ModDate = Convert.ToString(File.GetLastWriteTime(fileName));
                if (ModDate == "31-Dec-00 22:00:00")
                {
                    return "O caminho especificado não aponta para um arquivo válido";
                }
                else
                {
                    return ModDate;
                }

            }
            catch (IOException)
            {
                return "O caminho especificado não aponta para um arquivo válido";
            }
            catch (ArgumentException)
            {
                return String.Empty;
            }
        }
    }
}








