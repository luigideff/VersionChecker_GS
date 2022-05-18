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
using System.Diagnostics;

namespace Main
{
    class Program
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string SpreadsheetId = "1Wz3oQ19JQTBDAlhJLBVZBpqFcYOffpGp3iLPpnmAync";
        private const string GoogleCredentialsFileName = "gscred.json";
        private const string ReadRange = "A:B";
        private const string CORECONRJ = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONRJ01\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONPR = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONPR06\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONDF = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONDF11\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONSC = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONSC07\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONBA = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONBA05\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONGO = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONGO18\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONRR = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONRR27\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONRS = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONRS04\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CORECONPE = @"C:\HBSIS\Sistemas\COFECON\Producao\BRCTotal_CORECONPE03\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CREARN = @"C:\HBSIS\Sistemas\BackOffice_CREARN\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string CREFITOSE = @"C:\HBSIS\Sistemas\BRConselhos_CREFITO17_SE\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private const string ONR = @"C:\HBSIS\Sistemas\ONR\Producao\BRCTotal_ONR\bin\HBSIS.Conselho.BLL.Financeiro.dll";
        private static Timer TicTock;

        public static async Task Main(string[] args)
        {
            Console.WriteLine("Atualizando planilha de versão dos clientes...");
            await AtualizarPlanilha();
            //SetTimer();
            //Console.WriteLine("Planilha atualizada");
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
            string WriteRange = "A2:C2";
            var valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> 
            { 
                "CORECON RJ", 
                GetFileVersion(CORECONRJ),
                GetFileModDate(CORECONRJ) 
            } } };
            var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            var response = await update.ExecuteAsync();

            
            WriteRange = "A3:C3";
            valueRange = new ValueRange { Values = new List<IList<object>> { new List<object> 
            { 
                "CORECON PR", 
                GetFileVersion(CORECONPR),
                GetFileModDate(CORECONPR)
            } } };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A4:C4";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON DF",
                GetFileVersion(CORECONDF),
                GetFileModDate(CORECONDF)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A5:C5";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON SC",
                GetFileVersion(CORECONSC),
                GetFileModDate(CORECONSC)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A6:C6";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON BA",
                GetFileVersion(CORECONBA),
                GetFileModDate(CORECONBA)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A7:C7";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON GO",
                GetFileVersion(CORECONGO),
                GetFileModDate(CORECONGO)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A8:C8";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON RR",
                GetFileVersion(CORECONRR),
                GetFileModDate(CORECONRR)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A9:C9";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON RS",
                GetFileVersion(CORECONRS),
                GetFileModDate(CORECONRS)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A10:C10";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CORECON PE",
                GetFileVersion(CORECONPE),
                GetFileModDate(CORECONPE)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A11:C11";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CREA-RN",
                GetFileVersion(CREARN),
                GetFileModDate(CREARN)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A12:C12";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "CREFITO-SE",
                GetFileVersion(CREFITOSE),
                GetFileModDate(CREFITOSE)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();

            WriteRange = "A13:C13";
            valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>
            {
                "ONR",
                GetFileVersion(ONR),
                GetFileModDate(ONR)
            } }
            };
            update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            response = await update.ExecuteAsync();



        }
        private static void SetTimer()
        {
            
            TicTock = new Timer(10000);
            
            TicTock.Elapsed += OnTimedEvent;
            TicTock.AutoReset = true;
            TicTock.Enabled = true;
        }

        public static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            AtualizarPlanilha();
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








