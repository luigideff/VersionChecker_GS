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
using System.Runtime.InteropServices;

namespace Main
{
    class Program
    {
        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string SpreadsheetId = "1Wz3oQ19JQTBDAlhJLBVZBpqFcYOffpGp3iLPpnmAync";
        private const string GoogleCredentialsFileName = "gscred.json";
        private const string ReadRange = "A:B";
        private const string Cliente1 = @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenvdesc.dll";
        private const string Cliente2 = @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenvdesc.dll";
        private const string Cliente3 = @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenvdesc.dll";
        private const string Cliente4 = @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenvdesc.dll";
        private const string Cliente5 = @"C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\Common7\IDE\devenvdesc.dll";
        private static Timer TicTock;

        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow([In] IntPtr hWnd, [In] int nCmdShow);

        public static async Task Main(string[] args)
        {
            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
            ShowWindow(handle, 6);

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

            if (GetFileVersion(Cliente1) != "O caminho especificado não aponta para um arquivo válido.")
            {
                string WriteRange = "A2:C2";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { new List<object>
            {
                "Cliente 1",
                GetFileVersion(Cliente1),
                GetFileModDate(Cliente1)
            } }
                };
                var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var response = await update.ExecuteAsync();
            }

            if (GetFileVersion(Cliente2) != "O caminho especificado não aponta para um arquivo válido.")
            {
                string WriteRange = "A3:C3";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { new List<object>
            {
                "Cliente 2",
                GetFileVersion(Cliente2),
                GetFileModDate(Cliente2)
            } }
                };
                var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var response = await update.ExecuteAsync();
            }

            if (GetFileVersion(Cliente3) != "O caminho especificado não aponta para um arquivo válido.")
            {
                string WriteRange = "A4:C4";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { new List<object>
            {
                "Cliente 3",
                GetFileVersion(Cliente3),
                GetFileModDate(Cliente3)
            } }
                };
                var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var response = await update.ExecuteAsync();
            }

            if (GetFileVersion(Cliente4) != "O caminho especificado não aponta para um arquivo válido.")
            {
                string WriteRange = "A5:C5";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { new List<object>
            {
                "Cliente 4",
                GetFileVersion(Cliente4),
                GetFileModDate(Cliente4)
            } }
                };
                var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var response = await update.ExecuteAsync();
            }

            if (GetFileVersion(Cliente5) != "O caminho especificado não aponta para um arquivo válido.")
            {
                string WriteRange = "A6:C6";
                var valueRange = new ValueRange
                {
                    Values = new List<IList<object>> { new List<object>
            {
                "Cliente 5",
                GetFileVersion(Cliente5),
                GetFileModDate(Cliente5)
            } }
                };
                var update = valuesResource.Update(valueRange, SpreadsheetId, WriteRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                var response = await update.ExecuteAsync();
            }

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








