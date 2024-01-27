using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.VisualBasic;

namespace SituacaoAluno
{

    class Program
    {
        // Configure the data about the used sheet
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "SituacaoAluno";
        static readonly string SpreadsheetId = "1okvg-RTpvqI6l_E2-BWMFpXOZ5A1lVsWY2FPeZ2BpRQ";
        static readonly string sheet = "engenharia_de_software";
        static SheetsService? service;

        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("SheetApi.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            //Calls the function that reads and fill cells in the sheet
            Average();
        }




        static void Average()
        {
            //Variable declarations
            var range = $"{sheet}!C4:F27";
            var request = service?.Spreadsheets.Values.Get(SpreadsheetId, range);
            var response = request?.Execute();
            var values = response?.Values;

            var situacao = "";
            int totalSemesterClasses = 60;
            double finalApprovalNote = 0;
            //Checks if the sheet is not epmpty
            if (values != null && values.Count > 0)
            {
                //Pass trought the sheet
                for (int row = 0; row < values.Count; row++)
                {
                    //Stores the cells values
                    double col1 = double.Parse(values[row][0].ToString());
                    double col2 = double.Parse(values[row][1].ToString());
                    double col3 = double.Parse(values[row][2].ToString());
                    double col4 = double.Parse(values[row][3].ToString());

                    //Get the average Value and write it in the console for better visualisation
                    double rowAverage = (col2 + col3 + col4) / 3;
                    Console.WriteLine("row "+ (row+4) + " average: " + rowAverage);

                    //Checks the student situation
                    if ((totalSemesterClasses * 0.25) < (col1))
                    {
                        situacao = "Reprovado por Falta";
                    }
                    else if (rowAverage < 50)
                    {
                        situacao = "Reprovado por Nota";
                    }
                    else if ((rowAverage >= 50) && (rowAverage < 70))
                    {
                        situacao = "Exame Final";
                    }
                    else
                    {
                        situacao = "Aprovado";
                    }

                    // Calls the function to pass the situation values to the sheet in the "Situação" column
                    UpdateAverage(row + 4, situacao);
                    Console.WriteLine("Situation updated in the sheet");

                    //Checks if the student is in the "Exame Final" situation and calculate the grade required 
                    if (situacao == "Exame Final")
                    {
                        finalApprovalNote = Math.Ceiling(rowAverage / 2);
                    }
                    else
                    {
                        finalApprovalNote = 0;
                    }
                    // Calls the function to pass the situation values to the sheet in the "Nota para Aprovação Final" column
                    UpdateFinalApprovalNote(row + 4, finalApprovalNote);
                    Console.WriteLine("Final Approval Note Updated in the sheet");
                }


            }
            else
            {
                Console.WriteLine("No data was found");
            }
        }

        static void UpdateAverage(int row, string situacao)
        {
            //Configure the settings that will be used in the sheet
            var range = $"{sheet}!G{row}";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { situacao };
            valueRange.Values = new List<IList<object>> { objectList };
            //Update the given cell in the sheet
            var updateRequest = service?.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }

        static void UpdateFinalApprovalNote(int row, double finalApprovalNote)
        {
            //Configure the settings that will be used in the sheet
            var range = $"{sheet}!H{row}";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { finalApprovalNote };
            valueRange.Values = new List<IList<object>> { objectList };
            //Update the given cell in the sheet
            var updateRequest = service?.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();
        }
    }
}
