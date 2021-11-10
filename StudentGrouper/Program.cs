using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using System.Linq;

namespace StudentGrouper
{
	public class Program
	{
		static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static readonly string ApplicationName = "appNameTemp";
		static readonly string spreadSheetID = "1gf-drH_3LdHMqUEALLLmg9GxTPO_l9gvAkhHaxk3_Qw";
		static readonly string sheet = "LiveSheet";
		static readonly string historySheet = "History";
		static SheetsService service;


		static void Main(string[] args)
		{
            SetUpCreds();
			//WriteEntries();
			//ReadEntries($"{sheet}!A2:A26");
			//UpdateEntries($"{sheet}!A1:B1");
			//DeleteEntries($"{sheet}!A1:B1");
			List<string> devs = GetEntries($"{sheet}!A2:B26").ToList(); //dont hard code the end for the current total, get too much and just stop at the first blank entry
			List<string> arts = GetEntries($"{sheet}!C2:D17").ToList();

			devs.RemoveAll(x => x == null);
			arts.RemoveAll(x => x == null);
			//PrintGroup(devs);
			//PrintGroup(arts);



            Sorter sorter = new Sorter();

            sorter.SetAllStudents(devs, arts);

			sorter.FigureOutPairs();
			//check if should load saved file with records or start new grouping log

			Console.ReadKey();

        }
		static void PrintGroup(string[] group)
        {
			foreach(var s in group)
            {
				if (s == null)
					return;
                Console.WriteLine(s);
            }
        }
		static void SetUpCreds()
		{
			GoogleCredential credential;
			using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
			{
				credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
			}

			service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});
		}
		static void ReadEntries(string range)
		{
			//var range = $"{sheet}!A1:A1";
            var request = service.Spreadsheets.Values.Get(spreadSheetID, range);

            //var objectList = new List<Object>() { "TEST AF", "TEST 2 da Loo" };
            //valueRange.Values = new List<IList<object>> { objectList };

            //var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
            //appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            //var appendResponse = appendRequest.Execute()

            var response = request.Execute();
			var values = response.Values;
			if (values != null && values.Count >= 0)
			{
				foreach (var row in values)
				{
					Console.WriteLine("{0}", row[0]);
				}

			}
			else
			{
				Console.WriteLine("No data");
			}
		}
		static string[] GetEntries(string range)
		{
			string[] draftArray;
			//var range = $"{sheet}!A1:A1";
            var request = service.Spreadsheets.Values.Get(spreadSheetID, range);

            //var objectList = new List<Object>() { "TEST AF", "TEST 2 da Loo" };
            //valueRange.Values = new List<IList<object>> { objectList };

            //var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
            //appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            //var appendResponse = appendRequest.Execute()

            var response = request.Execute();
			var values = response.Values;
			if (values != null && values.Count >= 0)
			{
				draftArray = new string[values.Count]; //too big, but can be trimmed
				int count = 0;
                for (int i = 0; i < values.Count; i++)
                {
					if (values[i].Count > 1 && ((string)values[i][1]).Contains("x", StringComparison.OrdinalIgnoreCase)) //if X appears in any configuration, with or without spaces or other characters or in any case
						continue;

					draftArray[count] = (string)values[i][0];
					count++;
				}
				return draftArray;
			}
			else
			{
				Console.WriteLine("No data");
				return null;
			}
		}

		static void WriteEntries(string range)
		{
			//var range = $"{sheet}!A1:B1";
			var valueRange = new Google.Apis.Sheets.v4.Data.ValueRange();

			var objectList = new List<Object>() { "YAGEEEEEL", "NUUUUUUU" };
			valueRange.Values = new List<IList<object>> { objectList };

			var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
			appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

			var appendResponse = appendRequest.Execute();
		}
		static void UpdateEntries(string range)
		{
			//var range = $"{sheet}!A1:B1";
			var valueRange = new Google.Apis.Sheets.v4.Data.ValueRange();

			var objectList = new List<Object>() { "WHAT", "The fuck?"};
			valueRange.Values = new List<IList<object>> { objectList };

			var updateRequest = service.Spreadsheets.Values.Update(valueRange, spreadSheetID, range);
			updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

			var updateResponse = updateRequest.Execute();
		}
		static void DeleteEntries(string range)
		{
			//var range = $"{sheet}!A1:B1";
			var requestBody = new ClearValuesRequest();

			var deleteRequest = service.Spreadsheets.Values.Clear(requestBody,spreadSheetID, range);
			

			var deleteResponse = deleteRequest.Execute();
		}

	}



	public class Sorter
	{
		List<String> artStudents;
		List<String> devStudents;

		public Sorter()
		{
			artStudents = new List<string>();
			devStudents = new List<string>();
		}

		public void ReadAllStudents()
		{
			string text = System.IO.File.ReadAllText(@"C:\Users\alons\Documents\Student Sorter\Students.txt");

			string[] seperated = text.Split("\r\n");
			bool isDev = true;
			foreach (var name in seperated)
			{
				if (name == "_")
				{
					isDev = false;
					continue;
				}

				if(isDev)
                {
					devStudents.Add(name);
                }
				else
                {
					artStudents.Add(name);
                }
			}

   //         foreach (var item in artStudents)
   //         {
   //             Console.WriteLine(item + " art");
   //         }
			//foreach (var item in devStudents)
   //         {
   //             Console.WriteLine(item + " dev");
   //         }

		}
		public void SetAllStudents(List<string> dev, List<string> art)
        {
			devStudents = dev;
			artStudents = art;
		}
		public void PrintSingleConfiguration() //repeating groups not avoided yet!
        {
			Random rand = new Random();

			foreach (var art in artStudents)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(art);
                Console.ForegroundColor = ConsoleColor.Blue;
				int randomDevnik = rand.Next(0, devStudents.Count - 1);
                Console.WriteLine(devStudents[randomDevnik]);
				devStudents.RemoveAt(randomDevnik);
				randomDevnik = rand.Next(0, devStudents.Count - 1);
				Console.WriteLine(devStudents[randomDevnik]);
				devStudents.RemoveAt(randomDevnik);
			}
        }
		public void FigureOutPairs()
        {
			int devTotal = devStudents.Count;
			int artCount = 0;
			Random rand = new Random();

            for (int i = 0; i < devTotal; i++)
            {
				Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(artStudents[artCount]);

				int r = rand.Next(0, devStudents.Count);
				Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(devStudents[r]);
				devStudents.RemoveAt(r);

			    artCount++;

				if (artCount >= artStudents.Count)
				{
					artCount = 0;
                    Console.WriteLine();
				}
            }
        }
	}
}
