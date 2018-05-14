using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using hoatv.Models;
using System.IO;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.Drawing;
using Microsoft.AspNetCore.Http;
using System.Net.Http.Headers;
using Microsoft.Extensions.FileProviders;
using System.Text.RegularExpressions;
using System.Globalization;

namespace hoatv.Controllers
{
	public class HomeController : Controller
	{
		static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static readonly string ApplicationName = "Inventory reader";
		static readonly string SpreadsheetId = "1keRgQVxB2xRcbjWmQ7OP-8QnAm1VPMFbixU0Nihb1Ug";
		static readonly string sheet = "Duplicated";
		static SheetsService service;

		private IHostingEnvironment _hostingEnvironment;
		public HomeController(IHostingEnvironment hostingEnvironment)
		{
			_hostingEnvironment = hostingEnvironment;
		}

		public IActionResult Index()
		{
			return View();
		}

		[HttpPost]
		public IActionResult Index(IFormFile files)
		{
			// Check is correct file or not
			if (!files.ContentType.Contains("text"))
			{
				ViewBag.Error = "Please choose correct file type!!!";
				return View();
			}
			// Get data from inventory
			GoogleService();
			var listInventory = new List<String>();
			listInventory = ReadEntries(); // list phone number

			// Get data from log file and return list not difference contacts
			string compareDate = string.Empty;
			var listContactFromLog = GetListContactFromLogFile(listInventory, files, ref compareDate);

			string sWebRootFolder = _hostingEnvironment.WebRootPath;
			string fileName = @"ContactC3_"+compareDate+""+".xlsx";
			string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, fileName);

			String link = ExportListContact(listContactFromLog, sWebRootFolder, fileName, URL,compareDate);
			FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, link));

			return downloadFile(sWebRootFolder, fileName);
		}

		public FileResult downloadFile(string filePath, string fileName)
		{
			IFileProvider provider = new PhysicalFileProvider(filePath);
			IFileInfo fileInfo = provider.GetFileInfo(fileName);
			var readStream = fileInfo.CreateReadStream();
			return File(readStream, GetContentType(fileName), fileName);
		}

		private string GetContentType(string path)
		{
			var types = GetMimeTypes();
			var ext = Path.GetExtension(path).ToLowerInvariant();
			return types[ext];
		}

		private Dictionary<string, string> GetMimeTypes()
		{
			return new Dictionary<string, string>
			{
				{".txt", "text/plain"},
				{".pdf", "application/pdf"},
				{".doc", "application/vnd.ms-word"},
				{".docx", "application/vnd.ms-word"},
				{".xls", "application/vnd.ms-excel"},
				{".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
				{".png", "image/png"},
				{".jpg", "image/jpeg"},
				{".jpeg", "image/jpeg"},
				{".gif", "image/gif"},
				{".csv", "text/csv"}
			};
		}


		private static String ExportListContact(List<Contact> listContactFromLog, String sWebRootFolder, String sFileName, String URL, string compareDate)
		{
			FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
			if (file.Exists)
			{
				file.Delete();
				file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));
			}
			using (ExcelPackage package = new ExcelPackage(file))
			{
				// add a new worksheet to the empty workbook
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Contacts");
				//First add the headers
				Image img = Image.FromFile(@"topica_logo.jpg");
				ExcelPicture pic = worksheet.Drawings.AddPicture("Picture_Name", img);
				pic.SetSize(250, 65);
				pic.SetPosition(1, 0, 0, 0);
				worksheet.Cells[2, 4].Value = "Statistics list of contact";
				worksheet.Cells[2, 4].Style.Font.Name = "Arial";
				worksheet.Cells[2, 4].Style.Font.Size = 28;
				worksheet.Cells[2, 4].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#7F0F19"));
				worksheet.Cells[2, 4].Style.Font.Bold = true;

				worksheet.Cells[3, 5].Value = compareDate;
				worksheet.Cells[3, 5].Style.Font.Name = "Arial";
				worksheet.Cells[3, 5].Style.Font.Size = 15;
				//worksheet.Cells[3, 5].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#7F0F19"));
				//worksheet.Cells[3, 5].Style.Font.Bold = true;

				for (int i = 1; i < 25; i++)
				{
					worksheet.Cells[6, i].Style.Font.Name = "Arial";
					worksheet.Cells[6, i].Style.Font.Size = 12;
					worksheet.Cells[6, i].Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#FFFFFF"));
					worksheet.Cells[6, i].Style.Font.Bold = true;
					worksheet.Cells[6, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
					worksheet.Cells[6, i].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#7F0F19"));
				}

				worksheet.Cells[6, 1].Value = "No.";
				worksheet.Cells[6, 2].Value = "Name";
				worksheet.Cells[6, 3].Value = "Email";
				worksheet.Cells[6, 4].Value = "Phone";
				worksheet.Cells[6, 5].Value = "Address";
				worksheet.Cells[6, 6].Value = "Education level";
				worksheet.Cells[6, 7].Value = "Major";
				worksheet.Cells[6, 8].Value = "Time";
				worksheet.Cells[6, 9].Value = "Campaigns";
				worksheet.Cells[6, 10].Value = "Landing page";
				worksheet.Cells[6, 11].Value = "Chanels";
				worksheet.Cells[6, 12].Value = "Ads";
				worksheet.Cells[6, 13].Value = "Keyword";
				worksheet.Cells[6, 14].Value = "Site source";
				worksheet.Cells[6, 15].Value = "Teacher link";
				worksheet.Cells[6, 16].Value = "ContactID";
				worksheet.Cells[6, 17].Value = "DOB";
				worksheet.Cells[6, 18].Value = "Age";
				worksheet.Cells[6, 19].Value = "Line ID";
				worksheet.Cells[6, 20].Value = "Status_regis";
				worksheet.Cells[6, 21].Value = "Time_tuvan";
				worksheet.Cells[6, 22].Value = "Credit_card";
				worksheet.Cells[6, 23].Value = "Value_computer";
				worksheet.Cells[6, 24].Value = "Url tracking";

				worksheet.Column(2).Width = 25;
				worksheet.Column(3).Width = 25;
				worksheet.Column(4).Width = 20;
				worksheet.Column(5).Width = 20;
				worksheet.Column(6).Width = 20;
				worksheet.Column(7).Width = 20;
				worksheet.Column(8).Width = 20;
				worksheet.Column(9).Width = 20;
				worksheet.Column(10).Width = 20;
				worksheet.Column(11).Width = 20;
				worksheet.Column(12).Width = 20;
				worksheet.Column(13).Width = 20;
				worksheet.Column(14).Width = 20;
				worksheet.Column(15).Width = 20;
				worksheet.Column(16).Width = 20;
				worksheet.Column(17).Width = 20;
				worksheet.Column(18).Width = 20;
				worksheet.Column(19).Width = 20;
				worksheet.Column(20).Width = 20;
				worksheet.Column(21).Width = 20;
				worksheet.Column(22).Width = 20;
				worksheet.Column(23).Width = 20;
				worksheet.Column(24).Width = 30;

				int startIndex = 7;
				for (int i = 0; i < listContactFromLog.Count; i++)
				{
					String a = "A" + (startIndex + i);
					worksheet.Cells["A" + (startIndex + i)].Value = (i + 1);
					worksheet.Cells["B" + (startIndex + i)].Value = listContactFromLog[i].Name;
					worksheet.Cells["C" + (startIndex + i)].Value = listContactFromLog[i].Email;
					worksheet.Cells["D" + (startIndex + i)].Value = listContactFromLog[i].Phone;
					worksheet.Cells["R" + (startIndex + i)].Value = listContactFromLog[i].Age.Replace("??", "").Replace("ปี", "");
					worksheet.Cells["X" + (startIndex + i)].Value = listContactFromLog[i].CodeChanel;
				}


				package.Save(); //Save the workbook.
			}
			return URL;
		}


		public List<string> ReadAsList(IFormFile file)
		{
			var result = new List<string>();
			using (var reader = new StreamReader(file.OpenReadStream()))
			{
				while (reader.Peek() >= 0)
					result.Add(reader.ReadLine());
			}
			return result;
		}

		public List<Contact> GetListContactFromLogFile(List<String> listInventory, IFormFile file, ref string compareDate)
		{
			var listStr = new List<string>();
			listStr = ReadAsList(file);
			//string[] lines = System.IO.File.ReadAllLines(@"logfile.txt");
			var listResult = new List<Contact>();
			Contact contact = new Contact();

			int i = 0;
			foreach (string line in listStr)
			{
				if (i == 0)
				{
					var regex = new Regex(@"\b\d{4}\-\d{2}-\d{2}\b");
					foreach (Match m in regex.Matches(line))
					{
						DateTime dt;
						if (DateTime.TryParseExact(m.Value, "yyyy-MM-dd", null, DateTimeStyles.None, out dt))
						{
							compareDate = dt.ToString("yyyy-MM-dd");
						}
					}
				}
				//Console.WriteLine("\t" + line);
				foreach (var keyItem in getListKey())
				{
					if (line.Contains(keyItem))
					{
						switch (keyItem)
						{
							case "[name] =>":
								contact.Name = Remove(line, keyItem);
								break;
							case "[phone] =>":
								contact.Phone = Remove(line, keyItem);
								break;
							case "[email] =>":
								contact.Email = Remove(line, keyItem);
								break;
							case "[age] =>":
								contact.Age = Remove(line, keyItem);
								break;
							case "[id_camp_landingpage] =>":
								contact.IDCampainLandingpage = Remove(line, keyItem);
								break;
							case "[code_chanel] =>":
								contact.CodeChanel = Remove(line, keyItem);
								if (IsPhoneExistsInInventory(contact.Phone, listInventory))
								{
									// break out of loop
									contact = new Contact();
									goto BREAKHERE;
								}
								else
								{
									listResult.Add(contact);
									contact = new Contact();
									break;
								}
							default:
								break;
						}
					}

				}
				i++;
			BREAKHERE: continue;
			}
			return listResult;
		}

		private bool IsPhoneExistsInInventory(String phone, List<String> listPhoneInventory)
		{
			foreach (var item in listPhoneInventory)
			{
				if (item.Contains(phone) || phone.Contains(item))
				{
					return true;
				}
			}
			return false;
		}

		private String Remove(String input, String key)
		{
			return input.Replace(key, "").Trim();
		}

		private List<String> getListKey()
		{
			var listResult = new List<String>();
			listResult.Add("[name] =>");
			listResult.Add("[phone] =>");
			listResult.Add("[email] =>");
			listResult.Add("[age] =>");
			listResult.Add("[id_camp_landingpage] =>");
			listResult.Add("[code_chanel] =>");
			return listResult;
		}


		static List<String> ReadEntries()
		{
			var range = $"{sheet}!A:F";
			SpreadsheetsResource.ValuesResource.GetRequest request =
					service.Spreadsheets.Values.Get(SpreadsheetId, range);

			var response = request.Execute();
			IList<IList<object>> values = response.Values;
			var listResult = new List<String>();
			if (values != null && values.Count > 0)
			{
				foreach (var row in values)
				{
					//listResult.Add(row[3].ToString()); // Only get phone number from inventory
					//listResult.Add(row[0].ToString()); // Only get phone number from inventory
					if (!string.IsNullOrEmpty(row[0].ToString())){
						listResult.Add(row[0].ToString());
					}
				}
			}
			return listResult;
		}

		private void GoogleService()
		{
			UserCredential credential;

			using (var stream =
				new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
			{
				string credPath = System.Environment.GetFolderPath(
					System.Environment.SpecialFolder.Personal);
				credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.Load(stream).Secrets,
					Scopes,
					"user",
					CancellationToken.None,
					new FileDataStore(credPath, true)).Result;
				Console.WriteLine("Credential file saved to: " + credPath);
			}

			// Create Google Sheets API service.
			service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});
		}





		public IActionResult Error()
		{
			return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
		}
	}
}
