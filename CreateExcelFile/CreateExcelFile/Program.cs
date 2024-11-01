using OfficeOpenXml;
using OfficeOpenXml.Configuration;
using OfficeOpenXml.Style;
using System.Drawing;

namespace CreateExcelFile
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// File Path and Name
			string filePath = @"C:\\Documents\\Excel\\Data1.xlsx";
			string directoryPath = Path.GetDirectoryName(filePath);

			// Set EPPlus as NonComercial
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			// Verify Route Exists, if not create it
			if (!Directory.Exists(directoryPath))
			{
				Directory.CreateDirectory(directoryPath);
			}

			// Create a new Excel Package
			using (var package = new ExcelPackage())
			{
				// Create a new Worksheet
				var worksheet = package.Workbook.Worksheets.Add("Data");

				// Headers
				worksheet.Cells[1, 1].Value = "Id";
				worksheet.Cells[1, 2].Value = "Name";
				worksheet.Cells[1, 3].Value = "Age";

				// -- Data --
				// Row 1
				worksheet.Cells[2, 1].Value = 1;
				worksheet.Cells[2, 2].Value = "Mike";
				worksheet.Cells[2, 3].Value = 25;
				// Row 2
				worksheet.Cells[3, 1].Value = 2;
				worksheet.Cells[3, 2].Value = "Ana";
				worksheet.Cells[3, 3].Value = 30;
				// Row 3
				worksheet.Cells[4, 1].Value = 2;
				worksheet.Cells[4, 2].Value = "Peter";
				worksheet.Cells[4, 3].Value = 30;

				// Save File
				FileInfo excelFile = new FileInfo(filePath);
				package.SaveAs(excelFile);
			}

			// Create Second Excel File
			CrearArchivoExcel();

			// Print Message in Console
			Console.WriteLine("Archivo de Excel creado exitosamente.");
		}



		// Create Second File methond
		static void CrearArchivoExcel()
		{
			// File Path and Name
			string filePath = @"C:\\Users\\dotne\\Documents\\Excel\\miArchivoEstilizado.xlsx";

			// Data List
			List<(string Nombre, int Edad)> personas = new List<(string, int)>
			{
				("Marta", 25),
				("Ana", 30),
				("Luis", 22),
				("Carla", 28),
				("María", 35),
				("David", 22)
			};

			// Verify Route Exists, if not create it
			if (!Directory.Exists(Path.GetDirectoryName(filePath)))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(filePath));
			}

			// Create a new Excel Package
			using (var package = new ExcelPackage())
			{
				// Create a new Worksheet
				var worksheet = package.Workbook.Worksheets.Add("Datos");

				// Headers 
				worksheet.Cells[1, 1].Value = "ID";
				worksheet.Cells[1, 2].Value = "Nombre";
				worksheet.Cells[1, 3].Value = "Edad";

				// Apply Styles to headers
				using (var range = worksheet.Cells[1, 1, 1, 3])
				{
					range.Style.Font.Bold = true;
					range.Style.Font.Size = 12;
					range.Style.Font.Name = "Arial";
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
					range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
				}

				//  -- Data --
				int row = 2;
				foreach (var persona in personas)
				{
					worksheet.Cells[row, 1].Value = row - 1; // ID
					worksheet.Cells[row, 2].Value = persona.Nombre;
					worksheet.Cells[row, 3].Value = persona.Edad;

					// Appy borders to data
					using (var range = worksheet.Cells[row, 1, row, 3])
					{
						range.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Gray);
					}

					row++;
				}

				// Adjust widht
				worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

				// Save File
				FileInfo excelFile = new FileInfo(filePath);
				package.SaveAs(excelFile);
			}
		}


	}
}
