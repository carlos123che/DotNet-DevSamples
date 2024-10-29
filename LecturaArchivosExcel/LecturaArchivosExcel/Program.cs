using System;
using System.IO;
using OfficeOpenXml;

namespace LecturaArchivosExcel
{
	internal class Program
	{
		static void Main(string[] args)
		{
			// Excel File Route
			string filePath = @"C:\Users\dotne\Documents\Proyectos\Archivos\Planilla.xlsx";

			// LicenseContext of the ExcelPackage class to NonCommercial
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			// Validate File Exists
			if (File.Exists(filePath)) 
			{
				using (var package = new ExcelPackage(new FileInfo(filePath)))
				{
					// Read first sheet from excel
					var worksheet = package.Workbook.Worksheets[0];
					// Read Rows
					for (int row = 1; row <= worksheet.Dimension.Rows; row++) 
					{
						// Read Columns
						for (int column = 1; column <= worksheet.Dimension.Columns; column++) 
						{
							// Print content in console
							Console.Write(worksheet.Cells[row, column].Text + "\t\t");
						}
						Console.WriteLine();
					}
				}
			}
			else
			{
				// File Not Found
				Console.WriteLine("Error: File not found");
			}

			Console.ReadLine();
		}
	}
}
