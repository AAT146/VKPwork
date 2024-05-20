using System;
using System.Collections.Generic;
using System.IO;
using ASTRALib;
using OfficeOpenXml;
using MathNet.Numerics.Distributions;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Reflection;
using System.IO.Ports;


namespace VKPwork
{
	class Program
	{
		static void Main(string[] args)
		{
			// Укажите путь к вашему файлу Excel
			string filePath = "C:\\Users\\Анастасия\\Desktop\\2в10.xlsx";
			string outputPath = "C:\\Users\\Анастасия\\Desktop\\результат2в10.xlsx";

			// Проверка существования файла
			if (!File.Exists(filePath))
			{
				Console.WriteLine("Файл не найден.");
				return;
			}

			// Открытие и чтение файла Excel
			FileInfo fileInfo = new FileInfo(filePath);
			FileInfo outputFileInfo = new FileInfo(outputPath);

			using (ExcelPackage package = new ExcelPackage(fileInfo))
			{
				// Получение первого рабочего листа
				ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
				int rowCount = worksheet.Dimension.Rows;

				// Добавление заголовка для нового столбца
				//worksheet.Cells[1, 2].Value = "Decimal Value";

				// Перебор строк в столбце A
				for (int row = 1; row <= rowCount; row++) // Начинаем с 2, чтобы пропустить заголовок
				{
					string binaryString = worksheet.Cells[row, 1].Text;

					if (!string.IsNullOrEmpty(binaryString))
					{
						try
						{
							// Конвертация из двоичной в десятичную систему
							int decimalValue = Convert.ToInt32(binaryString, 2);
							worksheet.Cells[row, 2].Value = decimalValue;
						}
						catch (FormatException)
						{
							worksheet.Cells[row, 2].Value = "Invalid Format";
						}
					}
					else
					{
						worksheet.Cells[row, 2].Value = "Empty Cell";
					}
				}

				// Сохранение результата в новый файл
				package.SaveAs(outputFileInfo);
			}

			Console.WriteLine($"Результаты сохранены в {outputPath}");
		}
	}
}
