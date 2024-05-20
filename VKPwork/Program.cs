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
			string filePath = "path_to_your_excel_file.xlsx";

			// Проверка существования файла
			if (!File.Exists(filePath))
			{
				Console.WriteLine("Файл не найден.");
				return;
			}

			// Открытие и чтение файла Excel
			FileInfo fileInfo = new FileInfo(filePath);
			using (ExcelPackage package = new ExcelPackage(fileInfo))
			{
				// Получение первого рабочего листа
				ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
				int rowCount = worksheet.Dimension.Rows;

				// Перебор строк в столбце A
				for (int row = 1; row <= rowCount; row++)
				{
					string binaryString = worksheet.Cells[row, 1].Text;

					if (!string.IsNullOrEmpty(binaryString))
					{
						try
						{
							// Конвертация из двоичной в десятичную систему
							int decimalValue = Convert.ToInt32(binaryString, 2);
							Console.WriteLine($"Строка {row}: {binaryString} в десятичной системе = {decimalValue}");
						}
						catch (FormatException)
						{
							Console.WriteLine($"Строка {row}: Неверный формат числа \"{binaryString}\".");
						}
					}
					else
					{
						Console.WriteLine($"Строка {row}: Пустая ячейка.");
					}
				}
			}
		}
	}
}
