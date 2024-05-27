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
	/// <summary>
	/// Расчета ПБН на примере Бодайбинского ЭР Иркутской ОЗ.
	/// </summary>
	public class Program
	{
		/// <summary>
		/// Упрощенное моделирование.
		/// </summary>
		public static void Main()
		{
			// Создание объекта времени
			Stopwatch stopwatch = new Stopwatch();

			// Засекаем время начала операции
			stopwatch.Start();

			Console.WriteLine($"Работа алгоритма.\n");

			// Константы для з.распр. генерации ГЭС - ЛЕТО
			double gs1 = 0.6;
			double skoGS1 = 1.52;
			double moGS1 = 88;
			double gs2 = 0.4;
			double lowerS = 34;
			double upperS = 84;

			// Константы з.распр. генерации ГЭС - ЗИМА
			double gw1 = 0.10;
			double skoGW1 = 3.4;
			double moGW1 = 19;
			double gw2 = 0.05;
			double gw3 = 0.85;
			double skoGW3 = 3.4;
			double moGW3 = 14;
			double lowerW = 26;
			double upperW = 66;

			// Min&Max знаечния генерации
			double minGen = 8;
			double maxGen = 87;

			// Константы з.распр. нагрузки - ЛЕТО
			double ls1 = 0.27;
			double skoLS1 = 10;
			double moLS1 = 101;
			double ls2 = 0.50;
			double skoLS2 = 6;
			double moLS2 = 110;
			double ls3 = 0.23;
			double skoLS3 = 5.5;
			double moLS3 = 125;

			// Константы з.распр. нагрузки - ЗИМА
			double lw1 = 0.41;
			double skoLW1 = 8;
			double moLW1 = 110.5;
			double lw2 = 0.42;
			double skoLW2 = 9;
			double moLW2 = 117.8;
			double lw3 = 0.17;
			double skoLW3 = 5;
			double moLW3 = 113;

			// Min&Max знаечния нагрузки
			double minLoad = 10;
			double maxLoad = 167;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ генерации
			List<double> randValueGenSummer = new List<double>();
			List<double> randValueGenWinter = new List<double>();

			// Лист для хранения СВ нагрузки
			List<double> randValueLoadSummer = new List<double>();
			List<double> randValueLoadWinter = new List<double>();

			// СВ генерация ЛЕТО
			while (randValueGenSummer.Count < 45733)
			{
				double q = rand.NextDouble();

				if (q > 0 && q <= gs1)
				{
					Normal normalDistribution = new Normal(moGS1, skoGS1);
					double part1 = Math.Round(normalDistribution.Sample(), 0);
					if (part1 >= minGen && part1 < maxGen)
					{
						randValueGenSummer.Add(part1);
					}
				}
				else if (q > gs1 && q <= (gs1 + gs2))
				{
					ContinuousUniform uniformDist = new ContinuousUniform(lowerS, upperS);
					double part2 = Math.Round(uniformDist.Sample(), 0);
					if (part2 >= minGen && part2 < maxGen)
					{
						randValueGenSummer.Add(part2);
					}
				}
			}

			// СВ генерация ЗИМА
			while (randValueGenWinter.Count < 59676)
			{
				double q = rand.NextDouble();

				if (q > 0 && q <= gw1)
				{
					Normal normalDistribution = new Normal(moGW1, skoGW1);
					double part3 = Math.Round(normalDistribution.Sample(), 0);
					if (part3 >= minGen && part3 < maxGen)
					{
						randValueGenWinter.Add(part3);
					}
				}
				else if (q > gw1 && q <= (gw1 + gw2))
				{
					ContinuousUniform uniformDist = new ContinuousUniform(lowerW, upperW);
					double part4 = Math.Round(uniformDist.Sample(), 0);
					if (part4 >= minGen && part4 < maxGen)
					{
						randValueGenWinter.Add(part4);
					}
				}
				else if (q > (gw1 + gw2) && q <= (gw1 + gw2 + gw3))
				{
					Normal normalDistribution = new Normal(moGW3, skoGW3);
					double part5 = Math.Round(normalDistribution.Sample(), 0);
					if (part5 >= minGen && part5 < maxGen)
					{
						randValueGenWinter.Add(part5);
					}
				}
			}

			// СВ нагрузка ЛЕТО
			while (randValueLoadSummer.Count < 45733)
			{
				double q = rand.NextDouble();
				if (q > 0 && q <= ls1)
				{
					Normal normalDistribution = new Normal(moLS1, skoLS1);
					double part6 = Math.Round(normalDistribution.Sample(), 0);
					if (part6 >= minLoad && part6 < maxLoad)
					{
						randValueLoadSummer.Add(part6);
					}
				}
				else if (q > ls1 && q <= (ls1 + ls2))
				{
					Normal normalDistribution = new Normal(moLS2, skoLS2);
					double part7 = Math.Round(normalDistribution.Sample(), 0);
					if (part7 >= minLoad && part7 < maxLoad)
					{
						randValueLoadSummer.Add(part7);
					}
				}
				else if (q > (ls1 + ls2) && q <= (ls1 + ls2 + ls3))
				{
					Normal normalDistribution = new Normal(moLS3, skoLS3);
					double part8 = Math.Round(normalDistribution.Sample(), 0);
					if (part8 >= minLoad && part8 < maxLoad)
					{
						randValueLoadSummer.Add(part8);
					}
				}
			}

			// СВ нагрузка ЗИМА
			while (randValueLoadWinter.Count < 59676)
			{
				double q = rand.NextDouble();
				if (q > 0 && q <= lw1)
				{
					Normal normalDistribution = new Normal(moLW1, skoLW1);
					double part9 = Math.Round(normalDistribution.Sample(), 0);
					if (part9 >= minLoad && part9 < maxLoad)
					{
						randValueLoadWinter.Add(part9);
					}
				}
				else if (q > lw1 && q <= (lw1 + lw2))
				{
					Normal normalDistribution = new Normal(moLW2, skoLW2);
					double part10 = Math.Round(normalDistribution.Sample(), 0);
					if (part10 >= minLoad && part10 < maxLoad)
					{
						randValueLoadWinter.Add(part10);
					}
				}
				else if (q > (lw1 + lw2) && q <= (lw1 + lw2 + lw3))
				{
					Normal normalDistribution = new Normal(moLW3, skoLW3);
					double part11 = Math.Round(normalDistribution.Sample(), 0);
					if (part11 >= minLoad && part11 < maxLoad)
					{
						randValueLoadWinter.Add(part11);
					}
				}
			}

			// Путь до файла Excel Результат
			string folder = @"C:\Users\Анастасия\Desktop\NewWork\ResultRandom";
			string file1 = "Summer.xlsx";
			string xlsxFile1 = Path.Combine(folder, file1);

			// Создание книги и листа
			Application excelApp1 = new Application();
			Workbook workbook1 = excelApp1.Workbooks.Add();
			Worksheet worksheet1 = workbook1.Sheets.Add();
			worksheet1.Name = "Значения";

			// Запись значений в файл Excel
			for (int i = 0; i < 45733; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range = worksheet1.Range["A1"];

				// Запись случайной величины в столбец А - генерация
				range.Offset[0, 0].Value = "Генерация";
				range.Offset[i + 1, 0].Value = randValueGenSummer[i];

				// Запись случайной величины в столбец B - нагрузка
				range.Offset[0, 1].Value = "Нагрузка";
				range.Offset[i + 1, 1].Value = randValueLoadSummer[i];
			}

			workbook1.SaveAs(xlsxFile1);
			workbook1.Close();
			excelApp1.Quit();

			string file2 = "Winter.xlsx";
			string xlsxFile2 = Path.Combine(folder, file2);

			// Создание книги и листа
			Application excelApp2 = new Application();
			Workbook workbook2 = excelApp2.Workbooks.Add();
			Worksheet worksheet2 = workbook2.Sheets.Add();
			worksheet2.Name = "Значения";

			// Запись значений в файл Excel
			for (int i = 0; i < 59676; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range = worksheet2.Range["A1"];

				// Запись случайной величины в столбец А - генерация
				range.Offset[0, 0].Value = "Генерация";
				range.Offset[i + 1, 0].Value = randValueGenWinter[i];

				// Запись случайной величины в столбец B - нагрузка
				range.Offset[0, 1].Value = "Нагрузка";
				range.Offset[i + 1, 1].Value = randValueLoadWinter[i];
			}

			workbook2.SaveAs(xlsxFile2);
			workbook2.Close();
			excelApp2.Quit();

			// Останавливаем счетчик
			stopwatch.Stop();

			Console.WriteLine($"\nВремя расчета: {stopwatch.ElapsedMilliseconds} мс\n" +
				$"Файл ExcelSummer успешно сохранен по пути: {xlsxFile1}\n" +
				$"Файл ExcelWinter успешно сохранен по пути: {xlsxFile2}\n" +
				$"Количество СВ генерации ЗИМА: {randValueGenWinter.Count}\n" +
				$"Количество СВ генерации ЛЕТО: {randValueGenSummer.Count}\n" +
				$"Количество СВ нагрузки ЗИМА: {randValueLoadWinter.Count}\n" +
				$"Количество СВ нагрузки ЛЕТО: {randValueLoadSummer.Count}\n");

			Console.ReadKey();
		}
	}
}
