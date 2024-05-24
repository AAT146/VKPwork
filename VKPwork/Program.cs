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
			double gs1 = 0.993;
			double skoGS1 = 1.52;
			double moGS1 = 88;
			double gs2 = 0.007;
			double lowerS = 34;
			double upperS = 84;

			// Константы з.распр. генерации ГЭС - ЗИМА
			double gw1 = 0.148;
			double skoGW1 = 3.4;
			double moGW1 = 19;
			double gw2 = 0.002;
			double gw3 = 0.84;
			double skoGW3 = 3.4;
			double moGW3 = 14;
			double lowerW = 23;
			double upperW = 83.05;

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
			List<double> randValueLoadLoad = new List<double>();

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

			// Генерация случайного числа НАГРУЗКИ в цикле с условием
			while (randValueLoad.Count < 105409)
			{
				double q = rand.NextDouble();
				if (q >= 0 && q < v4)
				{
					Normal normalDistribution = new Normal(mo4, sko4);
					double part4 = Math.Round(normalDistribution.Sample(), 0);
					if (part4 >= minLoad && part4 < maxLoad)
					{
						randValueLoad.Add(part4);
					}
				}
				else if (q >= v4 && q < (v4 + v5))
				{
					Normal normalDistribution = new Normal(mo5, sko5);
					double part5 = Math.Round(normalDistribution.Sample(), 0);
					if (part5 >= minLoad && part5 < maxLoad)
					{
						randValueLoad.Add(part5);
					}
				}
				else if (q >= v5 && q < (v4 + v5 + v6))
				{
					Normal normalDistribution = new Normal(mo6, sko6);
					double part6 = Math.Round(normalDistribution.Sample(), 0);
					if (part6 >= minLoad && part6 < maxLoad)
					{
						randValueLoad.Add(part6);
					}
				}
			}

			// Путь до файла Excel Результат
			string folder = @"C:\Users\Анастасия\Desktop\NewWork\filesExcel";
			string file1 = "Summer.xlsx";
			string file2 = "Winter.xlsx";
			string xlsxFile1 = Path.Combine(folder, file1);

			// Создание книги и листа
			Application excelApp = new Application();
			Workbook workbook = excelApp.Workbooks.Add();
			Worksheet worksheet = workbook.Sheets.Add();
			worksheet.Name = "Случайные величины";

			// Запись значений в файл Excel
			for (int i = 0; i < 105409; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range = worksheet.Range["A1"];

				// Запись случайной величины в столбец А - генерация
				range.Offset[0, 0].Value = "Генерация";
				range.Offset[i + 1, 0].Value = randValueGen[i];

				// Запись случайной величины в столбец B - нагрузка
				range.Offset[0, 1].Value = "Нагрузка";
				range.Offset[i + 1, 1].Value = randValueLoad[i];

				// Запись случайной величины в столбец C - Пеледуй - Сухой Лог I цепь
				range.Offset[0, 2].Value = "Пеледуй - Сухой Лог I цепь";
				range.Offset[i + 1, 2].Value = v1PeledSyxLog[i];

				// Запись случайной величины в столбец D - Пеледуй - Сухой Лог II цепь
				range.Offset[0, 3].Value = "Пеледуй - Сухой Лог II цепь";
				range.Offset[i + 1, 3].Value = v2PeledSyxLog[i];

				// Запись случайной величины в столбец E - Таксимо - Мамакан I цепь
				range.Offset[0, 4].Value = "Таксимо - Мамакан I цепь";
				range.Offset[i + 1, 4].Value = v1TaksimoMamakan[i];

				// Запись случайной величины в столбец F - Таксимо - Мамакан II цепь
				range.Offset[0, 5].Value = "Таксимо - Мамакан II цепь";
				range.Offset[i + 1, 5].Value = v2TaksimoMamakan[i];

				// Запись случайной величины в столбец G - КС Пеледуй - Сухой Лог
				range.Offset[0, 6].Value = "КС Пеледуй - Сухой Лог";
				range.Offset[i + 1, 6].Value = ksPeledSyxLog[i];

				// Запись случайной величины в столбец H - КС Таксимо - Мамакан
				range.Offset[0, 7].Value = "КС Таксимо - Мамакан";
				range.Offset[i + 1, 7].Value = ksTaksimoMamakan[i];

				// Запись случайной величины в столбец I - СМЗУ КС_МДП (Пеледуй - Сухой Лог)
				range.Offset[0, 8].Value = "СМЗУ КС_МДП (П-СХ)";
				range.Offset[i + 1, 8].Value = smzyPSL[i];

				// Запись случайной величины в столбец J - СМЗУ КС_МДП (Таксимо - Мамакан)
				range.Offset[0, 9].Value = "СМЗУ КС_МДП (Т-М)";
				range.Offset[i + 1, 9].Value = smzyTM[i];

				// Запись случайной величины в столбец K - ПУР КС_МДП (Пеледуй - Сухой Лог)
				range.Offset[0, 10].Value = "ПУР КС_МДП (П-СХ)";
				range.Offset[i + 1, 10].Value = pyrPSL[i];

				// Запись случайной величины в столбец L - ПУР КС_МДП (Таксимо - Мамакан)
				range.Offset[0, 11].Value = "ПУР КС_МДП (Т-М)";
				range.Offset[i + 1, 11].Value = pyrTM[i];

				// Запись случайной величины в столбец M - ПУР1 КС_МДП (Таксимо - Мамакан)
				range.Offset[0, 12].Value = "ПУР1 КС_МДП (Т-М)";
				range.Offset[i + 1, 12].Value = pyrTM1[i];
			}

			workbook.SaveAs(xlsxFile);
			workbook.Close();
			excelApp.Quit();

			// Останавливаем счетчик
			stopwatch.Stop();

			Console.WriteLine($"\nВремя расчета: {stopwatch.ElapsedMilliseconds} мс\n" +
				$"Файл Excel успешно сохранен по пути: {xlsxFile}\n" +
				$"Количество СВ генерации: {randValueGen.Count}\n" +
				$"Количество СВ нагрузки: {randValueLoad.Count}\n" +
				$"Количество просчитанных режимов: {numberYR}\n");

			Console.ReadKey();
		}
	}
}
