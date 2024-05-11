﻿using System;
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
	/// Класс, содержащий метод сравнения двух величин.
	/// </summary>
	public class ComparisonHelper
	{
		/// <summary>
		/// Метод: сравнение значений P.КС с МДП.
		/// </summary>
		/// <param name="X">Лист со значениями перетока по КС.</param>
		/// <param name="Y">Лист со значениями МДП.</param>
		/// <returns>Новый список, содержащий в себе delta или 0.</returns>
		/// <exception cref="ArgumentException">Исключение при неравной длине
		/// исходных списков.</exception>
		public static List<double> CompareLists(List<double> X, List<double> Y)
		{
			//if (X.Count != Y.Count)
			//{
			//	throw new ArgumentException("Списки X и Y должны иметь одинаковую длину.");
			//}

			List<double> results = new List<double>();

			for (int i = 0; i < X.Count; i++)
			{
				// Если true, то возврат 1; Если false (< либо =), то возврат 0.
				double result = X[i] > Y[i] ? 1 : 0;
				results.Add(result);
			}

			return results;
		}
	}

	/// <summary>
	/// Расчета ПБН на примере Бодайбинского ЭР Иркутской ОЗ.
	/// </summary>
	public class Program
	{
		/// <summary>
		/// Метод: чтение файла формата Excel.
		/// </summary>
		/// <param name="filePath">Файл Excel.</param>
		/// <returns>Массив данных.</returns>
		public static List<double> ReadFileFromExcel(string filePath)
		{
			// Установка контекста лицензирования
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			using (var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var worksheet = package.Workbook.Worksheets[0];
				List<double> data = new List<double>(worksheet.Dimension.Rows);

				for (int i = 1; i <= worksheet.Dimension.Rows; i++)
				{
					data.Add(worksheet.Cells[i, 1].GetValue<double>());
				}

				return data;
			}
		}

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

			// Константы для искусственного з.распр. генерации ГЭС
			double v3 = 0.42;
			double sko3 = 3.7;
			double mo3 = 14.5;
			double v2 = 0.3;
			double sko2 = 1.3;
			double mo2 = 86.95;
			double v1 = 0.16;
			double lowerBound = 23;
			double upperBound = 83.05;
			double minGen = 8;
			double maxGen = 87;

			// Константы для искусственного з.распр. нагрузки
			double v4 = 0.43;
			double sko4 = 5.8;
			double mo4 = 110.55;
			double v5 = 0.41;
			double sko5 = 11.2;
			double mo5 = 107.8;
			double v6 = 0.16;
			double sko6 = 5.5;
			double mo6 = 123.5;
			double maxLoad = 167;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ генерации
			List<double> randValueGen = new List<double>();

			// Лист для хранения СВ нагрузки
			List<double> randValueLoad = new List<double>();

			// Генерация случайного числа ГЕНЕРАЦИИ в цикле с условием
			while (randValueGen.Count < 10)
			{
				double q = rand.NextDouble();

				if (q >= 0 && q < v3)
				{
					Normal normalDistribution = new Normal(mo3, sko3);
					double part3 = Math.Round(normalDistribution.Sample(), 0);
					if (part3 >= minGen)
					{
						randValueGen.Add(part3);
					}
				}

				else if (q >= v3 && q < (v3 + v1))
				{
					ContinuousUniform uniformDist = new ContinuousUniform(lowerBound, upperBound);
					double part1 = Math.Round(uniformDist.Sample(), 0);
					randValueGen.Add(part1);
				}

				else if (q >= (v3 + v1) && q < (v3 + v1 + v2))
				{
					Normal normalDistribution = new Normal(mo2, sko2);
					double part2 = Math.Round(normalDistribution.Sample(), 0);
					if (part2 < maxGen)
					{
						randValueGen.Add(part2);
					}
				}
			}

			// Генерация случайного числа НАГРУЗКИ в цикле с условием
			while (randValueLoad.Count < 10)
			{
				double q = rand.NextDouble();
				if (q >= 0 && q < v4)
				{
					Normal normalDistribution = new Normal(mo4, sko4);
					double part4 = Math.Round(normalDistribution.Sample(), 0);
					if (part4 < maxLoad)
					{
						randValueLoad.Add(part4);
					}
				}
				else if (q >= v4 && q < (v4 + v5))
				{
					Normal normalDistribution = new Normal(mo5, sko5);
					double part5 = Math.Round(normalDistribution.Sample(), 0);
					if (part5 < maxLoad)
					{
						randValueLoad.Add(part5);
					}
				}
				else if (q >= v5 && q < (v4 + v5 + v6))
				{
					Normal normalDistribution = new Normal(mo6, sko6);
					double part6 = Math.Round(normalDistribution.Sample(), 0);
					if (part6 < maxLoad)
					{
						randValueLoad.Add(part6);
					}
				}
			}

			// Создание указателя на экземпляр RastrWin и его запуск
			IRastr rastr = new Rastr();

			// Загрузка файла
			string fileRegim = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\Режим.rg2";
			string shablonRegim = @"C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\режим.rg2";

			rastr.Load(RG_KOD.RG_REPL, fileRegim, shablonRegim);

			string fileSechen = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\Сечения.sch";
			string shablonSechen = @"C:\Program Files (x86)\RastrWin3\RastrWin3\SHABLON\сечения.sch";

			rastr.Load(RG_KOD.RG_REPL, fileSechen, shablonSechen);

			// Объявление объекта, содержащего таблицу "Узлы"
			ITable tableNode = (ITable)rastr.Tables.Item("node");

			// Объявление объекта, содержащего таблицу "Ветви"
			ITable tableVetv = (ITable)rastr.Tables.Item("vetv");

			// Объявление объекта, содержащего таблицу "Генератор(УР)"
			ITable tableGenYR = (ITable)rastr.Tables.Item("Generator");

			// Объявление объекта, содержащего таблицу "Сечения"
			ITable tableSechen = (ITable)rastr.Tables.Item("sechen");

			// Узлы
			ICol numberNode = (ICol)tableNode.Cols.Item("ny");   // Номер
			ICol activeLoad = (ICol)tableNode.Cols.Item("pn");   // Акт. мощность нагрузки

			// Генераторы(УР)
			ICol nAgr = (ICol)tableGenYR.Cols.Item("Num"); // Номер агрегата
			ICol pGenYR = (ICol)tableGenYR.Cols.Item("P"); // Акт. мощность генерации

			// Сечения
			ICol nSech = (ICol)tableSechen.Cols.Item("ns"); // Номер сечения
			ICol valueSech = (ICol)tableSechen.Cols.Item("psech"); // Полученное значение

			// Лист для хранения перетока по КС
			List<double> ksPeledSyxLog = new List<double>();
			List<double> ksTaksimoMamakan = new List<double>();

			double numberYR = 0;

			// Цикл расчета перетоков в RastrWin3
			for (int i = 0; i < 10; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = randValueGen[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				//var PPP = activeLoad.Z[index2];
				activeLoad.Z[index2] = randValueLoad[i];

				// Расчет УР
				_ = rastr.rgm("");
				numberYR += 1;

				// Считывание перетоков по каждому КС
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index7 = tableSechen.FindNextSel[-1];
				ksPeledSyxLog.Add(Math.Round(valueSech.Z[index7], 0));

				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index8 = tableSechen.FindNextSel[-1];
				ksTaksimoMamakan.Add(Math.Round(valueSech.Z[index8], 0));
			}

			// Файл Excel значений МДП по КС
			string xlsxMdpPeledSyxLog = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\KsPeledSyxLog.xlsx";
			string xlsxMdpTaksimoMamakan = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\KsTaksimoMamakan.xlsx";
			string xlsxPYRPeledSyxLog = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\PYRPeledSyxLog.xlsx";
			string xlsxPYRTaksimoMamakan = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\PYRTaksimoMamakan.xlsx";
			string xlsxPYRTaksimoMamakan1 = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\PYRTaksimoMamakan1.xlsx";

			// Чтение данных из файла Excel
			List<double> mdpPeledSyxLog = ReadFileFromExcel(xlsxMdpPeledSyxLog);
			List<double> mdpTaksimoMamakan = ReadFileFromExcel(xlsxMdpTaksimoMamakan);
			List<double> pyrPeledSyxLog = ReadFileFromExcel(xlsxPYRPeledSyxLog);
			List<double> pyrTaksimoMamakan = ReadFileFromExcel(xlsxPYRTaksimoMamakan);
			List<double> pyrTaksimoMamakan1 = ReadFileFromExcel(xlsxPYRTaksimoMamakan1);

			// Определение разницы между КС и МДП
			List<double> smzyPSL = ComparisonHelper.CompareLists(ksPeledSyxLog, mdpPeledSyxLog);
			List<double> smzyTM = ComparisonHelper.CompareLists(ksTaksimoMamakan, mdpTaksimoMamakan);
			List<double> pyrPSL = ComparisonHelper.CompareLists(ksPeledSyxLog, pyrPeledSyxLog);
			List<double> pyrTM = ComparisonHelper.CompareLists(ksTaksimoMamakan, pyrTaksimoMamakan);
			List<double> pyrTM1 = ComparisonHelper.CompareLists(ksTaksimoMamakan, pyrTaksimoMamakan1);

			// Путь до файла Excel Результат
			string folder = @"C:\Users\Анастасия\Desktop\ПроизПрактика";
			string fileExcel = "Результат.xlsx";
			string xlsxFile = Path.Combine(folder, fileExcel);

			// Создание книги и листа
			Application excelApp = new Application();
			Workbook workbook = excelApp.Workbooks.Add();
			Worksheet worksheet = workbook.Sheets.Add();
			worksheet.Name = "Результат";

			// Запись значений в файл Excel
			for (int i = 0; i < 10; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range = worksheet.Range["A1"];

				// Запись случайной величины в столбец А - генерация
				range.Offset[0, 0].Value = "Генерация";
				range.Offset[i + 1, 0].Value = randValueGen[i];

				// Запись случайной величины в столбец B - нагрузка
				range.Offset[0, 1].Value = "Нагрузка";
				range.Offset[i + 1, 1].Value = randValueLoad[i];

				// Запись случайной величины в столбец C - КС Пеледуй - Сухой Лог
				range.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range.Offset[i + 1, 2].Value = ksPeledSyxLog[i];

				// Запись случайной величины в столбец D - КС Таксимо - Мамакан
				range.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range.Offset[i + 1, 3].Value = ksTaksimoMamakan[i];

				// Запись случайной величины в столбец E - СМЗУ КС_МДП (Пеледуй - Сухой Лог)
				range.Offset[0, 4].Value = "СМЗУ КС_МДП (П-СЛ)";
				range.Offset[i + 1, 4].Value = smzyPSL[i];

				// Запись случайной величины в столбец F - СМЗУ КС_МДП (Таксимо - Мамакан)
				range.Offset[0, 5].Value = "СМЗУ КС_МДП (Т-М)";
				range.Offset[i + 1, 5].Value = smzyTM[i];

				// Запись случайной величины в столбец J - ПУР КС_МДП (Пеледуй - Сухой Лог)
				range.Offset[0, 6].Value = "ПУР КС_МДП (П-СЛ)";
				range.Offset[i + 1, 6].Value = pyrPSL[i];

				// Запись случайной величины в столбец H - ПУР КС_МДП (Таксимо - Мамакан)
				range.Offset[0, 7].Value = "ПУР КС_МДП (Т-М)";
				range.Offset[i + 1, 7].Value = pyrTM[i];

				// Запись случайной величины в столбец I - ПУР1 КС_МДП (Таксимо - Мамакан)
				range.Offset[0, 8].Value = "ПУР1 КС_МДП (Т-М)";
				range.Offset[i + 1, 8].Value = pyrTM1[i];
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
