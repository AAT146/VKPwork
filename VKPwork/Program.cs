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
			if (X.Count != Y.Count)
			{
				throw new ArgumentException("Списки X и Y должны иметь одинаковую длину.");
			}

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

			// Константы вероятности состояния цепей ЛЭП
			double q0psl1 = 0.1328;
			double q1psl1 = 0.8672;
			double q0psl2 = 0.0407;
			double q1psl2 = 0.9593;
			double q0tm1 = 0.0322;
			double q1tm1 = 0.9678;
			double q0tm2 = 0.0358;
			double q1tm2 = 0.9642;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ генерации
			List<double> randValueGen = new List<double>();

			// Лист для хранения СВ нагрузки
			List<double> randValueLoad = new List<double>();

			// Лист для хранения СС Пеледуй-Сухой Лог №1
			List<double> randSostPeledSyxLog1 = new List<double>();

			// Лист для хранения СС Пеледуй-Сухой Лог №2
			List<double> randSostPeledSyxLog2 = new List<double>();

			// Лист для хранения СС Таксимо-Мамакан №1
			List<double> randSostTaksimoMamakan1 = new List<double>();

			// Лист для хранения СС Таксимо-Мамакан №2
			List<double> randSostTaksimoMamakan2 = new List<double>();

			// Генерация случайного числа ГЕНЕРАЦИИ в цикле с условием
			while (randValueGen.Count < 105409)
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
			while (randValueLoad.Count < 105409)
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

			// 1 - отключение; 0 -включение
			double s1 = 1;
			double s2 = 0;
			// Генерация случайного состояния ЦЕПИ линии
			while (randSostPeledSyxLog1.Count < 105409 && randSostPeledSyxLog2.Count < 105409 
				&& randSostTaksimoMamakan1.Count < 105409 && randSostTaksimoMamakan2.Count < 105409)
			{
				double q1 = rand.NextDouble();
				double q2 = rand.NextDouble();
				double q3 = rand.NextDouble();
				double q4 = rand.NextDouble();
				if (q1 >= 0 && q1 <= q0psl1)
				{
					randSostPeledSyxLog1.Add(s1);
				}
				if (q1 > q0psl1 && q1 <= (q0psl1 + q1psl1))
				{
					randSostPeledSyxLog1.Add(s2);
				}
				if (q2 >= 0 && q2 <= q0psl2)
				{
					randSostPeledSyxLog2.Add(s1);
				}
				if (q2 > q0psl2 && q2 <= (q0psl2 + q1psl2))
				{
					randSostPeledSyxLog2.Add(s2);
				}
				if (q3 >= 0 && q3 <= q0tm1)
				{
					randSostTaksimoMamakan1.Add(s1);
				}
				if (q3 > q0tm1 && q3 <= (q0tm1 + q1tm1))
				{
					randSostTaksimoMamakan1.Add(s2);
				}
				if (q4 >= 0 && q4 <= q0tm2)
				{
					randSostTaksimoMamakan2.Add(s1);
				}
				if (q4 > q0tm2 && q4 <= (q0tm2 + q1tm2))
				{
					randSostTaksimoMamakan2.Add(s2);
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

			// Ветви
			ICol staVetv = (ICol)tableVetv.Cols.Item("sta");   // Состояние
			ICol nStart = (ICol)tableVetv.Cols.Item("ip");   // Номер начала
			ICol nEnd = (ICol)tableVetv.Cols.Item("iq");   // Номер конца
			ICol nParall = (ICol)tableVetv.Cols.Item("np");   // Номер параллельности
			ICol nameVetv = (ICol)tableVetv.Cols.Item("name");   // Название

			// Генераторы(УР)
			ICol nAgr = (ICol)tableGenYR.Cols.Item("Num");   // Номер агрегата
			ICol pGenYR = (ICol)tableGenYR.Cols.Item("P");   // Акт. мощность генерации

			// Сечения
			ICol nSech = (ICol)tableSechen.Cols.Item("ns");   // Номер сечения
			ICol valueSech = (ICol)tableSechen.Cols.Item("psech");   // Полученное значение

			// Лист для хранения перетока по КС
			List<double> ksPeledSyxLog = new List<double>();
			List<double> ksTaksimoMamakan = new List<double>();

			double numberYR = 0;

			// Цикл расчета перетоков в RastrWin3
			for (int i = 0; i < 105409; i++)
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
				activeLoad.Z[index2] = randValueLoad[i];

				// Присваивание сгенерированного состояния цепям линий
				var setSelVetv1 = "ip=" + 3 + "&" + "iq=" + 2 + "&" + "np=" + 1;   // П-СХ № 1
				tableVetv.SetSel(setSelVetv1);
				var number1 = tableVetv.FindNextSel[-1];
				staVetv.Z[number1] = randSostPeledSyxLog1[i];

				var setSelVetv2 = "ip=" + 3 + "&" + "iq=" + 2 + "&" + "np=" + 2;   // П-СХ № 2
				tableVetv.SetSel(setSelVetv2);
				var number2 = tableVetv.FindNextSel[-1];
				staVetv.Z[number2] = randSostPeledSyxLog2[i];

				var setSelVetv3 = "ip=" + 4 + "&" + "iq=" + 2 + "&" + "np=" + 1;   // Т-М № 1
				tableVetv.SetSel(setSelVetv3);
				var number3 = tableVetv.FindNextSel[-1];
				staVetv.Z[number3] = randSostTaksimoMamakan1[i];

				var setSelVetv4 = "ip=" + 4 + "&" + "iq=" + 2 + "&" + "np=" + 2;   // Т-М № 2
				tableVetv.SetSel(setSelVetv4);
				var number4 = tableVetv.FindNextSel[-1];
				staVetv.Z[number4] = randSostTaksimoMamakan2[i];

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
			Worksheet worksheet1 = workbook.Sheets.Add();
			worksheet1.Name = "Случайные величины";
			Worksheet worksheet2 = workbook.Sheets.Add();
			worksheet2.Name = "Логическая операция";
			Worksheet worksheet3 = workbook.Sheets.Add();
			worksheet3.Name = "Состояние линий";

			// Запись значений в файл Excel
			for (int i = 0; i < 105409; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range1 = worksheet1.Range["A1"];
				Range range2 = worksheet2.Range["A1"];
				Range range3 = worksheet3.Range["A1"];

				// Запись случайной величины в столбец А листа 1 - генерация
				range1.Offset[0, 0].Value = "Генерация";
				range1.Offset[i + 1, 0].Value = randValueGen[i];

				// Запись случайной величины в столбец B листа 1 - нагрузка
				range1.Offset[0, 1].Value = "Нагрузка";
				range1.Offset[i + 1, 1].Value = randValueLoad[i];

				// Запись случайной величины в столбец C листа 1 - КС Пеледуй - Сухой Лог
				range1.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range1.Offset[i + 1, 2].Value = ksPeledSyxLog[i];

				// Запись случайной величины в столбец D листа 1 - КС Таксимо - Мамакан
				range1.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range1.Offset[i + 1, 3].Value = ksTaksimoMamakan[i];

				// Запись логической операции в столбец A листа 2 - СМЗУ КС_МДП (Пеледуй - Сухой Лог)
				range2.Offset[0, 0].Value = "СМЗУ КС_МДП (П-СЛ)";
				range2.Offset[i + 1, 0].Value = smzyPSL[i];

				// Запись логической операции в столбец B листа 2 - СМЗУ КС_МДП (Таксимо - Мамакан)
				range2.Offset[0, 1].Value = "СМЗУ КС_МДП (Т-М)";
				range2.Offset[i + 1, 1].Value = smzyTM[i];

				// Запись логической операции в столбец C листа 2 - ПУР КС_МДП (Пеледуй - Сухой Лог)
				range2.Offset[0, 2].Value = "ПУР КС_МДП (П-СЛ)";
				range2.Offset[i + 1, 2].Value = pyrPSL[i];

				// Запись логической операции в столбец D листа 2 - ПУР КС_МДП (Таксимо - Мамакан)
				range2.Offset[0, 3].Value = "ПУР КС_МДП (Т-М)";
				range2.Offset[i + 1, 3].Value = pyrTM[i];

				// Запись логической операции в столбец E листа 2 - ПУР1 КС_МДП (Таксимо - Мамакан)
				range2.Offset[0, 4].Value = "ПУР1 КС_МДП (Т-М)";
				range2.Offset[i + 1, 4].Value = pyrTM1[i];

				// Запись состояния линии в столбец А листа 3 - №1 П-СХ
				range3.Offset[0, 0].Value = "№1 П-СХ";
				range3.Offset[i + 1, 0].Value = randSostPeledSyxLog1[i];

				// Запись состояния линии в столбец B листа 3 - №2 П-СХ
				range3.Offset[0, 1].Value = "№2 П-СХ";
				range3.Offset[i + 1, 1].Value = randSostPeledSyxLog2[i];

				// Запись состояния линии в столбец C листа 3 - №1 Т-М
				range3.Offset[0, 2].Value = "№1 Т-М";
				range3.Offset[i + 1, 2].Value = randSostTaksimoMamakan1[i];

				// Запись состояния линии в столбец D листа 3 - №2 Т-М
				range3.Offset[0, 3].Value = "№2 Т-М";
				range3.Offset[i + 1, 3].Value = randSostTaksimoMamakan2[i];
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
