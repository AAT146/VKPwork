using System;
using System.Collections.Generic;
using System.IO;
using ASTRALib;
using OfficeOpenXml;
using MathNet.Numerics.Distributions;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using ClassLibrary;
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
		/// Метод: генерация СВ ГЭС (ЛЕТО)
		/// </summary>
		/// <returns>Список СВ по Ргэс(лето)</returns>
		public static List<double> RndValueGenSummer()
		{
			// Константы для з.распр. генерации ГЭС - ЛЕТО
			double gs1 = 0.85;
			double skoGS1 = 3;
			double moGS1 = 91;
			double gs2 = 0.075;
			double lowerS = 34;
			double upperS = 83;
			double minGen = 8;
			double maxGen = 89;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ генерации
			List<double> randValueGenSummer = new List<double>();

			// СВ генерация ЛЕТО
			while (randValueGenSummer.Count < 45733)
			{
				double q = rand.NextDouble();

				if (q >= 0 && q <= gs1)
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
					randValueGenSummer.Add(part2);
				}
			}

			return randValueGenSummer;
		}

		/// <summary>
		/// Метод: генерация СВ ГЭС (ЗИМА)
		/// </summary>
		/// <returns>Список СВ по Ргэс(зима)</returns>
		public static List<double> RndValueGenWinter()
		{
			// Константы з.распр. генерации ГЭС - ЗИМА
			double gw1 = 0.13;
			double skoGW1 = 3.2;
			double moGW1 = 19;
			double gw2 = 0.07;
			double gw3 = 0.85;
			double skoGW3 = 3.2;
			double moGW3 = 14;
			double lowerW = 26;
			double upperW = 66;
			double minGen = 8;
			double maxGen = 89;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ генерации
			List<double> randValueGenWinter = new List<double>();

			//СВ генерация ЗИМА
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

			return randValueGenWinter;
		}

		/// <summary>
		/// Метод: генерация СВ Нагрузки (ЛЕТО)
		/// </summary>
		/// <returns>Список СВ по Рнагр(лето)</returns>
		public static List<double> RndValueLoadSummer()
		{
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
			double minLoad = 10;
			double maxLoad = 167;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ нагрузки
			List<double> randValueLoadSummer = new List<double>();

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

			return randValueLoadSummer;
		}

		/// <summary>
		/// Метод: генерация СВ Нагрузки (ЗИМА)
		/// </summary>
		/// <returns>Список СВ по Рнагр(зима)</returns>
		public static List<double> RndValueLoadWinter()
		{
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
			double minLoad = 10;
			double maxLoad = 167;


			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ нагрузки
			List<double> randValueLoadWinter = new List<double>();

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

			return randValueLoadWinter;
		}

		/// <summary>
		/// Работа алгоритма.
		/// </summary>
		public static void Main()
		{
			// Создание объекта времени
			Stopwatch stopwatch = new Stopwatch();

			// Засекаем время начала операции
			stopwatch.Start();

			Console.WriteLine($"Работа алгоритма.\n");

			// Генерация случайных величин
			Random rand = new Random();

			// Создание указателя на экземпляр RastrWin и его запуск
			IRastr rastr = new Rastr();

			// Загрузка файла
			string fileRegim = @"C:\Users\aat146\Desktop\NewWork\Растр.rg2";
			string shablonRegim = @"C:\Users\aat146\Documents\RastrWin3\SHABLON\режим.rg2";

			rastr.Load(RG_KOD.RG_REPL, fileRegim, shablonRegim);

			string fileSechen = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\Сечения.sch";
			string shablonSechen = @"C:\Users\aat146\Documents\RastrWin3\SHABLON\сечения.sch";

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

			// Цикл расчета в RastrWin3 (ЗИМА)
			for (int i = 0; i < 59676; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = [i];

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
