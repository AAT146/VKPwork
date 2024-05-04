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
			double maxGen = 87
				;
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
			double minLoad = 10;
			double maxLoad = 167;

			// Генерация случайных величин (СВ)
			Random rand = new Random();

			// Лист для хранения СВ генерации
			List<double> randValueGen = new List<double>();

			// Лист для хранения СВ нагрузки
			List<double> randValueLoad = new List<double>();

			// Генерация случайного числа генерации в цикле с условием
			while (randValueGen.Count < 1000)
			{
				double q = rand.NextDouble();

				if (q >= 0 && q < v3)
				{
					Normal normalDistribution = new Normal(mo3, sko3);
					double part3 = Math.Round(normalDistribution.Sample(), 0);
					if (part3 >= minGen && part3 < maxGen)
					{
						randValueGen.Add(part3);
					}
				}

				else if (q >= v3 && q < (v3 + v1))
				{
					ContinuousUniform uniformDist = new ContinuousUniform(lowerBound, upperBound);
					double part1 = Math.Round(uniformDist.Sample(), 0);
					if (part1 >= minGen && part1 < maxGen)
					{
						randValueGen.Add(part1);
					}
				}

				else if (q >= (v3 + v1) && q < (v3 + v1 + v2))
				{
					Normal normalDistribution = new Normal(mo2, sko2);
					double part2 = Math.Round(normalDistribution.Sample(), 0);
					if (part2 >= minGen && part2 < maxGen)
					{
						randValueGen.Add(part2);
					}
				}
			}

			// Генерация случайного числа нагрузки в цикле с условием
			while (randValueLoad.Count < 1000)
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

			// Отсортированные списки по возрастанию
			var sotrValueGen = randValueGen.OrderBy(x => x).ToList();
			var sortValueLoad = randValueLoad.OrderBy(x => x).ToList();

			// Путь до файла Excel
			string folder = @"C:\Users\aat146\Desktop\ПроизПрактика";
			string fileExcel = "Результат.xlsx";
			string xlsxFile = Path.Combine(folder, fileExcel);

			// Создание книги и листа
			Application excelApp = new Application();
			Workbook workbook = excelApp.Workbooks.Add();
			Worksheet worksheet = workbook.Sheets.Add();
			worksheet.Name = "Случайные величины";

			// Создание указателя на экземпляр RastrWin и его запуск
			IRastr rastr = new Rastr();

			// Загрузка файла
			string file = @"C:\Users\aat146\Desktop\ПроизПрактика\Растр\Режим.rg2";
			string shablon = @"C:\Programs\RastrWin3\RastrWin3\SHABLON\режим.rg2";

			rastr.Load(RG_KOD.RG_REPL, file, shablon);

			// Объявление объекта, содержащего таблицу "Узлы"
			ITable tableNode = (ITable)rastr.Tables.Item("node");

			// Объявление объекта, содержащего таблицу "Генератор(УР)"
			ITable tableGenYR = (ITable)rastr.Tables.Item("Generator");

			// Объявление объекта, содержащего таблицу "Ветви"
			ITable tableVetv = (ITable)rastr.Tables.Item("vetv");

			// Объявление объекта, содержащего таблицу "Сечения"
			ITable tableSechen = (ITable)rastr.Tables.Item("sechen");

			// Узлы
			ICol numberNode = (ICol)tableNode.Cols.Item("ny");   // Номер
			ICol nameNode = (ICol)tableNode.Cols.Item("name");   // Название
			ICol activeGen = (ICol)tableNode.Cols.Item("pg");   // Акт. мощность генерации
			ICol activeLoad = (ICol)tableNode.Cols.Item("pn");   // Акт. мощность нагрузки

			// Генераторы(УР)
			ICol nAgr = (ICol)tableGenYR.Cols.Item("Num"); // Номер агрегата
			ICol nameGenYR = (ICol)tableGenYR.Cols.Item("Name"); // Название
			ICol pGenYR = (ICol)tableGenYR.Cols.Item("P"); // Акт. мощность генерации

			// Ветви
			ICol staVetv = (ICol)tableVetv.Cols.Item("sta");   // Состояние
			ICol tipVetv = (ICol)tableVetv.Cols.Item("tip");   // Тип
			ICol nStart = (ICol)tableVetv.Cols.Item("ip");   // Номер начала
			ICol nEnd = (ICol)tableVetv.Cols.Item("iq");   // Номер конца
			ICol nParall = (ICol)tableVetv.Cols.Item("np");   // Номер параллельности
			ICol nameVetv = (ICol)tableVetv.Cols.Item("name");   // Название
			ICol pVetvEnd = (ICol)tableVetv.Cols.Item("pl_iq");   // Поток P в конце ветви

			// Сечения
			ICol nSech = (ICol)tableSechen.Cols.Item("ns"); // Номер сечения
			ICol nameSech = (ICol)tableSechen.Cols.Item("name"); // Имя сечения
			ICol minSech = (ICol)tableSechen.Cols.Item("pmin"); // Минимальное значение
			ICol maxSech = (ICol)tableSechen.Cols.Item("pmax"); // Максимальное значение
			ICol valueSech = (ICol)tableSechen.Cols.Item("psech"); // Полученное значение

			// Лист для хранения перетока по Пеледуй - Сухой Лог I и II цепь
			List<double> v1PeledSyxLog = new List<double>();
			List<double> v2PeledSyxLog = new List<double>();

			// Лист для хранения перетока по Таксимо - Мамакан I и II цепь
			List<double> v1TaksimoMamakan = new List<double>();
			List<double> v2TaksimoMamakan = new List<double>();

			// Лист для хранения перетока по КС
			List<double> ksPeledSyxLog = new List<double>();
			List<double> ksTaksimoMamakan = new List<double>();

			//// Цикл расчета перетоков в RastrWin3
			for (int i = 0; i < 1000; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Nym=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = sotrValueGen[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index1] = sortValueLoad[i];

				// Расчет УР
				_ = rastr.rgm("");

				// Считывание перетоков по каждой ветви
				var setSelVetv1 = "ip=" + 2 + "&" + "iq=" + 3 + "&" + "np=" + 1;
				tableVetv.SetSel(setSelVetv1);
				var index3 = tableVetv.FindNextSel[-1];
				v1PeledSyxLog.Add(pVetvEnd.Z[index3]);

				var setSelVetv2 = "ip=" + 2 + "&" + "iq=" + 3 + "&" + "np=" + 2;
				tableVetv.SetSel(setSelVetv2);
				var index4 = tableVetv.FindNextSel[-1];
				v2PeledSyxLog.Add(pVetvEnd.Z[index4]);

				var setSelVetv3 = "ip=" + 2 + "&" + "iq=" + 4 + "&" + "np=" + 1;
				tableVetv.SetSel(setSelVetv3);
				var index5 = tableVetv.FindNextSel[-1];
				v1TaksimoMamakan.Add(pVetvEnd.Z[index5]);

				var setSelVetv4 = "ip=" + 2 + "&" + "iq=" + 4 + "&" + "np=" + 2;
				tableVetv.SetSel(setSelVetv4);
				var index6 = tableVetv.FindNextSel[-1];
				v1TaksimoMamakan.Add(pVetvEnd.Z[index6]);

				// Считывание перетоков по каждому КС
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index7 = tableNode.FindNextSel[-1];
				ksPeledSyxLog.Add(valueSech.Z[index7]);
				
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index8 = tableNode.FindNextSel[-1];
				ksTaksimoMamakan.Add(valueSech.Z[index8]);
			}






			Console.WriteLine($"Работа алгоритма.\n");
			//Console.WriteLine($"Сучайные числа генерации:\n");

			// Вывод значений на экран и в excel
			for (int i = 0; i < randValueGen.Count; i++)
			{
				//Console.WriteLine(randValueGen[i]);

				// Получаем диапазон ячеек начиная с ячейки A1
				Range range = worksheet.Range["A1"];

				// Запись случайной величины в столбец А
				range.Offset[0, 0].Value = "Генерация";
				range.Offset[i + 1, 0].Value = randValueGen[i];
			}

			Console.WriteLine("\nСучайные числа нагрузки:\n");

			// Вывод значений на экран и в excel
			for (int i = 0; i < randValueLoad.Count; i++)
			{
				//Console.WriteLine(randValueLoad[i]);

				// Получаем диапазон ячеек начиная с ячейки A1
				Range range = worksheet.Range["A1"];

				// Запись случайной величины в столбец А
				range.Offset[0, 1].Value = "Нагрузка";
				range.Offset[i + 1, 1].Value = randValueLoad[i];
			}



			workbook.SaveAs(xlsxFile);
			workbook.Close();
			excelApp.Quit();

			// Останавливаем счетчик
			stopwatch.Stop();

			Console.WriteLine($"\nВремя расчета: {stopwatch.ElapsedMilliseconds} мс\n" +
				$"Файл Excel успешно сохранен по пути: {xlsxFile}\n" +
				$"Количество СВ генерации: {randValueGen.Count}\n" +
				$"Количество СВ нагрузки: {randValueLoad.Count}\n");

			Console.ReadKey();

			//var setSelName = "ny=" + 5;   // Переменная ny = 5 (№ узла = 5)
			//tableNode.SetSel(setSelName);   // Выборка по переменной
			//var index = tableNode.FindNextSel[-1];   // Возврат индекса след.строки, удовл-ей выборке (искл: -1)
			//activeLoad.Z[index] = rdm;   // Переменная с найденным индексом в столбце Название
			//Console.WriteLine($"Узел № {index} || Нагрузка: {activeLoad.Z[index]}");


			//int p = 500;
			//powerActiveGeneration.Z[index] = p;

			//var setSelVetv = "ip=" + 2 + "&" + "iq=" + 3 + "&" + "np=" + 2;
			//tableVetv.SetSel(setSelVetv);
			//var number = tableVetv.FindNextSel[-1];
			//staVetv.Z[number] = 1;    // 1 - отключение; 0 -включение
			//var name1v = nameVetv.Z[number];
			//Console.WriteLine($"Название ветви: {name1v}");

			// Расчет УР
			//_ = rastr.rgm("");

			//// Сохранение результатов
			//string fileNew = @"C:\Users\aat146\Desktop\ПроизПрактика\Растр\Режим2.rg2";
			//rastr.Save(fileNew, shablon);
		}
	}
}
