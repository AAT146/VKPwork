using System;
using System.Collections.Generic;
using System.IO;
using ASTRALib;
using OfficeOpenXml;
using MathNet.Numerics.Distributions;


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
		public static double[] ReadFileFromExcel(string filePath)
		{
			using (var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var worksheet = package.Workbook.Worksheets[0];
				var data = new double[worksheet.Dimension.Rows];

				for (int i = 1; i <= worksheet.Dimension.Rows; i++)
				{
					data[i - 1] = worksheet.Cells[i, 1].GetValue<double>();
				}

				return data;
			}
		}

		/// <summary>
		/// Упрощенное моделирование.
		/// </summary>
		public static void Main()
		{
			// Создание указателя на экземпляр RastrWin и его запуск
			IRastr rastr = new Rastr();

			// Загрузка файл
			string file = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\Режим.rg2";
			string shablon = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\режим.rg2";

			rastr.Load(RG_KOD.RG_REPL, file, shablon);

			// Объявление объекта, содержащего таблицу "Узлы"
			ITable tableNode = (ITable)rastr.Tables.Item("node");

			// Объявление объекта, содержащего таблицу "Генератор(УР)"
			ITable tableGenYR = (ITable)rastr.Tables.Item("Generator");

			// Объявление объекта, содержащего таблицу "Ветви"
			ITable tableVetv = (ITable)rastr.Tables.Item("vetv");

			// Узлы
			ICol numberNode = (ICol)tableNode.Cols.Item("ny");   // Номер
			ICol nameNode = (ICol)tableNode.Cols.Item("name");   // Название
			ICol activeGen = (ICol)tableNode.Cols.Item("pg");   // Мощность генерации
			ICol activeLoad = (ICol)tableNode.Cols.Item("pn");   // Мощность нагрузки

			// Ветви
			ICol staVetv = (ICol)tableVetv.Cols.Item("sta");   // Состояние
			ICol tipVetv = (ICol)tableVetv.Cols.Item("tip");   // Тип
			ICol nStart = (ICol)tableVetv.Cols.Item("ip");   // Номер начала
			ICol nEnd = (ICol)tableVetv.Cols.Item("iq");   // Номер конца
			ICol nParall = (ICol)tableVetv.Cols.Item("np");   // Номер параллельности
			ICol nameVetv = (ICol)tableVetv.Cols.Item("name");   // Название

			// Файл Excel генеральной совопукности
			string xlsxLoad = "C:\\Users\\Анастасия\\Desktop\\ПроизПрактика\\Растр\\Load.xlsx";
			//string xlsxGenerator = "C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\\Generator.xlsx";

			// Чтение данных из файла Excel
			double[] dataLoad = ReadFileFromExcel(xlsxLoad);
			//double[] dataGenerator = ReadFileFromExcel(xlsxGenerator);

			// Константы для двух распределений
			double mo1 = 104;
			double sko1 = 9.6;
			double k1 = 0.4;
			double mo2 = 109;
			double sko2 = 5.6;
			double k2 = 0.6;

			// Создаем нормальное распределение
			Normal normalDistribution = new Normal(mean: mo1, stddev: sko1);
			double sample = normalDistribution.Sample();

			//NormalDistribution standardNormal1 = new NormalDistribution(mo1, sko1);
			//NormalDistribution standardNormal2 = new NormalDistribution(mo2, sko2);

			Console.WriteLine($"{sample}");
			
			//// Создаем нормальное распределение с МО = 0 и СКО = 1
			//NormalDistribution standardNormal = new NormalDistribution();

			//// Масштабируем и сдвигаем распределение
			//NormalDistribution normal1 = (standardNormal * sko1) + mo1;

			//// Генерация случайной величины по нормальному распределению
			//Random rand = new Random();
			//double a = rand.NextDouble();
			//double normal = Math.Sqrt(-2.0 * Math.Log(a)) * Math.Sin(2.0 * Math.PI * a);

			//// Формирование случайной величины
			//double rdm1 = k1 * (Math.Round((mo1 + (sko1 * normal)), 2));
			//double rdm2 = k2 * (Math.Round((mo2 + (sko2 * normal)), 2));
			//double rdm = rdm1 + rdm2;

			//Console.WriteLine($"МО 1: {mo1} || СКО 1: {sko1} || Коэфф 1: {k1} || СВ 1: {rdm1}\n" +
			//	$"МО 2: {mo2} || СКО 2: {sko2} || Коэфф 2: {k2} || СВ 2: {rdm2}\n" +
			//	$"СВ: {rdm}");

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
			_ = rastr.rgm("");

			// Сохранение результатов
			string fileNew = @"C:\Users\Анастасия\Desktop\ПроизПрактика\Растр\Режим2.rg2";
			rastr.Save(fileNew, shablon);
		}
	}
}
