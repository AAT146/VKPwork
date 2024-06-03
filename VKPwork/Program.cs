using ASTRALib;
using ClassLibrary;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;


namespace VKPwork
{
	/// <summary>
	/// Расчета ПБН на примере Бодайбинского ЭР Иркутской ОЗ.
	/// </summary>
	public class Program
	{
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

			// Создание экземпляра класса
			SchemeNumber numberScheme = new SchemeNumber();
			PowerLineBefore powerLineBefore = new PowerLineBefore();
			PowerLineAfter powerLineAfter = new PowerLineAfter();

			// Дельта снижения нагрузки
			const double deltaLoad = 1;

			//// Листы (ЗИМА || ДО || ПУР)
			List<double> ksPSLBeforeWinter1 = new List<double>();
			List<double> ksTMBeforeWinter1 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadWinterBefore1 = new List<(int, double)>();

			double numberYRwinter1 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 59676; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenWinter()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadWinter()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 53);

				// Топология сети
				for (int  j = 0; j < numberScheme.numberBefore.Length; j++)
				{
					if (numberScheme.numberBefore[j] == r)
					{
						int[][] workScheme = powerLineBefore.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpWinterBeforePurTM = Excel.mdpTMpurBeforeWinter[r - 1];
				double mdpWinterBeforePurPSL = Excel.mdpPSLpurBeforeWinter[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRwinter1++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLBeforeWinter1.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMBeforeWinter1.Add(ksTM);

				if (ksPSL > mdpWinterBeforePurPSL || ksTM > mdpWinterBeforePurPSL)
				{
					while (!(ksPSL <= mdpWinterBeforePurPSL && ksTM <= mdpWinterBeforePurPSL))
					{
						double newLoadWinterBefore1 = RandomValue.RndValueLoadWinter()[i] - deltaLoad;
						if (newLoadWinterBefore1 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadWinterBefore1;
							_ = rastr.rgm("");
							numberYRwinter1++;

							listNewLoadWinterBefore1.Add((i, newLoadWinterBefore1));
						}
					}
				}
			}

			//// Листы (ЗИМА || ДО || СМЗУ)
			List<double> ksPSLBeforeWinter2 = new List<double>();
			List<double> ksTMBeforeWinter2 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadWinterBefore2 = new List<(int, double)>();

			double numberYRwinter2 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 59676; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenWinter()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadWinter()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 53);

				// Топология сети
				for (int j = 0; j < numberScheme.numberBefore.Length; j++)
				{
					if (numberScheme.numberBefore[j] == r)
					{
						int[][] workScheme = powerLineBefore.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpWinterBeforeSmzyTM = Excel.mdpTMsmzyBeforeWinter[r - 1];
				double mdpWinterBeforeSmzyPSL = Excel.mdpPSLsmzyBeforeWinter[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRwinter2++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLBeforeWinter2.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMBeforeWinter2.Add(ksTM);

				if (ksPSL > mdpWinterBeforeSmzyPSL || ksTM > mdpWinterBeforeSmzyPSL)
				{
					while (!(ksPSL <= mdpWinterBeforeSmzyPSL && ksTM <= mdpWinterBeforeSmzyPSL))
					{
						double newLoadWinterBefore2 = RandomValue.RndValueLoadWinter()[i] - deltaLoad;
						if (newLoadWinterBefore2 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadWinterBefore2;
							_ = rastr.rgm("");
							numberYRwinter2++;

							listNewLoadWinterBefore2.Add((i, newLoadWinterBefore2));
						}
					}
				}
			}

			//// Листы (ЗИМА || ПОСЛЕ || ПУР)
			List<double> ksPSLAfterWinter1 = new List<double>();
			List<double> ksTMAfterWinter1 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadWinterAfter1 = new List<(int, double)>();

			double numberYRwinter3 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 59676; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenWinter()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadWinter()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 72);

				// Топология сети
				for (int j = 0; j < numberScheme.numberAfter.Length; j++)
				{
					if (numberScheme.numberAfter[j] == r)
					{
						int[][] workScheme = powerLineAfter.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpWinterAfterPurTM = Excel.mdpTMpurAfterWinter[r - 1];
				double mdpWinterAfterPurPSL = Excel.mdpPSLpurAfterWinter[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRwinter3++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLAfterWinter1.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMAfterWinter1.Add(ksTM);

				if (ksPSL > mdpWinterAfterPurTM || ksTM > mdpWinterAfterPurPSL)
				{
					while (!(ksPSL <= mdpWinterAfterPurTM && ksTM <= mdpWinterAfterPurPSL))
					{
						double newLoadWinterAfter1 = RandomValue.RndValueLoadWinter()[i] - deltaLoad;
						if (newLoadWinterAfter1 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadWinterAfter1;
							_ = rastr.rgm("");
							numberYRwinter3++;

							listNewLoadWinterAfter1.Add((i, newLoadWinterAfter1));
						}
					}
				}
			}

			//// Листы (ЗИМА || ПОСЛЕ || СМЗУ)
			List<double> ksPSLAfterWinter2 = new List<double>();
			List<double> ksTMAfterWinter2 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadWinterAfter2 = new List<(int, double)>();

			double numberYRwinter4 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 59676; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenWinter()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadWinter()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 72);

				// Топология сети
				for (int j = 0; j < numberScheme.numberAfter.Length; j++)
				{
					if (numberScheme.numberAfter[j] == r)
					{
						int[][] workScheme = powerLineAfter.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpWinterAfterSmzyTM = Excel.mdpTMsmzyAfterWinter[r - 1];
				double mdpWinterAfterSmzyPSL = Excel.mdpPSLsmzyAfterWinter[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRwinter4++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLAfterWinter2.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMAfterWinter2.Add(ksTM);

				if (ksPSL > mdpWinterAfterSmzyTM || ksTM > mdpWinterAfterSmzyPSL)
				{
					while (!(ksPSL <= mdpWinterAfterSmzyTM && ksTM <= mdpWinterAfterSmzyPSL))
					{
						double newLoadWinterAfter2 = RandomValue.RndValueLoadWinter()[i] - deltaLoad;
						if (newLoadWinterAfter2 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadWinterAfter2;
							_ = rastr.rgm("");
							numberYRwinter4++;

							listNewLoadWinterAfter2.Add((i, newLoadWinterAfter2));
						}
					}
				}
			}

			//// Листы (ЛЕТО || ДО || ПУР)
			List<double> ksPSLBeforeSummer1 = new List<double>();
			List<double> ksTMBeforeSummer1 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadSummerBefore1 = new List<(int, double)>();

			double numberYRsummer1 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 45733; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenSummer()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadSummer()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 53);

				// Топология сети
				for (int j = 0; j < numberScheme.numberBefore.Length; j++)
				{
					if (numberScheme.numberBefore[j] == r)
					{
						int[][] workScheme = powerLineBefore.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpSummerBeforePurTM = Excel.mdpTMpurBeforeSummer[r - 1];
				double mdpSummerBeforePurPSL = Excel.mdpPSLpurBeforeSummer[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRsummer1++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLBeforeSummer1.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMBeforeSummer1.Add(ksTM);

				if (ksPSL > mdpSummerBeforePurTM || ksTM > mdpSummerBeforePurPSL)
				{
					while (!(ksPSL <= mdpSummerBeforePurTM && ksTM <= mdpSummerBeforePurPSL))
					{
						double newLoadSummerBefore1 = RandomValue.RndValueLoadSummer()[i] - deltaLoad;
						if (newLoadSummerBefore1 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadSummerBefore1;
							_ = rastr.rgm("");
							numberYRsummer1++;

							listNewLoadSummerBefore1.Add((i, newLoadSummerBefore1));
						}
					}
				}
			}

			//// Листы (ЛЕТО || ДО || СМЗУ)
			List<double> ksPSLBeforeSummer2 = new List<double>();
			List<double> ksTMBeforeSummer2 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadSummerBefore2 = new List<(int, double)>();

			double numberYRsummer2 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 45733; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenSummer()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadSummer()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 53);

				// Топология сети
				for (int j = 0; j < numberScheme.numberBefore.Length; j++)
				{
					if (numberScheme.numberBefore[j] == r)
					{
						int[][] workScheme = powerLineBefore.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpSummerBeforeSmzyTM = Excel.mdpTMsmzyBeforeSummer[r - 1];
				double mdpSummerBeforeSmzyPSL = Excel.mdpPSLsmzyBeforeSummer[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRsummer2++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLBeforeSummer2.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMBeforeSummer2.Add(ksTM);

				if (ksPSL > mdpSummerBeforeSmzyTM || ksTM > mdpSummerBeforeSmzyPSL)
				{
					while (!(ksPSL <= mdpSummerBeforeSmzyTM && ksTM <= mdpSummerBeforeSmzyPSL))
					{
						double newLoadSummerBefore2 = RandomValue.RndValueLoadSummer()[i] - deltaLoad;
						if (newLoadSummerBefore2 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadSummerBefore2;
							_ = rastr.rgm("");
							numberYRsummer2++;

							listNewLoadSummerBefore2.Add((i, newLoadSummerBefore2));
						}
					}
				}
			}

			//// Листы (ЛЕТО || ПОСЛЕ || ПУР)
			List<double> ksPSLAfterSummer1 = new List<double>();
			List<double> ksTMAfterSummer1 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadSummerAfter1 = new List<(int, double)>();

			double numberYRsummer3 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 45733; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenSummer()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadSummer()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 72);

				// Топология сети
				for (int j = 0; j < numberScheme.numberAfter.Length; j++)
				{
					if (numberScheme.numberAfter[j] == r)
					{
						int[][] workScheme = powerLineAfter.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpSummerAfterPurTM = Excel.mdpTMpurAfterSummer[r - 1];
				double mdpSummerAfterPurPSL = Excel.mdpPSLpurAfterSummer[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRsummer3++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLAfterSummer1.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMAfterSummer1.Add(ksTM);

				if (ksPSL > mdpSummerAfterPurTM || ksTM > mdpSummerAfterPurPSL)
				{
					while (!(ksPSL <= mdpSummerAfterPurTM && ksTM <= mdpSummerAfterPurPSL))
					{
						double newLoadSummerAfter1 = RandomValue.RndValueLoadSummer()[i] - deltaLoad;
						if (newLoadSummerAfter1 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadSummerAfter1;
							_ = rastr.rgm("");
							numberYRsummer3++;

							listNewLoadSummerAfter1.Add((i, newLoadSummerAfter1));
						}
					}
				}
			}

			//// Листы (ЛЕТО || ПОСЛЕ || СМЗУ)
			List<double> ksPSLAfterSummer2 = new List<double>();
			List<double> ksTMAfterSummer2 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadSummerAfter2 = new List<(int, double)>();

			double numberYRsummer4 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 45733; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 6;
				tableGenYR.SetSel(setSelAgr);
				var index1 = tableGenYR.FindNextSel[-1];
				pGenYR.Z[index1] = RandomValue.RndValueGenSummer()[i];

				// Присвоение нового числа мощности нагрузки
				var setSelNy = "ny=" + 5;
				tableNode.SetSel(setSelNy);
				var index2 = tableNode.FindNextSel[-1];
				activeLoad.Z[index2] = RandomValue.RndValueLoadSummer()[i];

				// Определение номера схемы
				Random random = new Random();
				int r = random.Next(1, 72);

				// Топология сети
				for (int j = 0; j < numberScheme.numberAfter.Length; j++)
				{
					if (numberScheme.numberAfter[j] == r)
					{
						int[][] workScheme = powerLineAfter.schemes[r - 1];
						foreach (int[] array in workScheme)
						{
							var setSelVetv = "ip=" + array[0] + "&" + "iq=" + array[1] + "&" + "np=" + array[2];
							tableVetv.SetSel(setSelVetv);
							var number = tableVetv.FindNextSel[-1];
							staVetv.Z[number] = 1;
						}
					}
				}

				// Определние МДП по топологии сети
				double mdpSummerAfterSmzyTM = Excel.mdpTMsmzyAfterSummer[r - 1];
				double mdpSummerAfterSmzyPSL = Excel.mdpPSLsmzyAfterSummer[r - 1];

				// Расчет УР
				_ = rastr.rgm("");
				numberYRsummer4++;

				// КС Пеледуй - Сухой-Лог
				var setSelNs1 = "ns=" + 1;
				tableSechen.SetSel(setSelNs1);
				var index3 = tableSechen.FindNextSel[-1];
				var ksPSL = Math.Round(valueSech.Z[index3], 0);
				ksPSLAfterSummer2.Add(ksPSL);

				// КС Таксимо - Мамакан
				var setSelNs2 = "ns=" + 2;
				tableSechen.SetSel(setSelNs2);
				var index4 = tableSechen.FindNextSel[-1];
				var ksTM = Math.Round(valueSech.Z[index4], 0);
				ksTMAfterSummer2.Add(ksTM);

				if (ksPSL > mdpSummerAfterSmzyTM || ksTM > mdpSummerAfterSmzyPSL)
				{
					while (!(ksPSL <= mdpSummerAfterSmzyTM && ksTM <= mdpSummerAfterSmzyPSL))
					{
						double newLoadSummerAfter2 = RandomValue.RndValueLoadSummer()[i] - deltaLoad;
						if (newLoadSummerAfter2 == 0)
						{
							break;
						}
						else
						{
							activeLoad.Z[index2] = newLoadSummerAfter2;
							_ = rastr.rgm("");
							numberYRsummer4++;

							listNewLoadSummerAfter2.Add((i, newLoadSummerAfter2));
						}
					}
				}
			}


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
