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

			Console.WriteLine($"Выполнение процесса.\n");

			// Создание указателя на экземпляр RastrWin и его запуск
			IRastr rastr = new Rastr();

			// Загрузка файла
			string fileRegim = @"C:\Users\Анастасия\Desktop\NewWork\Растр\Режим.rg2";
			string shablonRegim = @"C:\Users\Анастасия\Documents\RastrWin3\SHABLON\режим.rg2";

			rastr.Load(RG_KOD.RG_REPL, fileRegim, shablonRegim);

			string fileSechen = @"C:\Users\Анастасия\Desktop\NewWork\Растр\Режим.sch";
			string shablonSechen = @"C:\Users\Анастасия\Documents\RastrWin3\SHABLON\сечения.sch";

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
			double deltaLoad = 10;

			//// Листы (ЗИМА || ДО || ПУР)
			List<double> ksPSLBeforeWinter1 = new List<double>();
			List<double> ksTMBeforeWinter1 = new List<double>();
			List<(int step, double newLoadWinterBefore)> listNewLoadWinterBefore1 = new List<(int, double)>();
			List<int> nScheme1 = new List<int>();

			double numberYRwinter1 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme1.Add(r);

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
			List<int> nScheme2 = new List<int>();

			double numberYRwinter2 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme2.Add(r);

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
			List<int> nScheme3 = new List<int>();

			double numberYRwinter3 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme3.Add(r);

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
			List<int> nScheme4 = new List<int>();

			double numberYRwinter4 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme4.Add(r);

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
			List<int> nScheme5 = new List<int>();

			double numberYRsummer1 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme5.Add(r);

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
			List<int> nScheme6 = new List<int>();

			double numberYRsummer2 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme6.Add(r);

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
			List<int> nScheme7 = new List<int>();

			double numberYRsummer3 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme7.Add(r);

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
			List<int> nScheme8 = new List<int>();

			double numberYRsummer4 = 0;

			// Цикл расчета в RastrWin3
			for (int i = 0; i < 2; i++)
			{
				// Присвоение нового числа мощности генерации
				var setSelAgr = "Num=" + 2;
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
				nScheme8.Add(r);

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

			// Путь до файла Excel
			string folder = @"C:\Users\Анастасия\Desktop\NewWork";
			string fileExcel1 = "Зима.xlsx";
			string xlsxFile1 = Path.Combine(folder, fileExcel1);
			string fileExcel2 = "Лето.xlsx";
			string xlsxFile2 = Path.Combine(folder, fileExcel2);

			// Создание книги и листа для файла ЗИМА
			Application excelApp1 = new Application();
			Workbook workbook1 = excelApp1.Workbooks.Add();
			Worksheet worksheet1 = workbook1.Sheets.Add();
			worksheet1.Name = "ДО_ПУР";
			Worksheet worksheet2 = workbook1.Sheets.Add();
			worksheet2.Name = "ДО_СМЗУ";
			Worksheet worksheet3 = workbook1.Sheets.Add();
			worksheet3.Name = "ПОСЛЕ_ПУР";
			Worksheet worksheet4 = workbook1.Sheets.Add();
			worksheet4.Name = "ПОСЛЕ_СМЗУ";

			// Запись значений в файл Excel ЗИМА
			for (int i = 0; i < 2; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range1 = worksheet1.Range["A1"];
				Range range2 = worksheet2.Range["A1"];
				Range range3 = worksheet3.Range["A1"];
				Range range4 = worksheet4.Range["A1"];

				// Запись случайной величины в столбец А листа 1 - генерация
				range1.Offset[0, 0].Value = "Генерация";
				range1.Offset[i + 1, 0].Value = RandomValue.RndValueGenWinter()[i];

				// Запись случайной величины в столбец B листа 1 - нагрузка
				range1.Offset[0, 1].Value = "Нагрузка";
				range1.Offset[i + 1, 1].Value = RandomValue.RndValueLoadWinter()[i];

				// Запись случайной величины в столбец C листа 1 - КС Пеледуй - Сухой Лог
				range1.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range1.Offset[i + 1, 2].Value = ksPSLBeforeWinter1[i];

				// Запись случайной величины в столбец D листа 1 - КС Таксимо - Мамакан
				range1.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range1.Offset[i + 1, 3].Value = ksTMBeforeWinter1[i];

				// Запись Номера схемы сети
				range1.Offset[0, 4].Value = "№ Схемы";
				range1.Offset[i + 1, 4].Value = nScheme1[i];

				// Запись шага итерации, на котором проищошло превышение 
				range1.Offset[0, 5].Value = "";
				range1.Offset[i + 1, 5].Value = listNewLoadWinterBefore1[i];

				// Запись случайной величины в столбец А листа 2 - генерация
				range2.Offset[0, 0].Value = "Генерация";
				range2.Offset[i + 1, 0].Value = RandomValue.RndValueGenWinter()[i];

				// Запись случайной величины в столбец B листа 2 - нагрузка
				range2.Offset[0, 1].Value = "Нагрузка";
				range2.Offset[i + 1, 1].Value = RandomValue.RndValueLoadWinter()[i];

				// Запись случайной величины в столбец C листа 1 - КС Пеледуй - Сухой Лог
				range2.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range2.Offset[i + 1, 2].Value = ksPSLBeforeWinter2[i];

				// Запись случайной величины в столбец D листа 1 - КС Таксимо - Мамакан
				range2.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range2.Offset[i + 1, 3].Value = ksTMBeforeWinter2[i];

				// Запись Номера схемы сети
				range2.Offset[0, 4].Value = "№ Схемы";
				range2.Offset[i + 1, 4].Value = nScheme2[i];

				// Запись шага итерации, на котором проищошло превышение 
				range2.Offset[0, 5].Value = "";
				range2.Offset[i + 1, 5].Value = listNewLoadWinterBefore2[i];

				// Запись случайной величины в столбец А листа 3 - генерация
				range3.Offset[0, 0].Value = "Генерация";
				range3.Offset[i + 1, 0].Value = RandomValue.RndValueGenWinter()[i];

				// Запись случайной величины в столбец B листа 3 - нагрузка
				range3.Offset[0, 1].Value = "Нагрузка";
				range3.Offset[i + 1, 1].Value = RandomValue.RndValueLoadWinter()[i];

				// Запись случайной величины в столбец C листа 3 - КС Пеледуй - Сухой Лог
				range3.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range3.Offset[i + 1, 2].Value = ksPSLAfterWinter1[i];

				// Запись случайной величины в столбец D листа 3 - КС Таксимо - Мамакан
				range3.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range3.Offset[i + 1, 3].Value = ksTMAfterWinter1[i];

				// Запись Номера схемы сети
				range3.Offset[0, 4].Value = "№ Схемы";
				range3.Offset[i + 1, 4].Value = nScheme3[i];

				// Запись шага итерации, на котором проищошло превышение 
				range3.Offset[0, 5].Value = "";
				range3.Offset[i + 1, 5].Value = listNewLoadWinterAfter1[i];

				// Запись случайной величины в столбец А листа 4 - генерация
				range4.Offset[0, 0].Value = "Генерация";
				range4.Offset[i + 1, 0].Value = RandomValue.RndValueGenWinter()[i];

				// Запись случайной величины в столбец B листа 4 - нагрузка
				range4.Offset[0, 1].Value = "Нагрузка";
				range4.Offset[i + 1, 1].Value = RandomValue.RndValueLoadWinter()[i];

				// Запись случайной величины в столбец C листа 4 - КС Пеледуй - Сухой Лог
				range4.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range4.Offset[i + 1, 2].Value = ksPSLAfterWinter2[i];

				// Запись случайной величины в столбец D листа 4 - КС Таксимо - Мамакан
				range4.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range4.Offset[i + 1, 3].Value = ksTMAfterWinter2[i];

				// Запись Номера схемы сети
				range4.Offset[0, 4].Value = "№ Схемы";
				range4.Offset[i + 1, 4].Value = nScheme4[i];

				// Запись шага итерации, на котором проищошло превышение 
				range4.Offset[0, 5].Value = "";
				range4.Offset[i + 1, 5].Value = listNewLoadWinterAfter2[i];
			}

			workbook1.SaveAs(xlsxFile1);
			workbook1.Close();
			excelApp1.Quit();

			// Создание книги и листа для файла ЛЕТО
			Application excelApp2 = new Application();
			Workbook workbook2 = excelApp2.Workbooks.Add();
			Worksheet worksheet5 = workbook2.Sheets.Add();
			worksheet5.Name = "ДО_ПУР";
			Worksheet worksheet6 = workbook2.Sheets.Add();
			worksheet6.Name = "ДО_СМЗУ";
			Worksheet worksheet7 = workbook2.Sheets.Add();
			worksheet7.Name = "ПОСЛЕ_ПУР";
			Worksheet worksheet8 = workbook2.Sheets.Add();
			worksheet8.Name = "ПОСЛЕ_СМЗУ";

			// Запись значений в файл Excel ЛЕТО
			for (int i = 0; i < 2; i++)
			{
				// Получаем диапазон ячеек начиная с ячейки A1
				Range range5 = worksheet5.Range["A1"];
				Range range6 = worksheet6.Range["A1"];
				Range range7 = worksheet7.Range["A1"];
				Range range8 = worksheet8.Range["A1"];

				// Запись случайной величины в столбец А листа 1 - генерация
				range5.Offset[0, 0].Value = "Генерация";
				range5.Offset[i + 1, 0].Value = RandomValue.RndValueGenSummer()[i];

				// Запись случайной величины в столбец B листа 1 - нагрузка
				range5.Offset[0, 1].Value = "Нагрузка";
				range5.Offset[i + 1, 1].Value = RandomValue.RndValueLoadSummer()[i];

				// Запись случайной величины в столбец C листа 1 - КС Пеледуй - Сухой Лог
				range5.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range5.Offset[i + 1, 2].Value = ksPSLBeforeSummer1[i];

				// Запись случайной величины в столбец D листа 1 - КС Таксимо - Мамакан
				range5.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range5.Offset[i + 1, 3].Value = ksTMBeforeSummer1[i];

				// Запись Номера схемы сети
				range5.Offset[0, 4].Value = "№ Схемы";
				range5.Offset[i + 1, 4].Value = nScheme5[i];

				// Запись шага итерации, на котором проищошло превышение 
				range5.Offset[0, 5].Value = "";
				range5.Offset[i + 1, 5].Value = listNewLoadSummerBefore1[i];

				// Запись случайной величины в столбец А листа 2 - генерация
				range6.Offset[0, 0].Value = "Генерация";
				range6.Offset[i + 1, 0].Value = RandomValue.RndValueGenSummer()[i];

				// Запись случайной величины в столбец B листа 2 - нагрузка
				range6.Offset[0, 1].Value = "Нагрузка";
				range6.Offset[i + 1, 1].Value = RandomValue.RndValueLoadSummer()[i];

				// Запись случайной величины в столбец C листа 1 - КС Пеледуй - Сухой Лог
				range6.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range6.Offset[i + 1, 2].Value = ksPSLBeforeSummer2[i];

				// Запись случайной величины в столбец D листа 1 - КС Таксимо - Мамакан
				range6.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range6.Offset[i + 1, 3].Value = ksTMBeforeSummer2[i];

				// Запись Номера схемы сети
				range6.Offset[0, 4].Value = "№ Схемы";
				range6.Offset[i + 1, 4].Value = nScheme6[i];

				// Запись шага итерации, на котором проищошло превышение 
				range6.Offset[0, 5].Value = "";
				range6.Offset[i + 1, 5].Value = listNewLoadSummerBefore2[i];

				// Запись случайной величины в столбец А листа 3 - генерация
				range7.Offset[0, 0].Value = "Генерация";
				range7.Offset[i + 1, 0].Value = RandomValue.RndValueGenSummer()[i];

				// Запись случайной величины в столбец B листа 3 - нагрузка
				range7.Offset[0, 1].Value = "Нагрузка";
				range7.Offset[i + 1, 1].Value = RandomValue.RndValueLoadSummer()[i];

				// Запись случайной величины в столбец C листа 3 - КС Пеледуй - Сухой Лог
				range7.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range7.Offset[i + 1, 2].Value = ksPSLAfterSummer1[i];

				// Запись случайной величины в столбец D листа 3 - КС Таксимо - Мамакан
				range7.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range7.Offset[i + 1, 3].Value = ksTMAfterSummer1[i];

				// Запись Номера схемы сети
				range7.Offset[0, 4].Value = "№ Схемы";
				range7.Offset[i + 1, 4].Value = nScheme7[i];

				// Запись шага итерации, на котором проищошло превышение 
				range7.Offset[0, 5].Value = "";
				range7.Offset[i + 1, 5].Value = listNewLoadSummerAfter1[i];

				// Запись случайной величины в столбец А листа 4 - генерация
				range8.Offset[0, 0].Value = "Генерация";
				range8.Offset[i + 1, 0].Value = RandomValue.RndValueGenSummer()[i];

				// Запись случайной величины в столбец B листа 4 - нагрузка
				range8.Offset[0, 1].Value = "Нагрузка";
				range8.Offset[i + 1, 1].Value = RandomValue.RndValueLoadSummer()[i];

				// Запись случайной величины в столбец C листа 4 - КС Пеледуй - Сухой Лог
				range8.Offset[0, 2].Value = "КС Пеледуй - Сухой Лог";
				range8.Offset[i + 1, 2].Value = ksPSLAfterSummer2[i];

				// Запись случайной величины в столбец D листа 4 - КС Таксимо - Мамакан
				range8.Offset[0, 3].Value = "КС Таксимо - Мамакан";
				range8.Offset[i + 1, 3].Value = ksTMAfterSummer2[i];

				// Запись Номера схемы сети
				range8.Offset[0, 4].Value = "№ Схемы";
				range8.Offset[i + 1, 4].Value = nScheme8[i];

				// Запись шага итерации, на котором проищошло превышение 
				range8.Offset[0, 5].Value = "";
				range8.Offset[i + 1, 5].Value = listNewLoadSummerAfter2[i];
			}

			// Останавливаем счетчик
			stopwatch.Stop();

			Console.WriteLine($"\nПроцесс завершен.\n" +
				$"Время расчета: {stopwatch.ElapsedMilliseconds} мс.\n" +
				$"Файл Excel 1 успешно сохранен по пути: {xlsxFile1}.\n" +
				$"Файл Excel 2 успешно сохранен по пути: {xlsxFile2}.\n" +
				$"Количество СВ генерации (зима): {RandomValue.RndValueGenWinter().Count}.\n" +
				$"Количество СВ генерации (лето): {RandomValue.RndValueGenSummer().Count}.\n" +
				$"Количество СВ нагрузки (зима): {RandomValue.RndValueLoadWinter().Count}.\n" +
				$"Количество СВ нагрузки (лето): {RandomValue.RndValueLoadSummer().Count}.\n" +
				$"Количество РУР (зима|до|пур): {numberYRwinter1}.\n" +
				$"Количество РУР (зима|до|смзу): {numberYRwinter2}.\n" +
				$"Количество РУР (зима|после|пур): {numberYRwinter3}.\n" +
				$"Количество РУР (зима|после|смзу): {numberYRwinter4}.\n" +
				$"Количество РУР (лето|до|пур): {numberYRsummer1}.\n" +
				$"Количество РУР (лето|до|смзу): {numberYRsummer2}.\n" +
				$"Количество РУР (лето|после|пур): {numberYRsummer3}.\n" +
				$"Количество РУР (лето|после|смзу): {numberYRsummer4}.\n");

			Console.ReadKey();

		}
	}
}
