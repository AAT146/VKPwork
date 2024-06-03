using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
	public class Excel
	{
		/// <summary>
		/// Метод: чтение файла формата Excel.
		/// </summary>
		/// <param name="filePath">Файл Excel.</param>
		/// <returns>Массив данных.</returns>
		public static List<double> ReadFileFromExcel(string filePath)
		{
			// Установка контекста лицензирования
			ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

			using (var package = new ExcelPackage(new FileInfo(filePath)))
			{
				var worksheet = package.Workbook.Worksheets[0];
				List<double> data = new List<double>(worksheet.Dimension.Rows);

				for (int i = 2; i <= worksheet.Dimension.Rows; i++)
				{
					data.Add(worksheet.Cells[i, 1].GetValue<double>());
				}

				return data;
			}
		}

		// КС Таксимо - Мамакан (После)
		public static string mdpTMpurAfterSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPpur_TM_After_Summer.xlsx";
		public static string mdpTMpurAfterWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPpur_TM_After_Winter.xlsx";
		public static string mdpTMsmzyAfterSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPsmzy_TM_After_Summer.xlsx";
		public static string mdpTMsmzyAfterWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPsmzy_TM_After_Winter.xlsx";

		public static List<double> mdpTMpurAfterSummer = ReadFileFromExcel(mdpTMpurAfterSummerxlsx);
		public static List<double> mdpTMpurAfterWinter = ReadFileFromExcel(mdpTMpurAfterWinterxlsx);
		public static List<double> mdpTMsmzyAfterSummer = ReadFileFromExcel(mdpTMsmzyAfterSummerxlsx);
		public static List<double> mdpTMsmzyAfterWinter = ReadFileFromExcel(mdpTMsmzyAfterWinterxlsx);

		// КС Таксимо - Мамакан (До)
		public static string mdpTMpurBeforeSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPpur_TM_Before_Summer.xlsx";
		public static string mdpTMpurBeforeWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPpur_TM_Before_Winter.xlsx";
		public static string mdpTMsmzyBeforeSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPsmzy_TM_Before_Summer.xlsx";
		public static string mdpTMsmzyBeforeWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\TM\MDPsmzy_TM_Before_Winter.xlsx";

		public static List<double> mdpTMpurBeforeSummer = ReadFileFromExcel(mdpTMpurBeforeSummerxlsx);
		public static List<double> mdpTMpurBeforeWinter = ReadFileFromExcel(mdpTMpurBeforeWinterxlsx);
		public static List<double> mdpTMsmzyBeforeSummer = ReadFileFromExcel(mdpTMsmzyBeforeSummerxlsx);
		public static List<double> mdpTMsmzyBeforeWinter = ReadFileFromExcel(mdpTMsmzyBeforeWinterxlsx);

		// КС Пеледуй - Сухой Лог (После)
		public static string mdpPSLpurAfterSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPpur_PLS_After_Summer.xlsx";
		public static string mdpPSLpurAfterWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPpur_PLS_After_Winter.xlsx";
		public static string mdpPSLsmzyAfterSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPsmzy_PLS_After_Summer.xlsx";
		public static string mdpPSLsmzyAfterWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPsmzy_PLS_After_Winter.xlsx";

		public static List<double> mdpPSLpurAfterSummer = ReadFileFromExcel(mdpTMpurAfterSummerxlsx);
		public static List<double> mdpPSLpurAfterWinter = ReadFileFromExcel(mdpTMpurAfterWinterxlsx);
		public static List<double> mdpPSLsmzyAfterSummer = ReadFileFromExcel(mdpTMsmzyAfterSummerxlsx);
		public static List<double> mdpPSLsmzyAfterWinter = ReadFileFromExcel(mdpTMsmzyAfterWinterxlsx);

		// КС Пеледуй - Сухой Лог (До)
		public static string mdpPSLpurBeforeSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPpur_PLS_Before_Summer.xlsx";
		public static string mdpPSLpurBeforeWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPpur_PLS_Before_Winter.xlsx";
		public static string mdpPSLsmzyBeforeSummerxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPsmzy_PLS_Before_Summer.xlsx";
		public static string mdpPSLsmzyBeforeWinterxlsx = @"C:\Users\aat146\Desktop\Чтение МДП\PLS\MDPsmzy_PLS_Before_Winter.xlsx";

		public static List<double> mdpPSLpurBeforeSummer = ReadFileFromExcel(mdpTMpurBeforeSummerxlsx);
		public static List<double> mdpPSLpurBeforeWinter = ReadFileFromExcel(mdpTMpurBeforeWinterxlsx);
		public static List<double> mdpPSLsmzyBeforeSummer = ReadFileFromExcel(mdpTMsmzyBeforeSummerxlsx);
		public static List<double> mdpPSLsmzyBeforeWinter = ReadFileFromExcel(mdpTMsmzyBeforeWinterxlsx);
	}
}
