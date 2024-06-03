using MathNet.Numerics.Distributions;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
	public class RandomValue
	{
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
	}
}
