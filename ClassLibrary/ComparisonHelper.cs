using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
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
}
