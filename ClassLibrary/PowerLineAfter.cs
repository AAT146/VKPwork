using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary
{
	public class PowerLineAfter
	{
		public int[][][] schemes = new int[][][] { scheme1, scheme2, scheme3, scheme4, scheme5,
			scheme6, scheme7, scheme8, scheme9, scheme10, scheme11, scheme12, scheme13,
			scheme14, scheme15, scheme16, scheme15, scheme18, scheme19, scheme20, scheme21,
			scheme22, scheme23, scheme24, scheme25, scheme26, scheme27, scheme28, scheme29,
			scheme30, scheme31, scheme32, scheme33, scheme34, scheme35, scheme36, scheme37,
			scheme38, scheme39, scheme40, scheme41, scheme42, scheme43, scheme44, scheme45,
			scheme46, scheme47, scheme48, scheme49, scheme50, scheme51, scheme52, scheme53,};

		// Создание массивов массивов для каждой схемы
		public static int[][] scheme1 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 5004, 2102, 2 },
			new int[] { 932, 934, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme2 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 5004, 2102, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme3 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 876, 831, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme4 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 876, 831, 2 },
			new int[] { 7, 1, 1 }
		};


		public static int[][] scheme5 = new int[][]
			{
				new int[] { 2102, 892, 2 },
				new int[] { 932, 933, 1 },
				new int[] { 7, 1, 1 } 
			};

		public static int[][] scheme6 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 932, 934, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme7 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme8 = new int[][]
		{
			new int[] { 2102, 892, 1 },
			new int[] { 835, 931, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme9 = new int[][]
		{
			new int[] { 2102, 892, 1 },
			new int[] { 2841, 935, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme10 = new int[][]
		{
			new int[] { 2102, 892, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme11 = new int[][]
		{
			new int[] { 5004, 2102, 1 },
			new int[] { 876, 831, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme12 = new int[][]
		{
			new int[] { 5004, 2102, 1 },
			new int[] { 936, 938, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme13 = new int[][]
		{
			new int[] { 5004, 2102, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme14 = new int[][]
		{
			new int[] { 2102, 892, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme15 = new int[][]
		{
			new int[] { 5004, 2102, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme16 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 879, 880, 1 },
			new int[] { 877, 878, 1 },
			new int[] { 876, 877, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme17 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 879, 880, 1 },
			new int[] { 877, 878, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme18 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 879, 880, 1 },
			new int[] { 876, 877, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme19 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 876, 877, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme20 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 932, 934, 2 },
			new int[] { 939, 940, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme21 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 939, 940, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme22 = new int[][]
		{
			new int[] { 5004, 880, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme23 = new int[][]
		{
			new int[] { 878, 879, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme24 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 932, 933, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme25 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 934, 2841, 2 },
			new int[] { 2841, 935, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme26 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 934, 2841, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme27 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 2841, 935, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme28 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 937, 939, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme29 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 940, 892, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme30 = new int[][]
		{
			new int[] { 876, 877, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme31 = new int[][]
		{
			new int[] { 876, 877, 2 },
			new int[] { 933, 2841, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme32 = new int[][]
		{
			new int[] { 876, 877, 2 },
			new int[] { 934, 2841, 2 },
			new int[] { 2841, 935, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme33 = new int[][]
		{
			new int[] { 876, 877, 2 },
			new int[] { 2841, 935, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme34 = new int[][]
		{
			new int[] { 876, 877, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme35 = new int[][]
		{
			new int[] { 876, 831, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme36 = new int[][]
		{
			new int[] { 831, 832, 2 },
			new int[] { 932, 933, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme37 = new int[][]
		{
			new int[] { 831, 832, 2 },
			new int[] { 932, 934, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme38 = new int[][]
		{
			new int[] { 831, 832, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme39 = new int[][]
		{
			new int[] { 833, 834, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme40 = new int[][]
		{
			new int[] { 931, 932, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme41 = new int[][]
		{
			new int[] { 834, 836, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme42 = new int[][]
		{
			new int[] { 836, 932, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme43 = new int[][]
		{
			new int[] { 932, 933, 1 },
			new int[] { 936, 938, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme44 = new int[][]
		{
			new int[] { 932, 933, 1 },
			new int[] { 939, 940, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme45 = new int[][]
		{
			new int[] { 932, 933, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme46 = new int[][]
		{
			new int[] { 933, 2841, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme47 = new int[][]
		{
			new int[] { 932, 934, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme48 = new int[][]
		{
			new int[] { 934, 2841, 2 },
			new int[] { 2841, 935, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme49 = new int[][]
		{
			new int[] { 937, 939, 2 },
			new int[] { 939, 940, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme50 = new int[][]
		{
			new int[] { 937, 939, 2 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme51 = new int[][]
		{
			new int[] { 939, 940, 1 },
			new int[] { 940, 892, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme52 = new int[][]
		{
			new int[] { 939, 940, 1 },
			new int[] { 7, 1, 1 }
		};

		public static int[][] scheme53 = new int[][]
		{
			new int[] { 7, 1, 1 }
		};

	}
}
