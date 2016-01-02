using System;
using InteropExcel=Microsoft.Office.Interop.Excel;

namespace Excel
{
	public class IOWrite
	{
		private DataStruct _data;
		private InteropExcel.Application excel;

		public IOWrite (DataStruct data)
		{
		}

		private bool exportTable ()
		{
			try
			{
				//Междинни проверки

				return true;
			}catch{
			}
			return false;

		}

		private void addRow (DataRow _row)
		{
			try{
				

			}catch{

			}
		}

		private void runFile()
		{
			try{

			}catch{
			}
		}

		private string getPath()
		{
			return System.IO.Path.Combine (AppDomain.CurrentDomain.BaseDirectory, "Table1.xlsx");
		}
	}
}

