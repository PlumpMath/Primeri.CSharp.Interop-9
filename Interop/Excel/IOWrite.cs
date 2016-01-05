using System;
using InteropExcel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel
{
	public class IOWrite
	{
		private DataStruct _data;
		private InteropExcel.Application excel;

		public IOWrite (DataStruct data)
		{
			_data = data;
		}

		public bool exportTable ()
		{
			try
			{
				//Подготовка
				excel = new InteropExcel.ApplicationClass();
				if(excel == null) return false;

				excel.Visible = false;

				InteropExcel.Workbook workbook=excel.Workbooks.Add();
				if(workbook==null) return false;

				InteropExcel.Worksheet sheet = (InteropExcel.Worksheet)workbook.Worksheets[1];
				sheet.Name="Таблица 1";

				//Попълване на таблицата

				int i=1;
				addRow (new DataRow ("Първо име", "Фамилия", "Години"), i++);

				foreach (DataRow row in _data.table)
				{
					addRow(row, i++);
				}

				//Запаметяване и затваряне
				workbook.SaveCopyAs(getPath());

				excel.DisplayAlerts=false; //Изключване на всички съобюения на Excel

				workbook.Close();
				excel.Quit();

				//Освобождаване на паметта от Excel
				if (workbook != null) Marshal.ReleaseComObject (workbook);
				if (sheet    != null) Marshal.ReleaseComObject (sheet);
				if (excel    != null) Marshal.ReleaseComObject (excel);

				workbook = null;
				sheet    = null;
				excel    = null;

				GC.Collect ();

				return true;
			}catch{
			}
			return false;

		}

		public void addRow (DataRow _datarow, int _indexRow)
		{
			try{
				
				InteropExcel.Range range;

				range=excel.Range["A" + _indexRow.ToString(), "A" + _indexRow.ToString()];
				range.Value2 = _datarow.firsName;

				range=excel.Range["B" + _indexRow.ToString(), "B" + _indexRow.ToString()];
				range.Value2 = _datarow.lastName;

				range=excel.Range["C" + _indexRow.ToString(), "C" + _indexRow.ToString()];
				range.Value2 = _datarow.age;

			}catch{

			}
		}

		public void runFile()
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

