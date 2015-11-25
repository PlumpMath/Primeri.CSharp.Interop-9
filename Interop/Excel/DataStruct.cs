using System;
using System.Collections.Generic;

namespace Excel
{
	public class DataStruct
	{
		public List<DataRow> table = new List<DataRow>();

		public DataStruct ()
		{
		}

		public void addRow (string _fName, String _lName, string _age)
		{
			table.Add (new DataRow (_fName, _lName, _age));
		}

		public void printTable()
		{
			try
			{
				foreach(DataRow row in table)
				{
					Console.WriteLine(row.firsName+" "+row.lastName+" "+row.age);
				}
			}catch{
			}
		}

	}

	public class DataRow
	{
		private string _firsName = "";
		private string _lastName = "";
		private string _age = "";

		public DataRow(string __firsName, string __lastName, string __age)
		{
			_firsName = __firsName;
			_lastName = __lastName;
			_age = __age;

		}

		public string firsName
		{
			set{ _firsName = value; }
			get{ return _firsName; }
		}

		public string lastName
		{
			set{ _lastName = value; }
			get{ return _lastName; }
		}

		public string age
		{
			set{ _age = value; }
			get{ return _age; }
		}
	}
}

