using System;

namespace Excel
{
	class MainClass
	{
		public static void Main (string[] args)
		{
			DataStruct data = new DataStruct ();
			IOWrite write = new IOWrite (data);

			//Набиране на данни в основната таблица
			data.addRow ("Мартин", "Симеонов", "33");
			data.addRow ("Геори", "Маринов", "37");

			//Проверка на таблицата
			data.printTable();

			write.exportTable ();
			write.runFile ();

		}
	}
}
