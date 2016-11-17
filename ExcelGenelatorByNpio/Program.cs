using System;
using System.IO;
using System.Reflection;
using System.Linq;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
//using NPOI.SS.UserModel;

namespace ExcelGenelator
{
	class MainClass
	{
		public static void Main(string[] args)
		{
			args = new String[] { "input.csv", "output.xlsx" };

			// アプリケーションのフルパス
			var appPath = Assembly.GetExecutingAssembly().Location;
			if (args.Length != 2)
			{
				Console.WriteLine($"use: mono {Path.GetFileName(appPath)} input.csv output.xlsx");
				return;
			}
			var appDirectory = Path.GetDirectoryName(appPath);

			// テンプレートExcel
			var templateExcelName = Path.Combine(appDirectory,"template.xlsx");
			if (!File.Exists(templateExcelName))
			{
				Console.WriteLine($"ERROR {templateExcelName} not Found.");
				return;
			}

			// 入力CSV
			var inputCsvName = Path.Combine(appDirectory, args[0]);
			if (!File.Exists(inputCsvName))
			{
				Console.WriteLine($"ERROR {inputCsvName} not Found.");
				return;
			}

			// 出力Excel
			var outputExcelName = Path.Combine(appDirectory, args[1]);
			using (FileStream file = new FileStream(templateExcelName, FileMode.Open, FileAccess.ReadWrite))
			{
				var wb = new XSSFWorkbook(file);
				ISheet sheet = wb.GetSheetAt(0);
				var lines = File.ReadAllLines(inputCsvName);
				foreach (var item in lines.Select((line, row) => new { line, row }))
				{
					// データを差し込むのは7行目以降
					var row = sheet.GetRow(item.row + 7);
					var values = item.line.Split(',');
					foreach (int i in Enumerable.Range(0, 3))
					{
						// データを差し込むカラムは3個目以降
						var cell = row.GetCell(i + 3);
						if (i == 0) // 品名は、文字として挿入
						{
							cell.SetCellValue(values[i]);
						}
						else
						{ // 数量・単価は、数値として挿入
							cell.SetCellValue(Int32.Parse(values[i]));
						}
					}
				}
				XSSFFormulaEvaluator.EvaluateAllFormulaCells(wb); // これで、数式を再計算します
				using (var fs = new FileStream(outputExcelName, FileMode.CreateNew))
				{
					wb.Write(fs);
				}
			}
		}
	}
}

