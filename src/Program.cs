using System;

namespace ExportExcel{
	class Program{
		static void Main(string[] args){
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			if(args.Length>=2){
				ExcelExporter.Export(args[0],args[1]);
			}
			else{
				Console.WriteLine("example: excel inputFolder outputFolder");
			}
		}
	}
}