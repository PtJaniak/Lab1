using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

using System;
using System.Globalization;
using System.IO;

namespace Lab1{
public class Dane{
   
        public int Id { get; set; }
        public string Segment { get; set; }
        public string Country { get; set; }
        public string Product { get; set; }
        public string DiscountBand { get; set; }
        public decimal UnitsSold { get; set; }
        public double MnfPrice { get; set; }
        public double SalePrice { get; set; }
        public double GrossSales { get; set; }
        public double Discount { get; set; }
        public double Sales { get; set; }
        public double COGS { get; set; }
        public double Profit { get; set; }
        public DateTime Date { get; set; }

        public int MonthNumber { get => _monthNumber; }
        public string MonthName { get => _monthName; }
        public int Year { get => _year; }

        private int _year => Date.Year;
        private int _monthNumber => Date.Month;
        private string _monthName => Date.ToString("MMMM", CultureInfo.InvariantCulture);
    

}
  public class Raport
    {
        public string Segment { get; set; }
        public string Country { get; set; }
        public decimal UnitsSold { get; set; }
    }

   

public class DataContainer {
    public List<Dane> NaszeDane {get;set;} = new List<Dane>();
    private const string DataFile = @"sample-xlsx-file-for-testing.xlsx";
   
  
        
    public DataContainer(){
        
         ReadDataFromExcel();
         
        
    }
    
    public void ReadDataFromExcel()
    {
        var xls = new ExcelPackage(new System.IO.FileInfo(DataFile));
        var wrk = xls.Workbook.Worksheets.First();
        var row = 2;

        NaszeDane = new List<Dane>();
        while($"{wrk.Cells[row, 1].Value}" !="")
        {
            var dane = new Dane();
            dane.Id = row;
            dane.Segment = wrk.Cells[row, 1].Value.ToString();
            dane.Country = wrk.Cells[row, 2].Value.ToString();
            dane.Product = wrk.Cells[row, 3].Value.ToString();
            dane.DiscountBand = wrk.Cells[row, 4].Text;
            dane.UnitsSold = decimal.Parse(wrk.Cells[row, 5].Value.ToString());
            dane.MnfPrice = (double)wrk.Cells[row, 6].Value;
            dane.SalePrice = (double)wrk.Cells[row, 7].Value;
            dane.GrossSales = (double)wrk.Cells[row, 8].Value;
            dane.Discount = (double)wrk.Cells[row, 9].Value;
            dane.Sales = (double)wrk.Cells[row, 10].Value;
            dane.COGS = (double)wrk.Cells[row, 11].Value;
            dane.Profit = (double)wrk.Cells[row, 12].Value;
            dane.Date = (DateTime)wrk.Cells[row, 13].Value;
            
            
            NaszeDane.Add(dane);
            row++;
            
        }
       

        }
    public void AddDataToExcel(Dane data)
        {
            using (var package = new ExcelPackage(new FileInfo(DataFile)))
            {
                var wrk = package.Workbook.Worksheets[0];

                var row = wrk.Dimension.End.Row + 1;

                wrk.Cells[row, 1].Value = data.Segment;
                wrk.Cells[row, 2].Value = data.Country;
                wrk.Cells[row, 3].Value = data.Product;
                wrk.Cells[row, 4].Value = data.DiscountBand;
                wrk.Cells[row, 5].Value = data.UnitsSold;
                wrk.Cells[row, 6].Value = data.MnfPrice;
                wrk.Cells[row, 7].Value = data.SalePrice;
                wrk.Cells[row, 8].Value = data.GrossSales;
                wrk.Cells[row, 9].Value = data.Discount;
                wrk.Cells[row, 10].Value = data.Sales;
                wrk.Cells[row, 11].Value = data.COGS;
                wrk.Cells[row, 12].Value = data.Profit;
                wrk.Cells[row, 13].Value = data.Date;

                package.Save();
               
            }
            
            ReadDataFromExcel();
             
        }
        public bool RemoveDataFromExcel(int id)
        {
            using (var package = new ExcelPackage(new FileInfo(DataFile)))
            {
                var wrk = package.Workbook.Worksheets[0];
                var endRow = wrk.Dimension.End.Row;
                if (endRow < id) return false;
                wrk.DeleteRow(id);
                package.Save();
                
                ReadDataFromExcel();
                
                return true;
            }
            
        }
        

    }

}
