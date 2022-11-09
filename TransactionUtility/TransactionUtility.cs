using System;
using System.IO;
using Microsoft.Extensions.Configuration;
using IronXL;


namespace TransactionUtility;

public class TransactionUtility
{
    public static void Main()
    {
        var athToCSV = new ConfigurationBuilder().AddJsonFile("appSettings.json").Build().GetSection("AppSettings")["pathToCSV"];
        var pathToCSV = Path.Combine(Environment.CurrentDirectory, "Checking1.csv");
        var lineItems = new List<List<string>>();
        try
        {
            if (File.Exists(pathToCSV))
            {
                using (StreamReader sr = new StreamReader(File.OpenRead(pathToCSV)))
                {
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        var values = line.Split(',').ToList();
                        lineItems.Add(values);
                    }
                }
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }

        var positiveValues = new List<List<string>>();
        var negativeValues = new List<List<string>>();
        var recurringPayments = new List<List<string>>();

        foreach (var transaction in lineItems)
        {
            if (Int32.Parse(transaction[2]) < 0)
            {
                negativeValues.Add(transaction);
            }

            if (Int32.Parse(transaction[2]) > 0)
            {
                positiveValues.Add(transaction);
            }

            if (transaction[4].Contains("RECURRING"))
            {
                recurringPayments.Add(transaction);
            }
        }
        
        WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);

        var sheet = workBook.CreateWorkSheet("Organized Transactions");

        int rowTracker = 1;
        int positiveTotal = 0;
        int negativeTotal = 0;
        foreach (var transaction in lineItems)
        {
            if (Int32.Parse(transaction[2]) < 0)
            {
                //D = amount, E = date, F = description
                sheet["F" + rowTracker].Value = transaction[4];
                sheet["E" + rowTracker].Value = transaction[0];
                sheet["D" + rowTracker].Value = transaction[1];
                sheet["D" + rowTracker].Style.SetBackgroundColor(IronSoftware.Drawing.Color.Green);
                positiveTotal += Int32.Parse(transaction[1]);
            }

            if (Int32.Parse(transaction[2]) > 0)
            {
                //A = description, B = date, C = amount
                sheet["A" + rowTracker].Value = transaction[4];
                sheet["B" + rowTracker].Value = transaction[0];
                sheet["C" + rowTracker].Value = transaction[1];
                sheet["C" + rowTracker].Style.SetBackgroundColor(IronSoftware.Drawing.Color.Red);
                negativeTotal += Int32.Parse(transaction[1]);
            }

            rowTracker++;
        }

        sheet["A" + rowTracker].Value = "Totals:";
        sheet["C" + rowTracker].Value = negativeTotal.ToString();
        sheet["D" + rowTracker].Value = positiveTotal.ToString();
        sheet["C" + rowTracker].Style.SetBackgroundColor(IronSoftware.Drawing.Color.Red);
        sheet["D" + rowTracker].Style.SetBackgroundColor(IronSoftware.Drawing.Color.Green);

        workBook.SaveAs("Transactions.xlsx");
    }
}