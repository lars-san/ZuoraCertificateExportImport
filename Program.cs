using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualBasic.FileIO;//.TextFieldParser;

class Program
{
    static void Main()
    {
        //Command entry
        bool exit = false;
        ConsoleKeyInfo keypress; // ready the variable used for accepting user input

        Console.WriteLine("Instructions");
        Console.WriteLine("This program will take export.csv and customer.csv in the local directory and write import.csv in the local directory, reformatting it to the import format and using the new account ID.");
        Console.WriteLine("Press \"I\" to create the import.csv file.");
        Console.WriteLine("");
        Console.WriteLine("You can press Esc to exit.");
        while (!exit) // This is the main loop and will end when the user presses the Esc key.
        {
            keypress = Console.ReadKey(true); //This is listed with true so the key press is not shown.
            if (keypress.Key == ConsoleKey.I)
            {
                ConvertExport();
            }
            if (keypress.Key == ConsoleKey.Escape) // This checks to see if the key pressed was Esc
            { exit = true; } // This variable change will end the main loop and close the program.
        }
    }

    static void ConvertExport()
    {
        int lineCount = File.ReadAllLines("export.csv").Length;
        Console.WriteLine("Processing " + lineCount + " lines.");
        string[,] data = new string[lineCount, 31];
        TextFieldParser parser = new TextFieldParser("export.csv");

        parser.HasFieldsEnclosedInQuotes = true;
        parser.SetDelimiters(",");
        string[] fields;
        int row = 0;
        int column = 0;
        while (!parser.EndOfData)
        {
            fields = parser.ReadFields();
            foreach (string field in fields)
            {
                //Console.WriteLine(field); // This can be used to see data as it loads into the data array
                //data[row, column] = field;
                switch (column)
                {
                    case 0:
                        data[row, 0] = field + "0";
                        break;
                    case 2:
                        data[row, 1] = field;
                        break;
                    case 3:
                        data[row, 2] = field;
                        break;
                    case 4:
                        data[row, 3] = field;
                        break;
                    case 5:
                        data[row, 4] = field;
                        break;
                    case 6:
                        data[row, 5] = field;
                        break;
                    case 7:
                        data[row, 6] = field;
                        break;
                    case 9:
                        data[row, 7] = field;
                        break;
                    case 10:
                        data[row, 8] = "0";
                        break;
                    case 12:
                        if (field == "1) ACCOMMODATIONS AND FOOD SERVICES")
                        { data[row, 10] = "1"; }
                        else if (field == "2) AGRICULTURAL, FORESTRY, FISHING AND HUNTING")
                        { data[row, 10] = "2"; }
                        else if (field == "3) CONSTRUCTION")
                        { data[row, 10] = "3"; }
                        else if (field == "4) FINANCE AND INSURANCE")
                        { data[row, 10] = "4"; }
                        else if (field == "5) INFORMATION, PUBLISHING AND COMMUNICATIONS")
                        { data[row, 10] = "5"; }
                        else if (field == "6) MANUFACTURING")
                        { data[row, 10] = "6"; }
                        else if (field == "7) MINING")
                        { data[row, 10] = "7"; }
                        else if (field == "8) REAL ESTATE")
                        { data[row, 10] = "8"; }
                        else if (field == "9) RENTAL AND LEASING")
                        { data[row, 10] = "9"; }
                        else if (field == "10) RETAIL TRADE")
                        { data[row, 10] = "10"; }
                        else if (field == "11) TRANSPORTATION AND WAREHOUSING")
                        { data[row, 10] = "11"; }
                        else if (field == "12) UTILITIES")
                        { data[row, 10] = "12"; }
                        else if (field == "13) WHOLESALE TRADE")
                        { data[row, 10] = "13"; }
                        else if (field == "14) BUSINESS SERVICES")
                        { data[row, 10] = "14"; }
                        else if (field == "15) PROFESSIONAL SERVICES")
                        { data[row, 10] = "15"; }
                        else if (field == "16) EDUCATIONS AND HEALTH-CARE SERVICES")
                        { data[row, 10] = "16"; }
                        else if (field == "17) NONPROFIT ORGANIZATION")
                        { data[row, 10] = "17"; }
                        else if (field == "18) GOVERNMENT")
                        { data[row, 10] = "18"; }
                        else if (field == "19) NOT A BUSINESS")
                        { data[row, 10] = "19"; }
                        else if (field == "20) OTHER")
                        { data[row, 10] = "20"; }
                        else
                        { data[row, 10] = field; }
                        break;
                    case 13:
                        data[row, 11] = field;
                        break;
                    case 14:
                        data[row, 12] = field.Substring(0, 1); // This takes only the first character
                        break;
                    case 15:
                        data[row, 13] = "\"" + field + "\"";
                        break;
                    case 16:
                        data[row, 14] = "\"" + field + "\"";
                        break;
                    case 17:
                        data[row, 15] = "\"" + field + "\"";
                        break;
                    case 18:
                        data[row, 16] = "\"" + field + "\"";
                        break;
                    case 19:
                        data[row, 17] = "\"" + field + "\"";
                        break;
                    case 20:
                        data[row, 18] = "\"" + field + "\"";
                        break;
                    case 21:
                        data[row, 19] = "\"" + field + "\"";
                        break;
                    case 22:
                        data[row, 20] = "\"" + field + "\"";
                        break;
                    case 23:
                        data[row, 21] = "\"" + field + "\"";
                        break;
                }
                column = column + 1;
            }
            column = 0;
            row = row + 1;
        }
        parser.Close();

        // Swaps out the Zuora Account Number for the Zuora Account Name using the customer.csv file data (Account Number, Account ID)
        data = ConvertAccNumToAccID(data, lineCount);

        // Writes the default import file header to row 0
        data = RedoDefaultHeader(data);

        // Writing to file
        WriteCSV("import.csv", data);
    }

    // The function above will need additional work. It was created when I was annoyed by Excel regarding a specific file.

    static string[,] ConvertAccNumToAccID(string[,] data, int records)
    {
        int lineCount = File.ReadAllLines("customer.csv").Length;
        Console.WriteLine("Processing " + lineCount + " customer records.");
        string[,] custdata = new string[lineCount, 2];
        TextFieldParser parser = new TextFieldParser("customer.csv");

        parser.HasFieldsEnclosedInQuotes = true;
        parser.SetDelimiters(",");
        string[] fields;
        int row = 0;
        int column = 0;
        while (!parser.EndOfData)
        {
            fields = parser.ReadFields();
            foreach (string field in fields)
            {
                //Console.WriteLine(field); // This can be used to see data as it loads into the data array
                //data[row, column] = field;
                switch (column)
                {
                    case 0:
                        custdata[row, 0] = field; // This is the Zuora Customer Number
                        break;
                    case 1:
                        custdata[row, 1] = field; // This is the Zuora Customer ID
                        break;
                }
                column = column + 1;
            }
            column = 0;
            row = row + 1;
        }
        parser.Close();

        Console.WriteLine("Correcting " + records + " customer codes.");
        string[,] CorrectedData = new string[records, 22];
        row = 0;
        column = 0;
        int SearchRow = 0;
        while (row < records - 1) // The " - 1" is to eliminate the blank row at the end of the Avalara export in our import
        {
            while (column <= 21)
            {
                //Console.WriteLine("Writing (" + row + "," + column +"): " + data[row, column]); // This line is for degugging
                CorrectedData[row, column] = data[row, column];
                if (column == 7)
                {
                    while (data[row, column] != custdata[SearchRow, 0] && SearchRow < lineCount - 1)
                    {
                        SearchRow = SearchRow + 1;
                    }
                    if (SearchRow != lineCount - 1)
                    {
                        Console.WriteLine("Fount it at line " + SearchRow + ".");
                        CorrectedData[row, column] = custdata[SearchRow, 1];
                    }
                    SearchRow = 0;
                }
                column = column + 1;
            }
            column = 0;
            row = row + 1;
        }

        return CorrectedData;
    }

    static string[,] RedoDefaultHeader(string[,] data)
    {
        data[0, 0] = "ECMS Cert ID"; data[0, 1] = "Issuing Country"; data[0, 2] = "Issuing Region"; data[0, 3] = "Exemption No"; data[0, 4] = "Exemption No Type"; data[0, 5] = "Eff Date"; data[0, 6] = "End Date"; data[0, 7] = "Customer Code"; data[0, 8] = "Cert Type"; data[0, 9] = "Invoice/PO No"; data[0, 10] = "Business Type"; data[0, 11] = "Business Type Desc"; data[0, 12] = "Exemption Reason"; data[0, 13] = "Exemption Reason Desc"; data[0, 14] = "Customer Name"; data[0, 15] = "Address 1"; data[0, 16] = "Address 2"; data[0, 17] = "Address 3"; data[0, 18] = "City"; data[0, 19] = "Region"; data[0, 20] = "Postal Code"; data[0, 21] = "Country";
        return data;
    }

    // Writing to file
    /*
     * 
     * Enhancement: Automatically split the files created into less than 100,000 line files.
     *
    */
    static void WriteCSV(string filename, string[,] data)
    {
        int column = 0;

        System.IO.File.WriteAllText(filename, string.Empty); // Before writing to the file, this empties the file. This way if there were previous contents with more lines than we are writing now, we will not have any of the old contents.
        try
        {
            var fs = File.Open(filename, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            var sw = new StreamWriter(fs);
            column = 0;
            foreach (string field in data)
            {
                if (column <= 20)
                {
                    sw.Write(field + ",");
                    column = column + 1;
                    //break;
                }
                else
                {
                    sw.WriteLine(field);
                    column = 0;
                    //break;
                }
            }
            sw.Flush();
            fs.Close();
        }
        catch (Exception e)
        {
            Console.WriteLine("Exception: " + e.Message);
        }
        Console.WriteLine("Done.");
    }
}

/*
 * ----------------Notes----------------
 * Import Template:
 * ProcessCode,DocCode,DocType,DocDate,CompanyCode,CustomerCode,EntityUseCode,LineNo,TaxCode,TaxDate,ItemCode,Description,Qty,Amount,Discount,Ref1,Ref2,ExemptionNo,RevAcct,DestAddress,DestCity,DestRegion,DestPostalCode,DestCountry,OrigAddress,OrigCity,OrigRegion,OrigPostalCode,OrigCountry,LocationCode,SalesPersonCode,PurchaseOrderNo,CurrencyCode,ExchangeRate,ExchangeRateEffDate,PaymentDate,TaxIncluded,DestTaxRegion,OrigTaxRegion,Taxable,TaxType,TotalTax,CountryName,CountryCode,CountryRate,CountryTax,StateName,StateCode,StateRate,StateTax,CountyName,CountyCode,CountyRate,CountyTax,CityName,CityCode,CityRate,CityTax,Other1Name,Other1Code,Other1Rate,Other1Tax,Other2Name,Other2Code,Other2Rate,Other2Tax,Other3Name,Other3Code,Other3Rate,Other3Tax,Other4Name,Other4Code,Other4Rate,Other4Tax,ReferenceCode,BuyersVATNo
 * 76 fields
*/