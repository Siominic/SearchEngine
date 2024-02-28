using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc;
using Aspose.Cells;
using System.Collections;
using System.Security.Cryptography.Pkcs;
using System.Linq;

public class IndexModel : PageModel
{
    public static bool Triggered = false;
    public static List<object> itemList = new List<object>();
    public static object Message { get; set; } = "HELP ME";
    public string Word => (string)TempData[nameof(Name)];
    public static List<string> history = new List<string>(); 

    public void OnGet()
    {
    }

    public void OnPost([FromForm] string Word)
    {
        history.Add(Word);
        itemList.Clear();
        Message = "If you can read this, something is wrong";
        List<string> list = new List<string>();

        Workbook wb = new Workbook("SearchEngine.xlsx");

        // Get all worksheets
        WorksheetCollection collection = wb.Worksheets;
        for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
        {
            Worksheet worksheet = collection[worksheetIndex];
            int rows = worksheet.Cells.MaxDataRow;
            int cols = worksheet.Cells.MaxDataColumn;
            for (int i = 0; i < rows; i++)
            {
                for (int ii = 0; ii < cols + 1; ii++)
                {
                    // Print cell value
                    //Console.Write(worksheet.Cells[i, ii].Value + " | ");
                    string temporary = (worksheet.Cells[i, ii]).Value.ToString(); // it works!!
                    //Console.WriteLine(temporary);
                    list.Add(temporary);

                    //work out something...
                    //use a temporary array to loop through the data, if the array contains the word, add it.
                    
                }
                UseArray(list, Word);


            }
        }//return here
        


    }
    public static void UseArray(List<string> list, string Word)
    {
        for (int i = 0; i < list.Count; i++)
        {
            //Console.WriteLine(list[i]);
            if (list[i].Contains(Word))// connect this to an input system
            {
                Triggered = true;
            }
            
        }
        if (Triggered == true)
        {
            Algorithm(list);
            Triggered = false;
        }
        list.Clear();
    }
    public static void Algorithm(List<string> list)
    {
        string[][] resultsTable = new string[2][];
        //Console.WriteLine("IT WORKS");
        int words = Int32.Parse(list[4]);
        int links = Int32.Parse(list[5]);
        string url = (list[2]);
        double result = (0.1 * words) + (0.3 * links);

        object[] array = { url};

        //add items to a list to be ranked
        Message = url;//this returns the URL
        itemList.AddRange(array);
    }
}