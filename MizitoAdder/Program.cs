using ClosedXML.Excel;
using MizitoAdder;
using MizitoAdder.Models;
using System.IO;
using System.Windows.Forms;

ShowMessage.Welcome();
ShowMessage.Waiter(1);
ShowMessage.Clear();

#region GetExcelFilePath
var MyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
Console.WriteLine("Your Path is : " + MyPath);
Console.WriteLine("Find Source");
var ExcelFilePath = MyPath + "\\MizitoSource.xlsx";
#if DEBUG
ExcelFilePath = @"C:\Users\Asadi-PC\Desktop\Mizito-Adder\Sources\MizitoSource.xlsx";
#endif
while (File.Exists(ExcelFilePath) == false)
{
    ExcelFilePath = "";
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("File Not Find");

    var thrd = new Thread(() =>
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "فایل دیتای مشتریان";
            ofd.Filter = "Excel Files|*.xlsx;*.xls";
            ofd.InitialDirectory = MyPath;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExcelFilePath = ofd.FileName;
            }
        }
    });
    thrd.SetApartmentState(ApartmentState.STA);
    thrd.Start();
    thrd.Join();

    if (string.IsNullOrEmpty(ExcelFilePath))
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("No file selected.");
    }
}
#endregion

ShowMessage.Clear();

#region OpenFile
ShowMessage.Warning("Opening File ...");
IEnumerable<Customer> SrcCustomers = new List<Customer>();
IEnumerable<int> ErrorRow = new List<int>();
IEnumerable<int> WarningRow = new List<int>();
try
{
    using var exlfile = new XLWorkbook(ExcelFilePath);
    var exlsheet = exlfile.Worksheet(1);
    var range = exlsheet.RangeUsed();
    if (range == null)
    {
        Console.WriteLine("شیت خالی است.");
        throw new ArgumentNullException("range", "The worksheet is empty.");
    }

    foreach (var row in range.Rows())
    {
        if (row.RowNumber() == 1) continue; // Skip header row

        try
        {
            Customer customer = new Customer
            {
                RowId = row.Cell("A").GetValue<int>(),
                Name = row.Cell("B").GetValue<string>(),
                ShopName = row.Cell("C").GetValue<string>(),
                Telephone = row.Cell("D").GetValue<string>(),
                Phone = row.Cell("E").GetValue<string>(),
                Email = row.Cell("F").GetValue<string>(),
                Address = row.Cell("G").GetValue<string>(),
                Tags = row.Cell("H").GetValue<string>(),
                RepresentativeName = row.Cell("I").GetValue<string>(),
                RepresentativePhone = row.Cell("J").GetValue<string>(),
            };

            if (string.IsNullOrEmpty(customer.Phone))
            {
                ((List<int>)WarningRow).Add(row.RowNumber());
            }
            if (customer.Phone.Length != 10)
            {
                ((List<int>)WarningRow).Add(row.RowNumber());
            }

            //Validation Data
            if (string.IsNullOrEmpty(customer.Name) && string.IsNullOrEmpty(customer.ShopName))
            {
                ((List<int>)ErrorRow).Add(row.RowNumber());
                ShowMessage.Error("Error In Row " + row.RowNumber().ToString() + " : Name and ShopName are required.");
                continue;
            }
            if(string.IsNullOrEmpty(customer.Phone) && string.IsNullOrEmpty(customer.Telephone))
            {
                ((List<int>)ErrorRow).Add(row.RowNumber());
                ShowMessage.Error("Error In Row " + row.RowNumber().ToString() + " : Phone or Telephone is required.");
                continue;
            }

            //Add Data To List
            ((List<Customer>)SrcCustomers).Add(customer);
            ShowMessage.Success("ROW  " + row.Cell(1).GetValue<int>().ToString() + "  Is Added");
        }
        catch (Exception dataEx)
        {
            ((List<int>)ErrorRow).Add(row.RowNumber());
            ShowMessage.Error("Error In Row " + row.RowNumber().ToString() + " : " + dataEx.Message);
        }
    }
}
catch (Exception ex)
{
    ShowMessage.Error(ex.Message);
    Console.WriteLine("\n");
    ShowMessage.Error("Can not Read Data !\nPress Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
#endregion

#region DataResult
ShowMessage.Clear();
foreach (var error in ErrorRow)
{
    ShowMessage.Error("Error In Row " + error.ToString());
}
if (SrcCustomers.Count() == 0)
{
    ShowMessage.Error("No Valid Data Found !\nPress Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
else
{
    ShowMessage.Success("Total Valid Rows : " + SrcCustomers.Count().ToString());
    if (ErrorRow.Count() > 0)
        ShowMessage.Error("Total Error Rows : " + ErrorRow.Count().ToString());
    if(WarningRow.Count() > 0)
        ShowMessage.Warning("Total Warning Rows : " + WarningRow.Count().ToString());
}
//Ask To Continue
ShowMessage.Message("Do You Want To Continue ? (Y/N)");
string? answer = Console.ReadLine();
if (answer == null || answer == "N" || answer == "n")
{
    ShowMessage.Message("Press Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
if (answer != "Y" || answer != "y")
{
    ShowMessage.Message("Press Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
#endregion


ShowMessage.Clear();


//Exit
ShowMessage.Message("Press Any Key To Exit ...");
Console.ReadLine();
Environment.Exit(0);