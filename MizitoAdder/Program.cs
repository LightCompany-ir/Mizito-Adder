using ClosedXML.Excel;
using MizitoAdder;
using MizitoAdder.Models;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Reflection;

ShowMessage.Welcome();
ShowMessage.Waiter(1);
ShowMessage.Clear();

#region GetExcelFilePath
var MyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
ShowMessage.Message("Your Path is : " + MyPath);
ShowMessage.Message("Find Source");
var ExcelFilePath = MyPath + "\\MizitoSource.xlsx";
#if DEBUG
ExcelFilePath = @"C:\Users\Asadi-PC\Desktop\Mizito-Adder\Sources\MizitoSource.xlsx";
#endif
while (File.Exists(ExcelFilePath) == false)
{
    ExcelFilePath = "";
    ShowMessage.Error("File Not Find");

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
        ShowMessage.Error("No file selected.");
    }
}
#endregion

#region OpenFile
ShowMessage.Clear();
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
        ShowMessage.Message("شیت خالی است.");
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
            if (string.IsNullOrEmpty(customer.Phone) && string.IsNullOrEmpty(customer.Telephone))
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
    ShowMessage.Error(ex.Message + "\n");
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
    if (WarningRow.Count() > 0)
        ShowMessage.Warning("Total Warning Rows : " + WarningRow.Count().ToString());
}
//Ask To Continue
ShowMessage.Message("Do You Want To Continue ? (Y/N)");
string? answer = Console.ReadLine();
if (string.IsNullOrEmpty(answer) || (answer != "Y" && answer != "y"))
{
    ShowMessage.Message("Press Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
#endregion

#region ImportDTO
ShowMessage.Clear();
ShowMessage.Warning("Converting Data To ImportDTO ...");
ShowMessage.Waiter();
IEnumerable<ImportDTO> SrcImport = new List<ImportDTO>();
IEnumerable<int> ImportError = new List<int>();
IEnumerable<int> ImportWarning = new List<int>();
string ImportTemplateFile = MyPath + "\\ImportTemplate.xlsx";
while (File.Exists(ImportTemplateFile) == false)
{
    ImportTemplateFile = "";
    ShowMessage.Error("File Not Find");

    var thrd = new Thread(() =>
    {
        using (var ofd = new OpenFileDialog())
        {
            ofd.Title = "فایل دیتای ورودی";
            ofd.Filter = "Excel Files|*.xlsx;*.xls";
            ofd.InitialDirectory = MyPath;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ImportTemplateFile = ofd.FileName;
            }
        }
    });
    thrd.SetApartmentState(ApartmentState.STA);
    thrd.Start();
    thrd.Join();

    if (string.IsNullOrEmpty(ImportTemplateFile))
    {
        ShowMessage.Error("No file selected.");
    }
}
try
{
    using var exlimport = new XLWorkbook(ImportTemplateFile);
    var exlsheetimp = exlimport.Worksheet(1);
    var imprange = exlsheetimp.RangeUsed();
    if (imprange == null)
    {
        ShowMessage.Message("شیت خالی است.");
        throw new ArgumentNullException("imprange", "The worksheet is empty.");
    }

    foreach (var row in imprange.Rows())
    {
        if (row.RowNumber() == 1) continue; // Skip header row

        try
        {
            ImportDTO ImportData = new ImportDTO
            {
                ImportId = row.RowNumber(),
                CustomerName = row.Cell("A").GetValue<string>(),
                ShopName = row.Cell("B").GetValue<string>(),
                Telephone = row.Cell("C").GetValue<string>(),
                Phone = row.Cell("D").GetValue<string>(),
                Address = row.Cell("E").GetValue<string>(),
                Info = row.Cell("F").GetValue<string>(),
                Website = row.Cell("G").GetValue<string>(),
                Email = row.Cell("H").GetValue<string>(),
                PostalCode = row.Cell("I").GetValue<string>(),
                FaxNumber = row.Cell("J").GetValue<string>(),
                EconomicCode = row.Cell("K").GetValue<string>(),
                NationalID = row.Cell("L").GetValue<string>(),
                Tags = row.Cell("M").GetValue<string>(),
                RepresentativeName = row.Cell("N").GetValue<string>(),
                RepresentativePosition = row.Cell("O").GetValue<string>(),
                RepresentativeEmail = row.Cell("P").GetValue<string>(),
                RepresentativePhone = row.Cell("Q").GetValue<string>(),
                RepresentativeTelephone = row.Cell("R").GetValue<string>()
            };

            if (string.IsNullOrEmpty(ImportData.Phone))
            {
                ((List<int>)ImportWarning).Add(row.RowNumber());
            }
            else
            {
                if (ImportData.Phone.StartsWith("09"))
                    ImportData.Phone = ImportData.Phone.Remove(0, 1);
            }
            if (ImportData.Phone.Length != 10)
            {
                ((List<int>)ImportWarning).Add(row.RowNumber());
            }

            //Validation Data
            if (string.IsNullOrEmpty(ImportData.CustomerName) && string.IsNullOrEmpty(ImportData.ShopName))
            {
                ((List<int>)ImportError).Add(row.RowNumber());
                ShowMessage.Error("Error In Row " + row.RowNumber().ToString() + " : Name and ShopName are required.");
                continue;
            }
            if (string.IsNullOrEmpty(ImportData.Phone) && string.IsNullOrEmpty(ImportData.Telephone))
            {
                ((List<int>)ImportError).Add(row.RowNumber());
                ShowMessage.Error("Error In Row " + row.RowNumber().ToString() + " : Phone or Telephone is required.");
                continue;
            }
            if (string.IsNullOrEmpty(ImportData.Phone))
                ImportData.Phone = ImportData.Telephone;
            //Add Data To List
            ((List<ImportDTO>)SrcImport).Add(ImportData);
            ShowMessage.Success("ROW  " + row.RowNumber().ToString() + "  Is Added");
        }
        catch (Exception dataEx)
        {
            ((List<int>)ImportError).Add(row.RowNumber());
            ShowMessage.Error("Error In Row " + row.RowNumber().ToString() + " : " + dataEx.Message);
        }
    }
}
catch (Exception ex)
{
    ShowMessage.Error(ex.Message + "\n");
    ShowMessage.Error("Can not Read Data !\nPress Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
#endregion

#region ImportResult
ShowMessage.Clear();
foreach (var error in ImportError)
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
    ShowMessage.Success("Total Valid Rows : " + SrcImport.Count().ToString());
    if (ErrorRow.Count() > 0)
        ShowMessage.Error("Total Error Rows : " + ImportError.Count().ToString());
    if (WarningRow.Count() > 0)
        ShowMessage.Warning("Total Warning Rows : " + ImportWarning.Count().ToString());
}
//Ask To Continue
answer = null;
ShowMessage.Message("Do You Want To Continue ? (Y/N)");
answer = Console.ReadLine();
if (string.IsNullOrEmpty(answer) || (answer != "Y" && answer != "y"))
{
    ShowMessage.Message("Press Any Key To Exit ...");
    Console.ReadLine();
    Environment.Exit(0);
}
#endregion

#region CheckDuplicate
ShowMessage.Clear();
ShowMessage.Warning("Checking Duplicate Data ...");
List<int> DublicateIds = new List<int>();
foreach (var item in SrcImport)
{
    ShowMessage.Message($"Check Data : {item.Phone ?? item.Telephone}");
    if (string.IsNullOrEmpty(item.Phone) == false)
    {
        if (SrcCustomers.Any(c => c.Phone?.Contains(item.Phone) ?? false))
        {
            ShowMessage.Error("Duplicate Data Found : " + item.Phone + " - " + item.Telephone);
            DublicateIds.Add(item.ImportId);
        }
        if (SrcCustomers.Any(c => c.Telephone?.Contains(item.Phone) ?? false))
        {
            ShowMessage.Error("Duplicate Data Found : " + item.Phone + " - " + item.Telephone);
            DublicateIds.Add(item.ImportId);
        }
    }
    if (string.IsNullOrEmpty(item.Telephone) == false)
    {
        if (SrcCustomers.Any(c => c.Phone?.Contains(item.Telephone) ?? false))
        {
            ShowMessage.Error("Duplicate Data Found : " + item.Phone + " - " + item.Telephone);
            DublicateIds.Add(item.ImportId);
        }
        if (SrcCustomers.Any(c => c.Telephone?.Contains(item.Telephone) ?? false))
        {
            ShowMessage.Error("Duplicate Data Found : " + item.Phone + " - " + item.Telephone);
            DublicateIds.Add(item.ImportId);
        }
    }
}
ShowMessage.Message("Press Enter To continue ...");
Console.ReadLine();
#endregion

#region CheckNumber
ShowMessage.Clear();
string? phonesearch = null;
while (phonesearch != "00")
{
    ShowMessage.Message("Enter a Phone Number To Check Duplicate Or Enter '00' to Next ... ");
    phonesearch = Console.ReadLine();
    if (string.IsNullOrEmpty(phonesearch) == false && phonesearch != "00")
    {
        if (SrcCustomers.Any(c => c.Phone?.Contains(phonesearch) ?? false))
        {
            ShowMessage.Error("Duplicate Data Found : " + phonesearch);
        }
        if (SrcCustomers.Any(c => c.Telephone?.Contains(phonesearch) ?? false))
        {
            ShowMessage.Error("Duplicate Data Found : " + phonesearch);
        }
    }
}
#endregion

#region ExportData
ShowMessage.Message("Do You Want To Export All Data? (y/n)");
answer = Console.ReadLine();
if (answer == "Y" || answer == "y")
{
    //Save File
    string TagText = string.Empty;
    string TagFile = MyPath + "\\tag.txt";
#if DEBUG
    TagFile = @"C:\Users\Asadi-PC\Desktop\Mizito-Adder\Sources\tag.txt";
#endif
    if (File.Exists(TagFile) == false)
    {
        ShowMessage.Error("Tag File Not Finded !\n");
        ShowMessage.Warning("Do You Want To Select Tag File ? (Y/N)");
        string? taganswer = Console.ReadLine();
        if (taganswer == "Y" || taganswer == "y")
        {
            var tagthrd = new Thread(() =>
            {
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Title = "فایل برچسب ها";
                    ofd.Filter = "Text Files|*.txt";
                    ofd.FileName = "tag.txt";
                    ofd.InitialDirectory = MyPath;
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        TagFile = ofd.FileName;
                    }
                }
            });
            tagthrd.SetApartmentState(ApartmentState.STA);
            tagthrd.Start();
            tagthrd.Join();

            if (string.IsNullOrEmpty(TagFile))
            {
                ShowMessage.Error("No file selected.");
            }
        }
    }
    if (File.Exists(TagFile) == true)
    {
        TagText = File.ReadAllText(TagFile);
    }
    //Check Export File
    string ExportTemplateFile = MyPath + "\\Export.xlsx";
    while (File.Exists(ExportTemplateFile) == true)
    {
        ShowMessage.Error("File Is Finded \nTry To Delete ...");
        File.Delete(ExportTemplateFile);
    }

    //Fill Export Data
    using var wb = new XLWorkbook();
    var ws = wb.Worksheets.Add("Sheet1");
    //Heders
    //typeof(ImportDTO).GetProperty("CustomerName")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("A1").Value = typeof(ImportDTO).GetProperty("CustomerName")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName ?? "Not Found";
    ws.Cell("B1").Value = typeof(ImportDTO).GetProperty("ShopName")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("C1").Value = typeof(ImportDTO).GetProperty("Telephone")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("D1").Value = typeof(ImportDTO).GetProperty("Phone")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("E1").Value = typeof(ImportDTO).GetProperty("Address")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("F1").Value = typeof(ImportDTO).GetProperty("Info")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("G1").Value = typeof(ImportDTO).GetProperty("Website")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("H1").Value = typeof(ImportDTO).GetProperty("Email")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("I1").Value = typeof(ImportDTO).GetProperty("PostalCode")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("J1").Value = typeof(ImportDTO).GetProperty("FaxNumber")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("K1").Value = typeof(ImportDTO).GetProperty("EconomicCode")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("L1").Value = typeof(ImportDTO).GetProperty("NationalID")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("M1").Value = typeof(ImportDTO).GetProperty("Tags")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("N1").Value = typeof(ImportDTO).GetProperty("RepresentativeName")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("O1").Value = typeof(ImportDTO).GetProperty("RepresentativePosition")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("P1").Value = typeof(ImportDTO).GetProperty("RepresentativeEmail")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("Q1").Value = typeof(ImportDTO).GetProperty("RepresentativePhone")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";
    ws.Cell("R1").Value = typeof(ImportDTO).GetProperty("RepresentativeTelephone")?.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName?? "Not Found";

    int rowid = 2;
    foreach (ImportDTO item in SrcImport)
    {
        if (DublicateIds.Contains(item.ImportId))
            continue;
        ws.Cell(rowid, "A").Value = item.CustomerName;
        ws.Cell(rowid, "B").Value = item.ShopName;
        ws.Cell(rowid, "C").Value = item.Telephone;
        ws.Cell(rowid, "D").Value = item.Phone;
        ws.Cell(rowid, "E").Value = item.Address;
        ws.Cell(rowid, "F").Value = item.Info + "افزوده شده توسط اپلیکیشن لایت کمپانی";
        ws.Cell(rowid, "G").Value = item.Website;
        ws.Cell(rowid, "H").Value = item.Email;
        ws.Cell(rowid, "I").Value = item.PostalCode;
        ws.Cell(rowid, "J").Value = item.FaxNumber;
        ws.Cell(rowid, "K").Value = item.EconomicCode;
        ws.Cell(rowid, "L").Value = item.NationalID;
        ws.Cell(rowid, "M").Value = TagText + "," + item.Tags;
        ws.Cell(rowid, "N").Value = item.RepresentativeName;
        ws.Cell(rowid, "O").Value = item.RepresentativePosition;
        ws.Cell(rowid, "P").Value = item.RepresentativeEmail;
        ws.Cell(rowid, "Q").Value = item.RepresentativePhone;
        ws.Cell(rowid, "R").Value = item.RepresentativeTelephone;
        rowid++;
        ShowMessage.Success("ROW  " + item.ImportId.ToString() + "  Is Exported");
    }
    wb.SaveAs(ExportTemplateFile);
    ShowMessage.Message("Data Exported To Template File ...");
    //Save File
    string NewExportTemplateFile = string.Empty;
    var thrd = new Thread(() =>
    {
        using (var sfd = new SaveFileDialog())
        {
            sfd.Title = "فایل دیتای خروجی";
            sfd.Filter = "Excel Files|*.xlsx";
            sfd.FileName = "Export.xlsx";
            sfd.InitialDirectory = MyPath;
            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                NewExportTemplateFile = sfd.FileName;
            }
        }
    });
    thrd.SetApartmentState(ApartmentState.STA);
    thrd.Start();
    thrd.Join();

    if (string.IsNullOrEmpty(NewExportTemplateFile))
    {
        ShowMessage.Error("No file selected.");
    }
    else if (ExportTemplateFile != NewExportTemplateFile)
    {
        File.Copy(ExportTemplateFile, NewExportTemplateFile);
    }
}
#endregion

//Exit
ShowMessage.Message("Press Enter To Exit ...");
Console.ReadLine();
Environment.Exit(0);