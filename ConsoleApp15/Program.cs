using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing.Chart;

namespace WasteManagementApp
{
    class Program
    {
        static List<UserBase> members = new List<UserBase>();
        static List<Officer> officers = new List<Officer>();
        static readonly string adminUsername = "andrew";
        static readonly string adminPassword = "nikoiIV2020";
        static readonly string excelFilePath = "waste_management_data.xlsx";

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                LoadData();
                while (true)
                {
                    Console.Clear();
                    Console.WriteLine("Waste Management System");
                    Console.WriteLine("1. Officer Login");
                    Console.WriteLine("2. Admin Login");
                    Console.WriteLine("3. Exit");
                    Console.Write("Select an option: ");

                    switch (Console.ReadLine())
                    {
                        case "1":
                            OfficerLogin();
                            break;
                        case "2":
                            AdminLogin();
                            break;
                        case "3":
                            SaveData();
                            return;
                        default:
                            Console.WriteLine("Invalid option. Please try again.");
                            PressEnterToContinue();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                PressEnterToContinue();
            }
        }

        static void AdminLogin()
        {
            Console.Clear();
            Console.Write("Admin Username: ");
            string username = Console.ReadLine();
            Console.Write("Admin Password: ");
            string password = ReadPassword();

            if (username == adminUsername && password == adminPassword)
            {
                AdminMenu();
            }
            else
            {
                Console.WriteLine("Invalid credentials.");
                PressEnterToContinue();
            }
        }

        static void AdminMenu()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Admin Menu");
                Console.WriteLine("1. View Member Materials");
                Console.WriteLine("2. Delete Member Waste");
                Console.WriteLine("3. Update Member Waste");
                Console.WriteLine("4. Add New Officer");
                Console.WriteLine("5. Remove Officer");
                Console.WriteLine("6. View Weekly Waste Logs");
                Console.WriteLine("7. Logout");
                Console.Write("Select an option: ");

                switch (Console.ReadLine())
                {
                    case "1":
                        ViewMemberMaterials();
                        break;
                    case "2":
                        DeleteMemberWaste();
                        break;
                    case "3":
                        UpdateMemberWaste();
                        break;
                    case "4":
                        AddOfficer();
                        break;
                    case "5":
                        RemoveOfficer();
                        break;
                    case "6":
                        ViewWeeklyLogs();
                        break;
                    case "7":
                        return;
                    default:
                        Console.WriteLine("Invalid option. Try again.");
                        PressEnterToContinue();
                        break;
                }
            }
        }

        static void RemoveOfficer()
        {
            Console.Clear();
            Console.WriteLine("Officers List:");
            foreach (var officer in officers)
            {
                Console.WriteLine(officer.Username);
            }

            Console.Write("\nEnter the username of the officer to remove: ");
            string username = Console.ReadLine();
            var officerToRemove = officers.FirstOrDefault(o => o.Username == username);

            if (officerToRemove != null)
            {
                officers.Remove(officerToRemove);
                Console.WriteLine("Officer removed successfully.");
                SaveData();
            }
            else
            {
                Console.WriteLine("Officer not found.");
            }

            PressEnterToContinue();
        }

        static void ViewWeeklyLogs()
        {
            Console.Clear();
            Console.WriteLine("Weekly Waste Logs\n");

            var currentWeekStart = DateTime.Now.StartOfWeek(DayOfWeek.Monday);
            var lastWeekStart = currentWeekStart.AddDays(-7);
            var lastWeekEnd = currentWeekStart.AddDays(-1);

            Console.WriteLine("THIS WEEK'S LOG:");
            DisplayLogs(currentWeekStart, DateTime.Now);

            Console.WriteLine("\nLAST WEEK'S LOG:");
            DisplayLogs(lastWeekStart, lastWeekEnd);

            PressEnterToContinue();
        }

        static void DisplayLogs(DateTime startDate, DateTime endDate)
        {
            Console.WriteLine($"{"Date",-12} {"Material Type",-20} {"Amount (kg)",-10} ");
            Console.WriteLine(new string('-', 45));

            foreach (var member in members)
            {
                foreach (var material in member.GetWasteEntries(startDate, endDate))
                {
                    Console.WriteLine($"{material.Date.ToShortDateString(),-12} {material.Type,-20} {material.Amount,-10:F2}");
                }
            }
        }

        static void ViewMemberMaterials()
        {
            Console.Clear();
            Console.WriteLine("Member Materials:");
            foreach (var member in members)
            {
                Console.WriteLine($"\nMember: {member.Username}");
                member.DisplayWasteTracking();
            }
            PressEnterToContinue();
        }

        static void DeleteMemberWaste()
        {
            Console.Clear();
            Console.Write("Enter Member Username: ");
            string username = Console.ReadLine();
            var member = members.FirstOrDefault(m => m.Username == username);

            if (member == null)
            {
                Console.WriteLine("Member not found.");
            }
            else
            {
                Console.Clear();
                Console.WriteLine($"Waste entries for {member.Username}:");
                member.DisplayWasteTracking();
                Console.Write("Enter material type to delete: ");
                string materialType = Console.ReadLine();

                member.DeleteWaste(materialType);
                Console.WriteLine("Waste entry deleted.");
                SaveData();
            }
            PressEnterToContinue();
        }

        static void UpdateMemberWaste()
        {
            Console.Clear();
            Console.Write("Enter Member Username: ");
            string username = Console.ReadLine();
            var member = members.FirstOrDefault(m => m.Username == username);

            if (member == null)
            {
                Console.WriteLine("Member not found.");
            }
            else
            {
                Console.Clear();
                Console.WriteLine($"Waste entries for {member.Username}:");
                member.DisplayWasteTracking();
                Console.Write("Enter material type to update: ");
                string materialType = Console.ReadLine();
                Console.Write("Enter new amount (kg): ");
                if (double.TryParse(Console.ReadLine(), out double newAmount))
                {
                    member.UpdateWaste(materialType, newAmount, "Admin");
                    Console.WriteLine("Waste entry updated.");
                    SaveData();
                }
                else
                {
                    Console.WriteLine("Invalid amount entered.");
                }
            }
            PressEnterToContinue();
        }

        static void AddOfficer()
        {
            Console.Clear();
            Console.Write("Enter new officer username: ");
            string username = Console.ReadLine();
            Console.Write("Enter officer password: ");
            string password = Console.ReadLine();

            if (officers.Any(o => o.Username == username))
            {
                Console.WriteLine("Officer already exists.");
            }
            else
            {
                officers.Add(new Officer(username, password));
                Console.WriteLine("Officer added successfully.");
                SaveData();
            }
            PressEnterToContinue();
        }

        static void OfficerLogin()
        {
            Console.Clear();
            Console.Write("Officer Username: ");
            string username = Console.ReadLine();
            Console.Write("Officer Password: ");
            string password = ReadPassword();

            var officer = officers.FirstOrDefault(o => o.Username == username && o.Password == password);
            if (officer != null)
            {
                OfficerMenu(officer);
            }
            else
            {
                Console.WriteLine("Invalid credentials.");
                PressEnterToContinue();
            }
        }

        static void OfficerMenu(Officer officer)
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("Officer Menu");
                Console.WriteLine("1. Register New Member");
                Console.WriteLine("2. Search Member by Username");
                Console.WriteLine("3. View Member's Recycled Logs");
                Console.WriteLine("4. Select Rewards for Member");
                Console.WriteLine("5. Log Waste");
                Console.WriteLine("6. Remove Member's Logged Materials");
                Console.WriteLine("7. Update Member's Logged Materials");
                Console.WriteLine("8. Logout");
                Console.Write("Select an option: ");

                switch (Console.ReadLine())
                {
                    case "1":
                        RegisterNewMember();
                        break;
                    case "2":
                        SearchMemberByUsername();
                        break;
                    case "3":
                        ViewMemberRecycledLogs();
                        break;
                    case "4":
                        SelectRewardsForMember();
                        break;
                    case "5":
                        LogWasteForOfficer(officer);
                        break;
                    case "6":
                        RemoveMemberLoggedMaterials(officer);
                        break;
                    case "7":
                        UpdateMemberLoggedMaterials(officer);
                        break;
                    case "8":
                        return;
                    default:
                        Console.WriteLine("Invalid option. Try again.");
                        PressEnterToContinue();
                        break;
                }
            }
        }

        static void RegisterNewMember()
        {
            Console.Clear();
            Console.Write("Enter new member username: ");
            string username = Console.ReadLine();

            if (members.Any(m => m.Username == username))
            {
                Console.WriteLine("Member already exists.");
            }
            else
            {
                members.Add(new UserBase(username, DateTime.Now));
                Console.WriteLine("Member added successfully.");
                SaveData();
            }

            PressEnterToContinue();
        }

        static void SearchMemberByUsername()
        {
            Console.Clear();
            Console.Write("Enter member username to search: ");
            string username = Console.ReadLine();

            var member = members.FirstOrDefault(m => m.Username == username);
            if (member != null)
            {
                Console.WriteLine($"Member found: {member.Username}");
                Console.WriteLine($"Creation Date: {member.CreationDate.ToShortDateString()}");
            }
            else
            {
                Console.WriteLine("Member not found.");
            }

            PressEnterToContinue();
        }

        static void ViewMemberRecycledLogs()
        {
            Console.Clear();
            Console.Write("Enter member username to view logs: ");
            string username = Console.ReadLine();

            var member = members.FirstOrDefault(m => m.Username == username);
            if (member != null)
            {
                Console.WriteLine($"Recycled logs for {member.Username}:");
                member.DisplayWasteTracking();
            }
            else
            {
                Console.WriteLine("Member not found.");
            }

            PressEnterToContinue();
        }

        static void SelectRewardsForMember()
        {
            Console.Clear();
            Console.Write("Enter member username to assign rewards: ");
            string username = Console.ReadLine();

            var member = members.FirstOrDefault(m => m.Username == username);
            if (member != null)
            {
                Console.WriteLine("Select rewards for the member:");
                Console.WriteLine("1. 5 Kilo Rice");
                Console.WriteLine("2. 10 Canned Goods");
                Console.WriteLine("3. Green T-Shirt");
                Console.Write("Select an option: ");

                string rewardChoice = Console.ReadLine();
                switch (rewardChoice)
                {
                    case "1":
                        Console.WriteLine("5 Kilo Rice reward given.");
                        break;
                    case "2":
                        Console.WriteLine("10 Canned Goods reward given.");
                        break;
                    case "3":
                        Console.WriteLine("Green T-Shirt reward given.");
                        break;
                    default:
                        Console.WriteLine("Invalid option.");
                        break;
                }
            }
            else
            {
                Console.WriteLine("Member not found.");
            }

            PressEnterToContinue();
        }

        static void RemoveMemberLoggedMaterials(Officer officer)
        {
            Console.Clear();
            Console.Write("Enter member username to remove logged materials: ");
            string username = Console.ReadLine();

            var member = members.FirstOrDefault(m => m.Username == username);
            if (member != null)
            {
                Console.WriteLine($"Logged materials for {member.Username}:");
                member.DisplayWasteTracking();

                Console.Write("Enter material type to remove: ");
                string materialType = Console.ReadLine();

                Console.Write("Enter amount to remove (kg): ");
                if (double.TryParse(Console.ReadLine(), out double amountToReduce))
                {
                    member.ReduceWaste(materialType, amountToReduce);
                    Console.WriteLine("Logged materials removed successfully.");
                    SaveData();
                }
                else
                {
                    Console.WriteLine("Invalid amount entered.");
                }
            }
            else
            {
                Console.WriteLine("Member not found.");
            }

            PressEnterToContinue();
        }

        static void UpdateMemberLoggedMaterials(Officer officer)
        {
            Console.Clear();
            Console.Write("Enter member username to update logged materials: ");
            string username = Console.ReadLine();

            var member = members.FirstOrDefault(m => m.Username == username);
            if (member != null)
            {
                Console.WriteLine($"Logged materials for {member.Username}:");
                member.DisplayWasteTracking();

                Console.Write("Enter material type to update: ");
                string materialType = Console.ReadLine();

                Console.Write("Enter new amount (kg): ");
                if (double.TryParse(Console.ReadLine(), out double newAmount))
                {
                    member.UpdateWaste(materialType, newAmount, officer.Username);
                    Console.WriteLine("Logged materials updated successfully.");
                    SaveData();
                }
                else
                {
                    Console.WriteLine("Invalid amount entered.");
                }
            }
            else
            {
                Console.WriteLine("Member not found.");
            }

            PressEnterToContinue();
        }

        static void LogWasteForOfficer(Officer officer)
        {
            Console.Clear();
            Console.Write("Enter member username to log materials for: ");
            string username = Console.ReadLine();

            var member = members.FirstOrDefault(m => m.Username == username);
            if (member != null)
            {
                try
                {
                    Console.WriteLine("Select Material Type:");
                    Console.WriteLine("1. Paper");
                    Console.WriteLine("2. Plastic");
                    Console.WriteLine("3. Glass");
                    Console.WriteLine("4. Aluminum");
                    Console.WriteLine("5. Organic");

                    Console.Write("Enter the number corresponding to the material type: ");
                    if (int.TryParse(Console.ReadLine(), out int materialChoice) && materialChoice >= 1 && materialChoice <= 5)
                    {
                        string materialType = materialChoice switch
                        {
                            1 => "Paper",
                            2 => "Plastic",
                            3 => "Glass",
                            4 => "Aluminum",
                            5 => "Organic",
                            _ => throw new InvalidOperationException("Invalid material type")
                        };

                        Console.Write("Enter amount of material (kg): ");
                        if (double.TryParse(Console.ReadLine(), out double amount))
                        {
                            member.LogWaste(new WasteEntry(materialType, amount, DateTime.Now, officer.Username));
                            Console.WriteLine("Material logged successfully.");
                            SaveData();
                        }
                        else
                        {
                            Console.WriteLine("Invalid amount entered. Please enter a valid number.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Invalid material type selected.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred while logging material: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Member not found.");
            }

            PressEnterToContinue();
        }

        static void LoadData()
        {
            if (File.Exists(excelFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var membersWorksheet = package.Workbook.Worksheets["Members"];
                    var officersWorksheet = package.Workbook.Worksheets["Officers"];
                    var loggedMaterialsWorksheet = package.Workbook.Worksheets["Logged Materials"];

                    // Load members
                    for (int row = 2; row <= membersWorksheet.Dimension.End.Row; row++)
                    {
                        string username = membersWorksheet.Cells[row, 1].Text;
                        DateTime creationDate = DateTime.Parse(membersWorksheet.Cells[row, 2].Text);
                        members.Add(new UserBase(username, creationDate));
                    }

                    // Load officers
                    for (int row = 2; row <= officersWorksheet.Dimension.End.Row; row++)
                    {
                        string username = officersWorksheet.Cells[row, 1].Text;
                        string password = officersWorksheet.Cells[row, 2].Text;
                        officers.Add(new Officer(username, password));
                    }

                    // Load logged materials
                    for (int row = 2; row <= loggedMaterialsWorksheet.Dimension.End.Row; row++)
                    {
                        string username = loggedMaterialsWorksheet.Cells[row, 1].Text;
                        string type = loggedMaterialsWorksheet.Cells[row, 2].Text;
                        double amount = double.Parse(loggedMaterialsWorksheet.Cells[row, 3].Text);
                        DateTime date = DateTime.Parse(loggedMaterialsWorksheet.Cells[row, 4].Text);
                        string loggedBy = loggedMaterialsWorksheet.Cells[row, 5].Text;
                        DateTime? updatedDate = string.IsNullOrEmpty(loggedMaterialsWorksheet.Cells[row, 6].Text) ? null : (DateTime?)DateTime.Parse(loggedMaterialsWorksheet.Cells[row, 6].Text);
                        string updatedBy = loggedMaterialsWorksheet.Cells[row, 7].Text;

                        var member = members.FirstOrDefault(m => m.Username == username);
                        if (member != null)
                        {
                            member.LogWaste(new WasteEntry(type, amount, date, loggedBy, updatedDate, updatedBy));
                        }
                    }
                }
            }
        }

        static void SaveData()
        {
            using (var package = new ExcelPackage())
            {
                // Members worksheet
                var membersWorksheet = package.Workbook.Worksheets.Add("Members");
                membersWorksheet.Cells[1, 1].Value = "Username";
                membersWorksheet.Cells[1, 2].Value = "Creation Date";
                for (int i = 0; i < members.Count; i++)
                {
                    membersWorksheet.Cells[i + 2, 1].Value = members[i].Username;
                    membersWorksheet.Cells[i + 2, 2].Value = members[i].CreationDate;
                }

                // Officers worksheet
                var officersWorksheet = package.Workbook.Worksheets.Add("Officers");
                officersWorksheet.Cells[1, 1].Value = "Username";
                officersWorksheet.Cells[1, 2].Value = "Password";
                for (int i = 0; i < officers.Count; i++)
                {
                    officersWorksheet.Cells[i + 2, 1].Value = officers[i].Username;
                    officersWorksheet.Cells[i + 2, 2].Value = officers[i].Password;
                }

                // Logged Materials worksheet
                var loggedMaterialsWorksheet = package.Workbook.Worksheets.Add("Logged Materials");
                loggedMaterialsWorksheet.Cells[1, 1].Value = "Username";
                loggedMaterialsWorksheet.Cells[1, 2].Value = "Type";
                loggedMaterialsWorksheet.Cells[1, 3].Value = "Amount";
                loggedMaterialsWorksheet.Cells[1, 4].Value = "Date";
                loggedMaterialsWorksheet.Cells[1, 5].Value = "Logged By";
                loggedMaterialsWorksheet.Cells[1, 6].Value = "Updated Date";
                loggedMaterialsWorksheet.Cells[1, 7].Value = "Updated By";

                int row = 2;
                foreach (var member in members)
                {
                    foreach (var entry in member.GetAllWasteEntries())
                    {
                        loggedMaterialsWorksheet.Cells[row, 1].Value = member.Username;
                        loggedMaterialsWorksheet.Cells[row, 2].Value = entry.Type;
                        loggedMaterialsWorksheet.Cells[row, 3].Value = entry.Amount;
                        loggedMaterialsWorksheet.Cells[row, 4].Value = entry.Date;
                        loggedMaterialsWorksheet.Cells[row, 5].Value = entry.LoggedBy;
                        loggedMaterialsWorksheet.Cells[row, 6].Value = entry.UpdatedDate;
                        loggedMaterialsWorksheet.Cells[row, 7].Value = entry.UpdatedBy;
                        row++;
                    }
                }

                // Summary worksheet
                var summaryWorksheet = package.Workbook.Worksheets.Add("Summary");
                summaryWorksheet.Cells[1, 1].Value = "Material Type";
                summaryWorksheet.Cells[1, 2].Value = "Total Amount";

                var summary = members.SelectMany(m => m.GetAllWasteEntries())
                                     .GroupBy(w => w.Type)
                                     .Select(g => new { Type = g.Key, TotalAmount = g.Sum(w => w.Amount) })
                                     .OrderByDescending(x => x.TotalAmount);

                row = 2;
                foreach (var item in summary)
                {
                    summaryWorksheet.Cells[row, 1].Value = item.Type;
                    summaryWorksheet.Cells[row, 2].Value = item.TotalAmount;
                    row++;
                }

                // Create a chart in the Summary worksheet
                var chart = summaryWorksheet.Drawings.AddChart("WasteChart", eChartType.ColumnClustered);
                chart.SetPosition(1, 0, 5, 0);
                chart.SetSize(800, 400);
                chart.Title.Text = "Total Waste by Material Type";
                chart.Series.Add(summaryWorksheet.Cells[2, 2, row - 1, 2], summaryWorksheet.Cells[2, 1, row - 1, 1]);
                chart.XAxis.Title.Text = "Material Type";
                chart.YAxis.Title.Text = "Total Amount (kg)";

                package.SaveAs(new FileInfo(excelFilePath));
            }
        }

        static void PressEnterToContinue()
        {
            Console.WriteLine("Press Enter to continue...");
            Console.ReadLine();
        }

        static string ReadPassword()
        {
            string password = "";
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    password += key.KeyChar;
                    Console.Write("*");
                }
                else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                {
                    password = password.Substring(0, password.Length - 1);
                    Console.Write("\b \b");
                }
            } while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();
            return password;
        }
    }

    interface IUser
    {
        string Username { get; }
    }

    abstract class User : IUser
    {
        public string Username { get; protected set; }

        protected User(string username)
        {
            Username = username;
        }
    }

    class Officer : User
    {
        public string Password { get; }

        public Officer(string username, string password) : base(username)
        {
            Password = password;
        }
    }

    class UserBase : User
    {
        public DateTime CreationDate { get; }
        private List<WasteEntry> WasteEntries { get; }

        public UserBase(string username, DateTime creationDate) : base(username)
        {
            CreationDate = creationDate;
            WasteEntries = new List<WasteEntry>();
        }

        public void LogWaste(WasteEntry material)
        {
            WasteEntries.Add(material);
        }

        public List<WasteEntry> GetWasteEntries(DateTime startDate, DateTime endDate)
        {
            return WasteEntries.Where(w => w.Date >= startDate && w.Date <= endDate).ToList();
        }

        public List<WasteEntry> GetAllWasteEntries()
        {
            return WasteEntries.ToList();
        }

        public void DeleteWaste(string materialType)
        {
            WasteEntries.RemoveAll(w => w.Type.Equals(materialType, StringComparison.OrdinalIgnoreCase));
        }

        public void ReduceWaste(string materialType, double amountToReduce)
        {
            var entries = WasteEntries.Where(w => w.Type.Equals(materialType, StringComparison.OrdinalIgnoreCase)).ToList();
            double remainingReduction = amountToReduce;

            foreach (var entry in entries)
            {
                if (entry.Amount <= remainingReduction)
                {
                    remainingReduction -= entry.Amount;
                    WasteEntries.Remove(entry);
                }
                else
                {
                    entry.ReduceAmount(remainingReduction);
                    break;
                }
            }
        }

        public void UpdateWaste(string materialType, double newAmount, string updatedBy)
        {
            var entry = WasteEntries.FirstOrDefault(w => w.Type.Equals(materialType, StringComparison.OrdinalIgnoreCase));
            if (entry != null)
            {
                entry.UpdateAmount(newAmount, updatedBy);
            }
        }

        public void DisplayWasteTracking()
        {
            if (WasteEntries.Any())
            {
                foreach (var entry in WasteEntries)
                {
                    Console.WriteLine($"{entry.Date.ToShortDateString()} - {entry.Type}: {entry.Amount} kg (Logged by: {entry.LoggedBy})");
                    if (entry.UpdatedDate.HasValue)
                    {
                        Console.WriteLine($"  Updated on: {entry.UpdatedDate.Value.ToString("yyyy-MM-dd HH:mm:ss")} by {entry.UpdatedBy}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No material logs.");
            }
        }
    }

    class WasteEntry
    {
        public string Type { get; }
        public double Amount { get; private set; }
        public DateTime Date { get; }
        public string LoggedBy { get; }
        public DateTime? UpdatedDate { get; private set; }
        public string UpdatedBy { get; private set; }

        public WasteEntry(string type, double amount, DateTime date, string loggedBy, DateTime? updatedDate = null, string updatedBy = null)
        {
            Type = type;
            Amount = amount;
            Date = date;
            LoggedBy = loggedBy;
            UpdatedDate = updatedDate;
            UpdatedBy = updatedBy;
        }

        public void ReduceAmount(double amountToReduce)
        {
            Amount = Math.Max(0, Amount - amountToReduce);
        }

        public void UpdateAmount(double newAmount, string updatedBy)
        {
            Amount = newAmount;
            UpdatedDate = DateTime.Now;
            UpdatedBy = updatedBy;
        }
    }

    static class DateTimeExtensions
    {
        public static DateTime StartOfWeek(this DateTime date, DayOfWeek dayOfWeek)
        {
            int diff = (7 + (date.DayOfWeek - dayOfWeek)) % 7;
            return date.AddDays(-diff).Date;
        }
    }
}

