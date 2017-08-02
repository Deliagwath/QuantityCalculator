using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuantityCalculator
{
    public partial class Form1 : Form
    {
        private DB theMallDB;
        private SortedDictionary<string, string> Files;
        private List<Data> CurrData;
        
        public Form1()
        {
            InitializeComponent();
            Files = new SortedDictionary<string, string>();
            this.theMallDB = new DB();
            loadDB();
            loadComboBox();
        }

        private void getData(string filename)
        {
            List<Data> fileContents = new List<Data>();
            Data fragment, duplicate;
            string[] args = Files[filename].Split(',');
            // Files[filename] :
            // > #,#        -> Denoting LID and BID for DB
            // > Directory  -> Pointing to the Excel File
            
            if (args.Length > 1)
            {
                // Grab DB Contents
                fileContents = this.theMallDB.getData(args[0], args[1], checkBox2.Checked);
            }
            else
            {
                // Grab Excel Contents
                // Initialize COM Components
                Excel.Application xlApp;
                Excel.Workbook xlWorkbook;
                Excel._Worksheet xlWorksheet = null;
                Excel.Range xlRange = null;
                int rowCount, colCount;

                // Open Excel File and grab contents
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(Files[filename], ReadOnly: true);

                List<Data>[] sheetContents = new List<Data>[xlWorkbook.Sheets.Count];
                for (int i = 0; i < sheetContents.Length; i++) { sheetContents[i] = new List<Data>(); }

                // Grab data from every sheet
                for (int i = 1; i <= xlWorkbook.Sheets.Count; i++)
                {
                    xlWorksheet = xlWorkbook.Sheets[i];
                    xlRange = xlWorksheet.UsedRange;

                    rowCount = xlRange.Rows.Count;
                    colCount = xlRange.Columns.Count;

                    // Iterate through the Excel Sheet
                    for (int row = 7; row < rowCount; row++)
                    {
                        // Reached Limit, break
                        if (xlRange.Cells[row, 2].Value2 == null) { break; }

                        // Create new Data fragment
                        fragment = new Data(xlRange.Cells[row, 2].Value2, normaliseValue(xlRange.Cells[row, 3].Value2));

                        // Rename Data fragment's name if there was a duplicate
                        duplicate = sheetContents[i - 1].FirstOrDefault(record => record.ID == fragment.ID);
                        if (duplicate != null) { fragment.ID += fragment.Value; }

                        // Add stock values to fragment
                        fragment.add((int)normaliseValue(xlRange.Cells[row, 4].Value2));
                        fragment.add((int)normaliseValue(xlRange.Cells[row, 7].Value2));
                        fragment.add((int)normaliseValue(xlRange.Cells[row, 10].Value2));
                        fragment.add((int)normaliseValue(xlRange.Cells[row, 13].Value2));

                        // Add finished fragment as Data
                        sheetContents[i - 1].Add(fragment);
                    }
                }

                // Merge Data from multiple sheets
                List<Data> merger = new List<Data>();

                // Merge Data Counts
                sheetContents.ToList().ForEach(sheet =>
                {
                    sheet.ForEach(atom =>
                    {
                        if (!merger.Any(merged => merged.ID == atom.ID)) { merger.Add(new Data(atom)); }
                        merger.First(merged => merged.ID == atom.ID).add(atom.Count);
                    });
                });
                fileContents = merger;
                
                // Filter 0 Stock on CheckBox condition
                if (!checkBox2.Checked) { fileContents = fileContents.Where(e => e.Count > 0).ToList(); }
                else { fileContents.Where(e => e.Count == 0).ToList().ForEach(e => e.add(1)); }

                // Cleanup COM Objects
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                xlWorkbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(xlWorkbook);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            // Assign retrieved data to CurrData
            // CurrData will then be later used to get the ProductList or to Calculate
            CurrData = fileContents;
        }

        // Functions to handle dynamic datatypes from Excel
        private decimal normaliseValue(string input)
        {
            if (input == null || input == "") { return 0; }
            return Decimal.Parse(input);
        }
        private decimal normaliseValue(double input)
        {
            return (decimal)input;
        }

        private void loadComboBox()
        {
            // Get Filenames from Excel
            string fileDirectory = @"\\BRSPNAS\Shared\ห้าง\Max valu\2560_เริ่ม\2560_รายงานการส่งของ แต่ละสาขา";
            string[] files = Directory.GetFiles(fileDirectory);
            string filename;
            files = files.Where(file => file.Contains("สาขาที่")).ToArray();      // Pick only relevant files
            files = files.Where(file => !file.Contains("~$")).ToArray();        // Ignore Opened
            files = files.Where(file => !file.Contains("Copy of")).ToArray();   // Ignore Duplicates
            foreach (string filePath in files)
            {
                filename = Path.GetFileNameWithoutExtension(filePath);
                Files.Add(filename, filePath);
            }
            this.theMallDB.tbl_location.ForEach(l =>                            // Get Filenames from DB -- Start
            {
                List<DB.Branch> branches = this.theMallDB.tbl_branch
                    .Where(b => b.lid == l.lid)
                    .ToList();
                if (branches.Count > 1) { branches.ForEach(b => Files.Add(String.Format("{0} {1}", l.lname, b.bname), String.Format("{0},{1}", l.lid, b.bid))); }
                else { Files.Add(l.lname, String.Format("{0},{1}", l.lid, 0)); }
            });                                                                 // Get Filenames from DB -- End

            comboBoxDestinations.DataSource = new BindingSource(Files, null);   // Databind to UI
            comboBoxDestinations.DisplayMember = "Key";
            comboBoxDestinations.ValueMember = "Value";
        }
        
        private void loadDB()
        {
            string cwd = Directory.GetCurrentDirectory();
            string theMallPath = Path.Combine(Path.Combine(cwd, "Data"), "brsp.sql");
            string[] theMallData = File.ReadAllLines(theMallPath);

            // A list of positions for each table -- Hardcoded to deliver, should be match to keywords instead
            // List<int> positions = new List<int>();
            Dictionary<string, int[]> tables = new Dictionary<string, int[]>();
            tables.Add("po_master",     new int[] { 1, 2, 3, 4, 5 });
            tables.Add("tbl_location",  new int[] { 0, 1 });
            tables.Add("tbl_branch",    new int[] { 0, 1, 2 });
            tables.Add("po_detail",     new int[] { 0, 1, 2, 3, 5 });
            tables.Add("prod_set",      new int[] { 0, 1, 6, 9 });
            string current = "";
            bool validData = false;
            foreach (string row in theMallData)
            {
                Match insert = Regex.Match(row, @"INSERT INTO `([\w]+)`");
                if (insert.Success) // New Insert Block
                {
                    // Get Insert ID
                    current = insert.Groups[1].Value;
                    // Switch States
                    if (tables.Keys.Contains(current)) { validData = true; }
                    // Ignore Irrelevant Data
                    else { validData = false; }
                    continue;
                }
                // If the row is matched with a certain table, pass the information to the parser
                else if (validData) { parseSQLRow(current, row, tables[current]); }
            }
            this.theMallDB.cleanDB();
        }

        private void parseSQLRow(string table, string sqlString, int[] positions)
        {
            // Cleaning the SQL String
            sqlString = sqlString.Trim();
            if (sqlString.Length == 0) { return; }
            // Remove ()
            sqlString = sqlString.Substring(1, sqlString.Length - 2);

            // Split with StringSplitOptions.None to include empty values, so the indexing doesn't get messed up
            string[] values = sqlString.Split(new string[] { ", " }, StringSplitOptions.None);

            // Trim and remove wrapping '' from each value
            values = values.Select(v =>
            {
                v = v.Trim();
                v = v.Trim('\'');
                return v;
            }).ToArray();

            // If there is an invalid number of arguments
            if (values.Length < positions.Length) { return; }

            // Match using a switch, then add to the fake DB
            switch (table)
            {
                case "po_master":
                    this.theMallDB.po_master.Add(new DB.POMaster(values[positions[0]], values[positions[1]], values[positions[2]], values[positions[3]], values[positions[4]]));
                    break;
                case "tbl_location":
                    this.theMallDB.tbl_location.Add(new DB.Location(values[positions[0]], values[positions[1]]));
                    break;
                case "tbl_branch":
                    this.theMallDB.tbl_branch.Add(new DB.Branch(values[positions[0]], values[positions[1]], values[positions[2]]));
                    break;
                case "po_detail":
                    this.theMallDB.po_detail.Add(new DB.PODetail(values[positions[0]], values[positions[1]], values[positions[2]], values[positions[3]], values[positions[4]]));
                    break;
                case "prod_set":
                    this.theMallDB.prod_set.Add(new DB.ProdSet(values[positions[0]], values[positions[1]], values[positions[2]], values[positions[3]]));
                    break;
            }
        }

        private void textBoxMoney_TextChanged(object sender, EventArgs e)
        {
            // Reject illegal characters
            // Doesn't stop two decimals however, but that's up to the user
            // Also doesn't break the program, since all non-digits are stripped
            textBoxMoney.Text = Regex.Replace(textBoxMoney.Text.Trim(), "[^0-9,.฿]", "");
        }

        private class Data
        {
            public string ID;
            public decimal Value;
            public int Count;
            public decimal Total;
            public int Ratio;
            public bool Enabled;
            public List<int> Factors;
            public Data(Data atom)
            {
                this.ID = atom.ID;
                this.Value = atom.Value;
                this.Count = atom.Count;
                this.Total = this.Count * this.Value;
                this.Ratio = atom.Ratio;
                this.Enabled = atom.Enabled;
            }
            public Data(string dataString)
            {
                string[] splitted = dataString.Split(',');
                this.ID = splitted[0];
                this.Value = decimal.Parse(splitted[1]);
                this.Count = 0;
                this.Total = 0;
                this.Enabled = false;
            }
            public Data(string ID, double Value)
            {
                this.ID = ID;
                this.Value = Convert.ToDecimal(Value);
                this.Count = 0;
                this.Total = 0;
                this.Enabled = false;
            }
            public Data(string ID, decimal Value)
            {
                this.ID = ID;
                this.Value = Value;
                this.Count = 0;
                this.Total = 0;
                this.Enabled = false;
            }
            public void add(int value)
            {
                this.Count += value;
                this.Total = this.Count * this.Value;
            }
            public void sub(int value)
            {
                this.Count -= value;
                if (this.Count < 0) { this.Count = 0; }
                this.Total = this.Count * this.Value;
            }
            public void getFactors()
            {
                List<int> factors = new List<int>();
                int num = this.Count;
                while (num > 1)
                {
                    for (int i = 2; i <= this.Count; i++)
                    {
                        if (num % i == 0)
                        {
                            factors.Add(i);
                            num /= i;
                            break;
                        }
                    }
                }
                this.Factors = factors;
            }
        }
        private class DB
        {
            public List<POMaster> po_master;
            public List<Location> tbl_location;
            public List<Branch> tbl_branch;
            public List<PODetail> po_detail;
            public List<ProdSet> prod_set;

            public DB()
            {
                this.po_master = new List<POMaster>();
                this.tbl_location = new List<Location>();
                this.tbl_branch = new List<Branch>();
                this.po_detail = new List<PODetail>();
                this.prod_set = new List<ProdSet>();
            }

            public List<Data> getData(string lid, string bid, bool zero)
            {
                // Filters based on location and branch
                List<POMaster> filtered_po_master = this.po_master.Where(l => l.lid == lid).ToList();
                if (int.Parse(bid) > 0) { filtered_po_master = this.po_master.Where(b => b.bid == bid).ToList(); }
                Dictionary<string, Data> filtered = new Dictionary<string, Data>();

                // Matching Records between po_master and po_detail
                List<PODetail> filtered_po_detail = this.po_detail.Where(p => 
                    filtered_po_master.Any(fpm =>
                        fpm.yy == p.yy &&
                        fpm.mm == p.mm &&
                        fpm.no == p.no)
                ).ToList();

                // Conversion of ProdSet to Data
                foreach (PODetail detail in filtered_po_detail)
                {
                    if (!filtered.ContainsKey(detail.setid))
                    {
                        ProdSet ps = this.prod_set.FirstOrDefault(p => p.id == detail.setid);
                        if (ps == null) { continue; }
                        filtered.Add(detail.setid, new Data(ps.name, decimal.Parse(ps.sprice)));
                    }
                    filtered[detail.setid].add(int.Parse(detail.setnum));
                }

                return zero ? filtered.Values.Where(atom => atom.Count > 0).ToList() : filtered.Values.ToList();
            }

            public void cleanDB()
            {
                // Filtering by Location
                string[] validLocationIDs = new string[] { "2", "3", "4", "5" };
                this.po_master = this.po_master.Where(p => validLocationIDs.Contains(p.lid)).ToList();
                this.tbl_location = this.tbl_location.Where(l => validLocationIDs.Contains(l.lid)).ToList();
                this.tbl_branch = this.tbl_branch.Where(b => validLocationIDs.Contains(b.lid)).ToList();
                this.tbl_branch = this.tbl_branch.Where(b => b.bname.Length > 0).ToList();

                // Get Relevant Records
                this.po_detail = this.po_detail.Where(p =>
                    this.po_master.Any(fpm =>
                        fpm.yy == p.yy &&
                        fpm.mm == p.mm &&
                        fpm.no == p.no)
                ).ToList();

                // Filter out old records
                this.po_detail = this.po_detail.Where(p => int.Parse(p.yy) >= 60).ToList();
                List<string> itemIDs = this.po_detail.Select(p => p.setid).Distinct().ToList();
                this.prod_set = this.prod_set.Where(p => itemIDs.Contains(p.id)).ToList();
            }

            public class POMaster
            {
                public string yy, mm, no, lid, bid;
                public POMaster(string yy, string mm, string no, string lid, string bid)
                {
                    this.yy = yy;
                    this.mm = mm;
                    this.no = no;
                    this.lid = lid;
                    this.bid = bid;
                }
            }
            public class Location
            {
                public string lid, lname;
                public Location(string lid, string lname)
                {
                    this.lid = lid;
                    this.lname = lname;
                }
            }
            public class Branch
            {
                public string bid, lid, bname;
                public Branch(string bid, string lid, string bname)
                {
                    this.bid = bid;
                    this.lid = lid;
                    this.bname = bname;
                }
            }
            public class PODetail
            {
                public string yy, mm, no, setid, setnum;
                public PODetail(string yy, string mm, string no, string setid, string setnum)
                {
                    this.yy = yy;
                    this.mm = mm;
                    this.no = no;
                    this.setid = setid;
                    this.setnum = setnum;
                }
            }
            public class ProdSet
            {
                public string id, name, sprice;
                public bool enabled;
                public ProdSet(string id, string name, string sprice, string enabled)
                {
                    this.id = id;
                    this.name = name;
                    this.sprice = sprice;
                    this.enabled = enabled == "0" ? false : true;
                }
            }
        }

        private void getProducts(object sender, EventArgs e)
        {
            // Grabs the relevant data and sets it to CurrData for further operation
            getData(((KeyValuePair<string, string>)comboBoxDestinations.SelectedItem).Key);

            // Getting and formatting the input value
            decimal value = decimal.Parse(Regex.Replace(textBoxMoney.Text, "[^0-9.]", ""));
            textBoxMoney.Text = toCurrency(value);

            // Create/Clear results table
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            DataTable table = new DataTable();
            table.Rows.Clear();
            table.Columns.Clear();

            // Setting Table Columns
            List<string> columns = new List<string> { "ชื่อ", "ราคา", "จำนวน" };
            columns.ForEach(col => table.Columns.Add(col));
            table.Columns.Add("เปิด/ปิด", typeof(bool));

            // Setting Table Rows
            CurrData.ForEach(item => table.Rows.Add(item.ID, toCurrency(item.Value), item.Count, item.Enabled));

            // Databinding
            dataGridView1.DataSource = table;
            dataGridView1.ReadOnly = false;
            button1.Enabled = true;
        }

        private void generate(object sender, EventArgs e)
        {
            // Getting and formatting the input value
            decimal value = decimal.Parse(Regex.Replace(textBoxMoney.Text, "[^0-9.]", ""));
            if (value == 0M) { return; }
            textBoxMoney.Text = toCurrency(value);

            // Taking current data to local variable to work on
            List<Data> chosen = CurrData.Select(item => new Data(item)).ToList();

            // Filtering by chosen user inputs
            List<string> chosenIDs = new List<string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if ((bool)row.Cells["เปิด/ปิด"].Value) { chosenIDs.Add(row.Cells["ชื่อ"].Value.ToString()); }
            }

            // Filtering out irrelevant data
            chosen = chosen.Where(item => chosenIDs.Contains(item.ID)).ToList();

            // Calculate Ratios
            chosen.ForEach(atom => atom.getFactors());
            List<int> commonFactors = getCommonFactors(chosen);
            setRatios(chosen, commonFactors);

            decimal difference = value - getTotal(chosen);
            bool zeroed = false, found = false;

            // Recursive Calculations
            while (difference != 0)
            {
                // Recursion works by working towards EachItem.Count
                // Each Recursive Step reduces the starting point by EachItem.Ratio
                // Once the starting point for all items == 0, it looks for all possible combinations
                // Zeroed is then toggled for short circuit

                // Short Circuit -- No Combinations
                if (zeroed)
                {
                    outputData(null);
                    return;
                }
                zeroed = chosen.All(item => item.Count <= 0);

                // Recursive Call
                found = toLimit(chosen, difference);

                // Found possible combination short circuit
                if (found) { break; }

                // Recursive Step -- Increase Recursive Depth
                chosen.ForEach(item => item.sub(item.Ratio));
                difference = value - getTotal(chosen);
            }

            // Displaying Results
            outputData(chosen);
        }

        // Helper Functions for Ratio Calculations
        private void setRatios(List<Data> data, List<int> commonFactors)
        {
            data.ForEach(atom =>
            {
                List<int> atomFactors = atom.Factors;

                // Removes the common prime factors to reduce recursion step size
                commonFactors.ForEach(factor => atomFactors.Remove(factor));

                // Re-set item Ratio to the new calculated Ratio
                atom.Ratio = atomFactors.Aggregate(1, (acc, factor) => acc * factor);
            });
        }
        private List<int> getCommonFactors(List<Data> data)
        {
            List<List<int>> factors = data.Select(atom => atom.Factors).ToList();
            List<int> seed = factors[0];
            factors.RemoveAt(0);
            return factors.Aggregate(seed, (agg, arr) => commonFactor(agg, arr));
        }
        private List<int> commonFactor(List<int> a, List<int> b)
        {
            // Returns an Intersection of the two lists, but includes duplicates
            List<int> intersect = new List<int>();

            List<int> subUnion;
            do
            {
                // Intersects removes duplicates, hence it has to be done multiple times
                subUnion = a.Intersect(b).ToList();
                subUnion.ForEach(sU =>
                {
                    a.Remove(sU);
                    b.Remove(sU);
                });
                intersect = intersect.Concat(subUnion).ToList();
            } while (subUnion.Count > 0);
            return intersect;
        }

        private decimal getTotal(List<Data> list)
        {
            return list.Aggregate(0M, (a, b) => a + b.Total);
        }

        private string toCurrency(decimal value)
        {
            if (!checkBox1.Checked) { return String.Format("{0:0.00}", value); }
            return value.ToString("C", System.Globalization.CultureInfo.GetCultureInfo("th-TH"));
        }
        private string toCurrency(string value)
        {
            if (!checkBox1.Checked) { return String.Format("{0:0.00}", value); }
            return decimal.Parse(value).ToString("C", System.Globalization.CultureInfo.GetCultureInfo("th-TH"));
        }

        private void outputData(List<Data> output)
        {
            // Create/Clear Table
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            DataTable table = new DataTable();
            table.Rows.Clear();
            table.Columns.Clear();

            // Create Columns
            List<string> columns = new List<string> {"ชื่อ", "ราคา", "จำนวน", "ราคารวม"};
            columns.ForEach(col => table.Columns.Add(col));

            // In case there are no results
            if (output == null)
            {
                table.Rows.Add("No", "possible", "combination", "found.");
                dataGridView1.DataSource = table;
                return;
            }

            // Ignore products that aren't considered
            output = output.Where(item => item.Count > 0).ToList();

            // Set Rows and bind data
            output.ForEach(item => table.Rows.Add(item.ID, toCurrency(item.Value), item.Count, toCurrency(item.Total)));
            dataGridView1.DataSource = table;
            dataGridView1.ReadOnly = true;
        }

        private bool toLimit(List<Data> choice, decimal goalValue)
        {
            // Base Cases
            if (goalValue < 0 || (choice.All(item => item.Ratio == 0) && goalValue != 0)) { return false; }
            if (goalValue == 0M) { return true; }

            // BFS(ish)
            foreach (Data item in choice)
            {
                if (goalValue - item.Value == 0 && item.Ratio > 0)
                {
                    item.add(1);
                    return true;
                }
            }

            // Recursive Call
            foreach (Data item in choice)
            {
                if (item.Ratio == 0) { continue; }
                item.add(1);
                item.Ratio--;
                if (toLimit(choice, goalValue - item.Value))
                {
                    return true;
                }
                item.sub(1);
                item.Ratio++;
            }
            return false;
        }

        private void comboBoxDestinations_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.button1.Enabled = false;
        }
    }
}