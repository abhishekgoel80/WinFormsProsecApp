
//using CrystalDecisions.ReportAppServer.DataDefModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;



namespace WinFormsProsecApp
{
    public partial class Form1 : Form
    {

        // Connection string for database
        string dbConnString = "Data Source=LAPTOP-TSKTCB99\\SQLEXPRESS;Initial Catalog=Prosec;persist security info=True;Integrated Security=SSPI;";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Visible = true;
                textBox1.Text = openFileDialog.FileName;
                textBox1.Focus();
            }

            // Connection string for Excel file
            string excelConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";

            // Create the connection objects

            OleDbConnection excelConn = new(excelConnString);
            //System.Data.IDbConnection excelConn = new IDbConnection(excelConnString);
            SqlConnection dbConn = new(dbConnString);
            //System.Data.IDbConnection dbConn = new IDbConnection(dbConnString);


            try
            {
                // Open the Excel file connection
                excelConn.Open();

                // Get the sheet name from the Excel file
                DataTable dtSheet = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = dtSheet.Rows[0]["TABLE_NAME"].ToString();

                // Read data from Excel file
                OleDbDataAdapter daExcel = new("SELECT * FROM [" + sheetName + "]", excelConn);
                DataTable dtExcel = new();
                daExcel.Fill(dtExcel);

                // Open the database connection
                dbConn.Open();

                // Create the SQL command to insert data into the database table
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dbConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO T_raw_input VALUES (@employee_id, @empName, @Day1, @Day2, @Day3, @Day4, @Day5, @Day6, @Day7, @Day8," +
                    " @Day9, @Day10, @Day11, @Day12, @Day13, @Day14, @Day15, @Day16, @Day17, @Day18, @Day19, @Day20, @Day21, @Day22, @Day23," +
                    " @Day24, @Day25, @Day26, @Day27, @Day28, @Day29, @Day30, @Day31, @month, @year, @total_as_user_input, @created_by_userId," +
                    " @created_at_timestamp,@lastUpdated_by_user_id,@lastUpdated_timestamp,@status)";

                // Add parameters to the SQL command
                cmd.Parameters.Add("@employee_id", SqlDbType.NVarChar, 10);
                cmd.Parameters.Add("@empName", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@Day1", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day2", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day3", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day4", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day5", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day6", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day7", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day8", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day9", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day10", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day11", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day12", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day13", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day14", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day15", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day16", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day17", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day18", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day19", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day20", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day21", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day22", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day23", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day24", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day25", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day26", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day27", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day28", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day29", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day30", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Day31", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@total_as_user_input", SqlDbType.Int);
                cmd.Parameters.Add("@month", SqlDbType.Int);
                cmd.Parameters.Add("@year", SqlDbType.Int);
                cmd.Parameters.Add("@created_by_userId", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@created_at_timestamp", SqlDbType.DateTime);
                cmd.Parameters.Add("@lastUpdated_by_user_id", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@lastUpdated_timestamp", SqlDbType.DateTime);
                cmd.Parameters.Add("@status", SqlDbType.Bit);

                // Loop through the rows in the Excel data table and insert into database table
                foreach (DataRow row in dtExcel.Rows)
                {
                    if (row["employee_id"].ToString() == "") break;

                    cmd.Parameters["@employee_id"].Value = row["employee_id"].ToString();
                    cmd.Parameters["@empName"].Value = row["empName"].ToString();
                    cmd.Parameters["@Day1"].Value = row["Day1"].ToString();
                    cmd.Parameters["@Day2"].Value = row["Day2"].ToString();
                    cmd.Parameters["@Day3"].Value = row["Day3"].ToString();
                    cmd.Parameters["@Day4"].Value = row["Day4"].ToString();
                    cmd.Parameters["@Day5"].Value = row["Day5"].ToString();
                    cmd.Parameters["@Day6"].Value = row["Day6"].ToString();
                    cmd.Parameters["@Day7"].Value = row["Day7"].ToString();
                    cmd.Parameters["@Day8"].Value = row["Day8"].ToString();
                    cmd.Parameters["@Day9"].Value = row["Day9"].ToString();
                    cmd.Parameters["@Day10"].Value = row["Day10"].ToString();
                    cmd.Parameters["@Day11"].Value = row["Day11"].ToString();
                    cmd.Parameters["@Day12"].Value = row["Day12"].ToString();
                    cmd.Parameters["@Day13"].Value = row["Day13"].ToString();
                    cmd.Parameters["@Day14"].Value = row["Day14"].ToString();
                    cmd.Parameters["@Day15"].Value = row["Day15"].ToString();
                    cmd.Parameters["@Day16"].Value = row["Day16"].ToString();
                    cmd.Parameters["@Day17"].Value = row["Day17"].ToString();
                    cmd.Parameters["@Day18"].Value = row["Day18"].ToString();
                    cmd.Parameters["@Day19"].Value = row["Day19"].ToString();
                    cmd.Parameters["@Day20"].Value = row["Day20"].ToString();
                    cmd.Parameters["@Day21"].Value = row["Day21"].ToString();
                    cmd.Parameters["@Day22"].Value = row["Day22"].ToString();
                    cmd.Parameters["@Day23"].Value = row["Day23"].ToString();
                    cmd.Parameters["@Day24"].Value = row["Day24"].ToString();
                    cmd.Parameters["@Day25"].Value = row["Day25"].ToString();
                    cmd.Parameters["@Day26"].Value = row["Day26"].ToString();
                    cmd.Parameters["@Day27"].Value = row["Day27"].ToString();
                    cmd.Parameters["@Day28"].Value = row["Day28"].ToString();
                    cmd.Parameters["@Day29"].Value = row["Day29"].ToString();
                    cmd.Parameters["@Day30"].Value = row["Day30"].ToString();
                    cmd.Parameters["@Day31"].Value = row["Day31"].ToString();
                    cmd.Parameters["@month"].Value = int.Parse(row["month"].ToString());
                    cmd.Parameters["@year"].Value = int.Parse(row["year"].ToString());
                    cmd.Parameters["@total_as_user_input"].Value = int.Parse(row["total_as_user_input"].ToString());
                    cmd.Parameters["@created_by_userId"].Value = "Default User";
                    cmd.Parameters["@created_at_timestamp"].Value = DateTime.Now.ToString();
                    cmd.Parameters["@lastUpdated_by_user_id"].Value = "Default User";
                    cmd.Parameters["@lastUpdated_timestamp"].Value = DateTime.Now.ToString();
                    cmd.Parameters["@status"].Value = 1;

                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show("Data imported successfully.");
                DataTable data = GetDataFromDatabase1();
                dataGridView1.DataSource = data;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Close the connections
                excelConn.Close();
                dbConn.Close();
            }
        }


        //Process salary 
        private void button2_Click(object sender, EventArgs e)
        {
            // Create the connection objects
            //Getting the raw_attendance data from the database of the month currently hardcoded
            DataTable data = GetDataFromDatabase1();

            // Example: Process and display the data
            foreach (DataRow row in data.Rows)
            {
                // Access data using column names or indexes
                string column1Value = row["employee_id"].ToString();
                string column2Value = row["empName"].ToString();
                int column3Value = int.Parse(row["total_as_user_input"].ToString());
                int column4Value = int.Parse(row["month"].ToString());
                int column5Value = int.Parse(row["year"].ToString());
                CalculateSalary(column1Value, column2Value, column3Value, column4Value, column5Value);

                // Process and display the data as needed
                // MessageBox.Show("employee_id: " + column1Value);
                // MessageBox.Show("empName: " + column2Value);
            }

            //Display datagrid
            //Display datarid

            DataTable data1 = GetDataFromDatabase2();
            dataGridView1.DataSource = data1;
        }
        private void CalculateSalary(string id, string name, int attendance, int month, int year)
        {
            //constant values
            // this to be obtained from DB based on the id of the person we need to get it designation
            //this has to generated based on the company remplated
            //double Wages = 897;//get it from table
            //double Wages_ 2 = 973;//get it from table
            //double Wages_3 = 1193.01;//get it from table
            double WagesB = 816;//get it from table
            double WagesB_2 = 897;//get it from table
            double WagesB_3 = 1085.28;//get it from table
            double WagesC = 695;//get it from table
            double WagesC_2 = 816;//get it from table
            double WagesC_3 = 924.35;//get it from table
            double basicPlusVDA = 26 * WagesB;
            double pf_percentage = 0.12;
            double pf_ceiling_amount = 15000;
            double edli_percentage = 0.005;
            double admin_charges_percentage = 0.005;
            double esi_percentage =  0.0325;
            double hra_percentage = 0.24;//.16, 0.08
            double hra1 = 215.28;
            double hra2 = 233.52;
            double hra3 = 286.32;
            double esicOnHRA_percentage = 0.0325;
            double bonus_percentage = 0.0833;
            double uniform_allowance_percentage = 0.05;
            double washing_allowance_percentage = 0.03;
            double relieving_charges_fraction = 0.16667;
            double ESIPercent = 0.0075;
            double EPF_13_percent = 0.13;
            int duty, LRD;
            if (attendance < 27)
            {
                duty = attendance;
                LRD = 0;
            }
            else
            {
                duty = 26;
                LRD = attendance - duty;
            }

            //Values c
            string empID = id;
            string Desig = "xyz";
            string ESM_CIV_FLAG = "xyz";
            string Name = name;
            string Father_Name = "xyz";
            int Month = month;
            int Year = year;
            string CityClass = "xyz";
            string combinedCode = "xyz";
            double BasicPlusVDA = basicPlusVDA;
            int AttndTotal = attendance;
            double Basic = AttndTotal * BasicPlusVDA / DateTime.DaysInMonth(Year, Month);
            double PF1 = pf_percentage * Basic;
            double EDLI = edli_percentage * Basic;
            double Admin_Charges = admin_charges_percentage * Basic;
            double ESI = ESIPercent * Basic;
            double HRA = 3600 * AttndTotal / DateTime.DaysInMonth(Year, Month); ;
            double ESIConHRA = esicOnHRA_percentage * HRA;
            double Bonus = bonus_percentage * Basic;
            double UniformAllowance = uniform_allowance_percentage * Basic;
            double Washing_Allowance = washing_allowance_percentage * Basic;
            double Total1 = Basic + PF1 + EDLI + Admin_Charges + ESI + HRA + ESIConHRA + Bonus + UniformAllowance + Washing_Allowance;
            double Releiving_Charges = relieving_charges_fraction * Total1;
            double Gross_Total = Releiving_Charges + Total1;
            double PFDed = PF1;
            double EPFDed = PF1 * 13 / 12;
            double ESIDed = ESI;
            double UniformDed = UniformAllowance;
            double TotalDed = PFDed + EPFDed + ESIDed + UniformDed;
            double Net_Amt_Paid = Gross_Total - TotalDed;

            SqlConnection dbConn = new(dbConnString);
            try
            {
                dbConn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = dbConn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO T_Calculation VALUES (@empID, @Desig, @ESM_CIV_FLAG, @Name, @Father_Name, @month, @year, @CityClass," +
                        " @basicPlusVDA, @AttndTotal, @Basic, @PF, @EDLI, @AdminCharges, @ESI, @HRA, @ESIConHRA, @Bonus, @UniformAllowance," +
                        " @WashingAllowance, @Total, @ReleivingCharges, @GrossTotal, @PFDeduction, @ESIDeduction, @UniformDeduction, @TotalDeduction, " +
                        " @NetAmtPaid, @created_by_userId, @created_at_timestamp,@lastUpdated_by_user_id,@lastUpdated_timestamp,@status)";

                // Add parameters to the SQL command
                cmd.Parameters.Add("@empID", SqlDbType.NVarChar, 10);
                cmd.Parameters.Add("@Desig", SqlDbType.NVarChar, 10);
                cmd.Parameters.Add("@ESM_CIV_FLAG", SqlDbType.NVarChar, 3);
                cmd.Parameters.Add("@Name", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@Father_Name", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@month", SqlDbType.Int);
                cmd.Parameters.Add("@year", SqlDbType.Int);
                cmd.Parameters.Add("@CityClass", SqlDbType.NVarChar, 10);
                cmd.Parameters.Add("@basicPlusVDA", SqlDbType.Decimal);
                cmd.Parameters.Add("@AttndTotal", SqlDbType.Int);
                cmd.Parameters.Add("@Basic", SqlDbType.Decimal);
                cmd.Parameters.Add("@PF", SqlDbType.Decimal);
                cmd.Parameters.Add("@EDLI", SqlDbType.Decimal);
                cmd.Parameters.Add("@AdminCharges", SqlDbType.Decimal);
                cmd.Parameters.Add("@ESI", SqlDbType.Decimal);
                cmd.Parameters.Add("@HRA", SqlDbType.Decimal);
                cmd.Parameters.Add("@ESIConHRA", SqlDbType.Decimal);
                cmd.Parameters.Add("@Bonus", SqlDbType.Decimal);
                cmd.Parameters.Add("@UniformAllowance", SqlDbType.Decimal);
                cmd.Parameters.Add("@WashingAllowance", SqlDbType.Decimal);
                cmd.Parameters.Add("@Total", SqlDbType.Decimal);
                cmd.Parameters.Add("@ReleivingCharges", SqlDbType.Decimal);
                cmd.Parameters.Add("@GrossTotal", SqlDbType.Decimal);
                cmd.Parameters.Add("@PFDeduction", SqlDbType.Decimal);
                cmd.Parameters.Add("@ESIDeduction", SqlDbType.Decimal);
                cmd.Parameters.Add("@UniformDeduction", SqlDbType.Decimal);
                cmd.Parameters.Add("@TotalDeduction", SqlDbType.Decimal);
                cmd.Parameters.Add("@NetAmtPaid", SqlDbType.Decimal);
                cmd.Parameters.Add("@created_by_userId", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@created_at_timestamp", SqlDbType.DateTime);
                cmd.Parameters.Add("@lastUpdated_by_user_id", SqlDbType.NVarChar, 50);
                cmd.Parameters.Add("@lastUpdated_timestamp", SqlDbType.DateTime);
                cmd.Parameters.Add("@status", SqlDbType.Bit);

                //Populate
                cmd.Parameters["@empID"].Value = empID;
                cmd.Parameters["@Desig"].Value = Desig;
                cmd.Parameters["@ESM_CIV_FLAG"].Value = ESM_CIV_FLAG;
                cmd.Parameters["@Name"].Value = Name;
                cmd.Parameters["@Father_Name"].Value = Father_Name;
                cmd.Parameters["@month"].Value = Month;
                cmd.Parameters["@year"].Value = Year;
                cmd.Parameters["@CityClass"].Value = CityClass;
                cmd.Parameters["@basicPlusVDA"].Value = basicPlusVDA;
                cmd.Parameters["@AttndTotal"].Value = AttndTotal;
                cmd.Parameters["@Basic"].Value = Basic;
                cmd.Parameters["@PF"].Value = PF1;
                cmd.Parameters["@EDLI"].Value = EDLI;
                cmd.Parameters["@AdminCharges"].Value = Admin_Charges;
                cmd.Parameters["@ESI"].Value = ESI;
                cmd.Parameters["@HRA"].Value = HRA;
                cmd.Parameters["@ESIConHRA"].Value = ESIConHRA;
                cmd.Parameters["@Bonus"].Value = Bonus;
                cmd.Parameters["@UniformAllowance"].Value = UniformAllowance;
                cmd.Parameters["@WashingAllowance"].Value = Washing_Allowance;
                cmd.Parameters["@Total"].Value = Total1;
                cmd.Parameters["@ReleivingCharges"].Value = Releiving_Charges;
                cmd.Parameters["@GrossTotal"].Value = Gross_Total;
                cmd.Parameters["@PFDeduction"].Value = PFDed;
                cmd.Parameters["@ESIDeduction"].Value = ESIDed; ;
                cmd.Parameters["@UniformDeduction"].Value = UniformDed;
                cmd.Parameters["@TotalDeduction"].Value = TotalDed;
                cmd.Parameters["@NetAmtPaid"].Value = Net_Amt_Paid;
                cmd.Parameters["@created_by_userId"].Value = "Default User";
                cmd.Parameters["@created_at_timestamp"].Value = DateTime.Now.ToString();
                cmd.Parameters["@lastUpdated_by_user_id"].Value = "Default User";
                cmd.Parameters["@lastUpdated_timestamp"].Value = DateTime.Now.ToString();
                cmd.Parameters["@status"].Value = 1;

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Close the connections
                dbConn.Close();
            }
        }

        private DataTable GetDataFromDatabase1()
        {

            using (SqlConnection connection = new SqlConnection(dbConnString))
            {
                string query = " SELECT * from T_raw_input where month = 6";
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();

                try
                {
                    connection.Open();
                    adapter.Fill(dataTable);
                    MessageBox.Show("Data ready for processing");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                return dataTable;
            }

        }
        private DataTable GetDataFromDatabase2()
        {

            using (SqlConnection connection = new SqlConnection(dbConnString))
            {
                string query = " SELECT * from T_Calculation where month = 6";
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();

                try
                {
                    connection.Open();
                    adapter.Fill(dataTable);
                    MessageBox.Show("Data ready for processing");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                return dataTable;
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            // Create the connection objects
            using (SqlConnection connection = new SqlConnection(dbConnString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand("Create_transfersheet", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;

                    // Add any required parameters to the command
                    //command.Parameters.AddWithValue("@Parameter1", "Value1");
                    //command.Parameters.AddWithValue("@Parameter2", 2);

                    // Execute the stored procedure
                    command.ExecuteNonQuery();
                }

            }
            DataTable data = GetDataFromDatabase3();

            //Display datarid
            dataGridView1.DataSource = data;
        }


        private DataTable GetDataFromDatabase3()
        {

            using (SqlConnection connection = new SqlConnection(dbConnString))
            {
                string query = " SELECT * from T_Bank_Transfer_Sheet where Month = 6";
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();

                try
                {
                    connection.Open();
                    adapter.Fill(dataTable);
                    MessageBox.Show("Transfer sheet Generated");

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                return dataTable;
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void fileSystemWatcher1_Changed(object sender, FileSystemEventArgs e)
        {

        }

        // This button is used to export the data to excel
        private void button4_Click(object sender, EventArgs e)
        {

            // Export data from DataGridView to Excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Export headers
            for (int i = 1; i <= dataGridView1.Columns.Count; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }

            // Export data rows
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }

            // Auto-fit columns
            worksheet.Columns.AutoFit();
        }
    }
}