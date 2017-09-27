using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Npgsql;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

namespace ACMReports2
{
    public partial class WorktimeForm: Form
    {
        public WorktimeForm()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            load_departments();
            load_divisions();
        }

        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        private BindingSource bs = new BindingSource();

        private string selectedDepartment = "";     // selected department name
        private string selectedDivision = "";       // selected division name
        private int rowNum = new int();             // total rows in DataTable
        private const int rowsPerPage = 23;         // the number of rows in table per PDF page



        /// <summary>
        /// Load department list.
        /// </summary>
        public void load_departments()
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            // Load department list
            try
            {
                // Read settings from application configuration file
                string ACMServerIP = Properties.Settings.Default["ACMServerIP"].ToString();
                string ACMServerPort = Properties.Settings.Default["ACMServerPort"].ToString();
                string ACMUsername = Properties.Settings.Default["ACMUsername"].ToString();
                string ACMPassword = Properties.Settings.Default["ACMPassword"].ToString();

                // PostgeSQL-style connection string
                // SSL encryption is required for ACM connection
                string connStr = String.Format("Server={0};Port={1};UserId={2};Password={3};" +
                            "Database=TransactionDB;SSL Mode=Require;Trust Server Certificate=true",
                            ACMServerIP, ACMServerPort, ACMUsername, ACMPassword);

                NpgsqlConnection conn = new NpgsqlConnection(connStr);
                conn.Open();

                // Prepare SQL Request
                string SQLRequest = "SELECT DISTINCT identitydepartment " +
                                    "FROM rpt_alltrx " +
                                    "ORDER BY identitydepartment DESC";

                // Execute SQL Request
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQLRequest, conn);

                // Display the result in ComboBox
                ds.Reset();
                da.Fill(ds);
                comboBox1.DataSource = ds.Tables[0];
                comboBox1.DisplayMember = "identitydepartment";
                comboBox1.ValueMember = "identitydepartment";
                conn.Close();
            }
            catch (Exception ex)
            {
                // Silent notification if department list could not be recieved.
            }
        }


        /// <summary>
        /// Load division list.
        /// </summary>
        public void load_divisions()
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();

            // Load division list
            try
            {
                // Read settings from application configuration file
                string ACMServerIP = Properties.Settings.Default["ACMServerIP"].ToString();
                string ACMServerPort = Properties.Settings.Default["ACMServerPort"].ToString();
                string ACMUsername = Properties.Settings.Default["ACMUsername"].ToString();
                string ACMPassword = Properties.Settings.Default["ACMPassword"].ToString();

                // PostgeSQL-style connection string
                // SSL encryption is required for ACM connection
                string connStr = String.Format("Server={0};Port={1};UserId={2};Password={3};" +
                            "Database=TransactionDB;SSL Mode=Require;Trust Server Certificate=true",
                            ACMServerIP, ACMServerPort, ACMUsername, ACMPassword);

                NpgsqlConnection conn = new NpgsqlConnection(connStr);
                conn.Open();

                // Prepare SQL Request
                string SQLRequest = "SELECT DISTINCT identitydivision " +
                                    "FROM rpt_alltrx " +
                                    "ORDER BY identitydivision DESC";

                // Execute SQL Request
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQLRequest, conn);

                // Display the result in ComboBox
                ds.Reset();
                da.Fill(ds);
                comboBox2.DataSource = ds.Tables[0];
                comboBox2.DisplayMember = "identitydivision";
                comboBox2.ValueMember = "identitydivision";
                conn.Close();
            }
            catch (Exception ex)
            {
                // Silent notification if division list could not be recieved.
            }
        }



        /// <summary>
        /// Generate PDF button action.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created by ACMReports Generator.";
            document.Info.Subject = "Attendance Report created by ACMReports v2.0";
            document.Info.Author = "ACMReports v2.0";
            document.Info.Keywords = "ACMReports v2.0, Attendance Report";

            // Define the total number of pages in document
            int pageLast = rowNum / rowsPerPage + 1;  // the last page number

            // Make the page(s)
            for (int pageNum = 0; pageNum < pageLast; pageNum++)
            {
                Page1(document, pageNum + 1);
            }

            // Make directory to store reports
            const string directory = "Reports";
            Directory.CreateDirectory(directory);

            // Save the document
            string selectedDate = dateTimePicker1.Value.Date.ToString("yyyyMMdd");
            string filename = "ACMReports_" + selectedDate +
                "_" + selectedDepartment + "_" + selectedDivision + ".pdf";
            document.Save(directory + Path.DirectorySeparatorChar + filename);

            // Start PDF viewer
            Process.Start(directory + Path.DirectorySeparatorChar + filename);
        }



        /// <summary>
        /// PDF File generation.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="pageNum"></param>
        private void Page1(PdfDocument document, int pageNum)
        {
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            // Unicode hack
            gfx.MUH = PdfFontEncoding.Unicode;

            // Get the A4 page size
            Unit width, height;
            PageSetup.GetPageSize(PageFormat.A4, out width, out height);

            XFont font = new XFont("Verdana", 13, XFontStyle.Bold);

            // You always need a MigraDoc document for rendering.
            Document doc = new Document();

            // Add a section to the document and configure it such
            // that it will be in the center of the page
            Section section = doc.AddSection();
            section.PageSetup.PageHeight = height;
            section.PageSetup.PageWidth = width;
            section.PageSetup.LeftMargin = 0;
            section.PageSetup.RightMargin = 0;
            section.PageSetup.TopMargin = 20;

            // Add a top paragraph with date
            Paragraph para = section.AddParagraph();
            para.Format.Alignment = ParagraphAlignment.Right;
            para.Format.Font.Name = "Times New Roman";
            para.Format.Font.Size = 12;
            para.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.DarkGray;
            para.AddFormattedText(dateTimePicker1.Value.Date.ToString("dddd, d MMMM yyyy"), TextFormat.Bold);

            // Add a top department and division names
            Paragraph department = section.AddParagraph();
            department.Format.Alignment = ParagraphAlignment.Left;
            department.Format.Font.Name = "Times New Roman";
            department.Format.Font.Size = 12;
            department.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.DarkGray;
            department.AddFormattedText(String.Format("{0} | {1}", selectedDepartment, selectedDivision), TextFormat.Bold);

            // Add a bottom paragraph with page number
            Paragraph pg = section.AddParagraph();
            pg.Format.Alignment = ParagraphAlignment.Right;
            pg.Format.Font.Name = "Times New Roman";
            pg.Format.Font.Size = 12;
            pg.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.DarkGray;
            pg.AddFormattedText("Page " + pageNum, TextFormat.Bold);

            // Create a table
            Table table = new Table();
            table.Borders.Visible = true;
            table.Borders.Width = 1; // Default to show borders 1 pixel wide Column
            table.TopPadding = 5;
            table.BottomPadding = 5;

            Column column = table.AddColumn(40);
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn(150);
            column.Format.Alignment = ParagraphAlignment.Left;

            column = table.AddColumn(120);
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn(120);
            column.Format.Alignment = ParagraphAlignment.Center;

            column = table.AddColumn(80);
            column.Format.Alignment = ParagraphAlignment.Center;

            table.Rows.Height = 20;

            Row row = table.AddRow();
            row.Shading.Color = Colors.PaleGoldenrod;
            row.VerticalAlignment = VerticalAlignment.Top;

            // Table header
            row.Cells[0].AddParagraph("#");
            row.Cells[1].AddParagraph("Name");
            row.Cells[2].AddParagraph("Start");
            row.Cells[3].AddParagraph("End");
            row.Cells[4].AddParagraph("Duration");

            // Define the last row number on the current page
            int lastNum;
            if (rowNum >= pageNum * rowsPerPage)
            {
                lastNum = pageNum * rowsPerPage;
            }
            else
            {
                lastNum = rowNum;
            }

            // Put the rows into the table
            for (int i = 1 + (pageNum - 1) * rowsPerPage; i <= lastNum; i++)
            {
                DataRow dr = dt.Rows[i - 1];
                row = table.AddRow();
                row.Cells[0].AddParagraph(i.ToString());
                row.Cells[1].AddParagraph(dr["name"].ToString());
                row.Cells[2].AddParagraph(dr["start"].ToString());
                row.Cells[3].AddParagraph(dr["end"].ToString());
                row.Cells[4].AddParagraph(dr["duration"].ToString());
            }

            table.SetEdge(0, 0, 5, 1, Edge.Box, MigraDoc.DocumentObjectModel.BorderStyle.Single, 1.5, Colors.Black);

            doc.LastSection.Add(table);

            // Create a renderer and prepare (=layout) the document
            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc);
            docRenderer.PrepareDocument();

            // Render the paragraph. You can render tables or shapes the same way.
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(1), "10cm", department);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(10), XUnit.FromCentimeter(1), "10cm", para);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(2), "12cm", table);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(10), XUnit.FromCentimeter(28), "10cm", pg);
        }



        /// <summary>
        /// "Connection Setting" button action.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            Form2 myForm2 = new Form2();
            myForm2.Show();
        }



        /// <summary>
        /// "Read data" button action.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // Change coursor to wait (hourglass)
                Cursor.Current = Cursors.WaitCursor;

                // Read settings from application configuration file
                string ACMServerIP = Properties.Settings.Default["ACMServerIP"].ToString();
                string ACMServerPort = Properties.Settings.Default["ACMServerPort"].ToString();
                string ACMUsername = Properties.Settings.Default["ACMUsername"].ToString();
                string ACMPassword = Properties.Settings.Default["ACMPassword"].ToString();
                string EntranceEvent = Properties.Settings.Default["EntranceEvent"].ToString();
                string EntranceDoor = Properties.Settings.Default["EntranceDoor"].ToString();
                string ExitEvent = Properties.Settings.Default["ExitEvent"].ToString();
                string ExitDoor = Properties.Settings.Default["ExitDoor"].ToString();

                // PostgeSQL-style connection string
                // SSL encryption is required for ACM connection
                string connStr = String.Format("Server={0};Port={1};UserId={2};Password={3};" +
                        "Database=TransactionDB;SSL Mode=Require;Trust Server Certificate=true",
                        ACMServerIP, ACMServerPort, ACMUsername, ACMPassword);

                NpgsqlConnection conn = new NpgsqlConnection(connStr);
                conn.Open();

                // Prepare date range
                string startDate = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
                string endDate = dateTimePicker1.Value.AddDays(1).Date.ToString("yyyy-MM-dd");

                // Prepare department filter
                selectedDepartment = comboBox1.Text;
                string filterDepartment = String.Format("AND identitydepartment LIKE '%{0}%'", selectedDepartment);

                // Prepare division filter
            
    selectedDivision = comboBox2.Text;
                string filterDivision = String.Format("AND identitydivision LIKE '%{0}%'", selectedDivision);

                // Prepare final SQL request
                string SQLRequest = String.Format(
                "SELECT " +
                "    row_number() OVER (ORDER BY s.fn1 NULLS LAST) AS number, " +
                "    s.fn1 AS name, " +
                "    s.start_date AS start, " +
                "    e.end_date AS end, " +
                "    e.end_date - s.start_date AS duration, " +
                "    s.department AS department, " +
                "    s.division AS division " +
                "FROM " +
                "( " +
                "    SELECT " +
                "        to_timestamp(min(trxdateutc)) AS start_date, " +
                "        fullname AS fn1, " +
                "        identitydepartment AS department, " +
                "        identitydivision AS division " +
                "    FROM rpt_alltrx " +
                "    WHERE eventname LIKE '%{0}%' " +
                "        AND sourcename LIKE '%{1}%' " +
                "        {6} " +
                "        {7} " +
                "        AND to_timestamp(trxdateutc) >= to_timestamp('{4}', 'YYYY-MM-DD') " +
                "        AND to_timestamp(trxdateutc) <= to_timestamp('{5}', 'YYYY-MM-DD') " +
                "    GROUP BY fullname, department, division " +
                ") AS s " +
                "LEFT JOIN " +
                "( " +
                "    SELECT " +
                "        to_timestamp(max(trxdateutc)) AS end_date, " +
                "        fullname AS fn2" +
                "    FROM rpt_alltrx " +
                "    WHERE eventname LIKE '%{2}%' " +
                "        AND sourcename LIKE '%{3}%' " +
                "        AND to_timestamp(trxdateutc) >= to_timestamp('{4}', 'YYYY-MM-DD') " +
                "        AND to_timestamp(trxdateutc) <= to_timestamp('{5}', 'YYYY-MM-DD') " +
                "    GROUP BY fullname " +
                ") AS e " +
                "ON s.fn1 = e.fn2 " +
                "ORDER BY name ASC;",
                EntranceEvent, EntranceDoor, ExitEvent, ExitDoor,
                startDate, endDate, filterDepartment, filterDivision);

                // Execute SQL Request
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQLRequest, conn);

                // Display the result in GridView
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                bs.DataSource = dt;
                dataGridView1.DataSource = bs;
                conn.Close();

                // Enable "Export to PDF" button
                rowNum = dt.Rows.Count;
                if (rowNum > 0)
                {
                    button2.Enabled = true;
                }
                else
                {
                    button2.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                // if something went wrong, and we want to know why
                //MessageBox.Show("Connection or internal error. " + ex.ToString());
                MessageBox.Show("No connection to Transaction database.\nPlease check Connection Settings.", "Connection error");
            }

            // Set cursor to default state (arrow)
            Cursor.Current = Cursors.Default;
        }



        /// <summary>
        /// Edit Request button action.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            // Open SQL Request form
            Form3 myForm3 = new Form3();
            myForm3.Show();
        }



        /// <summary>
        /// Change Department filter action.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem != null)
            {
                DataRowView view = comboBox1.SelectedItem as DataRowView;
                selectedDepartment = view["identitydepartment"].ToString();
            }
        }



        /// <summary>
        /// Change Division filter action.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedItem != null)
            {
                DataRowView view = comboBox2.SelectedItem as DataRowView;
                selectedDivision = view["identitydivision"].ToString();
            }
        }
    }
}
