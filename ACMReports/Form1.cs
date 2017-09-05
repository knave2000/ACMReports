using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Npgsql;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

// Creation date: 01.09.2017
// Author: Vladimir Artyukhov, knave2000@gmail.com

namespace ACMReports
{
    public partial class WorktimeForm : Form
    {
        public WorktimeForm()
        {
            InitializeComponent();

            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();
        private int rowNum = new int();             // total rows in DataTable
        private const int rowsPerPage = 30;         // the number of rows in table per page

        private void button2_Click(object sender, EventArgs e)
        {
            // Create a new PDF document
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Created by ACM Report Generator.";
            document.Info.Subject = "Worktime report created by ACM Report Generator.";
            document.Info.Author = "ACM Report";
            document.Info.Keywords = "ACM Report, Worktime";

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
            string filename = "ACMReports_" + dateTimePicker1.Value.Date.ToString("yyyyMMdd") + ".pdf";
            document.Save(directory + Path.DirectorySeparatorChar + filename);

            // Start PDF viewer
            Process.Start(directory + Path.DirectorySeparatorChar + filename);
        }

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
            DataRow dr = dt.Rows[0];
            for (int i = 1 + (pageNum - 1) * rowsPerPage; i <= lastNum; i++)
            {
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
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(10), XUnit.FromCentimeter(1), "10cm", para);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(2), XUnit.FromCentimeter(2), "12cm", table);
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(10), XUnit.FromCentimeter(28), "10cm", pg);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 myForm2 = new Form2();
            myForm2.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
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

                // Prepare SQL Request
                string startDate = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
                string endDate = dateTimePicker1.Value.AddDays(1).Date.ToString("yyyy-MM-dd");

                string SQLRequest = String.Format(
                "SELECT " +
                "    row_number() OVER (ORDER BY s.fn1 NULLS LAST) AS number, " +
                "    s.fn1 AS name, " +
                "    s.start_date AS start, " +
                "    e.end_date AS end, " +
                "    e.end_date - s.start_date AS duration " +
                "FROM " +
                "( " +
                "    SELECT " +
                "        to_timestamp(min(trxdateutc)) AS start_date, " +
                "        fullname AS fn1 " +
                "    FROM rpt_alltrx " +
                "    WHERE eventname like '%{0}' " +
                "        AND sourcename like '%{1}%' " +
                "        AND to_timestamp(trxdateutc) >= to_timestamp('{4}', 'YYYY-MM-DD') " +
                "        AND to_timestamp(trxdateutc) <= to_timestamp('{5}', 'YYYY-MM-DD') " +
                "    GROUP BY fullname " +
                ") AS s " +
                "LEFT JOIN " +
                "( " +
                "    SELECT " +
                "        to_timestamp(max(trxdateutc)) AS end_date, " +
                "        fullname AS fn2" +
                "    FROM rpt_alltrx " +
                "    WHERE eventname like '%{2}%' " +
                "        AND sourcename like '%{3}%' " +
                "        AND to_timestamp(trxdateutc) >= to_timestamp('{4}', 'YYYY-MM-DD') " +
                "        AND to_timestamp(trxdateutc) <= to_timestamp('{5}', 'YYYY-MM-DD') " +
                "    GROUP BY fullname " +
                ") AS e " +
                "ON s.fn1 = e.fn2 " +
                "ORDER BY name ASC;",
                EntranceEvent, EntranceDoor, ExitEvent, ExitDoor, startDate, endDate);

                // Execute SQL Request
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(SQLRequest, conn);

                // Display the result in GridView
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                dataGridView1.DataSource = dt;
                conn.Close();

                // Enable Export button
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
            catch
            {
                // if something went wrong, and we want to know why
                MessageBox.Show("Connection or internal error.");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Open SQL Request form
            Form3 myForm3 = new Form3();
            myForm3.Show();
        }
    }
}
