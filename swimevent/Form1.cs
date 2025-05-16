using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using OfficeOpenXml;
            

namespace swimevent
{
    public partial class Form1 : Form
    {

        string Data1 = "Form Date";
        string Data2 = "Player Name";
        string Data3 = "SFIID";
        string Data4 = "Gender";
        string Data5 = "Birth Date";
        string Data06 = "Unit";
        string Data6 = "AgeGroup";
        string Data7 = "Events";
        DataTable dt = new DataTable();
        string competitonName;



        public Form1()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            //LicenseManager.UsageMode = LicenseUsageMode.Runtime;
            InitializeComponent();
            
            //LicenseContext = System.ComponentModel.LicenseContext.NonCommercial;

            webBrowser2.DocumentText = "<html>" +

                "<HEAD>" +
                "<STYLE>" +
                ".Name1 {" +
                "position: absolute;" +
                "top: 725px;" +
                "left: 249px;" +
                "width: 800px;"+
                "color : black" +
                "}" +
                ".Unit1 {" +
                "position: absolute;" +
                "top: 1079px;" +
                "left: 449px;" +
                "width: 800px;" +
                "color : black" +
                "}" +
                ".AgeGroup1 {" +
                "position: absolute;" +
                "top: 990px;" +
                "left: 249px;" +
                "width: 800px;" +
                "color : black" +
                "}" +
                ".Event1 {" +
                "position: absolute;" +
                "top: 806px;" +
                "left: 450px;" +
                "width: 800px;" +
                "color : black" +
                "}" +
                ".Place1 {" +
                "position: absolute;" +
                "top: 896px;" +
                "left: 249px;" +
                "width: 800px;" +
                "color : black" +
                "}" +
                ".Performance1 {" +
                "position: absolute;" +
                "top: 896px;" +
                "left: 800px;" +
                "width: 800px;" +
                "color : black" +
                "}" +
                ".DATE1 {" +
                "position: absolute;" +
                "top: 1250px;" +
                "left: 209px;" +
                "width: 800px;" +
                "color : black" +
                "}" +
                "</style> " +
                "</HEAD>" +
                "<body >" +
                 "<P style= \"font-size: 35px\" class=\"Name1\"  id=\"txt1\"> <b>BHARGAV CHIRAG SAILOR" +
              "</b></p>" +
              "<P style= \"font-size: 35px\" class=\"Unit1\" id=\"txt2\"> <b>    SURAT " +
              "</b></P>" +
              "<p  style= \"font-size: 35px\" class=\"AgeGroup1\" id=\"txt3\" ><b> 25 YEAR - 29 YEAR" +

              "</b></p>" +
              "<P style= \"font-size: 35px\" class=\"Event1\" id=\"txt4\"><b> 200MT BUTTER FLY" +
              "</b></p>" +
              "<P style= \"font-size: 35px\" class=\"Place1\" id=\"txt5\"><b> FIRST" +
              "</b></p>" +
              "<P style= \"font-size: 35px\" class=\"Performance1\" id=\"txt6\"><b> 02:12:38" +
              "</b></p>" +
              "<P style= \"font-size: 35px\" class=\"DATE1\" id=\"txt7\"><b> 05/05/2025" +
              "</b></p>" +
              "</body>" +
                "</html>";
            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
            buttonColumn.Name = "ActionColumn"; // Give the column a name
            buttonColumn.Text = "Action"; // Button text
            buttonColumn.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(buttonColumn);
            // webBrowser1.DocumentText
            competitonName = "<h2 style=\"text-align: center\">  SENIOR AND PREAGE GROUP   </h2> <h2 style=\"text-align: center\">AQUATIC CHAMPIONSHIP - 2025</h2> ";
            // competitonName = "<h2 style=\"text-align: center\">  GUJARAT STATE MASTER </h2> <h2 style=\"text-align: center\">AQUATIC CHAMPIONSHIP - 2024</h2> ";



            //webBrowser2.DocumentText = "<html>" +

            //    "<HEAD>" +
            //    "<STYLE>" +
            //    ".Name1 {" +
            //    "position: absolute;" +
            //    "top: 679px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    ".Unit1 {" +
            //    "position: absolute;" +
            //    "top: 766px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    ".AgeGroup1 {" +
            //    "position: absolute;" +
            //    "top: 856px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    ".Event1 {" +
            //    "position: absolute;" +
            //    "top: 946px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    ".Place1 {" +
            //    "position: absolute;" +
            //    "top: 1043px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    ".Performance1 {" +
            //    "position: absolute;" +
            //    "top: 1129px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    ".DATE1 {" +
            //    "position: absolute;" +
            //    "top: 1220px;" +
            //    "left: 249px;" +
            //    "width: 800px;" +
            //    "color : black" +
            //    "}" +
            //    "</style> " +
            //    "</HEAD>" +
            //    "<body >" +
            //     "<P style= \"font-size: 35px\" class=\"Name1\"  id=\"txt1\"> <b>BHARGAV CHIRAG SAILOR" +
            //  "</b></p>" +
            //  "<P style= \"font-size: 35px\" class=\"Unit1\" id=\"txt2\"> <b>    SURAT " +
            //  "</b></P>" +
            //  "<p  style= \"font-size: 35px\" class=\"AgeGroup1\" id=\"txt3\" ><b> 25 YEAR - 29 YEAR" +

            //  "</b></p>" +
            //  "<P style= \"font-size: 35px\" class=\"Event1\" id=\"txt4\"><b> 200MT BUTTER FLY" +
            //  "</b></p>" +
            //  "<P style= \"font-size: 35px\" class=\"Place1\" id=\"txt5\"><b> FIRST" +
            //  "</b></p>" +
            //  "<P style= \"font-size: 35px\" class=\"Performance1\" id=\"txt6\"><b> 02:12:38" +
            //  "</b></p>" +
            //  "<P style= \"font-size: 35px\" class=\"DATE1\" id=\"txt7\"><b> &nbsp" +
            //  "</b></p>" +
            //  "</body>" +
            //    "</html>";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.Title = "Select an Excel File";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                LoadExcelData(filePath);
            }


        }


       


        private void LoadExcelData(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Assuming the data is in the first sheet

            // Read data from the Excel file and populate the DataGridView
            
            dt.Columns.Add(Data1);
            dt.Columns.Add(Data2);
            dt.Columns.Add(Data3);
            dt.Columns.Add(Data4);
            dt.Columns.Add(Data5);
            dt.Columns.Add(Data06);
            dt.Columns.Add(Data6);
            dt.Columns.Add(Data7);
            for (int i = 7; i <= worksheet.UsedRange.Columns.Count; i++)
            {
                dt.Columns.Add("Column " + i);

            }

            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; row++)
            {
                DataRow dataRow = dt.NewRow();
                for (int col = 1; col <= worksheet.UsedRange.Columns.Count; col++)
                {
                    dataRow[col - 1] = worksheet.Cells[row, col].Value;
                }
                dt.Rows.Add(dataRow);
            }


            DataTable dtdropdowngender= new DataTable();
            dtdropdowngender = dt.AsEnumerable().GroupBy(r => new { Col1 = r[Data4] }).Select(g => g.OrderByDescending(r => r[Data4]).First()).CopyToDataTable();

            foreach (DataRow row in dtdropdowngender.Rows)
            {
                comboBox3.Items.Add(row[Data4].ToString());
            }


            DataTable dtdropdownSchool = new DataTable();
            dtdropdownSchool = dt.AsEnumerable().GroupBy(r => new { Col1 = r[Data06] }).Select(g => g.OrderByDescending(r => r[Data06]).First()).CopyToDataTable();

            foreach (DataRow row in dtdropdownSchool.Rows)
            {
                comboBox4.Items.Add(row[Data06].ToString());
            }

            DataTable dtdropdown = new DataTable();
            dtdropdown = dt.AsEnumerable().GroupBy(r => new { Col1 = r[Data6] }).Select(g => g.OrderByDescending(r => r[Data6]).First()).CopyToDataTable();

            foreach (DataRow row in dtdropdown.Rows)
            {
                comboBox1.Items.Add(row[Data6].ToString());
                comboBox5.Items.Add(row[Data6].ToString());
            }



           // DataTable dtdropdownevent = new DataTable();
           // dtdropdownevent = dt.AsEnumerable().GroupBy(r => new { Col1 = r[Data7] }).Select(g => g.OrderBy(r => r[Data7]).First()).CopyToDataTable();

           // foreach (DataRow row in dtdropdownevent.Rows)
           // {


            // event combo box 
            comboBox2.Items.Add("50mt Free Style");
            comboBox2.Items.Add("100mt Free Style");
            comboBox2.Items.Add("200mt Free Style");
            comboBox2.Items.Add("400mt Free Style");
            comboBox2.Items.Add("800mt Free Style");
            comboBox2.Items.Add("1500mt Free Style");
            comboBox2.Items.Add("50mt Breast Stroke");
            comboBox2.Items.Add("100mt Breast Stroke");
            comboBox2.Items.Add("200mt Breast Stroke");
            comboBox2.Items.Add("50mt Butter Fly Stroke");
            comboBox2.Items.Add("100mt Butter Fly Stroke");
            comboBox2.Items.Add("200mt Butter Fly Stroke");
            comboBox2.Items.Add("50mt Back Stroke");
            comboBox2.Items.Add("100mt Back Stroke");
            comboBox2.Items.Add("200mt Back Stroke");
            comboBox2.Items.Add("200mt Individual Medlay");
            comboBox2.Items.Add("400mt Individual Medlay");

            foreach (var item in comboBox2.Items)
            {
                comboBox6.Items.Add(item);
            }
            // }
            // Display the data in the DataGridView
            dataGridView1.DataSource = dt;

            // Close Excel objects
            workbook.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Boolean Flagoprater = false;
            Boolean super = true;
            string selectbox1 = comboBox1.SelectedItem?.ToString(); // Get the selected value from the ComboBox
            string selectbox2 = comboBox2.SelectedItem?.ToString();
            string selectbox3 = comboBox3.SelectedItem?.ToString();
            string selectbox4 = comboBox4.SelectedItem?.ToString();
            StringBuilder filterQuery = new StringBuilder();
            if (!string.IsNullOrEmpty(selectbox1))
            {
                
                filterQuery.Append (Data6 + " LIKE '%" + selectbox1 + "%'");
                Flagoprater = true;
            }
            if (!string.IsNullOrEmpty(selectbox2))
            {
                if (Flagoprater)
                    filterQuery.Append(" AND ");
                filterQuery.Append(Data7 + " LIKE '%" + selectbox2 + "%'");
                Flagoprater = true;
            }
            if (!string.IsNullOrEmpty(selectbox3))
            {
                if (Flagoprater)
                    filterQuery.Append(" AND ");
                filterQuery.Append(Data4 + " = '" + selectbox3 + "'");
                Flagoprater = true;
            }
            if (!string.IsNullOrEmpty(selectbox4))
            {
                if (Flagoprater)
                    filterQuery.Append(" AND ");
                filterQuery.Append(Data06 + " LIKE '%" + selectbox4 + "%'");
                Flagoprater = true;
            }

            if(Flagoprater)
            {
                if (dataGridView1.DataSource != null)
                    {
                    if(dt.Select(filterQuery.ToString()).ToList().Count > 0)
                    {
                        dataGridView1.DataSource = dt.Select(filterQuery.ToString()).CopyToDataTable();
                        
                    }
                    else
                    {
                        super = false;
                    }
                
                }
            }
            else
            {
                // If no value is selected, show all data
                dataGridView1.DataSource = dt;
            }


            StringBuilder crateTemplet = new StringBuilder();
            DataTable templetDt = new DataTable();
            DataTable temp = (DataTable)dataGridView1.DataSource;
            templetDt.Columns.Add(Data2);
            templetDt.Columns.Add(Data06);
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                templetDt.Rows.Add(new string[] { temp.Rows[i][1].ToString(), temp.Rows[i][5].ToString() });
            }

            if (super)
            {


                int headno = 1;
                crateTemplet.Append("<html><HEAD><style>table{border-collapse: collapse} tr {height : 32px}</style></HEAD><body>");
                crateTemplet.Append(competitonName);
                crateTemplet.Append("<H3 style=\"text-align: center\">" + selectbox2 + "(" + selectbox3 + ")</H3>");
                crateTemplet.Append("<H5 style=\"text-align: center\">" + selectbox1 + "                HEAT NO " + headno + "</H5>");
                crateTemplet.Append("<table style=\"empty-cells: show\" border = \"1px solid black \">");
                crateTemplet.Append("<tr>\r\n");
                //crateTemplet.Append(" <th> Sr.No</th> <th > Player Name </th> <th width='10%'>LANE NO</th><th> MIN . </th><th> SEC . </th><th> MIC . </th><th width='50px'>REMARK</th>");
                crateTemplet.Append(" <th> Sr.No</th> <th > Player Name </th> <th width='50px'>UNIT</th><th width='10%'>LANE NO</th><th> MIN . </th><th> SEC . </th><th> MIC . </th>");
                crateTemplet.Append("</tr>\r\n");
                // for (int i = 0; i < templetDt.Rows.Count; i++)
                // {
                int playerCount = 0;
                foreach (DataRow r in templetDt.Rows)
                {
                    playerCount++;
                    if (playerCount == 9)
                    {
                        playerCount = 1;
                        headno++;
                        crateTemplet.Append("</table>");
                        crateTemplet.Append(competitonName);
                        crateTemplet.Append("<H3 style=\"text-align: center\">" + selectbox2 + "(" + selectbox3 + ")</H3>");
                        crateTemplet.Append("<H5 style=\"text-align: center\">" + selectbox1 + "                HEAT NO " + headno + "</H5>");
                        crateTemplet.Append("<table style = \"empty-cells: show\" border = \"1px solid black \">");
                        crateTemplet.Append("<tr>\r\n");
                        crateTemplet.Append(" <th> Sr.No</th> <th> Player Name </th> <th width='50px'>UNIT</th><th width='10%'> LANE NO </th><th> MIN. </th><th> SEC. </th><th> MIC. </th>");
                        crateTemplet.Append("</tr>\r\n");
                    }
                    crateTemplet.Append("<tr>");
                    crateTemplet.Append("<td>" + playerCount + "</td>");
                    foreach (DataColumn c in templetDt.Columns)
                    {
                        crateTemplet.Append("<td >");
                        crateTemplet.Append(r[c.ColumnName]);
                        crateTemplet.Append("</td>");

                        //crateTemplet.Append("<td  width='50px'>");
                        //crateTemplet.Append(r["Unit"]);
                        //crateTemplet.Append("</td>");
                    }
                    crateTemplet.Append("<td width='10%'>&nbsp</td>");
                    crateTemplet.Append("<td>&nbsp</td>");
                    crateTemplet.Append("<td>&nbsp</td>");
                    crateTemplet.Append("<td width='50px'>&nbsp</td>");
                    crateTemplet.Append("</tr>");



                }
                //}
                crateTemplet.Append("</table>");


                crateTemplet.Append(" </body></html>");

                webBrowser1.DocumentText = crateTemplet.ToString();
            }
            else
            {
                MessageBox.Show("NO DATA");
            }
        }



        private void button3_Click(object sender, EventArgs e)
        {
            //StringBuilder filterQuery = new StringBuilder();
            PrintDocument printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);

            // Show the print dialog to allow the user to select a printer
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                webBrowser1.Print();
            }
        }
        private Bitmap bitmapToPrint;
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (bitmapToPrint == null)
            {
                // Create a Bitmap to capture the content of the WebBrowser control
                bitmapToPrint = new Bitmap(webBrowser1.Width, webBrowser1.Height);
                webBrowser1.DrawToBitmap(bitmapToPrint, new Rectangle(0, 0, webBrowser1.Width, webBrowser1.Height));
            }

            // Draw the Bitmap onto the printed page
            e.Graphics.DrawImage(bitmapToPrint, e.MarginBounds);
        }




        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Check if the clicked cell is in the action column
            if (e.ColumnIndex == dataGridView1.Columns["ActionColumn"].Index && e.RowIndex >= 0)
            {

                textBox1.Text = "";
                textBox2.Text = "RUSTAMPURA SWIMMING POOL";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox5.Text = "";
                if (string.IsNullOrEmpty(comboBox5.SelectedItem?.ToString()))
                    comboBox5.Text = " ";
                else
                    comboBox5.SelectedText = " ";

                if (string.IsNullOrEmpty(comboBox6.SelectedItem?.ToString()))
                    comboBox6.Text = " ";
                else
                    comboBox6.SelectedText = " " +
                        "";
                //comboBox5.SelectedText = "";
                //comboBox6.SelectedText = "";
                // Perform the action based on the clicked button
                int rowIndex = e.RowIndex;
                DataGridViewRow selectedRow = dataGridView1.Rows[rowIndex];


                textBox1.Text = selectedRow.Cells[Data2].Value.ToString();
                textBox2.Text = selectedRow.Cells[Data06].Value.ToString();
                comboBox6.SelectedText = comboBox2.SelectedItem?.ToString(); 
                comboBox5.SelectedText = selectedRow.Cells[Data6].Value.ToString();
                textBox4.Text = selectedRow.Cells[Data5].Value.ToString();
                textBox5.Text = selectedRow.Cells[Data3].Value.ToString();
                // Example: Display a message with the row's data
                //string message = $"Action button clicked for row {rowIndex}, Data: {selectedRow.Cells["DataColumn"].Value}";
                //MessageBox.Show(message);
            }
        }


        string Place = "";
        private void button4_Click(object sender, EventArgs e)
        {
            

            
            foreach (RadioButton radioButton in groupBox1.Controls.OfType<RadioButton>())
            {
                // Check if the radio button is checked
                if (radioButton.Checked)
                {
                    Place = radioButton.Text;
                    // Use the selected value as needed
                    break; // Exit the loop after finding the selected radio button
                }
            }

            webBrowser2.Document.GetElementById("txt1").InnerText = textBox1.Text.ToString();
            webBrowser2.Document.GetElementById("txt2").InnerText = textBox2.Text.ToString();
            webBrowser2.Document.GetElementById("txt3").InnerText = string.IsNullOrEmpty(comboBox5.SelectedItem?.ToString()) ? comboBox5.Text.ToString() : comboBox5.SelectedItem?.ToString();
            webBrowser2.Document.GetElementById("txt4").InnerText = string.IsNullOrEmpty(comboBox6.SelectedItem?.ToString()) ? comboBox6.Text.ToString() : comboBox6.SelectedItem?.ToString();
            webBrowser2.Document.GetElementById("txt5").InnerText = Place;
            webBrowser2.Document.GetElementById("txt6").InnerText = textBox3.Text.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            StringBuilder crateTemplet = new StringBuilder();
            crateTemplet.Append("<html><HEAD><style>table{border-collapse: collapse}</style></HEAD><body>");
            crateTemplet.Append(competitonName);
            crateTemplet.Append("<table style=\"empty-cells: show\" border = \"1px solid black \">");
            DataTable pivotdt = new DataTable();
            pivotdt = (DataTable)dataGridView1.DataSource;
            crateTemplet.Append("<tr>");
            crateTemplet.Append("<th>&nbsp");
            crateTemplet.Append("</th>");
            foreach (var C in comboBox2.Items)
            {
                crateTemplet.Append("<th>");
                crateTemplet.Append(C);
                crateTemplet.Append("</th>");
            
            }
            crateTemplet.Append("</tr>");
           
                crateTemplet.Append("<tr>");
                foreach (var r in comboBox1.Items)
                {
                    crateTemplet.Append("<td>");
                    crateTemplet.Append(r);
                    crateTemplet.Append("</td>");
                    foreach (var C in comboBox2.Items)
                    {
                    crateTemplet.Append("<td>");
                    StringBuilder filterQuery = new StringBuilder();
                    filterQuery.Append(Data7 + " LIKE '%" + C + "%' AND " + Data6 + " LIKE '%" + r + "%' AND " + Data4 + " = 'Male'");
                    crateTemplet.Append(pivotdt.Select(filterQuery.ToString()).Count());
                    crateTemplet.Append("</td>");
                }
                crateTemplet.Append("</tr>");
                }
            crateTemplet.Append("</table>");






            crateTemplet.Append(competitonName);
            crateTemplet.Append("<table style=\"empty-cells: show\" border = \"1px solid black \">");
            DataTable pivotfemaildt = new DataTable();
            pivotfemaildt = (DataTable)dataGridView1.DataSource;
            crateTemplet.Append("<tr>");
            crateTemplet.Append("<th>&nbsp");
            crateTemplet.Append("</th>");
            foreach (var C in comboBox2.Items)
            {
                crateTemplet.Append("<th>");
                crateTemplet.Append(C);
                crateTemplet.Append("</th>");
                foreach (var r in comboBox1.Items)
                {

                }
            }
            crateTemplet.Append("</tr>");

            crateTemplet.Append("<tr>");
            foreach (var r in comboBox1.Items)
            {
                crateTemplet.Append("<td>");
                crateTemplet.Append(r);
                crateTemplet.Append("</td>");
                foreach (var C in comboBox2.Items)
                {
                    crateTemplet.Append("<td>");
                    StringBuilder filterQuery = new StringBuilder();
                    filterQuery.Append(Data7 + " LIKE '%" + C + "%' AND " + Data6 + " LIKE '%" + r + "%' AND " + Data4 + " = 'Female'");
                    crateTemplet.Append(pivotfemaildt.Select(filterQuery.ToString()).Count());
                    crateTemplet.Append("</td>");
                }
                crateTemplet.Append("</tr>");
            }
            crateTemplet.Append("</table>");
            crateTemplet.Append(" </body></html>");
            webBrowser1.DocumentText = crateTemplet.ToString();
          
        }

        int row_count;
        int NO = 1;
        private void button6_Click(object sender, EventArgs e)
        {
            //StringBuilder filterQuery = new StringBuilder();
            PrintDocument printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage_doc);

            // Show the print dialog to allow the user to select a printer
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                webBrowser2.ShowPageSetupDialog();
                webBrowser2.Print();
            }

            if (Place != "PARTICIPATE")
            {
                string filePath = "D:/Bhargav/SDAA/doc/SENIOR2025RESULT.xlsx";

            row_count++;
            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                
                    var sheetname = "SENIOR2025RESULT";

                    switch (comboBox3.SelectedItem?.ToString().ToLower())
                    {
                        case "male":
                            switch (comboBox1.SelectedItem?.ToString().ToLower())
                            {
                                case "group-1":
                                    sheetname = "GROUP1_MALE";
                                    break;
                                case "group-2":
                                    sheetname = "GROUP2_MALE";
                                    break;
                                case "group-3":
                                    sheetname = "GROUP3_MALE";
                                    break;
                                case "group-4":
                                    sheetname = "GROUP4_MALE";
                                    break;
                                case "senior":
                                    sheetname = "SENIOR_MALE";
                                    break;
                            }

                            break;
                        case "female":
                            switch (comboBox1.SelectedItem?.ToString().ToLower())
                            {
                                case "group-1":
                                    sheetname = "GROUP1_FEMALE";
                                    break;
                                case "group-2":
                                    sheetname = "GROUP2_FEMALE";
                                    break;
                                case "group-3":
                                    sheetname = "GROUP3_FEMALE";
                                    break;
                                case "group-4":
                                    sheetname = "GROUP4_FEMALE";
                                    break;
                                case "senior":
                                    sheetname = "SENIOR_FEMALE";
                                    break;
                            }
                            break;
                    }



                    //switch (comboBox3.SelectedItem?.ToString().ToLower())
                    //{
                    //    case "male":
                    //        switch (comboBox1.SelectedItem?.ToString().ToLower())
                    //        {
                    //            case "under 8":
                    //                sheetname = "under8_MALE";
                    //                break;
                    //            case "under 10":
                    //                sheetname = "under10_MALE";
                    //                break;
                    //            case "under 12":
                    //                sheetname = "under12_MALE";
                    //                break;
                    //            case "senior":
                    //                sheetname = "senior_MALE";
                    //                break;

                    //        }

                    //        break;
                    //    case "female":
                    //        switch (comboBox1.SelectedItem?.ToString().ToLower())
                    //        {
                    //            case "under 8":
                    //                sheetname = "under8_FEMALE";
                    //                break;
                    //            case "under 10":
                    //                sheetname = "under10_FEMALE";
                    //                break;
                    //            case "under 12":
                    //                sheetname = "under12_FEMALE";
                    //                break;
                    //            case "senior":
                    //                sheetname = "senior_FEMALE";
                    //                break;
                    //        }
                    //        break;
                    //}



                    //switch (comboBox3.SelectedItem?.ToString().ToLower())
                    //{
                    //    case "male":
                    //        switch (comboBox1.SelectedItem?.ToString().ToLower())
                    //        {
                    //            case "25Year-29Year":
                    //                sheetname = "25-29_MALE";
                    //                break;
                    //            case "30Year-34Year":
                    //                sheetname = "30-34_MALE";
                    //                break;
                    //            case "35Year-39Year":
                    //                sheetname = "35-39_MALE";
                    //                break;
                    //            case "40Year-44Year":
                    //                sheetname = "40-44_MALE";
                    //                break;
                    //            case "45Year-49Year":
                    //                sheetname = "45-49_MALE";
                    //                break;
                    //            case "50Year-54Year":
                    //                sheetname = "50-54_MALE";
                    //                break;
                    //            case "55Year-59Year":
                    //                sheetname = "55-59_MALE";
                    //                break;
                    //            case "60Year-64Year":
                    //                sheetname = "60-64_MALE";
                    //                break;
                    //            case "65Year-69Year":
                    //                sheetname = "65-69_MALE";
                    //                break;
                    //            case "70Year-74Year":
                    //                sheetname = "70-74_MALE";
                    //                break;
                    //            case "75Year-79Year":
                    //                sheetname = "75-79_MALE";
                    //                break;
                    //            case "80Year & ABOVE":
                    //                sheetname = "80_MALE";
                    //                break;

                    //        }

                    //        break;
                    //    case "female":
                    //        switch (comboBox1.SelectedItem?.ToString().ToLower())
                    //        {
                    //            case "25Year-29Year":
                    //                sheetname = "25-29_FEMALE";
                    //                break;
                    //            case "30Year-34Year":
                    //                sheetname = "30-34_FEMALE";
                    //                break;
                    //            case "35Year-39Year":
                    //                sheetname = "35-39_FEMALE";
                    //                break;
                    //            case "40Year-44Year":
                    //                sheetname = "40-44_FEMALE";
                    //                break;
                    //            case "45Year-49Year":
                    //                sheetname = "45-49_FEMALE";
                    //                break;
                    //            case "50Year-54Year":
                    //                sheetname = "50-54_FEMALE";
                    //                break;
                    //            case "55Year-59Year":
                    //                sheetname = "55-59_FEMALE";
                    //                break;
                    //            case "60Year-64Year":
                    //                sheetname = "60-64_FEMALE";
                    //                break;
                    //            case "65Year-69Year":
                    //                sheetname = "65-69_FEMALE";
                    //                break;
                    //            case "70Year-74Year":
                    //                sheetname = "70-74_FEMALE";
                    //                break;
                    //            case "75Year-79Year":
                    //                sheetname = "75-79_FEMALE";
                    //                break;
                    //            case "80Year & ABOVE":
                    //                sheetname = "80_FEMALE";
                    //                break;
                    //        }
                    //        break;
                    //}

                    var worksheet = package.Workbook.Worksheets[sheetname];

                    if (Place == "FIRST")
                    {
                        NO = 1;
                        worksheet.Cells["A" + row_count].Value = comboBox2.SelectedItem?.ToString() + "(" + comboBox1.SelectedItem?.ToString() + " - " + comboBox3.SelectedItem?.ToString() + ")";
                        worksheet.Cells["A" + row_count + ":H" + row_count].Merge = true;
                        row_count++;
                        worksheet.Cells["A" + row_count].Value = "NO";
                        worksheet.Cells["B" + row_count].Value = "PLAYERS NAME";
                        worksheet.Cells["C" + row_count].Value = "SFI ID";
                        worksheet.Cells["D" + row_count].Value = "GENDER";
                        worksheet.Cells["E" + row_count].Value = "BIRTHDAY";
                        worksheet.Cells["F" + row_count].Value = "UNIT";
                        worksheet.Cells["G" + row_count].Value = "TIME";
                        row_count++;
                    }
                    worksheet.Cells["A" + row_count].Value = NO.ToString();
                    worksheet.Cells["B" + row_count].Value = textBox1.Text.ToString();
                    worksheet.Cells["C" + row_count].Value = textBox5.Text.ToString();
                    worksheet.Cells["D" + row_count].Value = comboBox3.SelectedItem?.ToString();
                    worksheet.Cells["E" + row_count].Value = textBox4.Text.ToString();
                    worksheet.Cells["F" + row_count].Value = textBox2.Text.ToString();
                    worksheet.Cells["G" + row_count].Value = textBox3.Text.ToString();
                    NO++;


                    package.Save();
                }
            }
        }
        private Bitmap bitmapToPrints;
        private void PrintDocument_PrintPage_doc(object sender, PrintPageEventArgs e)
        {
            if (bitmapToPrints == null)
            {
                // Create a Bitmap to capture the content of the WebBrowser control

                bitmapToPrints = new Bitmap(webBrowser2.Width, webBrowser2.Height);
                webBrowser2.DrawToBitmap(bitmapToPrints, new Rectangle(0, 0, webBrowser2.Width, webBrowser2.Height));
            }

            // Draw the Bitmap onto the printed page
            e.Graphics.DrawImage(bitmapToPrints, e.MarginBounds);
        }

        private void button7_Click(object sender, EventArgs e)
        {
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Boolean Flagoprater = false;
            string selectbox1 = comboBox1.SelectedItem?.ToString(); // Get the selected value from the ComboBox
            string selectbox2 = comboBox2.SelectedItem?.ToString();
            string selectbox3 = comboBox3.SelectedItem?.ToString();
            string selectbox4 = comboBox4.SelectedItem?.ToString();
            StringBuilder filterQuery = new StringBuilder();
            if (!string.IsNullOrEmpty(selectbox1))
            {

                filterQuery.Append(Data6 + " LIKE '%" + selectbox1 + "%'");
                Flagoprater = true;
            }
            if (!string.IsNullOrEmpty(selectbox2))
            {
                if (Flagoprater)
                    filterQuery.Append(" AND ");
                filterQuery.Append(Data7 + " LIKE '%" + selectbox2 + "%'");
                Flagoprater = true;
            }
            if (!string.IsNullOrEmpty(selectbox3))
            {
                if (Flagoprater)
                    filterQuery.Append(" AND ");
                filterQuery.Append(Data4 + " = '" + selectbox3 + "'");
                Flagoprater = true;
            }
            if (!string.IsNullOrEmpty(selectbox4))
            {
                if (Flagoprater)
                    filterQuery.Append(" AND ");
                filterQuery.Append(Data06 + " LIKE '%" + selectbox4 + "%'");
                Flagoprater = true;
            }

            if (Flagoprater)
            {
                if (dataGridView1.DataSource != null)
                {
                    dataGridView1.DataSource = dt.Select(filterQuery.ToString()).CopyToDataTable();
                }
            }
            else
            {
                // If no value is selected, show all data
                dataGridView1.DataSource = dt;
            }


            StringBuilder crateTemplet = new StringBuilder();
            DataTable templetDt = new DataTable();
            DataTable temp = (DataTable)dataGridView1.DataSource;
            templetDt.Columns.Add(Data2);
            templetDt.Columns.Add(Data06);
            for (int i = 0; i < temp.Rows.Count; i++)
            {
                templetDt.Rows.Add(new string[] { temp.Rows[i][1].ToString(), temp.Rows[i][5].ToString() });
            }

            int headno = 1;
            crateTemplet.Append("<html><HEAD><style>table{border-collapse: collapse} tr {height : 32px}</style></HEAD><body>");
            crateTemplet.Append(competitonName);
            crateTemplet.Append("<H3 style=\"text-align: center\">" + selectbox2 + "(" + selectbox3 + ")</H3>");
            crateTemplet.Append("<H5 style=\"text-align: center\">" + selectbox1 + "                HEAT NO " + headno + "</H5>");
            crateTemplet.Append("<table style=\"empty-cells: show\" border = \"1px solid black \">");
            crateTemplet.Append("<tr>\r\n");
            crateTemplet.Append(" <th> Sr.No</th> <th > Player Name </th> <th width='10%'>LANE NO</th><th> MIN . </th><th> SEC . </th><th> MIC . </th><th width='50px'>REMARK</th>");
            //crateTemplet.Append(" <th> Sr.No</th> <th > Player Name </th> <th width='50px'>UNIT</th><th width='10%'>LANE NO</th><th> MIN . </th><th> SEC . </th><th> MIC . </th>");
            crateTemplet.Append("</tr>\r\n");
            // for (int i = 0; i < templetDt.Rows.Count; i++)
            // {
            int playerCount = 0;
            foreach (DataRow r in templetDt.Rows)
            {
                playerCount++;
                if (playerCount == 9)
                {
                    playerCount = 1;
                    headno++;
                    crateTemplet.Append("</table>");
                    crateTemplet.Append(competitonName);
                    crateTemplet.Append("<H3 style=\"text-align: center\">" + selectbox2 + "(" + selectbox3 + ")</H3>");
                    crateTemplet.Append("<H5 style=\"text-align: center\">" + selectbox1 + "                HEAT NO " + headno + "</H5>");
                    crateTemplet.Append("<table style = \"empty-cells: show\" border = \"1px solid black \">");
                    crateTemplet.Append("<tr>\r\n");
                    crateTemplet.Append(" <th> Sr.No</th> <th> Player Name </th> <th width='50px'>UNIT</th><th width='10%'> LANE NO </th><th> MIN. </th><th> SEC. </th><th> MIC. </th>");
                    crateTemplet.Append("</tr>\r\n");
                }
                crateTemplet.Append("<tr>");
                crateTemplet.Append("<td>" + playerCount + "</td>");
                foreach (DataColumn c in templetDt.Columns)
                {
                    crateTemplet.Append("<td >");
                    crateTemplet.Append(r[c.ColumnName]);
                    crateTemplet.Append("</td>");

                    //crateTemplet.Append("<td  width='50px'>");
                    //crateTemplet.Append(r["Unit"]);
                    //crateTemplet.Append("</td>");
                }
                crateTemplet.Append("<td width='10%'>&nbsp</td>");
                crateTemplet.Append("<td>&nbsp</td>");
                crateTemplet.Append("<td>&nbsp</td>");
                crateTemplet.Append("<td width='50px'>&nbsp</td>");
                crateTemplet.Append("</tr>");



            }
            //}
            crateTemplet.Append("</table>");


            crateTemplet.Append(" </body></html>");

            webBrowser1.DocumentText = crateTemplet.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
