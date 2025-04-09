using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tools_LayHinh
{
    public partial class Form1 : Form
    {
        DataSet ds;
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_import_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls", ValidateNames = true })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                        {
                            IExcelDataReader reader;
                            if (ofd.FilterIndex == 2)
                            {
                                reader = ExcelReaderFactory.CreateBinaryReader(stream);
                            }
                            else
                            {
                                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                            }

                            ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                {
                                    UseHeaderRow = true
                                }
                            });

                            cb_sheet.Items.Clear();
                            foreach (DataTable dt in ds.Tables)
                            {
                                cb_sheet.Items.Add(dt.TableName);
                            }
                            reader.Close();

                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error 404");
            }
        }

        private void cb_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DataSource = ds.Tables[cb_sheet.SelectedIndex];
            }
            catch (Exception)
            {
                MessageBox.Show("Error 8080000x");
            }
        }

        private void btn_exportimage_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    SaveFileDialog savefile = new SaveFileDialog();
            //    savefile.FileName = "Book.xls";
            //    //savefile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //    savefile.Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls";
            //    if (ds.Tables[0].Rows.Count > 0)
            //    {
            //        if (savefile.ShowDialog() == DialogResult.OK)
            //        {
            //            StreamWriter wr = new StreamWriter(savefile.FileName);
            //            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
            //            {
            //                wr.Write(ds.Tables[0].Columns[i].ToString().ToUpper() + "\t");
            //            }

            //            wr.WriteLine();

            //            //write rows to excel file
            //            for (int i = 0; i < (ds.Tables[0].Rows.Count); i++)
            //            {
            //                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
            //                {
            //                    if (ds.Tables[0].Rows[i][j] != null)
            //                    {
            //                        wr.Write(Convert.ToString(ds.Tables[0].Rows[i][j]) + "\t");
            //                    }
            //                    else
            //                    {
            //                        wr.Write("\t");
            //                    }
            //                }
            //                //go to next line
            //                wr.WriteLine();
            //            }
            //            //close file
            //            wr.Close();
            //            MessageBox.Show(this, "Data saved in Excel format at location " + savefile.FileName, "Successfully Saved", MessageBoxButtons.OK, MessageBoxIcon.Question);
            //        }
            //    }
            //    else
            //    {
            //        MessageBox.Show(this, "Zero record to export , perform a operation first", "Can't export file", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }



            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Error 404");
            //}



            try
            {
                SaveFileDialog savefile = new SaveFileDialog();
                savefile.FileName = "Book.xls";
                savefile.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                //savefile.Filter = "Excel Workbook|*.xlsx|Excel Workbook 97-2003|*.xls";
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (savefile.ShowDialog() == DialogResult.OK)
                    {
                        StreamWriter wr = new StreamWriter(savefile.FileName);
                        for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                        {
                            wr.Write(ds.Tables[0].Columns[i].ToString().ToUpper() + "\t");
                        }

                        wr.WriteLine();

                        //write rows to excel file
                        for (int i = 0; i < (ds.Tables[0].Rows.Count); i++)
                        {
                            for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                            {
                                if (ds.Tables[0].Rows[i][j] != null)
                                {
                                    wr.Write(Convert.ToString(ds.Tables[0].Rows[i][j]) + "\t");
                                }
                                else
                                {
                                    wr.Write("\t");
                                }
                            }
                            //go to next line
                            wr.WriteLine();
                        }
                        //close file
                        wr.Close();
                        MessageBox.Show(this, "Data saved in Excel format at location " + savefile.FileName, "Successfully Saved", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    }
                }
                else
                {
                    MessageBox.Show(this, "Zero record to export , perform a operation first", "Can't export file", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }



            }
            catch (Exception)
            {
                MessageBox.Show("Error 404");
            }

        }
    }
}
