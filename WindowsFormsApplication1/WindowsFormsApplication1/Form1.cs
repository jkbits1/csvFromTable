using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Data.Common;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");

            var mdbPath = @"";
            var mdbFileName = "x.mdb";            
            var connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath + mdbFileName;

            var origTableName = "x";
            var csvFileName = "x";

            using (var file =
                new System.IO.StreamWriter(mdbPath + csvFileName + ".csv"))
            {
                using (DbConnection conn = factory.CreateConnection())
                {
                    conn.ConnectionString = connString;

                    conn.Open();

                    DbCommand comm = conn.CreateCommand();

                    comm.CommandText = "Select * from " + origTableName;

                    DbDataReader reader = comm.ExecuteReader(CommandBehavior.Default);

                    DataTable dt = new DataTable(csvFileName);

                    dt.Load(reader);

                    StringBuilder sbColumns = new StringBuilder();

                    foreach (DataColumn dc in dt.Columns)
                    {                   
                        if (!String.IsNullOrEmpty(sbColumns.ToString()))
                        {
                            sbColumns.Append(",");
                        }

                        sbColumns.Append(dc.ColumnName.ToString());
                    }

                    file.WriteLine(sbColumns.ToString());

                    foreach (DataRow row in dt.Rows)
                    {
                        StringBuilder sb = new StringBuilder();

                        foreach (DataColumn dc in row.Table.Columns)
                        {
                            if (!String.IsNullOrEmpty(sb.ToString()))
                            {
                                sb.Append(",");
                            }

                            sb.Append(row[dc.ColumnName].ToString());
                        }

                        //sb.Append("\n");

                        file.WriteLine(sb.ToString());
                    }
                }
            }       
        }
    }
}
