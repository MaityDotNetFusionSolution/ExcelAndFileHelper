using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelLib.CsvHelper
{
    public class CsvHelper
    {
        public DataTable ReadCsv(Stream stream)
        {
            DataTable dt;
            //Regex for handing csv and special characters
            Regex regex = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            try
            {
                dt = new DataTable();
                using (StreamReader sr = new StreamReader(stream))
                {
                    //collecting header value
                    List<string> headerData = sr.ReadLine().Split(',').ToList();

                    //Add column  for heading
                    headerData.ForEach(x => dt.Columns.Add(x));

                    //Add Rows for inserting data
                    while(!sr.EndOfStream)
                    {
                        string[] rows = regex.Split(sr.ReadLine());

                        DataRow dr = dt.NewRow();
                        for(int i = 0; i<headerData.Count;i++)
                        {
                            dr[i] = rows[i].Replace("\"", string.Empty);
                        }
                        dt.Rows.Add(dr);
                    }
                }
            }catch (Exception ex)
            {
                throw;
            }
            return dt;
        }

        public byte[] CreateCsv(DataTable table)
        {
            StringBuilder sb = new StringBuilder();

            byte[] result;
            try
            {
                if(table != null && table.Rows.Count>0)
                {
                    //Header Row
                    int colLength = table.Columns.Count;
                    for(int i = 0;i< colLength;i++)
                    {
                        sb.Append(table.Columns[i].ColumnName.ToString());

                        if(i < table.Columns.Count -1)
                        {
                            sb.Append(",");
                        }
                    }
                    sb.Append("\n");

                    //Row part
                    foreach(DataRow row in table.Rows) 
                    { 
                        for(int colIndex = 0; colIndex < colLength ; colIndex++)
                        {
                            sb.Append(row[colIndex].ToString());
                            if(colIndex < colLength - 1)
                                sb.Append(',');
                        }
                        sb.Append("\n");
                    }

                }
                result = Encoding.ASCII.GetBytes(sb.ToString());
            }catch(Exception )
            {
                throw;
            }
            return result;
        }
    }
}
