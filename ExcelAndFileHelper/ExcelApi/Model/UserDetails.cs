using Newtonsoft.Json;
using System.Data;

namespace ExcelApi.Model
{
    
    public class UserDetails
    {
        public string? ID { get; set; }
        public string? Name { get; set; }
        public string? City { get; set; }
        public string? Country { get; set; }

        public List<UserDetails> GetEmployeeDummyData()
        {
            return new List<UserDetails> {
                new UserDetails { ID = "1001", Name = "ABCD", City = "City1", Country = "USA" },
                new UserDetails { ID = "1002", Name = "PQRS", City = "City2", Country = "INDIA" },
                new UserDetails { ID = "1003", Name = "XYZZ", City = "City3", Country = "CHINA" },
                new UserDetails { ID = "1004", Name = "LMNO", City = "City4", Country = "UK" },
            };
        }

        public List<List<UserDetails>> GetEmployeeDummyDataWithMultiList()
        {
            List<UserDetails> employees = new List<UserDetails> {
                new UserDetails { ID = "1001", Name = "ABCD", City = "City1", Country = "USA" },
                new UserDetails { ID = "1002", Name = "PQRS", City = "City2", Country = "INDIA" },
                new UserDetails { ID = "1003", Name = "XYZZ", City = "City3", Country = "CHINA" },
                new UserDetails { ID = "1004", Name = "LMNO", City = "City4", Country = "UK" },
            };
            List<UserDetails> employees1 = new List<UserDetails> {
                new UserDetails { ID = "1005", Name = "ABC1", City = "City5", Country = "USA" },
                new UserDetails { ID = "1006", Name = "PQR2", City = "City6", Country = "INDIA" },
                new UserDetails { ID = "1007", Name = "XYZ3", City = "City7", Country = "CHINA" },
                new UserDetails { ID = "1008", Name = "LMN4", City = "City8", Country = "UK" },
            };

            List<List<UserDetails>> emp = new List<List<UserDetails>>();
            emp.Add(employees);
            emp.Add(employees1);

            return emp;
        }

        public DataTable ConvertModelToDataTable(dynamic model)
        {
            return JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(model));
        }

        public DataSet ConvertModelToDataSet(dynamic model)
        {
            DataSet ds = new DataSet();
            foreach (var modelItem in model)
            {
                ds.Tables.Add(JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(modelItem)));
            }

            return ds;
        }
    }
}
