using Cms.WebPublisher.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cms.WebPublisher.DB
{
    public class WebPublisherRepository
    {
        private string _connString;

        public WebPublisherRepository(string connString)
        {
            _connString = connString;
        }

        public GenericCellData GetLatestPositions(string sheetName)
        {
            string spName = "dbo.GetLatestPositions";
            using (var conn = new SqlConnection(_connString))
            {
                try
                {
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = spName;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 300;
                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = sheetName
                    });

                    GenericCellData cellData = null;
                    var reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        var bytes = (byte[])reader["data"];
                        var json = Encoding.UTF8.GetString(bytes);
                        cellData = JsonConvert.DeserializeObject<GenericCellData>(json);

                        cellData.SheetName = reader["sheetName"].ToString();
                        cellData.Timestamp = (DateTime)reader["publishTime"];
                        cellData.Rows = (int)reader["rows"];
                        cellData.Columns = (int)reader["columns"];
                    }

                    if (cellData == null)
                        return new GenericCellData
                        {
                            SheetName = $"[{sheetName}] Not Found"
                        };

                    return cellData;
                }
                catch
                {
                    throw;
                }
                finally
                {
                    conn.Close();
                }
            }
        }
    }
}
