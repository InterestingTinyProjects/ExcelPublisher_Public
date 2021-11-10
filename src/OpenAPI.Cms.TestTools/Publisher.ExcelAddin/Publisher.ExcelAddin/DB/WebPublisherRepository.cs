using Newtonsoft.Json;
using OpenApi.Cms.TestTools.Client.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenApi.Cms.TestTools.Client.DB
{
    public class WebPublisherRepository
    {
        private string _connString;

        public WebPublisherRepository(string connString)
        {
            _connString = connString;
        }

        public bool TestDBCOnnection()
        {
            using (var conn = new SqlConnection(_connString))
            {
                try
                {
                    conn.Open();
                    return true;
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

       
        public int PublishPositions(GenericCellData cellData)
        {
            string spName = "dbo.PublishPositions";
            using (var conn = new SqlConnection(_connString))
            {
                try
                {                    
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = spName;
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandTimeout = 300;
                    command.Parameters.Add(new SqlParameter("@publishTime", SqlDbType.DateTime2)
                    {
                        Value = cellData.Timestamp
                    });

                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = cellData.SheetName
                    });

                    command.Parameters.Add(new SqlParameter("@rows", SqlDbType.Int)
                    {
                        Value = cellData.Rows
                    });

                    command.Parameters.Add(new SqlParameter("@columns", SqlDbType.Int)
                    {
                        Value = cellData.Columns
                    });

                    var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(cellData.Data));
                    command.Parameters.Add(new SqlParameter("@data", SqlDbType.VarBinary)
                    {
                        Value = bytes
                    });

                    return (int)command.ExecuteScalar();
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
