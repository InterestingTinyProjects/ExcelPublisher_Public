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

        public int TestDBConnection(string reportName)
        {
            using (var conn = new SqlConnection(_connString))
            {
                try
                {
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = @"SELECT COUNT(row_id) FROM [dbo].[Record] WHERE sheetName=@sheetName";
                    command.Parameters.Add(new SqlParameter("@sheetName", SqlDbType.NVarChar)
                    {
                        Value = reportName
                    });
                    command.CommandType = CommandType.Text;
                    command.CommandTimeout = 300;
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

                    var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(cellData));
                    command.Parameters.Add(new SqlParameter("@data", SqlDbType.VarBinary)
                    {
                        Value = bytes
                    });

                    if (cellData.MaxStoredRecords.HasValue)
                    {
                        command.Parameters.Add(new SqlParameter("@maxStoredRecords", SqlDbType.BigInt)
                        {
                            Value = cellData.MaxStoredRecords
                        });
                    }

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
