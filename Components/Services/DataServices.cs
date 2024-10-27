using System.Data;
using System.Dynamic;
using Microsoft.Data.SqlClient;

namespace importacionBono.Components.Services;
public class DataService
{
    private readonly IConfiguration _configuration;

    public DataService(IConfiguration configuration) => _configuration = configuration;

    public async Task<List<ExpandoObject>> GetDataAsync(string query)
    {
        try
        {
            var connectionString = _configuration.GetSection("ConnectionStrings:BusinessConnection").Value;
            var resultList = new List<ExpandoObject>();

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                await con.OpenAsync();
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.CommandType = CommandType.Text;
                cmd.CommandTimeout = 0;

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        dynamic expandoObject = new ExpandoObject();
                        var dataRow = expandoObject as IDictionary<string, object>;

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            string columnName = reader.GetName(i);
                            object value = reader.IsDBNull(i) ? null : reader.GetValue(i);
                            dataRow.Add(columnName, value);
                        }

                        resultList.Add(expandoObject);
                    }
                }
            }

            return resultList;
        }
        catch (Exception w)
        {
            Console.WriteLine("error GetDataAsync:" + w);
            return null;
        }
    }


    public async Task<bool> CrudAsync(string query, SqlParameter[] parametrosSP)
    {
        try
        {
            var connectionString = _configuration.GetSection("ConnectionStrings:BusinessConnection").Value;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = connection.CreateCommand())
                {
                    await connection.OpenAsync();
                    cmd.CommandText = query;
                    foreach (SqlParameter parSP in parametrosSP) cmd.Parameters.Add(parSP);
                    var rows = await cmd.ExecuteNonQueryAsync();
                    connection.Close();
                    return rows > 0 ? true : false;
                }
            }
        }
        catch (Exception w)
        {
            Console.WriteLine("error CrudAsync:" + w);
            return false;
        }
    }


    public SqlParameter ParameterSql(string Nombre, SqlDbType tipo, Object valor)
    {
        SqlParameter sqlParametro = new SqlParameter(Nombre, tipo);
        sqlParametro.Value = valor;
        return sqlParametro;
    }

}