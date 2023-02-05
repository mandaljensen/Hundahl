dotnet
{
    assembly("System.Data")
    {
        Version = '4.0.0.0';
        type("System.Data.SqlClient.SqlConnection"; NewSqlConnection) { }
        type("System.Data.SqlClient.SqlCommand"; NewSqlCommand) { }
        type("System.Data.SqlClient.SqlParameter"; NewSqlParameter) { }
        type("System.Data.SqlClient.SqlDataReader"; NewSqlDataReader) { }
        type("System.Data.CommandType"; CommandType) { }
        type("System.Data.SqlDbType"; newSqlDbType) { }
    }
}