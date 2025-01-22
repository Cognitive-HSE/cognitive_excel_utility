using Npgsql;

namespace Cognitive.ExcelUtility.DalPg.Base
{
    public static class PostgresExtensions
    {
        public static T GetFieldValue<T>(this NpgsqlDataReader reader, string fieldName)
        {
            var ordinal = reader.GetOrdinal(fieldName);
            return reader.GetFieldValue<T>(ordinal);
        }

        public static async Task<T?> GetFieldValueAsync<T>(this NpgsqlDataReader reader, string fieldName)
        {
            var ordinal = reader.GetOrdinal(fieldName);
            if (reader.IsDBNull(ordinal))
                return default;
            return await reader.GetFieldValueAsync<T>(ordinal);
        }
    }
}
