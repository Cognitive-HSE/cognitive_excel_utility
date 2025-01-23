using Npgsql;

namespace Cognitive.ExcelUtility.DalPg.Base
{
    public static class DbHelper
    {
        //todo: спрятать строку подключения
        private static NpgsqlDataSource? DataSource { get; set; }

        public static void BuildDataSource(bool devDb = false)
        {
            var db = devDb ? "cognitive_dev" : "cognitive";
            DataSource = NpgsqlDataSource.Create($"Host=79.137.204.140;Port=5000;Username=cognitive_excel_utility;Password=cognitive_excel_utility;Database={db}");
        }

        public static async ValueTask<NpgsqlConnection> CreateOpenedConnectionAsync() => DataSource != null
            ? await DataSource.OpenConnectionAsync()
            : throw new NullReferenceException("DataSource is null");
    }
}
