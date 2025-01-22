using Npgsql;

namespace Cognitive.ExcelUtility.DalPg.Base
{
    public static class DbHelper
    {
        //todo: спрятать строку подключения
        private static NpgsqlDataSource DataSource { get; set; } = NpgsqlDataSource.Create("Host=79.137.204.140;Port=5000;Username=cognitive_excel_utility;Password=cognitive_excel_utility;Database=cognitive_dev");

        public static async ValueTask<NpgsqlConnection> CreateOpenedConnectionAsync() => await DataSource.OpenConnectionAsync();
    }
}
