using Npgsql;
using NpgsqlTypes;
using System.Data;
using System.Reflection;
using System.Text;

namespace Cognitive.ExcelUtility.DalPg.Base
{
    /// <summary>
    /// Родитель для всех менеджеров
    /// </summary>
    public abstract class PostgresManagerBase
    {
        /// <summary>
        /// Выполнить функцию, возвращающую курсор
        /// </summary>
        /// <typeparam name="T">Cущность, в которую будет мапиться результат</typeparam>
        /// <param name="funcName">Имя функции из <see cref="PgNames"/></param>
        /// <param name="mapper">Функция-мапер/></param>
        /// <param name="parameters">Входные параметры функции. ПОРЯДОК ВАЖЕН!</param>
        /// <returns></returns>
        protected static async Task<List<T>> ExecuteCursorFunction<T>(string funcName,
            Func<NpgsqlDataReader, Task<T>> mapper,
            params (string Name, object? Value, NpgsqlDbType Type)[] parameters)
        {
            var query = $"select {funcName}{CreateParametersQuery(parameters)}";
            var result = new List<T>();

            var connection = await DbHelper.CreateOpenedConnectionAsync();

            // Делаем все в транзакции
            await using var transaction = await connection.BeginTransactionAsync(IsolationLevel.ReadCommitted);

            var cursorValue = "";
            // Вызываем функцию
            await using (var cmd = new NpgsqlCommand(query, connection, transaction))
            {
                cmd.Parameters.AddRange(CreateParameters(parameters));
                await using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    cursorValue = reader.GetValue(0).ToString();
                }
                await reader.CloseAsync();
            }

            // Извлекаем данные из курсора и мапим
            await using (var cmd = new NpgsqlCommand($"fetch all in \"{cursorValue}\"", connection, transaction))
            {
                await using var reader = await cmd.ExecuteReaderAsync();
                while (reader.Read())
                {
                    result.Add(await mapper(reader));
                }
                await reader.CloseAsync();
            }

            await transaction.CommitAsync();
            return result;
        }

        /// <summary>
        /// Выполнить функцию, возвращающую простой тип
        /// </summary>
        /// <typeparam name="T">Тип возвращаемого значения</typeparam>
        /// <param name="funcName">Имя функции из <see cref="PgNames"/></param>
        /// <param name="parameters">Параметры функции</param>
        /// <returns></returns>
        protected static async Task<T> ExecuteFunction<T>(string funcName,
            params (string Name, object? Value, NpgsqlDbType Type)[] parameters)
        {
            var query = $"select {funcName}{CreateParametersQuery(parameters)}";
            var connection = await DbHelper.CreateOpenedConnectionAsync();
            await using var cmd = new NpgsqlCommand(query, connection);
            cmd.Parameters.AddRange(CreateParameters(parameters));
            await using var reader = await cmd.ExecuteReaderAsync();

            T res = default!;

            while (await reader.ReadAsync())
            {
                res = reader.GetFieldValue<T>(0);
            }

            return res;
        }

        protected static async Task<List<T>> SelectView<T>()
        {
            var entityType = typeof(T);
            var viewName = entityType.GetCustomAttribute<ViewNameAttribute>()?.ViewName;
            if (viewName == null)
                throw new NullReferenceException(
                    $"Для сущности {entityType.Name} не задано название вью {nameof(ViewNameAttribute)}");

            var props = entityType.GetProperties();

            var query = $"select * from {viewName}";
            var connection = await DbHelper.CreateOpenedConnectionAsync();
            await using var cmd = new NpgsqlCommand(query, connection);
            await using var reader = await cmd.ExecuteReaderAsync();

            var res = new List<T>();

            while (await reader.ReadAsync())
            {
                var ent = Activator.CreateInstance<T>();
                foreach (var p in props)
                {
                    var attr = p.GetCustomAttribute<ViewColumnAttribute>();
                    if (attr == null)
                        throw new NullReferenceException(
                            $"Для поля {entityType.Name}.{p.Name} не задан атрибут {nameof(ViewColumnAttribute)}");

                    var colName = attr.ColumnName;
                    var setMethod = p.SetMethod;

                    if (setMethod == null)
                        throw new NullReferenceException(
                            $"Для поля {entityType.Name}.{p.Name} не задан сеттер");

                    var value = reader.GetValue(reader.GetOrdinal(colName));
                    setMethod.Invoke(ent, [value is DBNull ? null : value]);
                }
                res.Add(ent);
            }

            return res;
        }

        #region private

        private static NpgsqlParameter[] CreateParameters((string Name, object? Value, NpgsqlDbType Type)[] parameters)
        {
            return parameters.Select(el => new NpgsqlParameter { ParameterName = el.Name, Value = el.Value ?? DBNull.Value, NpgsqlDbType = el.Type }).ToArray();
        }

        private static string CreateParametersQuery((string Name, object? Value, NpgsqlDbType Type)[] parameters)
        {
            if (parameters.Length == 0)
                return "()";

            var res = new StringBuilder("(");

            for (var i = 0; i < parameters.Length; ++i)
            {
                res.Append($"@{parameters[i].Name}");
                res.Append(i != parameters.Length - 1 ? "," : "");
            }
            res.Append(')');

            return res.ToString();
        }

        #endregion
    }
}
