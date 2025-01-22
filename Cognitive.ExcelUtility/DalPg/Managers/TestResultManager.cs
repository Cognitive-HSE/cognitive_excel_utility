using Cognitive.ExcelUtility.DalPg.Base;
using Cognitive.ExcelUtility.DalPg.Entities;

namespace Cognitive.ExcelUtility.DalPg.Managers
{
    public class TestResultManager : PostgresManagerBase
    {
        public async Task<List<TestResultEntity>> ReadAll()
        {
            return await ExecuteCursorFunction("cognitive.f$test_results__read_all", async reader => new TestResultEntity
            {
                UserId = await reader.GetFieldValueAsync<long?>("user_id"),
                TestId = await reader.GetFieldValueAsync<long?>("test_id"),
                TryNumber = await reader.GetFieldValueAsync<short?>("try_number"),
                Accuracy = await reader.GetFieldValueAsync<decimal?>("accuracy"),
                CompleteTime = await reader.GetFieldValueAsync<TimeSpan?>("complete_time"),
                NumberCorrectAnswers = await reader.GetFieldValueAsync<int?>("number_correct_answers"),
                NumberAllAnswers = await reader.GetFieldValueAsync<int?>("number_all_answers"),
            });
        }
    }
}
