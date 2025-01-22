using Cognitive.ExcelUtility.DalPg.Base;
using Cognitive.ExcelUtility.DalPg.Entities;

namespace Cognitive.ExcelUtility.DalPg.Managers
{
    public class TestNsiManager : PostgresManagerBase
    {
        public async Task<List<TestNsiEntity>> ReadAll()
        {
            return await ExecuteCursorFunction("cognitive.f$test_nsi__read_all", async reader => new TestNsiEntity
            {
                TestId = await reader.GetFieldValueAsync<long?>("test_id"),
                TestName = await reader.GetFieldValueAsync<string?>("test_name"),
                TitleAll = await reader.GetFieldValueAsync<string?>("title_all"),
                TitleCorrect = await reader.GetFieldValueAsync<string?>("title_correct"),
            });
        }
    }
}
