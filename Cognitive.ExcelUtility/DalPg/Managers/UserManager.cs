using Cognitive.ExcelUtility.DalPg.Base;
using Cognitive.ExcelUtility.DalPg.Entities;

namespace Cognitive.ExcelUtility.DalPg.Managers
{
    public class UserManager : PostgresManagerBase
    {
        public async Task<List<UserEntity>> ReadAll()
        {
            return await ExecuteCursorFunction("cognitive.f$users__read_all", async reader => new UserEntity
            {
                UserId = await reader.GetFieldValueAsync<long?>("user_id"),
                Age = await reader.GetFieldValueAsync<short?>("age"),
                Education = await reader.GetFieldValueAsync<string?>("education"),
                Speciality = await reader.GetFieldValueAsync<string?>("speciality"),
                Residence = await reader.GetFieldValueAsync<string?>("residence"),
                Height = await reader.GetFieldValueAsync<short?>("height"),
                Weight = await reader.GetFieldValueAsync<short?>("weight"),
                LeadHand = await reader.GetFieldValueAsync<string?>("lead_hand"),
                Diseases = await reader.GetFieldValueAsync<string?>("diseases"),
                Smoking = await reader.GetFieldValueAsync<bool?>("smoking"),
                Alcohol = await reader.GetFieldValueAsync<string?>("alcohol"),
                Sport = await reader.GetFieldValueAsync<string?>("sport"),
                Insomnia = await reader.GetFieldValueAsync<bool?>("insomnia"),
                CurrentHealth = await reader.GetFieldValueAsync<short?>("current_health"),
                Gaming = await reader.GetFieldValueAsync<bool?>("gaming")
            });
        }
    }
}
