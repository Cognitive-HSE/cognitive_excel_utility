using System.Globalization;
using System.Text;
using Cognitive.ExcelUtility.DalPg.Entities;
using Cognitive.ExcelUtility.DalPg.Managers;

namespace Cognitive.ExcelUtility
{
    public static class CsvBuilder
    {
        public static string BuildDataset()
        {
            var sb = new StringBuilder();

            sb.AppendLine(BuildHeader());

            var tests = GetTestResults().Result;
            var users = GetUsers().Result;

            foreach (var test in tests)
            {
                sb.AppendLine(BuildLine(users.FirstOrDefault(el => el.UserId == test.UserId), test));
            }

            return sb.ToString();
        }

        private static string BuildHeader()
        {
            string[] headers =
            [
                "user_id",
                "age",
                "education",
                "speciality",
                "residence",
                "height",
                "weight",
                "lead_hand",
                "diseases",
                "smoking",
                "alcohol",
                "sport",
                "insomnia",
                "current_health",
                "gaming",
                "test_id",
                "try_number",
                "accuracy",
                "complete_time",
                "number_correct_answers",
                "number_all_answers"
            ];
            return string.Join(';', headers);
        }

        private static string BuildLine(UserEntity? user, TestResultEntity test)
        {
            string?[] line =
            [
                user?.UserId.ToString(),
                user?.Age.ToString(),
                user?.Education,
                user?.Speciality,
                user?.Residence,
                user?.Height.ToString(),
                user?.Weight.ToString(),
                user?.LeadHand,
                user?.Diseases,
                user?.Smoking.ToCsvFormat(),
                user?.Alcohol,
                user?.Sport,
                user?.Insomnia.ToCsvFormat(),
                user?.CurrentHealth.ToString(),
                user?.Gaming.ToCsvFormat(),
                test.TestId.ToString(),
                test.TryNumber.ToString(),
                test.Accuracy.ToString(),
                test.CompleteTime.ToCsvFormat(),
                test.NumberCorrectAnswers.ToString(),
                test.NumberAllAnswers.ToString()
            ];

            return string.Join(';', line).ToString(CultureInfo.InvariantCulture);
        }

        private static async Task<List<TestResultEntity>> GetTestResults()
        {
            return await new TestResultManager().ReadAll();
        }
        private static async Task<List<UserEntity>> GetUsers()
        {
            return await new UserManager().ReadAll();
        }

        private static string ToCsvFormat(this bool? val) => val != null && val.Value ? "1" : val != null ? "0" : "";

        private static string ToCsvFormat(this TimeSpan? val) => val != null ? val.Value.TotalSeconds.ToString(CultureInfo.InvariantCulture) : "";
    }
}
