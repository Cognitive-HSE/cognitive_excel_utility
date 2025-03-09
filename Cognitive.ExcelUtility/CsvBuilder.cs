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
            {
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
            };
            return string.Join(',', headers);
        }

        private static string BuildLine(UserEntity? user, TestResultEntity test)
        {
            string?[] line =
            {
                user?.UserId.ToCsvFormat(),
                user?.Age.ToCsvFormat(),
                user?.Education.ToCsvFormat(),
                user?.Speciality.ToCsvFormat(),
                user?.Residence.ToCsvFormat(),
                user?.Height.ToCsvFormat(),
                user?.Weight.ToCsvFormat(),
                user?.LeadHand.ToCsvFormat(),
                user?.Diseases.ToCsvFormat(),
                user?.Smoking.ToCsvFormat(),
                user?.Alcohol.ToCsvFormat(),
                user?.Sport.ToCsvFormat(),
                user?.Insomnia.ToCsvFormat(),
                user?.CurrentHealth.ToCsvFormat(),
                user?.Gaming.ToCsvFormat(),
                test.TestId.ToCsvFormat(),
                test.TryNumber.ToCsvFormat(),
                test.Accuracy.ToCsvFormat(),
                test.CompleteTime.ToCsvFormat(),
                test.NumberCorrectAnswers.ToCsvFormat(),
                test.NumberAllAnswers.ToCsvFormat()
            };

            return string.Join(',', line).ToString(CultureInfo.InvariantCulture);
        }

        private static async Task<List<TestResultEntity>> GetTestResults()
        {
            return await new TestResultManager().ReadAll();
        }

        private static async Task<List<UserEntity>> GetUsers()
        {
            return await new UserManager().ReadAll();
        }

        private static string ToCsvFormat(this object? val)
        {
            if (val == null) return "-";

            if (val is bool b) return b ? "1" : "0";

            if (val is TimeSpan ts) return ts.TotalSeconds.ToString(CultureInfo.InvariantCulture);

            if (val is string s) return string.IsNullOrEmpty(s) ? "-" : $"\"{s}\"";

            if (val is short sh) return sh.ToString();
            if (val is int i) return i.ToString();
            if (val is long l) return l.ToString();
            if (val is decimal d) return d.ToString(CultureInfo.InvariantCulture);

            return "-";
        }
    }
}
