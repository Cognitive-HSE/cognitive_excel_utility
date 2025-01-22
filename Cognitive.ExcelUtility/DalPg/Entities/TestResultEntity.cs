namespace Cognitive.ExcelUtility.DalPg.Entities
{
    public class TestResultEntity
    {
        public long? UserId { get; set; }

        public long? TestId { get; set; }

        public short? TryNumber { get; set; }

        public decimal? Accuracy { get; set; }

        public TimeSpan? CompleteTime { get; set; }

        public int? NumberCorrectAnswers { get; set; }

        public int? NumberAllAnswers { get; set; }
    }
}
