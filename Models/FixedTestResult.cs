using Newtonsoft.Json;
using System.Collections.Generic;

namespace EFsetWidgetFix.Models
{
    public class FixedTestResult
    {
        [JsonProperty(PropertyName = "test_results")]
        public List<TestResultF> TestResults { get; set; }
    }

    public class TestResultF
    {
        [JsonProperty(PropertyName = "test_name")]
        public string TestName { get; set; }

        [JsonProperty(PropertyName = "scores")]
        public Score Score { get; set; }

        [JsonProperty(PropertyName = "test_start_time")]
        public string TestStartTime { get; set; }

        [JsonProperty(PropertyName = "test_finish_time")]
        public string TestFinishTime { get; set; }

        [JsonProperty(PropertyName = "status")]
        public string Status { get; set; }
    }

    public class Score
    {
        [JsonProperty(PropertyName = "level")]
        public string Level { get; set; }

        [JsonProperty(PropertyName = "raw_score")]
        public string Raw_Score { get; set; }

        [JsonProperty(PropertyName = "max_raw_score")]
        public string Max_Raw_Score { get; set; }

        [JsonProperty(PropertyName = "score")]
        public string Actual_Score { get; set; }

        [JsonProperty(PropertyName = "max_score")]
        public string Actual_Max_Score { get; set; }

        [JsonProperty(PropertyName = "reading")]
        public Reading Reading { get; set; }

        [JsonProperty(PropertyName = "listening")]
        public Listening Listening { get; set; }
    }

    public class Reading
    {
        [JsonProperty(PropertyName = "score")]
        public string Score { get; set; }

        [JsonProperty(PropertyName = "max_score")]
        public string Max_Score { get; set; }
    }

    public class Listening
    {
        [JsonProperty(PropertyName = "score")]
        public string Score { get; set; }

        [JsonProperty(PropertyName = "max_score")]
        public string Max_Score { get; set; }
    }
}