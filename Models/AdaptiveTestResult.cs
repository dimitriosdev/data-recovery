using Newtonsoft.Json;
using System.Collections.Generic;

namespace EFsetWidgetFix.Models
{
    public class AdaptiveTestResult
    {
        [JsonProperty(PropertyName = "test_results")]
        public List<TestResultA> AdaptiveResults { get; set; }
    }

    public class TestResultA
    {
        [JsonProperty(PropertyName = "test_name")]
        public string TestName { get; set; }

        [JsonProperty(PropertyName = "scores")]
        public AdaptiveScore Score { get; set; }

        [JsonProperty(PropertyName = "test_start_time")]
        public string TestStartTime { get; set; }

        [JsonProperty(PropertyName = "test_finish_time")]
        public string TestFinishTime { get; set; }
    }

    public class AdaptiveScore
    {
        public string Combined { get; set; }

        public string Cefr { get; set; }

        public AdaptiveReadingScore Reading { get; set; }

        public AdaptiveListeningScore Listening { get; set; }
    }

    public class AdaptiveReadingScore
    {
        public string Score { get; set; }

        public string Cefr { get; set; }
    }

    public class AdaptiveListeningScore
    {
        public string Score { get; set; }

        public string Cefr { get; set; }
    }
}