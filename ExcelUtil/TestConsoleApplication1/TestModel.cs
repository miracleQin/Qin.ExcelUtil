using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsoleApplication1
{
    [DisplayName("测试实体")]
    public class TestModel
    {
        [DisplayName("整形字段")]
        public int? IntValue { get; set; }
        [DisplayName("字符串类型字段")]
        public string StrValue { get; set; }
        [DisplayName("时间类型字段")]
        public DateTime? DTValue { get; set; }

    }
}
