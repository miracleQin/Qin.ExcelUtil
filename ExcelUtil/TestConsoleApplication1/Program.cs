using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            ToList();
        }
        static void ToExcel() 
        {
            List<TestModel> list = new List<TestModel>();
            list.Add(new TestModel() { DTValue=DateTime.Now, StrValue="test" });
            list.Add(new TestModel() { IntValue=2});

            Qin.ExcelUtil.ExcelUtil.ListToExcel<TestModel>(list, System.AppDomain.CurrentDomain.BaseDirectory+@"\test.xls");

        }
        static void ToList() 
        {
            var list = Qin.ExcelUtil.ExcelUtil.ExcelToList<TestModel>(@"C:\Users\user\Desktop\test.xls");
            foreach (var item in list) 
            {
                Console.WriteLine(string.Format("int:{0} dateTime:{1} string:{2}", item.IntValue, item.DTValue, item.StrValue));
            }
            Console.ReadLine();
        }
    }
}
