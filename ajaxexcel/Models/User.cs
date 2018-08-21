using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace ajaxexcel.Models
{
    public class User
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public int Age { get; set; }

        public string Address { get; set; }

        public static List<User> DefaultSection => new List<User>
        {
            new User { Name = "心惊", Address = "更健康",Age = 50},
            new User { Name = "伐柯伐柯", Address = "方法还",Age = 52},
            new User { Name = "付款", Address = "法规开关机",Age = 53}
        };
    }
}