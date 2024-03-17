using System;
using System.Collections.Generic;
using System.Threading.Tasks;




namespace CRUD_Test.Modules
{
    public class Response
    {
        public int StatusCode { get; set; }
        public string Message { get; set; }
        public UserData UserData { get; set; }
        public List<UserData> UserDatas { get; set;}

    }
}
