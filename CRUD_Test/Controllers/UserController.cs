using CRUD_Test.Modules;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Data.SqlClient;


namespace CRUD_Test.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UserController : ControllerBase
    {

        private readonly IConfiguration _configuration;
        public UserController(IConfiguration Configuration)
        {
            _configuration = Configuration;
        }

        [HttpGet]
        [Route("GetAllUsers")]

        public Response GetAllUsers()
        {
            SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection").ToString());
            Response response = new Response();

            UserQuery New = new UserQuery();
            response = New.GetAllUsers(connection);
            return response;
        }

        [HttpGet]
        [Route("GetUsersByName/{Name}")]
        public Response GetUsersByName(string Name)
        {
            SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection").ToString());
            Response response = new Response();

            UserQuery New = new UserQuery();
            response = New.GetUserByName(connection,Name);
            return response;
        }

        [HttpPost]
        [Route("AddUser")]
        public Response AddUser(UserData User)
        {
            SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection").ToString());
            Response response = new Response();

            UserQuery New = new UserQuery();
            response = New.AddUser(connection, User);
            return response;
        }

        [HttpPut]
        [Route("UpdateUser")]
        public Response UpdateUser(UserData User)
        {
            SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection").ToString());
            Response response = new Response();

            UserQuery New = new UserQuery();
            response = New.UpdateUser(connection, User);
            return response;
        }

        [HttpDelete]
        [Route("DeleteUser")]
        public Response DeleteUser(UserData User)
        {
            SqlConnection connection = new SqlConnection(_configuration.GetConnectionString("DefaultConnection").ToString());
            Response response = new Response();

            UserQuery New = new UserQuery();
            response = New.DeleteUser(connection, User);
            return response;
        }
    }
}
