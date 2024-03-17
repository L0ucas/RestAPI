using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;


namespace CRUD_Test.Modules
{
    public class UserQuery
    {
        public Response GetAllUsers(SqlConnection Connection)
        {
            Response Response = new Response();
            string strSqlQueryOrSpName = "SELECT * From tblTestUser";

            SqlDataAdapter cmd = new SqlDataAdapter(strSqlQueryOrSpName, Connection);

            //SqlDataAdapter da = new SqlDataAdapter(cmd);
            List<UserData> Users = new List<UserData>();
            DataTable dt = new DataTable();
            cmd.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    UserData user = new UserData();
                    user.Id = dt.Rows[i]["Id"].ToString();
                    user.Username = dt.Rows[i]["Username"].ToString();
                    user.Email = dt.Rows[i]["Email"].ToString();
                    user.Number = dt.Rows[i]["Number"].ToString();
                    user.Skillsets = dt.Rows[i]["Skillsets"].ToString();
                    user.Hobby = dt.Rows[i]["Hobby"].ToString();
                    user.Active = dt.Rows[i]["Active"].ToString();

                    Users.Add(user);
                }
            }


            if (Users.Count > 0)
            {
                Response.StatusCode = 200;
                Response.Message = "Users Retrieved.";
                Response.UserDatas = Users;
            }
            else
            {
                Response.StatusCode = 100;
                Response.Message = "Failed to retrieve users.";
                Response.UserDatas = Users;
            }



            return Response;
        }

        public Response GetUserByName(SqlConnection Connection, string Name)
        {
            Response Response = new Response();
            string strSqlQueryOrSpName = $"SELECT * From tblTestUser where Username = '{Name}' AND Active = 1";

            SqlDataAdapter cmd = new SqlDataAdapter(strSqlQueryOrSpName, Connection);

            //SqlDataAdapter da = new SqlDataAdapter(cmd);
            UserData Users = new UserData();
            DataTable dt = new DataTable();
            cmd.Fill(dt);

            if (dt.Rows.Count > 0)
            {
         
                    UserData user = new UserData();
                    user.Id = dt.Rows[0]["Id"].ToString();
                    user.Username = dt.Rows[0]["Username"].ToString();
                    user.Email = dt.Rows[0]["Email"].ToString();
                    user.Number = dt.Rows[0]["Number"].ToString();
                    user.Skillsets = dt.Rows[0]["Skillsets"].ToString();
                    user.Hobby = dt.Rows[0]["Hobby"].ToString();
                    user.Active = dt.Rows[0]["Active"].ToString();
                    Response.StatusCode = 200;
                    Response.Message = "Users Retrieved.";
                    Response.UserData = user;

            }

            else
            {
                Response.StatusCode = 100;
                Response.Message = "Failed to retrieve users.";
                Response.UserData = null;
            }



            return Response;
        }

        public Response AddUser(SqlConnection Connection, UserData userData)
        {
            Response Response = new Response();
            string strSqlQueryOrSpName = $"Insert into tblTestUser values(newID(),'{userData.Username}','{userData.Email}','{userData.Number}','{userData.Skillsets}','{userData.Hobby}','{userData.Active}')";


            SqlCommand cmd = new SqlCommand(strSqlQueryOrSpName, Connection);

            Connection.Open();
            int j = cmd.ExecuteNonQuery();
            Connection.Close();

            //SqlDataAdapter da = new SqlDataAdapter(cmd);

            if (j > 0)
            {
                Response.StatusCode = 200;
                Response.Message = "User added.";
            }

            else
            {
                Response.StatusCode = 100;
                Response.Message = "Failed to add user.";
            }

            return Response;
        }

        public Response UpdateUser(SqlConnection Connection, UserData userData)
        {
            Response Response = new Response();
            string strSqlQueryOrSpName = $"Update tblTestUser set Email = '{userData.Email}', Number = '{userData.Number}', Skillsets = '{userData.Skillsets}', Hobby = '{userData.Hobby}', Active = '{userData.Active}' where Username = '{userData.Username}' ";


            SqlCommand cmd = new SqlCommand(strSqlQueryOrSpName, Connection);

            Connection.Open();
            int j = cmd.ExecuteNonQuery();
            Connection.Close();

            //SqlDataAdapter da = new SqlDataAdapter(cmd);

            if (j > 0)
            {
                Response.StatusCode = 200;
                Response.Message = "User Updated.";
            }

            else
            {
                Response.StatusCode = 100;
                Response.Message = "Failed to Update user.";
            }

            return Response;
        }

        public Response DeleteUser(SqlConnection Connection, UserData userData)
        {
            Response Response = new Response();
            string strSqlQueryOrSpName = $"Delete from tblTestUser where Username = '{userData.Username}'";


            SqlCommand cmd = new SqlCommand(strSqlQueryOrSpName, Connection);

            Connection.Open();
            int j = cmd.ExecuteNonQuery();
            Connection.Close();

            //SqlDataAdapter da = new SqlDataAdapter(cmd);

            if (j > 0)
            {
                Response.StatusCode = 200;
                Response.Message = "User Deleted.";
            }

            else
            {
                Response.StatusCode = 100;
                Response.Message = "Failed to Delete user.";
            }

            return Response;
        }
    }
}
