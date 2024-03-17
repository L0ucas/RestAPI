# RestAPI
This project consist of a simple rest API that performs the following verbs: @Get, @Post, @Put and @Delete 
in conjunction with the data storage udner Microsoft SQL Server Management Studio.
- @Get retrieves all of the active users inside the table, you may retrieve a specific user via its unique Id.
- @Post Registers the new user and pumps its details into the database.
- @Put Updates current user details based on its user's name.
- @Delete Removes user details based on the user's name.

## Installation and Run
Download the Code above and run the Crud_test application under the bin/release folder.
![image](https://github.com/L0ucas/RestAPI/assets/50651727/549b5cf5-cfea-4c76-b188-c2c1690317a8)

Before Running the API, please run the database script below to set up the necessary table inside Microsoft SQL Server.
```
IF EXISTS (SELECT * FROM sysobjects WHERE name='tblTestUser' and xtype='U')
BEGIN
	DROP TABLE [dbo].[tblTestUser]
END
GO
CREATE TABLE [dbo].[tblTestUser]
(
	[Id] [uniqueidentifier] NOT NULL,
	[Username] [nvarchar](50) NOT NULL,
	[Email] [varchar](100) NOT NULL,
	[Number] [varchar](100) NULL,
	[Skillsets] [varchar](100) NOT NULL,
	[Hobby] [varchar](100) NOT NULL,
	[Active] [bit] NOT NULL,
 CONSTRAINT [PK_TestUser_Id] PRIMARY KEY NONCLUSTERED 
)
GO
```
*Please replace your database connection string inside appsettings.json under default connections.
![image](https://github.com/L0ucas/RestAPI/assets/50651727/4b49128b-88a0-4025-8d9f-935a2c35ce5c)

## Test API
GetAllUsers
- Endpoint: http://localhost:5000/api/User/GetAllUsers
- Method: Get

AddUser
- Endpoint: http://localhost:5000/api/User/AddUser
- Method: Post
- Content Type: application/json
- Request Parameter:
- ![image](https://github.com/L0ucas/RestAPI/assets/50651727/c36de655-7180-4b61-80b1-101f048e27bb)

UpdateUser
- Endpoint: http://localhost:5000/api/User/UpdateUser
- Method: Put
- Content Type: application/json
- Request Parameter: 
- ![image](https://github.com/L0ucas/RestAPI/assets/50651727/f5575aa2-6ddf-4453-b5f9-f0eceaecd580)

DeleteUser
- Endpoint: http://localhost:5000/api/User/DeleteUser
- Method: Post
- Content Type: application/json
- Request Parameter:
- ![image](https://github.com/L0ucas/RestAPI/assets/50651727/4ef8bdbc-9324-442a-b036-4b0ab58165a1)

