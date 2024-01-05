using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace ExcelReader;

internal class DatabaseManager
{
    public void PopulateTable(List<People> people)
    {
        try
        {
            CreateTable();
            Console.WriteLine("Inserting the data from the excel file into the database table..");

            using var connection = new SqlConnection(GetConnectionString());

            connection.Open();

            foreach (var person in people)
            {
                string sqlCommandText = 
                    $"use ExcelReader; " +
                    $"INSERT INTO People VALUES" +
                    $"({person.Id}, '" + person.FirstName + "', '" + person.LastName + "', '" + 
                    person.Sex + "', '" + person.Email + "', '" + person.Phone + "', '" + person.BirthDate + "', '" + person.JobTitle + "')";

                SqlCommand sqlCommand = new(sqlCommandText, connection);

                sqlCommand.ExecuteNonQuery();
            }

            connection.Close();

            Console.WriteLine("The data has been successfully inserted.");
        } catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    void CreateTable()
    {
        try
        {
            CreateDatabase();
            Console.WriteLine("Creating the table..");

            using var connection = new SqlConnection(GetConnectionString());

            connection.Open();

            string sqlCommandText =
                @"
                    use ExcelReader;
                    CREATE TABLE People(
                    PeopleId INT PRIMARY KEY,
                    FirstName VARCHAR(64),
                    LastName VARCHAR(64),
                    Sex VARCHAR(6),
                    Email VARCHAR(64),
                    Phone VARCHAR(10),
                    BirthDate Date,
                    JobTitle VARCHAR(64))
                ";

            SqlCommand sqlCommand = new(sqlCommandText, connection);

            if (sqlCommand.ExecuteNonQuery() == -1)
                Console.WriteLine("The table has been successfully created.\n");

            connection.Close();
        } catch (Exception ex)
        {
            Console.WriteLine($"The table couldn't be created: {ex.Message}");
        }
    }

    void CreateDatabase()
    {
        try
        {
            DeleteDatabase();
            Console.WriteLine("Creating the database..");

            using var connection = new SqlConnection(GetConnectionString());

            connection.Open();

            string sqlCommandText =
                @"
                    CREATE DATABASE ExcelReader;
                ";

            SqlCommand sqlCommand = new(sqlCommandText, connection);

            if (sqlCommand.ExecuteNonQuery() == -1)
                Console.WriteLine("The database has been successfully created.\n");

            connection.Close();
        } catch (Exception ex)
        {
            Console.WriteLine($"The database couldn't be created: {ex.Message}");
        }
    }

    void DeleteDatabase()
    {
        try
        {
            Console.WriteLine("Checking to see if the database already exists and deleting it if true..\n");

            using var connection = new SqlConnection(GetConnectionString());

            connection.Open();

            string sqlCommandText =
                @"
                    DROP DATABASE IF EXISTS ExcelReader;
                ";

            SqlCommand sqlCommand = new(sqlCommandText, connection);
            sqlCommand.ExecuteNonQuery();

            connection.Close();
        } catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }

    static string GetConnectionString()
    {
        IConfigurationBuilder configurationBuilder = new ConfigurationBuilder().AddJsonFile("appSettings.json");
        IConfigurationRoot configuration = configurationBuilder.Build();

        return configuration.GetConnectionString("DefaultConnectionString");
    }
}
