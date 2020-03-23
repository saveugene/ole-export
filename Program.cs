using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;

namespace ole_export
{
    public class FileObj
    {
        public FileObj(string title, byte[] data)
        {
            Title = title;
            Data = data;
        }
        public string Title { get; private set; }
        public byte[] Data { get; private set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            IDictionary<string, string> answersDefault = new Dictionary<string, string>(){
                {"dbName", "master"},
                {"server", "localhost"},
                {"login", "SA"},
                {"password", "<YourNewStrong!Passw0rd>"},
                {"table", "dbo.OleDocument"},
                {"oleColumn", "OLEContainer"},
                {"titleColumn", "Title"},
            };

            IDictionary<string, string> questions = new Dictionary<string, string>(){
                {"dbName", $@"Database name: (default: {answersDefault["dbName"]})"},
                {"server", $@"Server: (default: {answersDefault["server"]})"},
                {"login", $@"User Id: (default: {answersDefault["login"]})"},
                {"password", $@"Password: (default: {answersDefault["password"]})"},
                {"table", $@"Table: (default: {answersDefault["table"]})"},
                {"oleColumn", $@"OLE column: (default: {answersDefault["oleColumn"]})"},
                {"titleColumn", $@"Title column: (default: {answersDefault["titleColumn"]})"},
            };

            IDictionary<string, string> answers = new Dictionary<string, string>();

            foreach (KeyValuePair<string, string> question in questions)
            {
                string answer;
                if (question.Key != "password")
                    answer = AskQuesiton(question.Value);
                else answer = AskQuesiton(question.Value, true);

                if (!string.IsNullOrWhiteSpace(answer))
                    answers.Add(question.Key, answer);
                else answers.Add(question.Key, answersDefault[question.Key]);
            }

            string connectionString = $@"Server={answers["server"]}; Database={answers["dbName"]}; User Id={answers["login"]}; Password={answers["password"]};";
            string sql = $"SELECT {answers["oleColumn"]}, {answers["titleColumn"]} FROM {answers["dbName"]}.{answers["table"]}";
            WriteFilesOnDrive(connectionString, sql);
        }

        static string AskQuesiton(string question, bool isSecure = false)
        {
            Console.WriteLine(question);
            if (!isSecure)
                return Console.ReadLine();
            else
            {
                string password = "";
                while (true)
                {
                    var key = System.Console.ReadKey(true);
                    if (key.Key == ConsoleKey.Enter)
                        break;
                    if (key.Key == ConsoleKey.Backspace)
                    {
                        if (password.Length > 0)
                        {
                            password = password.Remove(password.Length - 1);
                            continue;
                        }
                    }
                    password += key.KeyChar;
                }
                return password;
            }

        }

        static void WriteFilesOnDrive(string connectionString, string sql)
        {
            List<FileObj> files = new List<FileObj>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(sql, connection);
                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    byte[] data;
                    string title;
                    if (!Convert.IsDBNull(reader.GetValue(0)) && !Convert.IsDBNull(reader.GetValue(1)))
                    {
                        data = (byte[])reader.GetValue(0);
                        title = reader.GetString(1);
                    }
                    else continue;

                    files.Add(new FileObj(title, data));
                }
            }

            if (files.Count > 0)
            {
                string dir = @"files";

                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                for (int i = 0; i < files.Count; i++)
                {
                    using (FileStream fs = new FileStream(Path.Combine(dir, files[i].Title.Replace("/", "")), FileMode.OpenOrCreate))
                    {
                        fs.Write(files[i].Data, 0, files[i].Data.Length);
                        Console.WriteLine($@"File '{files[i].Title}' is saved!");
                    }
                }
            }
            Console.WriteLine("Done!");
        }
    }
}
