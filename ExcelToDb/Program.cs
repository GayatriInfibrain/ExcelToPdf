using OfficeOpenXml;
using System;
using System.IO;
using System.Data.SqlClient;
namespace ExcelToDb
{
    class Program
    {
        static void Main(string[] args)
        {
            string connectionString = "Server=192.168.0.240;Database=ALMSDB;UID=sa;PWD=sa@123;"; // Db connection string
            string excelFilePath = "C:\\Users\\gayatri.rajput\\Desktop\\ResumeCopy.xlsx"; // excel file path
            string resumeFilePath = "I:\\SVN\\ExcelToDB\\ResumePDF";//resume File Path

            FileInfo existingFile = new FileInfo(excelFilePath);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];

                if (worksheet == null)
                {
                    throw new Exception("Worksheet 'Resume' not found.");
                }

                int rowCount = worksheet.Dimension.Rows;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {


                    for (int row = 2; row <= rowCount; row++)
                    {
                        int position = 0;
                        int source = 0;
                        string firstname = worksheet.Cells[row, 1].Text;
                        string lastname = worksheet.Cells[row, 2].Text;
                        string positionname = worksheet.Cells[row, 3].Text;
                        string sourcename = worksheet.Cells[row, 4].Text;
                        string oldfilename = worksheet.Cells[row, 5].Text;
                        string resumefilename = firstname + lastname;
                        resumefilename = resumefilename.Replace(" ", string.Empty);
                        string newresumefilename = positionname + "/" + resumefilename;
                        DateTime date = DateTime.Now;
                        string oldFilePath = resumeFilePath + "\\" + positionname + "\\" + oldfilename;
                        string newFilePath = resumeFilePath + "\\" + positionname + "\\" + resumefilename + ".pdf"; 
                        if (File.Exists(oldFilePath))
                        {
                            if (File.Exists(newFilePath))
                            {
                                string folderPath = resumeFilePath + "\\" + positionname;
                                newresumefilename = GenerateUniqueFileName(resumefilename, folderPath);
                                newFilePath = resumeFilePath + "\\" + positionname + "\\" + newresumefilename;
                                newresumefilename = positionname + "/" + newresumefilename;

                            }
                            File.Move(oldFilePath, newFilePath);
                            Console.WriteLine("File renamed successfully!");
                        }
                        else
                        {
                            Console.WriteLine("File not found: " + oldFilePath);
                        }
                        switch (positionname)
                        {
                            case "Business Development Executive":
                                position = 128;
                                break;
                            case "Dotnet Developer":
                                position = 130;
                                break;
                            case "Graphic Designer":
                                position = 157;
                                break;
                            case "SEO Manager":
                                position = 138;
                                break;
                            case "Project Manager":
                                position = 166;
                                break;
                            case "Network Engineer":
                                position = 170;
                                break;
                            default:
                                continue;
                        }

                        switch (sourcename)
                        {
                            case "Naukari":
                                source = 128;
                                break;
                            default:
                                continue;
                        }
                        string query = "INSERT INTO tblResumeBankManagement (FirstName, LastName, PositionId, SourceId, FileName,UploadedDate) VALUES (@FirstName, @LastName, @PositionId, @SourceId, @FileName, @Date)";
                        connection.Open();
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
    
                            command.Parameters.AddWithValue("@FirstName", firstname);
                            command.Parameters.AddWithValue("@LastName", lastname);
                            command.Parameters.AddWithValue("@PositionId", position);
                            command.Parameters.AddWithValue("@SourceId", source);
                            command.Parameters.AddWithValue("@FileName", newresumefilename);
                            command.Parameters.AddWithValue("@Date", date);
                            //int rowsAffected = command.ExecuteNonQuery();
                            //Console.WriteLine($"{rowsAffected} row(s) inserted for {firstname}.");
                            Console.WriteLine($"row(s) renamed to {oldfilename} to {newresumefilename}.");
                        }
                        connection.Close();
                    }
                }
            }
            static string GenerateUniqueFileName(string baseFileName, string folderPath)
            {
                string newFileName = baseFileName;
                string fullPath = Path.Combine(folderPath, newFileName + ".pdf");
                int count = 1;
                // Check if file exists and append numbers until a unique name is found
                while (File.Exists(fullPath))
                {
                    newFileName = $"{baseFileName}_{count}";
                    fullPath = Path.Combine(folderPath, newFileName + ".pdf");
                    count++;
                }
                return newFileName + ".pdf";
            }
        }
    }
}
