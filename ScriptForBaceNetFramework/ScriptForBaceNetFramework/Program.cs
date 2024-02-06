using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using ScriptForBaceNetFramework.Class;
using ScriptForBaceNetFramework.Entity;
using static System.Net.Mime.MediaTypeNames;

namespace ScriptForBaceNetFramework
{
    internal class Program
    {
        

        static void Main(string[] args)
        {
            string pathFroImage = $"C:\\Users\\{Environment.UserName}\\OneDrive\\Документы\\NBA\\";
            
            //AddPicturesEntry(pathFroImage);
            //AddPlayerEntry(pathFroImage);
            //AddTeamEntry(pathFroImage);
            InsertToMatchLog(pathFroImage);


        }
        
        static void AddPicturesEntry(string pathFroImage)
        {
            var context = GetDB.GetInstance().DB();
            Workbook wr = new Workbook(pathFroImage + "Data\\Pictures.xlsx");
            WorksheetCollection collection = wr.Worksheets;
            Worksheet worksheet = collection[0];
            for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
            {
                var newPicture = new Pictures();
                for (int j = 0; j <= worksheet.Cells.MaxDataColumn; j++)
                {
                    switch (j)
                    {
                        case 0:
                            newPicture.Id = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 1:
                            var pic = File.ReadAllText(pathFroImage + "Image\\Pictures\\" + worksheet.Cells[i, j].Value);
                            newPicture.Img = Encoding.Default.GetBytes(pic);
                            break;
                        case 2:
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                newPicture.Description = worksheet.Cells[i, j].Value.ToString();
                            }
                            break;
                        case 3:
                            newPicture.NumberOfLike = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 4:
                            string dateString = worksheet.Cells[i, j].Value.ToString();
                            string format = "dd.MM.yyyy H:mm:ss";
                            if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture,
                                    DateTimeStyles.None, out DateTime result))
                            {
                                newPicture.CreateTime = result;
                            }
                            break;
                    }


                }
                //Console.WriteLine(newPicture.CreateTime.ToString());
                context.Pictures.Add(newPicture);
                context.SaveChanges();
            }
        }

        static void AddTeamEntry(string pathFroImage)
        {
            var context = GetDB.GetInstance().DB();
            Workbook wr = new Workbook(pathFroImage + "Data\\Team.xlsx");
            WorksheetCollection collection = wr.Worksheets;
            Worksheet worksheet = collection[0];
            context.Division.Load();
            context.Conference.Load();
           

            for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
            {
                var newTeam = new Team();
                for (int j = 0; j <= worksheet.Cells.MaxDataColumn; j++)
                {
                    switch (j)
                    {
                        case 0:
                            newTeam.TeamId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 1:
                            newTeam.TeamName = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 2:
                            newTeam.DivisionId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 3:
                            newTeam.Abbr = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 4:
                            newTeam.Coach = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 5:
                            newTeam.Stadium = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 6:
                            var pic = File.ReadAllText(pathFroImage + "Image\\Teams\\" + worksheet.Cells[i, j].Value);
                            newTeam.Logo = Encoding.Default.GetBytes(pic);
                            break;
                    }
                }
                context.Team.Add(newTeam);
                context.SaveChanges();
            }
        }

        static void AddPlayerEntry(string pathFroImage)
        {
            var context = GetDB.GetInstance().DB();
            Workbook wr = new Workbook(pathFroImage + "Data\\Player.xlsx");
            WorksheetCollection collection = wr.Worksheets;
            Worksheet worksheet = collection[0];
            context.Position.Load();
            context.Country.Load();

            for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
            {
                var newPlayer = new Player();
                for (int j = 0; j <= worksheet.Cells.MaxDataColumn; j++)
                {
                    switch (j)
                    {
                        case 0:
                            newPlayer.PlayerId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 1:
                            newPlayer.Name = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 2:
                            newPlayer.PositionId = (int)worksheet.Cells[i,j].Value;
                            break;
                        case 3:
                            string dateString = worksheet.Cells[i, j].Value.ToString();
                            string format = "dd.MM.yyyy";
                            if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture,
                                    DateTimeStyles.None, out DateTime result))
                            {
                                newPlayer.JoinYear = result;
                            }
                            break;
                        case 4:
                            newPlayer.Height = decimal.Parse(worksheet.Cells[i, j].Value.ToString());
                            break;
                        case 5:
                            newPlayer.Weight = decimal.Parse(worksheet.Cells[i, j].Value.ToString());
                            break;
                        case 6:
                            if (worksheet.Cells[i, j].Value != null){
                                dateString = worksheet.Cells[i, j].Value.ToString();
                                format = "dd.MM.yyyy";
                                if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture,
                                        DateTimeStyles.None, out  result))
                                {
                                    newPlayer.DateOfBirth = result;
                                }
                            }
                            break;
                        case 7:
                            if(worksheet.Cells[i, j].Value != null)
                                newPlayer.College = worksheet.Cells[i, j].Value.ToString();
                            break;
                        case 8:
                            newPlayer.CountryCode = worksheet.Cells[i,j].Value.ToString();
                            break;
                        case 9:
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                var pic = File.ReadAllText(pathFroImage + "Image\\Players\\" + worksheet.Cells[i, j].Value);
                                newPlayer.Img = Encoding.Default.GetBytes(pic);
                            }
                            break;
                        case 10:
                            if (worksheet.Cells[i, j].Value.ToString() == "ЛОЖЬ")
                            {
                                newPlayer.IsRetirment = false;
                            }
                            else
                            {
                                newPlayer.IsRetirment = true;
                            }
                            break;  
                        case 11:
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                dateString = worksheet.Cells[i, j].Value.ToString();
                                format = "dd.MM.yyyy";
                                if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture,
                                        DateTimeStyles.None, out result))
                                {
                                    newPlayer.RetirementTime = result;
                                }
                            }
                            break;

                    }
                }
                context.Player.Add(newPlayer);
                context.SaveChanges();
            }
        }

        static void InsertToMatchLog(string pathFroImage)
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            var context = GetDB.GetInstance().DB();
            Workbook wr = new Workbook(pathFroImage + "Data\\MatchupLog.xlsx");
            WorksheetCollection collection = wr.Worksheets;
            Worksheet worksheet = collection[0];
            context.Player.Load();
            context.Team.Load();
            context.Matchup.Load();
            context.ActionType.Load();
            for (int i = 2; i <= 3; i++)
            {
                var newMatchupLog = new MatchupLog();
                for (int j = 0; j <= worksheet.Cells.MaxDataColumn; j++)
                {
                    switch (j)
                    {
                        case 0:
                            newMatchupLog.Id = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 1:
                            newMatchupLog.MatchupId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 2:
                            newMatchupLog.Quarter = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 3:
                            string str = worksheet.Cells[i, j].Value.ToString().Split(' ')[1];
                            newMatchupLog.OccurTime = str;
                            break;
                        case 4:
                            newMatchupLog.TeamId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 5:
                            newMatchupLog.PlayerId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 6:
                            newMatchupLog.ActionTypeId = (int)worksheet.Cells[i, j].Value;
                            break;
                        case 7:
                            newMatchupLog.Remark = worksheet.Cells[i, j].Value.ToString();
                            break;
                        
                    }
                }
                context.MatchupLog.Add(newMatchupLog);
                context.SaveChanges();
            }
            stopwatch.Stop();
            TimeSpan ts = stopwatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            Console.WriteLine("RunTime " + elapsedTime);
        }
    }
}
