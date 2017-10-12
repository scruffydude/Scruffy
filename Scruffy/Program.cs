using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;


namespace Scruffy
{
    class Program
    {
        static void Main(string[] args)
        {
            //setup log file information
            string logPath = @"\\cfc1afs01\Operations-Analytics\Log_Files\";
            StreamWriter logging = null;
            logging = new StreamWriter(logPath + "Scruffy's_Task_list.txt");
            logging.WriteLine("Scruffy gonna start cleaning things up at " + DateTime.Now + " mmhmm");
            Console.WriteLine("Scruffy gonna start cleaning things up at " + DateTime.Now + " mmhmm");


            //set up flags
            bool moveFlag = false;
            //determine how we plan on cleaning up either moving or deleting
            if (moveFlag)
            {
                logging.WriteLine("Looks like we're gonna move stuff..... mmhmm");
                Console.WriteLine("Looks like we're gonna move stuff..... mmhmm");
            }
            else
            {
                logging.WriteLine("Looks like we're gonna copy stuff..... mmhmm");
                Console.WriteLine("Looks like we're gonna copy stuff..... mmhmm");
            }

            //define what locations we wish to clean up
            string archivePath = @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\NotProcessed\";
            string processedPath = @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\ProcessingCompleted\";

            string[] warehouses = { "AVP1", "CFC1", "DFW1", "EFC3", "WFC2" };

            foreach (string wh in warehouses)
            {
                switch (wh)
                {
                    case "AVP1":
                        archivePath = @"\\avp1afs01\Outbound\AVP Flow Plan\FlowPlanArchive\";
                        break;
                    case "CFC1":
                        archivePath = @"\\cfc1afs01\Outbound\Master Wash\FlowPlanArchive\";
                        break;
                    case "DFW1":
                        archivePath = @"\\dfw1afs01\Outbound\FlowPlanArchive\";
                        break;
                    case "EFC3":
                        archivePath = @"\\wh-pa-fs-01\OperationsDrive\FlowPlanArchive\";
                        break;
                    case "WFC2":
                        archivePath = @"\\wfc2afs01\Outbound\Outbound Flow Planner\FlowPlanArchive\";
                        break;
                    default:
                        Console.WriteLine("Warehouse not found please add to structure" + wh);
                        logging.WriteLine("Warehouse not found please add to structure" + wh);
                        break;
                }
                archivePath = archivePath + @"2017\"; //change default location of Archive to combat old information.
                processedPath = archivePath + @"ProcessingCompleted\";

                if (Directory.Exists(archivePath))
                {
                    if (!Directory.Exists(processedPath))
                    {
                        System.IO.Directory.CreateDirectory(processedPath);
                    }

                    ProcessDirectory(archivePath, processedPath, moveFlag);
                }
                else
                {
                    Console.WriteLine("Looks like she's empty " + archivePath);
                    logging.WriteLine("Looks like she's empty " + archivePath);
                }
                //shout success
                Console.WriteLine("Works done now mmhmm");
                logging.WriteLine("Works done now mmhmm");
                //Console.ReadLine();
            }
        }

        public static void ProcessDirectory(string archivePath, string processedPath, bool moveFlag)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(archivePath);
            ProcessFile(fileEntries, processedPath, moveFlag);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(archivePath);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory, processedPath, moveFlag);
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string[] archiveFiles, string processedPath, bool moveFlag)
        {
            DateTime lastModified = new DateTime(1900,1,1);
            string fileName = "";
            string destFile = "";
            string version = "";
            string destloc = "";

            //setting the destination for the roll up data.
            string Lvl1rollup = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\RollUpInfo\LVL1rollup.xlsx";
            string Lvl2rollup = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\RollUpInfo\LVL2rollup.xlsx";
            string Lvl3rollup = @"\\CFC1AFS01\Operations-Analytics\Projects\Flow Plan\RollUpInfo\LVL3rollup.xlsx";
            string[] lvl1rollupCellsSOS = 
                {
                "I2", "I3", "I4", "I5", "C9", "C10", "C11", "C12",
                "C14", "C15", "C16", "C17", "E9", "E10", "E11",
                "E12", "G9", "G10","G11", "G12", "I9", "I10", "I11",
                "I12" 
                }; // here we list ever cell we need informatino from for the version 2 or greater
            string[] lvl1rollupCellsEOS =
            {
                "D30", "D31", "D33", "D34", "D35","D39"
            };
            string[] lvl1rollupCellsEOSolderversion =
            {
                "E23","E24","E26","E27","E28", "E20"
            };
            string[] lvl2rollupCells =
            {

            };
            string[] lvl3rollupCells =
            {

            };
            string[] lvl1processableCellsSOS = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};
            string[] lvl1processableCellsEOS = { "", "", "", "", "", "" };
            
            //open file
            //setup excel application and define the destination file we are going to output.
            Excel.Application app = null;
            app = new Excel.Application();
            app.Application.Visible = true;

            //create the blank workbooks holder
            Excel.Workbooks workbook = null;
            Excel.Workbook lvl1rollupWorkbook = null;
            Excel.Workbook lvl2rollupWorkbook = null;
            Excel.Workbook lvl3rollupWorkbook = null;
            Excel.Workbook file = null;
            workbook = app.Workbooks;

            //emptycheck if rollups exist if not create them then open them
            var lvl1file = new FileInfo(Lvl1rollup);
            if (!lvl1file.Exists)
            {
                FileStream fs = new FileStream(Lvl1rollup, FileMode.CreateNew);
            }
            lvl1rollupWorkbook = workbook.Open(Lvl1rollup, false, false);

            var lvl2file = new FileInfo(Lvl2rollup);
            if (!lvl2file.Exists)
            {
                FileStream fs = new FileStream(Lvl2rollup, FileMode.CreateNew);
            }
            lvl2rollupWorkbook = workbook.Open(Lvl2rollup, false, false);

            var lvl3file = new FileInfo(Lvl3rollup);
            if (!lvl3file.Exists)
            {
                FileStream fs = new FileStream(Lvl3rollup, FileMode.CreateNew);
            }
            lvl3rollupWorkbook = workbook.Open(Lvl3rollup, false, false);

            foreach (string archiveSrcFile in archiveFiles)
            {
                // Use system Path methods to extract only the file name from the path.
                Console.WriteLine("There's another at '{0}'.", archiveSrcFile);
                fileName = System.IO.Path.GetFileName(archiveSrcFile);
                lastModified = System.IO.File.GetLastWriteTime(archiveSrcFile);
                string date = lastModified.ToString("yyyy-MM-dd");
                date = date.Replace('-', '\\');
                destloc = System.IO.Path.Combine(processedPath, date);
                destFile = System.IO.Path.Combine(destloc, fileName);



                var fi = new FileInfo(archiveSrcFile);
                if (fi.Exists)
                {
                    try
                    {
                        file = workbook.Open(archiveSrcFile, false, false);
                    }
                    catch
                    {
                        continue;
                    }
                    
                }

                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.DisplayAlerts = false;
                Excel.Worksheet SOS = file.Worksheets.Item["SOS"];
                Excel.Worksheet EOS = file.Worksheets.Item["EOS"];
                Excel.Worksheet HourlyStaffing = file.Worksheets.Item["Hourly ST"];
                //parse info, only if  a valid version number
                version = file.Worksheets.Item["SOS"].cells(2, 9).value;

                //get last row of destinatino files
                int lvl1lastRow = lvl1rollupWorkbook.Worksheets[1].UsedRange.Rows.Count;
                int lvl2lastRow = lvl2rollupWorkbook.Worksheets[1].UsedRange.Rows.Count;
                int lvl3lastRow = lvl3rollupWorkbook.Worksheets[1].UsedRange.Rows.Count;

                var emptycheck = SOS.Cells[9, 5].value;
                
                if(emptycheck == null)
                {
                    file.Close(false);
                    System.IO.File.Delete(archiveSrcFile);
                }
                else
                {
                    if(version.Contains("2.0.1"))
                    {
                        //process newest version file
                        lvl1rollupCellsSOS.CopyTo(lvl1processableCellsSOS, 0);
                        lvl1rollupCellsEOS.CopyTo(lvl1processableCellsEOS, 0);
                    }
                    else if(version.Contains("1.9") || version.Contains("2.0"))
                    {
                        //process older version file
                        lvl1rollupCellsSOS.CopyTo(lvl1processableCellsSOS, 0);
                        lvl1processableCellsSOS[22] = "I8";
                        lvl1rollupCellsEOSolderversion.CopyTo(lvl1processableCellsEOS, 0);
                        SOS.Cells[17, 3].value = SOS.Cells[12, 3].value - SOS.Cells[22, 9].value;
                    }
                    else
                    {
                        Console.WriteLine("Incorrect Version Information: " + version);
                        file.Close(false);
                    }

                    //process all of the lvl1 roll up versions

                    int i = 1;

                    foreach (string cell in lvl1processableCellsSOS)
                    {
                        Excel.Range rng = SOS.Range[cell];
                        lvl1rollupWorkbook.Worksheets.Item[1].cells(lvl1lastRow + 1, i).value = rng.Value;
                        i++;
                    }
                    foreach (string cell in lvl1processableCellsEOS)
                    {
                        Excel.Range rng = EOS.Range[cell];
                        lvl1rollupWorkbook.Worksheets.Item[1].cells(lvl1lastRow + 1, i).value = rng.Value;
                        i++;
                    }

                    ////process all of the lvl2 roll up values
                    //for( int x = 1; x<13;  x++)// work down the categories
                    //{
                    //    string currentCategoryProcessing = "";
                    //    string categoryCheck = HourlyStaffing.Cells[x+4, 3].value;
                    //    if(categoryCheck!= null)
                    //    {

                    //    }

                    //}

                }


                if (Directory.Exists(destloc))
                {
                    if (File.Exists(destFile))
                    {
                        System.IO.File.Delete(destFile);
                    }
                    if (moveFlag)
                    {
                        System.IO.File.Move(archiveSrcFile, destFile);
                    }
                    else
                    {
                        try
                        {
                            System.IO.File.Copy(archiveSrcFile, destFile, true);
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                else
                {
                    System.IO.Directory.CreateDirectory(destloc);
                    if (moveFlag)
                    {
                        System.IO.File.Move(archiveSrcFile, destFile);
                    }
                    else
                    {
                        System.IO.File.Copy(archiveSrcFile, destFile, true);
                    }

                }
            }
            //end
            lvl1rollupWorkbook.Close(true, Lvl1rollup);
            lvl1rollupWorkbook = null;
            lvl2rollupWorkbook = null;
            lvl3rollupWorkbook = null;
            file = null;
            workbook = null;
            app.Quit();
            app = null;
        }
        //test
    }

}

