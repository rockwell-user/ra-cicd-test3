// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:     DeploymentProgram.cs
// FileType:     Visual C# Source file
// Author:       Rockwell Automation Engineering
// Created:      2024
// Description:  This script programmatically verifies and flashes the firmware of target modules from an input excel sheet.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using OfficeOpenXml;
using RockwellAutomation.LogixDesigner;
using System.Diagnostics;
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;

namespace CD_Deployment
{
    /// <summary>
    /// This class contains the methods and logic to programmatically verify and flash the firmware of target modules from an input excel sheet.
    /// </summary>
    public class FlashControllers
    {
        // Static variable for the character length limit of each line printed to the console.
        public static readonly int consoleCharLengthLimit = 110;
        public static readonly DateTime startTime = DateTime.Now; /* --------------------- The time during which this test was first initiated. 
                                                                                               (Used at end of test to calculate unit test length.) */
        public static readonly string currentDateTime = startTime.ToString("yyyyMMddHHmmss"); /* Time during which test was first initiated, as a string.
                                                                                                     (Used to name generated files & test reports.)*/
        /// <summary>
        /// A structure containing all the information needed to verify and flash a madule using the LDSDK and CF+SDK.
        /// </summary>
        public struct ModuleInfo
        {
            public string Type { get; set; }
            public string CommPath { get; set; }
            public string TargetRevision { get; set; }
            public ModuleInfo()
            {
                Type = string.Empty;
                CommPath = string.Empty;
                TargetRevision = string.Empty;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputArg_inputExcelFilePath">The input excel workbook file path. <br/>
        /// (This file defines what modules will have firmware verified and flashed if needed).</param>
        /// <returns></returns>
        static async Task Main(string[] args)
        {
            // Print unit test banner to the console.
            Console.WriteLine("\n  ".PadRight(consoleCharLengthLimit - 2, '='));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            string bannerContents = "MODULE FIRMWARE VERIFICATION & FLASHING | " + DateTime.Now + " " + TimeZoneInfo.Local;
            int padding = (consoleCharLengthLimit - bannerContents.Length) / 2;
            Console.WriteLine(bannerContents.PadLeft(bannerContents.Length + padding).PadRight(consoleCharLengthLimit));
            Console.WriteLine("".PadRight(consoleCharLengthLimit, '='));
            Console.WriteLine("  ".PadRight(consoleCharLengthLimit - 2, '=') + "\n");

            // Parse the input excel sheet needed to determine which modules are to have their firmware verified & flashed if needed.
            string inputArg_inputExcelFilePath = args[0];

            string inputArg_uploadedFilePath = args[1];

            // This list will store all the modules passed in from the excel sheet.
            List<ModuleInfo> moduleList = new List<ModuleInfo>();

            int moduleCount = GetPopulatedCellsInColumnCount(inputArg_inputExcelFilePath, 2) - 2;


            // Parse excel sheet for modules to flash.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputArg_inputExcelFilePath)))
            {
                ExcelWorksheet inputExcelWorksheet = package.Workbook.Worksheets.FirstOrDefault()!;

                for (int row = 7; row < moduleCount + 7; row++)
                {
                    string currentModuleType = inputExcelWorksheet.Cells[row, 2].Value.ToString()!.Trim()!;
                    string currentModuleCommPath = inputExcelWorksheet.Cells[row, 3].Value.ToString()!.Trim()!;
                    string currentModuleTargetRevision = inputExcelWorksheet.Cells[row, 4].Value.ToString()!.Trim()!;
                    ModuleInfo currentModule = new ModuleInfo
                    {
                        Type = currentModuleType,
                        CommPath = currentModuleCommPath,
                        TargetRevision = currentModuleTargetRevision,
                    };
                    moduleList.Add(currentModule);
                }
            }

            LogixProject project;
            string currentmode = string.Empty;
            string newVersionACDFilePath = string.Empty;

            foreach (var module in moduleList)
            {
                ConsoleMessage($"STARTING the verification & flashing of the '{module.Type}' module located at '{module.CommPath}'.", "NEWSECTION", false);

                #region Controller firmware verification & download.
                if (CheckIfModuleIsController(module.Type))
                {
                    // Create the file path used to upload a project.
                    string uploadACDFilePath = inputArg_uploadedFilePath + @"\" + currentDateTime + "_" + ReplaceSpecialCharacters(module.CommPath) + ".acd";

                    ConsoleMessage($"START uploading the target controller's application to '{uploadACDFilePath}'.", "STATUS");

                    // Upload the project in the currently target controller
                    try
                    {
                        LogixProject logixproject = await LogixProject.UploadToNewProjectAsync(uploadACDFilePath, module.CommPath);
                        ConsoleMessage("Upload to New Project Complete.", "STATUS");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        ConsoleMessage("Upload Failed. Aborting.", "ERROR");
                    }

                    // Change controller to 'Program' mode if it isn't already set to program.
                    ConsoleMessage("Verifying if controller mode set to 'Program'.", "STATUS");
                    try
                    {
                        project = await LogixProject.OpenLogixProjectAsync(uploadACDFilePath);
                        await project.SetCommunicationsPathAsync(module.CommPath);
                        LogixProject.ControllerMode controllerMode = await project.ReadControllerModeAsync();
                        if (controllerMode != LogixProject.ControllerMode.Program)
                        {
                            ConsoleMessage("Setting controller mode to 'Program'.", "STATUS");
                            await project.ChangeControllerModeAsync(LogixProject.RequestedControllerMode.Program);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        ConsoleMessage("\r\nError changing controller mode! Aborting!", "ERROR");
                        break;
                    }


                    // Pack the ControlFlash_SDK executable arguments into an array.
                    string[] parameters = { module.CommPath, module.TargetRevision };

                    ConsoleMessage("Start executing the ControlFLASH Plus SDK executable that varifies and flashes the module.", "STATUS");
                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        FileName = @"C:\Lab Files\Device Management\ControlFlash_SDK\ControlFlash_SDK\bin\x86\Release\net48\ControlFlash_SDK.exe",
                        Arguments = string.Join(" ", parameters),
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true,
                    };

                    using (Process process = new Process { StartInfo = startInfo })
                    {
                        process.OutputDataReceived += (sender, e) =>
                        {
                            if (!string.IsNullOrEmpty(e.Data))
                            {
                                Console.WriteLine(e.Data);
                            }
                        };
                        process.ErrorDataReceived += (sender, e) =>
                        {
                            if (!string.IsNullOrEmpty(e.Data))
                            {
                                Console.WriteLine("Error: " + e.Data);
                            }
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();
                        process.WaitForExit();
                    }

                    // Convert project to the targetted version.
                    ConsoleMessage($"Converting the application at '{uploadACDFilePath}' to version '{GetStringPart(module.TargetRevision, "LEFT")}'.", "STATUS");
                    try
                    {
                        project = await LogixProject.ConvertAsync(uploadACDFilePath, Convert.ToInt32(GetStringPart(module.TargetRevision, "LEFT")));
                        newVersionACDFilePath = inputArg_uploadedFilePath + @"\" + currentDateTime + "_" + ReplaceSpecialCharacters(module.CommPath) + "_v" + GetStringPart(module.TargetRevision, "LEFT") + ".acd";
                        await project.SaveAsAsync(newVersionACDFilePath, true);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        ConsoleMessage("Project Conversion Failed. Aborting.", "ERROR");
                    }

                    // Download the converted project to the newly flashed controller.
                    ConsoleMessage($"Downloading the newly converted project at '{newVersionACDFilePath}' to '{module.CommPath}'.", "STATUS");
                    try
                    {
                        project = await LogixProject.OpenLogixProjectAsync(newVersionACDFilePath);
                        await project.SetCommunicationsPathAsync(module.CommPath);
                        LogixProject.ControllerMode controllerMode = await project.ReadControllerModeAsync();

                        if (controllerMode != LogixProject.ControllerMode.Program)
                            await project.ChangeControllerModeAsync(LogixProject.RequestedControllerMode.Program);

                        await project.DownloadAsync();
                        await project.SaveAsync();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        ConsoleMessage("Download Failed. Aborting.", "ERROR");
                    }

                    ConsoleMessage($"Setting the '{module.Type}' controller at '{module.CommPath}' back to 'Run' mode.", "STATUS");
                    try
                    {
                        await project.ChangeControllerModeAsync(LogixProject.RequestedControllerMode.Run);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        ConsoleMessage("Change to 'Run' mode Failed. Aborting.", "ERROR");
                    }

                }
                #endregion

                #region Non-controller firmware verification & download.
                else
                {
                    // Pack the ControlFlash_SDK executable arguments into an array.
                    string[] parameters = { module.CommPath, module.TargetRevision };

                    ProcessStartInfo startInfo = new ProcessStartInfo
                    {
                        FileName = @"C:\Lab Files\Device Management\ControlFlash_SDK\ControlFlash_SDK\bin\x86\Release\net48\ControlFlash_SDK.exe",
                        Arguments = string.Join(" ", parameters),
                        UseShellExecute = false,
                        RedirectStandardError = true,
                        RedirectStandardOutput = true,
                    };

                    using (Process process = new Process { StartInfo = startInfo })
                    {
                        process.OutputDataReceived += (sender, e) =>
                        {
                            if (!string.IsNullOrEmpty(e.Data))
                            {
                                Console.WriteLine(e.Data);
                            }
                        };
                        process.ErrorDataReceived += (sender, e) =>
                        {
                            if (!string.IsNullOrEmpty(e.Data))
                            {
                                Console.WriteLine("Error: " + e.Data);
                            }
                        };

                        process.Start();
                        process.BeginOutputReadLine();
                        process.BeginErrorReadLine();
                        process.WaitForExit();
                    }
                }
                #endregion
            }
        }

        #region METHODS
        /// <summary>
        /// Check if the type of module is a controller.
        /// </summary>
        /// <param name="moduleType">A string with the module type.</param>
        /// <returns>A boolean 'true' if the input module type was for a controller, otherwise 'false'.</returns>
        private static bool CheckIfModuleIsController(string moduleType)
        {
            if (moduleType == "1756-L75")
                return true;
            else if (moduleType == "1756-L85E")
                return true;
            else
                return false;
        }

        /// <summary>
        /// In the first worksheet of an Excel workbook, get the number of populated cells in the specified column.
        /// </summary>
        /// <param name="excelFilePath">The excel workbook file path.</param>
        /// <param name="columnNumber">The column in which the populated cell count is derived.</param>
        /// <returns>The number of populated cells in the specified column.</returns>
        private static int GetPopulatedCellsInColumnCount(string excelFilePath, int columnNumber)
        {
            int returnCellCount = 0;
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault()!;
                int maxRowNum = worksheet.Dimension.End.Row;

                for (int row = 1; row <= maxRowNum; row++)
                {
                    var cellValue = worksheet.Cells[row, columnNumber].Value;

                    if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        returnCellCount++;
                }
            }
            return returnCellCount;
        }

        /// <summary>
        /// Helper method to remove the special characters in the module communication path.<br/>
        /// </summary>
        /// <param name="input">Input string from which to replace '!', '\', '-', or '.' characters with '_'.</param>
        /// <returns>A string that can be used as a componenet of a file name.</returns>
        private static string ReplaceSpecialCharacters(string input)
        {
            // Define the characters to be replaced
            char[] specialChars = { '!', '\\', '-', '.' };

            // Replace each special character with '_'
            foreach (char c in specialChars)
            {
                input = input.Replace(c, '_');
            }

            return input;
        }

        /// <summary>
        /// Helper method to split the software revision string as needed.
        /// </summary>
        /// <param name="inputString">The string to be separated.</param>
        /// <param name="side">A string specifying the 'LEFT' or 'RIGHT' side of string to be returned.</param>
        /// <returns>A string that returns the 'LEFT' or 'RIGHT' side of the input string, split at the first period.</returns>
        private static string GetStringPart(string inputString, string side)
        {
            int periodIndex = inputString.IndexOf('.');
            side = side.ToUpper().Trim();

            if (side == "LEFT")
                return inputString.Substring(0, periodIndex);
            else if (side == "RIGHT")
                return inputString.Substring(periodIndex + 1);
            else
                return inputString;
        }
        #endregion
    }
}