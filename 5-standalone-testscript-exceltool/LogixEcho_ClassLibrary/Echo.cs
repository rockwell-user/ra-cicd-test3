// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:    Echo.cs
// FileType:    Visual C# Source File
// Author:      Rockwell Automation
// Created:     2024
// Description: This script sets up an emulated controller & chassis using the Factory Talk Logix Echo SDK.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using RockwellAutomation.FactoryTalkLogixEcho.Api.Client;
using RockwellAutomation.FactoryTalkLogixEcho.Api.Interfaces;
using System.Globalization;

namespace LogixEcho_ClassLibrary
{
    /// <summary>
    /// Class containing Factory Talk Logix Echo SDK methods needed for CI/CD execution.
    /// </summary>
    public class LogixEchoMethods
    {
        /// <summary>
        /// Script that programmatically sets up an emulated controller and chassis based on the user inputs.<br/>
        /// If no emulated controller based on the ACD file path yet exists, create one, and then return the communication path.<br/>
        /// If an emulated controller based on the ACD file path exists, only return the communication path.
        /// </summary>
        /// <param name="acdFilePath">The file path pointing to the ACD project used for testing.</param>
        /// <param name="chassisName">The name of the chassis to be created if it does not yet exist or to be checked for the specified controller.</param>
        /// <param name="controllerName">The name of the controller to be either created or whose communication path is to be returned.</param>
        /// <returns>A string containing the communication path of the emulated controller that the ACD file will go online with during testing.</returns>
        public static async Task<string> Main(string acdFilePath, string chassisName, string controllerName)
        {
            // Set up initial variables:
            var serviceClient = ClientFactory.GetServiceApiClientV2("CLIENT_TestStage_CICDExample"); // Factory Talk Logix Echo SDK service client
            serviceClient.Culture = new CultureInfo("en-US");                                        // client language set to english
            ChassisData? chassisCICD = new ChassisData();                                            // Chassis variable

            // Check if an emulated chassis exists with the Main method's input 'chassisName'. If not, run through the if statement contents to create one.
            if (CheckCurrentChassis_Async(chassisName, serviceClient).GetAwaiter().GetResult() == false)
            {
                // Set up new emulated chassis information.
                var chassisUpdate = new ChassisUpdate
                {
                    Name = chassisName,
                    Description = "Test chassis for CI/CD demonstration."
                };

                chassisCICD = await serviceClient.CreateChassis(chassisUpdate);
            }

            // Check if an emulated controller exists within an emulated chassis. If not, run through the if statement contents to create one.
            if (CheckCurrentController_Async(chassisName, controllerName, serviceClient).GetAwaiter().GetResult() == false)
            {
                // Get chassis data to be used for the controller specified
                chassisCICD = await GetChassisData_Async(chassisName, serviceClient);

                // Set up emulated controller information.
                using (var fileHandle = await serviceClient.SendFile(acdFilePath))
                {
                    var controllerUpdate = await serviceClient.GetControllerInfoFromAcd(fileHandle);

                    if (chassisCICD != null)
                        controllerUpdate.ChassisGuid = chassisCICD.ChassisGuid;

                    await serviceClient.CreateController(controllerUpdate);
                }
            }

            // Get emulated controller information.
            ControllerData testControllerDate = await GetControllerData_Async(chassisName, controllerName, serviceClient);

            // Create the communication path needed for proper emmulated controller connnections.
            string commPath = @"EmulateEthernet\" + testControllerDate!.IPConfigurationData.Address.ToString() ?? "";

            return commPath;
        }

        #region METHODS: FactoryTalk Logix Echo SDK methods
        /// <summary>
        /// Asynchronously get the ChassisData chassis variable using the chassis's name.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to get the ChassisData variable from.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>The variable ChassisData for the chassis name specified.</returns>
        private static async Task<ChassisData> GetChassisData_Async(string chassisName, IServiceApiClientV2 serviceClient)
        {

            ChassisData returnChassisData = new();
            // Get the list of chassis currently created and iterate through them until the desired chassis is selected, then return it. 
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                if (chassisList[i].Name == chassisName)
                    returnChassisData = chassisList[i];
            }
            return returnChassisData;
        }

        /// <summary>
        /// Asynchronously get the ControllerData controller variable using the desired controller name & chassis.
        /// </summary>
        /// <param name="chassisName">The emulated chassis to the emulatedcontroller information from.</param>
        /// <param name="controllerName">The emulated controller name.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>The variable ControllerData for the chassis name specified.</returns>
        private static async Task<ControllerData> GetControllerData_Async(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            ControllerData returnControllerData = new();

            // Get the list of chassis currently created and iterate through them until the desired chassis is selected.
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                if (chassisList[i].Name == chassisName)
                {
                    // Get the list of chassis currently created and iterate through them until the desired controller is selected, then return.
                    var controllerList = (await serviceClient.ListControllers(chassisList[i].ChassisGuid)).ToList();
                    for (int j = 0; j < controllerList.Count; j++)
                    {
                        if (controllerList[j].ControllerName == controllerName)
                            returnControllerData = controllerList[j];
                    }
                }
            }
            return returnControllerData;
        }

        /// <summary>
        /// Asynchronously check to see if a specific controller exists in a specific chassis.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to check the emulated controler in.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>A Task that returns a boolean value <c>true</c> if the emulated controller already exists and a <c>false</c> if it does not.</returns>
        private static async Task<bool> CheckCurrentChassis_Async(string chassisName, IServiceApiClientV2 serviceClient)
        {
            // Get the list of chassis currently created and iterate through them. If a chassis in the list has the same name as the input string, return true.
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                if (chassisList[i].Name == chassisName)
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Asynchronously check to see if a specific controller exists in a specific chassis.
        /// </summary>
        /// <param name="chassisName">The name of the emulated chassis to check the emulated controler in.</param>
        /// <param name="controllerName">The name of the emulated controller to check.</param>
        /// <param name="serviceClient">The Factory Talk Logix Echo interface.</param>
        /// <returns>A Task that returns a boolean value <c>true</c> if the emulated controller already exists and a <c>false</c> if it does not.</returns>
        private static async Task<bool> CheckCurrentController_Async(string chassisName, string controllerName, IServiceApiClientV2 serviceClient)
        {
            // Get the list of chassis currently created and iterate through them.
            var chassisList = (await serviceClient.ListChassis()).ToList();
            for (int i = 0; i < chassisList.Count; i++)
            {
                // If a chassis in the list has the same name as the input string, continue to checking controllers. 
                if (chassisList[i].Name == chassisName)
                {
                    // Get the list of controllers currently created and iterate through them.
                    var controllerList = (await serviceClient.ListControllers(chassisList[i].ChassisGuid)).ToList();
                    for (int j = 0; j < controllerList.Count; j++)
                    {
                        // If a controller in the list has the same name as the input string, return true. 
                        if (controllerList[j].ControllerName == controllerName)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Delete the specified chassis and any controllers that exist within it.
        /// </summary>
        /// <param name="chassisName">The name of the chassis to be deleted</param>
        public static async Task DeleteChassis_Async(string chassisName)
        {
            // Create a new instance of the Echo service client to use for chassis deletion.
            var serviceClient = ClientFactory.GetServiceApiClientV2("Deleter Client");

            // Get the specified chassis' information in the ChassisData variable.
            ChassisData chassisToDelete = await GetChassisData_Async(chassisName, serviceClient);

            // Get the list of controllers from the specified chassis.
            var controllerList = (await serviceClient.ListControllers(chassisToDelete!.ChassisGuid)).ToList();

            // Iterate through any existing controllers and delete them. A chassis can't be deleted if it still has controllers in it.
            for (int i = 0; i < controllerList.Count; i++)
            {
                await serviceClient.DeleteController(controllerList[i].ControllerGuid);
            }

            await serviceClient.DeleteChassis(chassisToDelete.ChassisGuid);
        }
        #endregion
    }
}