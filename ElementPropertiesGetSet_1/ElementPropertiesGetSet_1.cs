/*
****************************************************************************
*  Copyright (c) 2024,  Skyline Communications NV  All Rights Reserved.    *
****************************************************************************

By using this script, you expressly agree with the usage terms and
conditions set out below.
This script and all related materials are protected by copyrights and
other intellectual property rights that exclusively belong
to Skyline Communications.

A user license granted for this script is strictly for personal use only.
This script may not be used in any way by anyone without the prior
written consent of Skyline Communications. Any sublicensing of this
script is forbidden.

Any modifications to this script by the user are only allowed for
personal use and within the intended purpose of the script,
and will remain the sole responsibility of the user.
Skyline Communications will not be responsible for any damages or
malfunctions whatsoever of the script resulting from a modification
or adaptation by the user.

The content of this script is confidential information.
The user hereby agrees to keep this confidential information strictly
secret and confidential and not to disclose or reveal it, in whole
or in part, directly or indirectly to any person, entity, organization
or administration without the prior written consent of
Skyline Communications.

Any inquiries can be addressed to:

	Skyline Communications NV
	Ambachtenstraat 33
	B-8870 Izegem
	Belgium
	Tel.	: +32 51 31 35 69
	Fax.	: +32 51 31 01 29
	E-mail	: info@skyline.be
	Web		: www.skyline.be
	Contact	: Ben Vandenberghe

****************************************************************************
Revision History:

DATE		VERSION		AUTHOR			COMMENTS

24/10/2024	1.0.0.1		NVC, Skyline	Initial version
****************************************************************************
*/

namespace ElementPropertiesGetSet_1
{
    using Skyline.DataMiner.Automation;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    /// <summary>
    /// Represents a DataMiner Automation script.
    /// </summary>
    public class Script
    {
        /// <summary>
        /// The script entry point.
        /// </summary>
        /// <param name="engine">Link with SLAutomation process.</param>
        public void Run(IEngine engine)
        {
            try
            {
                RunSafe(engine);
            }
            catch (ScriptAbortException)
            {
                // Catch normal abort exceptions (engine.ExitFail or engine.ExitSuccess)
                throw; // Comment if it should be treated as a normal exit of the script.
            }
            catch (ScriptForceAbortException)
            {
                // Catch forced abort exceptions, caused via external maintenance messages.
                throw;
            }
            catch (ScriptTimeoutException)
            {
                // Catch timeout exceptions for when a script has been running for too long.
                throw;
            }
            catch (InteractiveUserDetachedException)
            {
                // Catch a user detaching from the interactive script by closing the window.
                // Only applicable for interactive scripts, can be removed for non-interactive scripts.
                throw;
            }
            catch (Exception e)
            {
                engine.ExitFail("Run|Something went wrong: " + e);
            }
        }

        private void RunSafe(IEngine engine)
        {
            string path = @"c:\Skyline_Data\elementProperties.csv";
            var elementsFromCsv = ElementsFromCsv(path, engine);

            foreach (var element in elementsFromCsv)
            {
                var elementOnCurrentDMA = engine.FindElement(element);

                if (elementOnCurrentDMA.IsActive == false)
                {
                    continue;
                }
                else
                {
                    var propertyForSelectedElement = GetPropertyForElement(element, engine, path);
                    FillPropertiesFields(engine, elementOnCurrentDMA, propertyForSelectedElement);
                }
            }
        }

        private void FillPropertiesFields(IEngine engine, Element element, List<string> properties)
        {
            foreach (string property in properties)
            {
                if (property == null)
                {
                    continue;
                }

                if (property.Contains("Info"))
                {
                    element.SetPropertyValue("info", property.Split('/')[1]);
                }
                else if (property.Contains("Location"))
                {
                    element.SetPropertyValue("location", property.Split('/')[1]);
                }
                else if (property.Contains("EMB Matrix Label"))
                {
                    element.SetPropertyValue("EMB Matrix Label", property.Split('/')[1]);
                }
                else if (property.Contains("FrontRear"))
                {
                    element.SetPropertyValue("FrontRear", property.Split('/')[1]);
                }
                else if (property.Contains("RU Height"))
                {
                    element.SetPropertyValue("RU Height", property.Split('/')[1]);
                }
                else if (property.Contains("RU Position"))
                {
                    element.SetPropertyValue("RU Position", property.Split('/')[1]);
                }
                else
                {
                    continue;
                }
            }
        }

        private List<string> GetPropertyForElement(string element, IEngine engine, string path)
        {
            string[] readText = File.ReadAllLines(path);
            List<string> properties = new List<string>();

            foreach (string line in readText)
            {
                if (line.Split(':')[0] == element)
                {
                    properties.Add(line.Split(':')[1]);
                }
            }

            return properties;
        }

        private List<string> ElementsFromCsv(string path, IEngine engine)
        {
            string[] readText = File.ReadAllLines(path);
            List<string> elements = new List<string>();

            foreach (string line in readText)
            {
                var elementName = line.Split(':')[0];
                elements.Add(elementName);
            }

            var distinctList = elements.Distinct().ToList();

            return distinctList;
        }
    }
}