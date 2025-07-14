using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Core = Microsoft.Office.Core;


/*This class approaches the purpose of prepal using single-shot voice recognition implemented 
 inside a while loop. */
namespace Prepal_Back_End_
{
    internal class LoopedRecognition
    {
        /*This is an internal method to use for proccessing the recieved transcribed text
         and find whether the anticipated keyword was said or not. */

        public string returnKeyword (string resultText)
        {
            resultText = resultText.Trim().ToLowerInvariant();
            string lookingFor = "prepal next";
            string orFor = "prepal end";
            if(resultText.IndexOf(lookingFor, StringComparison.Ordinal) != 1)
            {
                return resultText.Substring(resultText.IndexOf(lookingFor, StringComparison.Ordinal) + 1, lookingFor.Length);
            }
            else if (resultText.IndexOf(orFor, StringComparison.Ordinal) != 1)
            {
                return resultText.Substring(resultText.IndexOf(orFor, StringComparison.Ordinal) + 1, lookingFor.Length);
            }
            else
            {
                return "";
            }
        }

        //The main method, where the main action happens.
        async static Task Main(string[] args)
        {
            var program = new LoopedRecognition();

            //Create an instance of the speech configration to use the Azure API 
            var speechConfig = SpeechConfig.FromSubscription("key", "area");

            //Calling the PowerPoint application listner
            var pptApp = (PowerPoint.Application)
            Marshal.GetActiveObject("PowerPoint.Application");

            //An if statement that checks if the listner is running and if the slideshow started, if not the application is terminated
            if (pptApp == null && pptApp.SlideShowWindows.Count == 0)
            {
                Console.WriteLine("Mo slideshow is running.");
                return;
            }
            var slideShow = pptApp.SlideShowWindows[1].View;
            Console.WriteLine("Listening for the keyword...");


            // We use the transcribtion API and the previously defined internal method to look for the moving and endinig the slideshow keywords.
            using (var audioConfig = AudioConfig.FromDefaultMicrophoneInput())
            using (var recognizer = new SpeechRecognizer(speechConfig, audioConfig))
            {
                recognizer.Recognized += (s, e) =>
                {
                    string result = e.Result.Text;
                    string NeededResult = program.returnKeyword(result);

                    if (result == null || result == "")
                    {
                        return;
                    }
                    switch (result)
                    {
                        case "prepal next":
                            slideShow.Next();
                            Console.WriteLine("Moved to the net slide successfully");
                            break;
                        case "prepal end":
                            slideShow.Exit();
                            Console.WriteLine("Ended the slideshow successfully");
                            break;
                    }
                };
                await recognizer.StartContinuousRecognitionAsync();
                await recognizer.StopContinuousRecognitionAsync();
            }

        }
    }
}
