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



namespace Prepal_Back_End_
{
    /*This class approaches the purpose of prepal using continuos recognition.*/
    internal class ContinuousRecognition
    {
        async static Task Main(string[] args)
        {
            var program = new LoopedRecognition();

            var speechConfig = SpeechConfig.FromSubscription("key", "area");
            var pptApp = (PowerPoint.Application)
            Marshal.GetActiveObject("PowerPoint.Application");

            //We first check if the slideshow is running and ready for prepal.
            if (pptApp == null && pptApp.SlideShowWindows.Count == 0)
            {
                Console.WriteLine("No slideshow is running right now.");
                return;
            }
            else
            {
                var slideShow = pptApp.SlideShowWindows[1].View;
                using (var audioConfig = AudioConfig.FromDefaultMicrophoneInput())
                using (var recognizer = new SpeechRecognizer(speechConfig, audioConfig))
                {
                    //In continuous recognition we check for every possible outocome. Starting with text being recognized.
                    var stopRecognition = new TaskCompletionSource<int>();
                    recognizer.Recognizing += (s, e) =>
                    {
                        string textBeingRecognized = e.Result.Text;
                    };

                    //Recognized text, where prepal will take action.
                    recognizer.Recognized += (s, e) =>
                    {
                        if (e.Result.Reason == ResultReason.RecognizedSpeech)
                        {
                            string recognizedText = e.Result.Text;
                            string test = program.returnKeyword(recognizedText);
                            if(test== "prepal next")
                            {
                                slideShow.Next();
                                Console.WriteLine("Moved to the next slide.");
                            }
                            else if(test== "prepal end")
                            {
                                slideShow.Exit();
                                Console.WriteLine("Ended the slideshow.");
                            }
                        }
                        else if (e.Result.Reason == ResultReason.NoMatch)
                        {
                            Console.WriteLine("Speech couldn't be recognized.");
                        }
                    };
                    //Cancelation of recognition for all different reasons.
                    recognizer.Canceled += (s, e) =>
                    {
                        Console.WriteLine($"Canceed: Reason= {e.Reason}");
                        if (e.Reason == CancellationReason.Error)
                        {
                            Console.WriteLine($"Error Code = {e.ErrorCode}");
                            Console.WriteLine($"Error Details = {e.ErrorDetails}");
                        }

                        stopRecognition.TrySetResult(0);
                    };

                    //When the API stops recognizing.
                    recognizer.SessionStopped += (s, e) =>
                    {
                        Console.WriteLine("Session Stopped");
                        stopRecognition.TrySetResult(0);
                    };

                    await recognizer.StartContinuousRecognitionAsync();
                    Task.WaitAny(new[] { stopRecognition.Task });

                }

            }
        }
    }
}
