﻿
namespace PowerPointFromImageFolder
{
    
    
    class Program
    {
        
        static void Main(string[] args)
        {
            // ImageAdd.TestEnumerator();
            
            string imageFolder = @"D:\username\Pictures\majdu\Steve";
            if(System.Environment.OSVersion.Platform == System.PlatformID.Unix)
                imageFolder = @"/home/username/Pictures/";
            
            string outputFile = @"D:\PictureGallery.pptx";
            if (System.Environment.OSVersion.Platform == System.PlatformID.Unix)
                outputFile = @"/home/username/PictureGallery.pptx";
            
            PowerPointHelper.CreatePresentationFromImageFolder(outputFile, imageFolder, EmuPaperSize.Film35mm);
            
            System.Console.WriteLine(" --- Press any key to continue --- ");
            System.Console.ReadKey();
        } // End Sub Main 
        
        
    } // End Class Program 
    
    
} // End Namespace 
