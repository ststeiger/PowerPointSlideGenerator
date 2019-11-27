
using System.Drawing;

namespace PowerPointFromImageFolder
{
    
    
    class Program
    {
        
        static void Main(string[] args)
        {
            // ImageAdd.TestEnumerator();
            
            string imageFolder = @"D:\username\Pictures\majdu\Steve";
            imageFolder = @"/root/Pictures/";
            
            string outputFile = @"D:\PictureGallery.pptx";
            outputFile = @"/root/Pictures/PictureGallery.pptx";
            
            PowerPointHelper.CreatePresentationFromImageFolder(outputFile, imageFolder);
            
            System.Console.WriteLine(" --- Press any key to continue --- ");
            System.Console.ReadKey();
        } // End Sub Main 
        
        
    } // End Class Program 
    
    
} // End Namespace 
