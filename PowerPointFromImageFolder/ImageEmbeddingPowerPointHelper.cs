
namespace PowerPointFromImageFolder
{

    using System.Linq;
    
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using a = DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Presentation;
    
    
    // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/ee412267(v=office.12)?redirectedfrom=MSDN
    public partial class PowerPointHelper
    {
        
        // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
        // A DrawingML graphic's dimensions are specified in English Metric Units (EMUs).
        // It is so called because it allows an exact common representation of dimensions originally in either English or Metric units.
        // This unit is defined as 1/360,000 of a centimeter and thus there are 914,400 EMUs per inch, and 12,700 EMUs per point. 
        public static int ToEmu(int pixel, int ppi)
        {
            return pixel * 914400 / ppi; 
        }    

        
        // 10692000 = 29.7cm
        // 7560000 = 21cm
        
        public static int ToEmu(int pixel)
        {
            return ToEmu(pixel, 96);
        }
        
        
        public static void CreatePresentationFromImageFolder(string outputFile, string imageFolder)
        {
            // string presentationFolder = @"D:\";
            // string presentationTemplate = "PresentationTemplate.pptx";
            // PowerPointHelper.CreatePresentation(presentationFolder + presentationTemplate);
            // Make a copy of the template presentation. 
            // File.Copy(presentationFolder + presentationTemplate, outputFile, true);
            
            // Just create an empty powerpoint-file instead 
            PowerPointHelper.CreatePresentation(outputFile);
            
            
            string[] imageFileExtensions = new string[]
            {
                "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.png", "*.tif"
            };
            
            // Get the image files in the image folder.
            System.Collections.Generic.List<string> imageFileNames = GetImageFileNames(imageFolder,
              imageFileExtensions);
            
            // Create new slides for the images.
            if (imageFileNames.Count > 0)
                CreateSlides(imageFileNames, outputFile);
            
            // Validate the new presentation.
            DocumentFormat.OpenXml.Validation.OpenXmlValidator validator = 
                new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
            
            using (DocumentFormat.OpenXml.Packaging.PresentationDocument presentation =
                DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputFile, true))
            {
                System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> errors = validator.Validate(presentation);
                
                if (errors.Count() > 0)
                {
                    System.Console.WriteLine("The deck creation process completed but " +
                      "the created presentation failed to validate.");
                    System.Console.WriteLine("There are " + errors.Count() +
                      " errors:\r\n");
                    
                    DisplayValidationErrors(errors);
                }
                else
                    System.Console.WriteLine("The deck creation process completed and " +
                      "the created presentation validated with 0 errors.");
            } // End Using presentation 
            
        } // End Sub Test 
        
        
        static void CreateSlides(
              System.Collections.Generic.List<string> imageFileNames
            , string newPresentation)
        {
            string relId;
            SlideId slideId;

            // Slide identifiers have a minimum value of greater than or
            // equal to 256 and a maximum value of less than 2147483648.
            // Assume that the template presentation being used has no slides.
            uint currentSlideId = 256;
            // uint currentSlideId = 256+1;

            string imageFileNameNoPath;

            long imageWidthEMU = 0;
            long imageHeightEMU = 0;

            // Open the new presentation.
            using (PresentationDocument newDeck =
              PresentationDocument.Open(newPresentation, true))
            {
                PresentationPart presentationPart = newDeck.PresentationPart;

                // Reuse the slide master part. This code assumes that the
                // template presentation being used has at least one
                // master slide.
                SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();

                // Reuse the slide layout part. This code assumes that the
                // template presentation being used has at least one
                // slide layout.
                SlideLayoutPart slideLayoutPart = slideMasterPart.SlideLayoutParts.First();

                // If the new presentation doesn't have a SlideIdList element
                // yet then add it.
                if (presentationPart.Presentation.SlideIdList == null)
                    presentationPart.Presentation.SlideIdList = new SlideIdList();

                // Loop through each image file creating slides
                // in the new presentation.
                foreach (string imageFileNameWithPath in imageFileNames)
                {
                    imageFileNameNoPath =
                      System.IO.Path.GetFileNameWithoutExtension(imageFileNameWithPath);

                    // Create a unique relationship id based on the current
                    // slide id.
                    relId = "rel" + currentSlideId;

                    // Get the bytes, type and size of the image.
                    ImagePartType imagePartType = ImagePartType.Png;
                    byte[] imageBytes = GetImageData(imageFileNameWithPath,
                      ref imagePartType, ref imageWidthEMU, ref imageHeightEMU);

                    // Create a slide part for the new slide.
                    SlidePart slidePart = presentationPart.AddNewPart<SlidePart>(relId);
                    GenerateSlidePart(imageFileNameNoPath, imageFileNameNoPath,
                      imageWidthEMU, imageHeightEMU).Save(slidePart);

                    // Add the relationship between the slide and the
                    // slide layout.
                    slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);

                    // Create an image part for the image used by the new slide.
                    // A hardcoded relationship id is used for the image part since
                    // there is only one image per slide. If more than one image
                    // was being added to the slide an approach similar to that
                    // used above for the slide part relationship id could be
                    // followed, where the image part relationship id could be
                    // incremented for each image part.
                    ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg, "relId1");
                    

                    GenerateImagePart(imagePart, imageBytes);

                    // Add the new slide to the slide list.
                    slideId = new SlideId();
                    slideId.RelationshipId = relId;
                    slideId.Id = currentSlideId;
                    presentationPart.Presentation.SlideIdList.Append(slideId);

                    // Increment the slide id;
                    currentSlideId++;
                } // Next imageFileNameWithPath 

                // Save the changes to the slide master part.
                slideMasterPart.SlideMaster.Save();

                // Save the changes to the new deck.
                presentationPart.Presentation.Save();
            } // End Using newDeck 

        } // End Sub CreateSlides 


        private static System.Collections.Generic.List<string> GetImageFileNames(string imageFolder,
          string[] imageFileExtensions)
        {
            // Create a list to hold the names of the files with the
            // requested extensions.
            System.Collections.Generic.List<string> fileNames = 
                new System.Collections.Generic.List<string>();

            // Loop through each file extension.
            foreach (string extension in imageFileExtensions)
            {
                // Add all the files that match the current extension to the
                // list of file names.
                fileNames.AddRange(System.IO.Directory.GetFiles(imageFolder, extension,
                    System.IO.SearchOption.TopDirectoryOnly));
            } // Next extension 

            // Return the list of file names.
            return fileNames;
        } // End Function GetImageFileNames 


        private static byte[] GetImageData(string imageFilePath,
          ref ImagePartType imagePartType, ref long imageWidthEMU,
          ref long imageHeightEMU)
        {
            byte[] imageFileBytes;
            // Bitmap imageFile;

            // Open a stream on the image file and read it's contents. The
            // following code will generate an exception if an invalid file
            // name is passed.
            using (System.IO.FileStream fsImageFile = System.IO.File.OpenRead(imageFilePath))
            {
                imageFileBytes = new byte[fsImageFile.Length];
                fsImageFile.Read(imageFileBytes, 0, imageFileBytes.Length);

                using (System.Drawing.Bitmap imageFile = new System.Drawing.Bitmap(fsImageFile))
                {
                    // Determine the format of the image file. This sample code
                    // supports working with the following types of image files:
                    //
                    // Bitmap (BMP)
                    // Graphics Interchange Format (GIF)
                    // Joint Photographic Experts Group (JPG, JPEG)
                    // Portable Network Graphics (PNG)
                    // Tagged Image File Format (TIFF)

                    if (imageFile.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Bmp.Guid)
                        imagePartType = ImagePartType.Bmp;
                    else if (imageFile.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Gif.Guid)
                        imagePartType = ImagePartType.Gif;
                    else if (imageFile.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Jpeg.Guid)
                        imagePartType = ImagePartType.Jpeg;
                    else if (imageFile.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Png.Guid)
                        imagePartType = ImagePartType.Png;
                    else if (imageFile.RawFormat.Guid == System.Drawing.Imaging.ImageFormat.Tiff.Guid)
                        imagePartType = ImagePartType.Tiff;
                    else
                    {
                        throw new System.ArgumentException(
                          "Unsupported image file format: " + imageFilePath);
                    }

                    // Get the dimensions of the image in English Metric Units
                    // (EMU) for use when adding the markup for the image to the
                    // slide.
                    imageWidthEMU =
                    (long)
                    ((imageFile.Width / imageFile.HorizontalResolution) * 914400L);

                    imageHeightEMU =
                    (long)
                    ((imageFile.Height / imageFile.VerticalResolution) * 914400L);
                } // End Using imageFile 

            } // End Using fsImageFile 

            return imageFileBytes;
        } // End Function GetImageData 
        
        
        private static Slide GenerateSlidePart(string imageName,
          string imageDescription, long imageWidthEMU, long imageHeightEMU)
        {
            Slide element =
              new Slide(
                new CommonSlideData(
                  new ShapeTree(
                    new NonVisualGroupShapeProperties(
                      new NonVisualDrawingProperties()
                      {
                          Id = (UInt32Value)1U 
                        , Name = ""
                      },
                      new NonVisualGroupShapeDrawingProperties(),
                      new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(
                      new a.TransformGroup(
                        new a.Offset() { X = 0L, Y = 0L },
                        new a.Extents() { Cx = 0L, Cy = 0L },
                        new a.ChildOffset() { X = 0L, Y = 0L },
                        new a.ChildExtents() { Cx = 0L, Cy = 0L })),
                    new Picture(
                      new NonVisualPictureProperties(
                        new NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)4U,
                            Name = imageName,
                            Description = imageDescription
                        },
                        new NonVisualPictureDrawingProperties(
                          new a.PictureLocks() { NoChangeAspect = true }),
                        new ApplicationNonVisualDrawingProperties()),
                        new BlipFill(
                          new a.Blip() { Embed = "relId1" },
                          new a.Stretch(
                            new a.FillRectangle()))
                            
                            ,
                        new ShapeProperties(
                          new a.Transform2D(
                            new a.Offset() { X = 0L, Y = 0L },
                            new a.Extents()
                            {
                                Cx = imageWidthEMU,
                                Cy = imageHeightEMU
                            }),
                          new a.PresetGeometry(
                            new a.AdjustValueList())
                          { Preset = a.ShapeTypeValues.Rectangle }
                        )))),
                new ColorMapOverride(
                  new a.MasterColorMapping()));

            return element;
        } // End Function GenerateImagePart 

        private static void GenerateImagePart(OpenXmlPart part,
          byte[] imageFileBytes)
        {
            // Write the contents of the image to the ImagePart.
            using (System.IO.BinaryWriter writer = new System.IO.BinaryWriter(part.GetStream()))
            {
                writer.Write(imageFileBytes);
                writer.Flush();
            }
        }

        static void DisplayValidationErrors(
            System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> errors)
        {
            int errorIndex = 1;

            foreach (DocumentFormat.OpenXml.Validation.ValidationErrorInfo errorInfo in errors)
            {
                System.Console.WriteLine(errorInfo.Description);
                System.Console.WriteLine(errorInfo.Path.XPath);

                if (++errorIndex <= errors.Count())
                    System.Console.WriteLine("================");
            } // Next errorInfo 

        } // End Sub DisplayValidationErrors 


    } // End Class FolderToPpt 

}
