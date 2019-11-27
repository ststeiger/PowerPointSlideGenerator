
namespace PowerPointFromImageFolder
{

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.IO;
    using System.Drawing;
    using System.Drawing.Imaging;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Presentation;
    using a = DocumentFormat.OpenXml.Drawing;
    using DocumentFormat.OpenXml.Validation;


    // https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/ee412267(v=office.12)?redirectedfrom=MSDN
    class FolderToPpt
    {


        public static void Test()
        {
            string newPresentation = "DeckFromImages.pptx";
            string presentationTemplate = "PresentationTemplate.pptx";
            string presentationFolder = @"C:\Temp\";
            string imageFolder = @"C:\Temp";

            imageFolder = @"D:\username\Pictures\majdu\Steve";
            presentationFolder = @"D:\";


            string[] imageFileExtensions =
              new[] { "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.png", "*.tif" };



            // EmptyPpt.CreatePresentation(presentationFolder + presentationTemplate);

            // Make a copy of the template presentation. This will throw an
            // exception if the template presentation does not exist.
            // File.Copy(presentationFolder + presentationTemplate, presentationFolder + newPresentation, true);

            EmptyPpt.CreatePresentation(presentationFolder + newPresentation);



            // Get the image files in the image folder.
            List<string> imageFileNames = GetImageFileNames(imageFolder,
              imageFileExtensions);

            // Create new slides for the images.
            if (imageFileNames.Count() > 0)
                CreateSlides(imageFileNames,
                  presentationFolder + newPresentation);

            // Validate the new presentation.
            OpenXmlValidator validator = new OpenXmlValidator();

            using (DocumentFormat.OpenXml.Packaging.PresentationDocument presentation =
                DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(presentationFolder + newPresentation, true))
            {
                IEnumerable<ValidationErrorInfo> errors = validator.Validate(presentation);

                if (errors.Count() > 0)
                {
                    Console.WriteLine("The deck creation process completed but " +
                      "the created presentation failed to validate.");
                    Console.WriteLine("There are " + errors.Count() +
                      " errors:\r\n");

                    DisplayValidationErrors(errors);
                }
                else
                    Console.WriteLine("The deck creation process completed and " +
                      "the created presentation validated with 0 errors.");
            } // End Using presentation 

        } // End Sub Test 


        static void CreateSlides(List<string> imageFileNames, string newPresentation)
        {
            string relId;
            SlideId slideId;

            // Slide identifiers have a minimum value of greater than or
            // equal to 256 and a maximum value of less than 2147483648.
            // Assume that the template presentation being used has no slides.
            // uint currentSlideId = 256;
            uint currentSlideId = 256+1;

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
                      Path.GetFileNameWithoutExtension(imageFileNameWithPath);

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


        public static List<string> GetImageFileNames(string imageFolder,
          string[] imageFileExtensions)
        {
            // Create a list to hold the names of the files with the
            // requested extensions.
            List<string> fileNames = new List<string>();

            // Loop through each file extension.
            foreach (string extension in imageFileExtensions)
            {
                // Add all the files that match the current extension to the
                // list of file names.
                fileNames.AddRange(Directory.GetFiles(imageFolder, extension,
                  SearchOption.TopDirectoryOnly));
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
            using (FileStream fsImageFile = File.OpenRead(imageFilePath))
            {
                imageFileBytes = new byte[fsImageFile.Length];
                fsImageFile.Read(imageFileBytes, 0, imageFileBytes.Length);

                using (Bitmap imageFile = new Bitmap(fsImageFile))
                {
                    // Determine the format of the image file. This sample code
                    // supports working with the following types of image files:
                    //
                    // Bitmap (BMP)
                    // Graphics Interchange Format (GIF)
                    // Joint Photographic Experts Group (JPG, JPEG)
                    // Portable Network Graphics (PNG)
                    // Tagged Image File Format (TIFF)

                    if (imageFile.RawFormat.Guid == ImageFormat.Bmp.Guid)
                        imagePartType = ImagePartType.Bmp;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Gif.Guid)
                        imagePartType = ImagePartType.Gif;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Jpeg.Guid)
                        imagePartType = ImagePartType.Jpeg;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Png.Guid)
                        imagePartType = ImagePartType.Png;
                    else if (imageFile.RawFormat.Guid == ImageFormat.Tiff.Guid)
                        imagePartType = ImagePartType.Tiff;
                    else
                    {
                        throw new ArgumentException(
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


        private static System.Random s_rand = new System.Random();

        private static Slide GenerateSlidePart(string imageName,
          string imageDescription, long imageWidthEMU, long imageHeightEMU)
        {
            UInt32Value randomValue = new UInt32Value((uint)s_rand.Next(0, int.MaxValue));

            Slide element =
              new Slide(
                new CommonSlideData(
                  new ShapeTree(
                    new NonVisualGroupShapeProperties(
                      new NonVisualDrawingProperties()
                      {
                          Id = (UInt32Value)1U
                          // Id = randomValue
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
                            //Id = randomValue,
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
            using (BinaryWriter writer = new BinaryWriter(part.GetStream()))
            {
                writer.Write(imageFileBytes);
                writer.Flush();
            }
        }

        static void DisplayValidationErrors(
          IEnumerable<ValidationErrorInfo> errors)
        {
            int errorIndex = 1;

            foreach (ValidationErrorInfo errorInfo in errors)
            {
                Console.WriteLine(errorInfo.Description);
                Console.WriteLine(errorInfo.Path.XPath);

                if (++errorIndex <= errors.Count())
                    Console.WriteLine("================");
            } // Next errorInfo 

        } // End Sub DisplayValidationErrors 


    } // End Class FolderToPpt 

}
