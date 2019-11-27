
namespace PowerPointFromImageFolder
{

    // https://secretweaponsdigital.wordpress.com/2016/11/30/openxml-add-image-to-presentation/
    public class ImageAdd
    {


        public static void TestEnumerator()
        {
            System.Collections.Generic.List<string> ls = new System.Collections.Generic.List<string>();
            ls.Add("file1.pptx");
            ls.Add("file2.pptx");

            System.Collections.Generic.IEnumerator<string> e = ls.GetEnumerator();

            // e.Current is NULL here 
            if (!e.MoveNext())
            {
                throw new System.Exception("No elements !");
            }
            
            string file = file = e.Current;
            
            if (System.IO.File.Exists(file))
            {
                
                using (DocumentFormat.OpenXml.Packaging.PresentationDocument presentation =
                    DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(file, true))
                {
                    System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Packaging.SlidePart> slidePart = presentation
                       .PresentationPart
                       .SlideParts
                        // .First() // Requires using System.Linq; 
                    ;
                } // End Using presentation
                
            } // End if (System.IO.File.Exists(file)) 
            
        } // End Sub TestEnumerator 
        
        
        public static void AddImage(string file, string image)
        {
            using (DocumentFormat.OpenXml.Packaging.PresentationDocument presentation = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(file, true))
            {

                using (System.Collections.Generic.IEnumerator<DocumentFormat.OpenXml.Packaging.SlidePart> slidePartEnum = 
                    presentation.PresentationPart.SlideParts.GetEnumerator())
                {
                    if (!slidePartEnum.MoveNext())
                    {
                        throw new System.Exception("No elements");
                        //return e.Current;
                    }

                    DocumentFormat.OpenXml.Packaging.SlidePart slidePart = slidePartEnum.Current;


                    DocumentFormat.OpenXml.Packaging.ImagePart part = slidePart
                        .AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);

                    using (System.IO.Stream stream = System.IO.File.OpenRead(image))
                    {
                        part.FeedData(stream);
                    }

                    System.Collections.Generic.IEnumerator<DocumentFormat.OpenXml.Presentation.ShapeTree> treeEnum = 
                        slidePart.Slide
                        .Descendants<DocumentFormat.OpenXml.Presentation.ShapeTree>()
                        //.First();
                        .GetEnumerator();


                    if (!treeEnum.MoveNext())
                    {
                        throw new System.Exception("No elements");
                    }


                    DocumentFormat.OpenXml.Presentation.ShapeTree tree = treeEnum.Current;

                    DocumentFormat.OpenXml.Presentation.Picture picture = new DocumentFormat.OpenXml.Presentation.Picture();

                    picture.NonVisualPictureProperties = new DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties();
                    picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                    {
                        Name = "My Shape",
                        Id = (uint)tree.ChildElements.Count - 1
                    });

                    DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = 
                        new DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties();
                    nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
                    {
                        NoChangeAspect = true
                    });

                    picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
                    picture.NonVisualPictureProperties.Append(new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties());

                    DocumentFormat.OpenXml.Presentation.BlipFill blipFill = new DocumentFormat.OpenXml.Presentation.BlipFill();
                    DocumentFormat.OpenXml.Drawing.Blip blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
                    {
                        Embed = slidePart.GetIdOfPart(part)
                    };

                    DocumentFormat.OpenXml.Drawing.BlipExtensionList blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
                    DocumentFormat.OpenXml.Drawing.BlipExtension blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
                    {
                        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                    };

                    DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
                    {
                        Val = false
                    };

                    useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                    blipExtension1.Append(useLocalDpi1);
                    blipExtensionList1.Append(blipExtension1);
                    blip1.Append(blipExtensionList1);
                    DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                    stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
                    blipFill.Append(blip1);
                    blipFill.Append(stretch);
                    picture.Append(blipFill);

                    picture.ShapeProperties = new DocumentFormat.OpenXml.Presentation.ShapeProperties();
                    picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
                    picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
                    {
                        X = 0,
                        Y = 0,
                    });
                    picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
                    {
                        Cx = 1000000,
                        Cy = 1000000,
                    });
                    picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
                    {
                        Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                    });

                    tree.Append(picture);
                } // End Using slidePart 

            } // End Using presentation 

        } // End Sub AddImage 


    } // End Class ImageAdd 


} // End Namespace 
