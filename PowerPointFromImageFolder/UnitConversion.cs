namespace PowerPointFromImageFolder
{
    public class UnitConversion
    {
        
        public static int MilimeterToEmu(int mm)
        {
            // 1 cm = 360000
            // 1 mm = 36000
            return mm * 36000; 
        }
        
        
        // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
        // A DrawingML graphic's dimensions are specified in English Metric Units (EMUs).
        // It is so called because it allows an exact common representation of dimensions originally in either English or Metric units.
        // This unit is defined as 1/360,000 of a centimeter and thus there are 914,400 EMUs per inch, and 12,700 EMUs per point. 
        public static int ToEmu(int pixel, int ppi)
        {
            // centimeters = pixels * 2.54 / 96
            // 1 inch = 914400 emu
            return pixel * 914400 / ppi; 
        }
        
        
        public static int ToEmu(int pixel)
        {
            return ToEmu(pixel, 96);
        }
        
        
        public static int MilimeterToPoint(int mm)
        {
            // 1 inch = 72 pt
            // 1 cm = 28.34645669291339
            // 1 mm = 0.03937007874015748 
            return (int) System.Math.Ceiling(mm * 0.03937007874015748);
        }
        
        
        // https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
        // Twentieths of a point (dxa)
        // The main unit in OOXML is a twentieth of a point. 
        // This is used for specifying page dimensions, margins, tabs, etc.
        public static int MilimeterToDxa(int mm)
        {
            // 1 inch = 1440 dxa
            // 1 cm = 566.9291338582677 dxa
            // 1 mm = 56.69291338582677 dxa 
            return (int) System.Math.Ceiling(mm * 56.69291338582677);
        }


    }
}