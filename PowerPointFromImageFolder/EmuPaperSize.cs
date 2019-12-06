
namespace PowerPointFromImageFolder
{


    public static class EmuPaperSizeExtension
    {


        public static EmuPaperSize FromSlideSize(this DocumentFormat.OpenXml.Presentation.SlideSizeValues slideSize)
        {
            switch (slideSize)
            {
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Screen4x3: return EmuPaperSize.Screen4x3;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Letter: return EmuPaperSize.Letter;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.A4: return EmuPaperSize.A4;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Film35mm: return EmuPaperSize.Film35mm;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Overhead: return EmuPaperSize.Overhead;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Banner: return EmuPaperSize.Banner;
                // case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Custom: return EmuPaperSize.Custom;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Ledger: return EmuPaperSize.Ledger;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.A3: return EmuPaperSize.A3;
                //case DocumentFormat.OpenXml.Presentation.SlideSizeValues.B4ISO: return EmuPaperSize.B4ISO;
                //case DocumentFormat.OpenXml.Presentation.SlideSizeValues.B5ISO: return EmuPaperSize.B5ISO;
                //case DocumentFormat.OpenXml.Presentation.SlideSizeValues.B4JIS: return EmuPaperSize.B4JIS;
                //case DocumentFormat.OpenXml.Presentation.SlideSizeValues.B5JIS: return EmuPaperSize.B5JIS;
                //case DocumentFormat.OpenXml.Presentation.SlideSizeValues.HagakiCard: return EmuPaperSize.HagakiCard;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Screen16x9: return EmuPaperSize.Screen16x9;
                case DocumentFormat.OpenXml.Presentation.SlideSizeValues.Screen16x10: return EmuPaperSize.Screen16x10;
                default:break;
            }
            
            return EmuPaperSize.A4;
        }
    }




    public class EmuPaperSize
    {
        public int EmuX;
        public int EmuY;


        public EmuPaperSize()
        { }

        public EmuPaperSize(int x, int y)
        {
            this.EmuX = x;
            this.EmuY = y;
        }


        // https://www.digitalcitizen.life/sites/default/files/gdrive/powerpoint_slide_size/slide_size_4.jpg
        // http://lcorneliussen.de/raw/dashboards/ooxml/
        // Wrong sizes ???
        // https://phpoffice.github.io/PHPPresentation/coverage/develop/DocumentLayout.php.html

        // A4        
        // 21cm = 7560000 emu
        // 29.7cm = 10692000 emu
        public static EmuPaperSize A4 = new EmuPaperSize(7560000, 10692000);

        // A3
        // 29.7 cm = 10692000 emu
        // 42.0 cm = 15120000 emu
        public static EmuPaperSize A3 = new EmuPaperSize(10692000, 15120000);

        // 7772400 ==> 8.5 inch 
        // 10058400 ==> 11 inch
        public static EmuPaperSize Letter = new EmuPaperSize(7772400, 10058400);

        // 10058400 ==> 11 inch 
        // 15544800 ==> 17 inch
        public static EmuPaperSize Ledger = new EmuPaperSize(10058400, 15544800);

        // 9000000 ==> 250 mm
        // 12708000 ==> 353 mm
        public static EmuPaperSize B4 = new EmuPaperSize(9000000, 12708000);

        // 6336000 ==> 176 mm
        // 9000000 ==> 250 mm
        public static EmuPaperSize B5 = new EmuPaperSize(6336000, 9000000);

        // https://www.brightcarbon.com/blog/powerpoint-2013-widescreen-by-default/
        // normal 4:3 view which had an area of 25.4cm high x 19.05cm high.
        // 9144000 ==> 25.4 cm
        // 6858000 ==> 19.05 cm
        public static EmuPaperSize Overhead = new EmuPaperSize(9144000, 6858000);
        public static EmuPaperSize Screen4x3 = new EmuPaperSize(9144000, 6858000);

        // https://www.brightcarbon.com/blog/powerpoint-2013-widescreen-by-default/
        // 12192120 ==> 33.867
        // 6858000 ==> 19.05 cm
        public static EmuPaperSize Screen16x9 = new EmuPaperSize(12192120, 6858000);

        // 9144000 = 25.4 cm
        // 5715000 = 15.875 cm
        public static EmuPaperSize Screen16x10 = new EmuPaperSize(9144000, 5715000);

        // 10287000 ==> 28.575 cm
        // 6858000 ==> 19.05 cm
        public static EmuPaperSize Film35mm = new EmuPaperSize(10287000, 6858000);

        // 7315200 ==> 20.32 cm
        //  914400 ==>  2.54 cm
        public static EmuPaperSize Banner = new EmuPaperSize(7315200, 914400);

        // 12192120 ==> 33.867 cm
        //  6858000 ==>  19.05 cm
        public static EmuPaperSize WideScreen = new EmuPaperSize(12192120, 6858000);
    }


}
