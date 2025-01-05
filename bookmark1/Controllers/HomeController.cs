using System.Diagnostics;
using bookmark1.Models;
using BookMarks;
using Microsoft.AspNetCore.Mvc;
using System.Globalization;

namespace bookmark1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {


            //////////////////////////////////////////////////////////////////Bookmarks namespace
            //string docPath = @"C:\mydocs\f1.docx";
            //var bookmarksContent = new Dictionary<string, string>
            //{
            //    //{ "تصویر_واحد_سازمانی",  @"C:\mydocs\img\1.png" },
            //    { "واحد_سازمانی", "بازرسی کل استان تست" },
            //    { "تاریخ_خورشیدی",  DateTime.Now.ToString("dd-MM-yyyy") },
            //    { "پیوست", "ندارد"  },
            //    { "گیرندگان_رونوشت","تستی تستیان"  },
            //    { "رونوشت","تستی تستیان"  },
            //    { "شماره_ثبت","127126"  },
            //    { "طبقه_بندی","غیر محرمانه"  },
            //    { "عنوان_محترمانه_کامل_گیرندگان_رونوشت","تستی تستیان"  },
            //    { "فوریت", "فوری" },
            //    { "نام_و_نام_خانوادگی_فرستنده","تستی تستیان"  },
            //    //{ "نوع_جایگاه_امضاکننده_اصلی","تستی تستیان4"  },
            //    { "آدرس_جایگاه_فرستنده", "تهران - خیابان طالقانی سازمان بازرسی کشور"  },
            //    { "اهمیت_", "مهم"  },
            //    { "بارکد_شمس", "QRCode"  },
            //    { "امضای_اصلی", "@Binary:6930"  },
            //    { "گیرندگان_اصلی","جناب آقای تستی تستیان \r\n رییس محترم شورای اسلامی شهر"  },
            //    { "نوع_جایگاه_امضاکننده_اصلی", "تستی تستیان"  },
            //};
            //BookmarkOpenxml.UpdateBookmarks(docPath, bookmarksContent);
            //////////////////////////////////////Example simple text//////////////////////////////////////////
            string docPath = @"C:\mydocs\f1.docx";
            var bookmarksContent = new Dictionary<string, string>
            {
                { "واحد_سازمانی", "111" },
                { "تصویر_واحد_سازمانی", "2222" },
                { "بارکد_شمس", "333" },
                { "تاریخ_خورشیدی", "4444" },
                //{ "بارکد_شمس", "QRCode:6162" },
                //{ "تصویر_واحد_سازمانی", @"C:\mydocs\img\1.png" },
                //{ "تصویر_واحد_سازمانی2", "Binary:6930" },
            };
            //BookmarkOpenxml.UpdateBookmarks(docPath, bookmarksContent);
            BMH.BMH.UpdateTextBookmarks(docPath, bookmarksContent);




            //Console.WriteLine("Bookmarks updated successfully!");
            //try
            //{
            //    // Process binary images before updating bookmarks
            //    BookMarks.BookmarkOpenxml.ProcessBinaryImages(bookmarksContent);

            //    // Call the UpdateBookmarks method
            //    BookMarks.BookmarkOpenxml.UpdateBookmarks(filePath, bookmarksContent);
            //    Console.WriteLine("Bookmarks updated successfully!");
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"An error occurred: {ex.Message}");
            //}


            //////////////////////////////////////BM Helper Docx To Pdf//////////////////////////////////////////


            //OfficeHandler.WordHandler.ConvertWordToPdfWithLibreOffice(@"C:\mydocs\f1.docx", @"C:\mydocs\pdf\aaa.pdf");
            //OfficeHandler.WordHandler.ConvertWordToPdf();
            //OfficeHandler.WordHandler.docxtopdfapose();
            //OfficeHandler.WordHandler.ConvertWordToPdf();
            //////////////////////////////////////BM Helper Docx To Pdf//////////////////////////////////////////
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
