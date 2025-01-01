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
            /////////////////////////////////insert image/////////////////////////////////
            //string connectionString = "Server=.;Database=CentralUserInfo;User Id=sa;Password=Aa@12345;";
            //string userName = "222";
            //string pngpath = @"C:\mydocs\img\1.png";
            //UserImageHandler.UserImageHandler.InsertUserWithImage(userName, pngpath, connectionString);
            /////////////////////////////////test qrcode/////////////////////////////////
            //string qrImagePath = BookMarks.BookmarkOpenxml.GenerateQRCodeImage("Test QR Code");
            //Console.WriteLine($"Generated QR code image at: {qrImagePath}");
            /////////////////////////////////bookmark/////////////////////////////////


            //PersianCalendar pc = new PersianCalendar();
            //Console.WriteLine();
            //DateTime thisDate = DateTime.Now;





            //////////////////////////////////////////////////////////////////bm4 namespace
            //bm4.Program.Main();





            //////////////////////////////////////////////////////////////////Bookmarks namespace
            string docPath = @"C:\mydocs\f1.docx";
            var bookmarksContent = new Dictionary<string, string>
            {
                //{ "تصویر_واحد_سازمانی",  @"C:\mydocs\img\1.png" },
                { "واحد_سازمانی", "بازرسی کل استان تست" },
                { "تاریخ_خورشیدی",  DateTime.Now.ToString("dd-MM-yyyy") },
                { "پیوست", "ندارد"  },
                { "گیرندگان_رونوشت","تستی تستیان"  },
                { "رونوشت","تستی تستیان"  },
                { "شماره_ثبت","127126"  },
                { "طبقه_بندی","غیر محرمانه"  },
                { "عنوان_محترمانه_کامل_گیرندگان_رونوشت","تستی تستیان"  },
                { "فوریت", "فوری" },
                { "نام_و_نام_خانوادگی_فرستنده","تستی تستیان"  },
                //{ "نوع_جایگاه_امضاکننده_اصلی","تستی تستیان4"  },
                { "آدرس_جایگاه_فرستنده", "تهران - خیابان طالقانی سازمان بازرسی کشور"  },
                { "اهمیت_", "مهم"  },
                { "بارکد_شمس", "QRCode"  },
                { "امضای_اصلی", "@Binary:6930"  },
                { "گیرندگان_اصلی","جناب آقای تستی تستیان \r\n رییس محترم شورای اسلامی شهر"  },
                { "نوع_جایگاه_امضاکننده_اصلی", "تستی تستیان"  },
            };
            BookmarkOpenxml.UpdateBookmarks(docPath, bookmarksContent);






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
