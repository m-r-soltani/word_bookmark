using System.Diagnostics;
using bookmark1.Models;
using BookMarks;
using Microsoft.AspNetCore.Mvc;

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
            string docPath = @"C:\mydocs\f1.docx";
            var bookmarksContent = new Dictionary<string, string>
            {
                { "تصویر_واحد_سازمانی",  @"C:\mydocs\img\1.png" },
                { "واحد_سازمانی",  @"C:\mydocs\img\1.png" },
                { "تاریخ_خورشیدی",  @"C:\mydocs\img\1.png" },
                { "پیوست",@"C:\mydocs\img\1.png"  },
                //{ "رونوشت","@QRCode:This is a QR code"  },
                { "شماره_ثبت","111"  },
                //{ "طبقه_بندی","@Binary:6942"  },
                { "عنوان_محترمانه_کامل_گیرندگان_اصلی", @"C:\mydocs\img\1.png"  },
                { "عنوان_محترمانه_کامل_گیرندگان_رونوشت","111"  },
                { "فوریت",@"C:\mydocs\img\1.png"  },
                { "نام_و_نام_خانوادگی_فرستنده","111"  },
                //{ "آدرس_جایگاه_فرستنده","@Binary:6930"  },
                { "اهمیت_",@"C:\mydocs\img\1.png"  },
                { "بارکد_شمس",@"C:\mydocs\img\1.png"  },
                //{ "امضای_اصلی", "@Binary:6930"  },
                //{ "امضای_اصلی_و_فرستنده","@Binary:6942"  },
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
