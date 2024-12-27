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
            string connectionString = "Server=.;Database=CentralUserInfo;User Id=sa;Password=Aa@12345;";
            string userName = "222";
            string pngpath = @"C:\mydocs\img\1.png";
            UserImageHandler.UserImageHandler.InsertUserWithImage(userName, pngpath, connectionString);
            
            
            /////////////////////////////////bookmark/////////////////////////////////
            string docPath = @"C:\mydocs\f1.docx";
            var bookmarksContent = new Dictionary<string, string>
            {
                { "Bookmark1", "123123123" },
                { "Bookmark2", "@Binary:111" },
                //{ "Bookmark3", "@QRCode:This is a QR code" },
            };
            BookmarkOpenxml.ProcessBinaryImages(bookmarksContent);
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
