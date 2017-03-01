using System.IO;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using DataProcessing.Web.Controllers;
using DataProcessing.Web.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace DataProcessing.Web.Tests.Controllers
{
    [TestClass]
    public class HomeControllerTests
    {
        [TestMethod]
        public async Task UploadErrorAsync()
        {
            var postedFile = new Mock<HttpPostedFileBase>();

            ViewResult result;
            using (var stream = new MemoryStream())
            {
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write("");
                    writer.Flush();
                    stream.Position = 0;
                    postedFile.Setup(p => p.InputStream).Returns(stream);
                    
                    var homeController = new HomeController();
                    result = (ViewResult)await homeController.Convert(postedFile.Object);
                }
            }

            var model = result.Model as HomeViewModel;
            Assert.IsNotNull(model);
            Assert.IsNotNull(model.UploadErrorMessage);
        }
    }
}
