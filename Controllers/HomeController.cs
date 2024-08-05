using Microsoft.AspNetCore.Mvc;

namespace Authentication.API.Controllers;

public class HomeController : ControllerBase
{
    public IActionResult Index()
    {
        return Redirect("~/swagger");
    }
}
