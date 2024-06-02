using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Convert
{
    [Transaction(TransactionMode.Manual)]
    public class Ribbon : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication uiapp)
        {
            Assembly assembly = typeof(Ribbon).Assembly;
            string assemblyPath = assembly.Location;

            uiapp.CreateRibbonTab("智能转换");
            RibbonPanel panel = uiapp.CreateRibbonPanel("智能转换", "生成");

            PushButton create_auto = panel.AddItem(new PushButtonData("auto_define", "智能匹配", assemblyPath, "ConvertCAD.ConvertCAD2Revit")) as PushButton;
            create_auto.LargeImage = GetEmbeddedImage(assembly, "ConvertCAD.Icons.auto.png");
            create_auto.ToolTip = "按照图层名称信息自动匹配并创建三维模型。";

            panel.AddSeparator();

            PushButton create_wall = panel.AddItem(new PushButtonData("wall_define", "墙", assemblyPath, "ConvertCAD.ConvertWall2Revit")) as PushButton;
            create_wall.LargeImage = GetEmbeddedImage(assembly, "ConvertCAD.Icons.wall.png");
            create_wall.ToolTip = "选取墙图层，匹配并创建三维模型。";

            panel.AddSeparator();

            PushButton create_door = panel.AddItem(new PushButtonData("door_define", "门", assemblyPath, "ConvertCAD.ConvertDoor2Revit")) as PushButton;
            create_door.LargeImage = GetEmbeddedImage(assembly, "ConvertCAD.Icons.door.png");
            create_door.ToolTip = "先选取门图层，再选取墙图层，匹配并创建三维模型。";

            panel.AddSeparator();

            PushButton create_window = panel.AddItem(new PushButtonData("window_define", "窗", assemblyPath, "ConvertCAD.ConvertWindow2Revit")) as PushButton;
            create_window.LargeImage = GetEmbeddedImage(assembly, "ConvertCAD.Icons.window.png");
            create_window.ToolTip = "选取窗图层，匹配并创建三维模型。";

            panel.AddSeparator();

            PushButton create_column = panel.AddItem(new PushButtonData("column_define", "柱", assemblyPath, "ConvertCAD.ConvertColumn2Revit")) as PushButton;
            create_column.LargeImage = GetEmbeddedImage(assembly, "ConvertCAD.Icons.column.png");
            create_column.ToolTip = "选取柱图层，匹配并创建三维模型。";

            panel.AddSeparator();

            PushButton create_axes = panel.AddItem(new PushButtonData("axex_define", "轴", assemblyPath, "ConvertCAD.ConvertAxis2Revit")) as PushButton;            
            create_axes.LargeImage = GetEmbeddedImage(assembly, "ConvertCAD.Icons.axis.png");
            create_axes.ToolTip = "选取轴线图层，匹配并创建三维模型。";

            panel.AddSeparator();

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication uiapp)
        {
            return Result.Succeeded;
        }
        private ImageSource GetEmbeddedImage(Assembly assembly, string imageName)
        {
            System.IO.Stream file = assembly.GetManifestResourceStream(imageName);
            PngBitmapDecoder bd = new PngBitmapDecoder(file, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            return bd.Frames[0];
        }
    }
}