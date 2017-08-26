# Ribbon Sample Ribbon 示例

This sample shows how to add a ribbon UI extension with the Excel-DNA add-in, and how to use the Excel COM object model from C# to write some information to a workbook.<br\>
此示例演示如何添加Excel DNA添加在Ribbon界面的扩展，以及如何使用从C #写一些信息到一个工作簿的Excel COM对象模型。<br\>

## Initial setup 初始设置

The initial setup will create a new add-in with a simple test function (it's a useful indicator to show that the add-in is loaded into the Excel session).<br\>
初始设置将用一个简单的测试函数创建一个新的外接程序（它是一个有用的指示器，用来显示加载项被加载到Excel会话中）。<br\>

1. Create new Class Library project. 创建新类库项目
2. Install `ExcelDna.AddIn` package. 安装ExcelDna.AddIn Nuget包
3. Add a small test function: 添加一个小测试函数：

```cs
namespace Ribbon
{
    public static class Functions
    {
        public static string dnaRibbonTest()
        {
            return "Hello from the Ribbon Sample!";
        }
    }
}
```

4. Press F5 to load in Excel, and then test `=dnaRibbonTest()` in a cell.<br\>
4、按F5加载在Excel，然后测试，在单元格输入` = dnaribbontest() 。<br\>

## Add the ribbon controller 添加Ribbon

Next we add a class to implement the ribbon UI extension, with a simple button.<br\>
接下来，我们添加一个类来实现ribbonUI扩展，只使用一个简单的按钮。<br\>

1. Add a reference to the `System.Windows.Forms` assembly (we'll use that for showing our messages).<br\>
1、添加引用System.Windows.Forms（我们将使用它来显示我们的消息）。<br\>

2. Add a new class for the ribbon controller (maybe `RibbonController.cs`), with this code for a button and handler:<br\>
2、添加一个Ribbon的新类（如RibbonController.CS `），这个代码按钮处理程序：<br\>

```cs
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='My Tab'>
            <group id='group1' label='My Group'>
              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
        }
    }
}
```

3. Press F5 to load and test. 按F5调试


#### Notes 注意

* The ribbon class derives from the `ExcelDna.Integration.CustomUI.ExcelRibbon` base class. This is how Excel-DNA itentifies the class a defining a ribbon controller.<br\>
* 新建的Ribbon类继承于ExcelDna.Integration.CustomUI.ExcelRibbon。这是Excel-DNA对Ribbon的管理<br\>

* The ribbon class must be 'COM visible'. Either the class must be marked as `[ComVisible(true)]` (the default class library template in Visual Studio markes the assembly as `[assembly:ComVisible(false)]`).<br\>
* Ribbon类必须设置为 COM可见。[ComVisible(true)] 。Visual Studio中默认的类库模板是[assembly:ComVisible(false)]<br\>

* The xml namespace is important. Excel 2007 introduced the ribbon, and support only the original xml namespace as shown in this example - `xmlns='http://schemas.microsoft.com/office/2006/01/customui'`. Further enhancements to the ribbon was made in Excel 2010, including using the ribbon for worksheet context menus and adding the backstage area. To indicate the extended Excel 2010 xml schema, this version and later supports an update namespace - `xmlns='http://schemas.microsoft.com/office/2009/07/customui'`.<br\>
* xml命名空间很重要！！Excel 2007介绍了丝带,只支持原始xml名称空间如本例所示,“xmlns = ' http://schemas.microsoft.com/office/2006/01/customui ‘“ 对ribbon的进一步增强是在Excel 2010中完成的，包括在工作表上下文菜单中使用Ribbon，并添加后台区域。显示扩展的Excel 2010 xml模式,这个版本,后来支持更新名称空间——“xmlns = ' http://schemas.microsoft.com/office/2009/07/customui ' '。<br\>

* The Office applications have a debugging setting to assist in finding any errors in the ribbon xml, which would prevent the ribbon from loading. In Excel 2013, this setting can be found under `File -> Options -> Advanced`, then under `General` find 'Show add-in user interface errors'. Note that this setting applied to all installed Office applications, and can reveal unexpected errors that are present in other add-ins too.<br\>
* Office应用程序有一个调试设置，以帮助查找ribbon xml中的任何错误，这将阻止ribbon加载。在Excel 2013中，可以在文件-选项-选项-高级选项中找到这个设置，然后通常会发现“显示插件的用户界面错误”。请注意，此设置适用于所有已安装的Office应用程序，并且可以显示其他插件中出现的意外错误。<br\>

* There are different options for providing the ribbon xml. In this sample it is embedded as a string in the code and returned from the `ExcelRibbon.GetCustomUI` overload. Excel-DNA also supports placing the xml inside the .dna add-in configuration file (this is where the base class implementation of `GetCustomUI` looks for it). The ribbon xml can also be put in an assembly resource (either as a string or from a separate file) and extracted at runtime with some extra code in `GetCustomUI`.<br\>
* 提供ribbon xml有不同的选项。在这个示例中，它作为一个字符串嵌入到代码中，并从ExcelRibbon返回。GetCustomUI过载。excel-dna还支持将xml放入其中。dna插件配置文件(这是GetCustomUI的基类实现查找的地方)。ribbon xml还可以放在一个集合资源(作为字符串或单独的文件)，并在运行时在GetCustomUI中提取一些额外的代码。<br\>

* The callback methods, like `OnButtonPressed` in the example, are found by Excel using the COM `IDispatch` interface that is implicitly implemented by the COM visible .NET class.<br\>
* 回调方法，例如例子中的OnButtonPressed，是由Excel使用COM IDispatch接口发现的，该接口是由COM可见的隐式实现的.Net类。<br\>

* Behind the scenes, Excel-DNA registers and loads a COM helper add-in that provides the ribbon support. This COM helper add-in should load even if the user does not have administrator rights, but it might be blocked by some Excel-specific security settings.<br\>
* 在幕后，excel-dna注册并加载一个提供ribbon支持的COM助手插件。即使用户没有管理员权限，这个COM助手插件也应该加载，但是它可能会被某些特定于excel的安全设置所阻塞。<br\>

* Errors in the ribbon methods can cause Excel to mark the ribbon COM helper add-in as a 'Disabled Add-in'. This will reflect in the 'Disabled Add-ins' list under `File-> Options -> Add-Ins` under the `Manage` dropdown.<br\>
* ribbon方法中的错误可以导致Excel将ribbon COM助手插件标记为“禁用的插件”。这将反映在管理下拉菜单下的“禁用插件”列表中。<br\>

#### Ribbon xml and callback documentation
#### Ribbon xml和回调文档


Excel-DNA is responsible for loading the ribbon helper add-in, but is not otherwise involved in the ribbon extension. This means that the custom UI xml schema, and the signatures for the callback methods are exactly as documented by Microsoft. The best documentation for these aspects can be found in the three-part series on 'Customizing the 2007 Office Fluent Ribbon for Developers':<br\>
excel-dna负责加载ribbon助手插件，但与ribbon扩展不相关。这意味着定制的UI xml模式和回调方法的签名都是由Microsoft记录的。关于这些方面的最好的文档可以在“定制2007年的Office流畅的开发者”系列文章中找到。<br\>

* [Part 1 - Overview](https://msdn.microsoft.com/en-us/library/aa338202.aspx)
* [Part 2 - Controls and callback reference](https://msdn.microsoft.com/en-us/library/aa338199.aspx)
* [Part 3 - Frequently asked questions, including C# and VB.NET callback signatures](https://msdn.microsoft.com/en-us/library/aa722523.aspx)
* [Part 1 - 概述](https://msdn.microsoft.com/en-us/library/aa338202.aspx)
* [Part 2 - 控件和回调引用](https://msdn.microsoft.com/en-us/library/aa338199.aspx)
* [Part 3 - 经常被问到的问题，包括c#和VB.net 回调签名](https://msdn.microsoft.com/en-us/library/aa722523.aspx)

Information related to the Excel 2010 extensions to the ribbon can be found here:<br\>
有关2010年Excel 2010扩展的信息可以在这里找到:<br\>

* [Customizing Context Menus in Office 2010](https://msdn.microsoft.com/en-us/library/office/ee691832.aspx)
* [Customizing the Office 2010 Backstage View](https://msdn.microsoft.com/en-us/library/office/ee815851.aspx)
* [Ribbon Extensibility in Office 2010: Tab Activation and Auto-Scaling](https://msdn.microsoft.com/en-us/library/office/ee691834.aspx)
* [在Office 2010中定制上下文菜单](https://msdn.microsoft.com/en-us/library/office/ee691832.aspx)
* [自定义Office 2010后台查看](https://msdn.microsoft.com/en-us/library/office/ee815851.aspx)
* [Office 2010 Ribbon的可扩展性:标签激活和自动扩展](https://msdn.microsoft.com/en-us/library/office/ee691834.aspx)

Creating Dynamic Ribbon Customizations<br\>
创建动态Ribbon <br\>
* [Part 1](https://msdn.microsoft.com/en-us/library/dd548010%28v=office.12%29.aspx)
* [Part 2](https://msdn.microsoft.com/en-us/library/dd548011%28v=office.12%29.aspx)

Other ribbon-related resources:其他ribbon-related资源:

* [Ron de Bruin's Excel Tips](http://www.rondebruin.nl/win/s2/win003.htm)
* [Andy Pope's RibbonX Visual Designer](http://www.andypope.info/vba/ribboneditor.htm)

## Add access to the Excel COM object model
## 添加对Excel COM对象模型的访问

In this step we add access to the Excel COM object model, to show how C# code can use the familiar object model to manipulate Excel.<br\>
在此步骤中，我们添加了对Excel COM对象模型的访问，以展示ccode如何使用熟悉的对象模型来操作Excel。<br\>

1. Add a reference to the Primary Interop Assembly (PIA) for Excel. The easiest way to do this is to install the `ExcelDna.Interop` NuGet package. This will install the interop assemblies that correspond to Excel 2010, so are suitable for add-ins that support Excel 2010 and later.<br\>
1、为Excel添加一个主要的Interop程序集(PIA)的引用。最简单的方法是安装ExcelDna。互操作的NuGet包。这将安装与Excel 2010相对应的互操作程序集，因此适合于支持Excel 2010和以后的插件。<br\>

2. Add a class for our data writer:<br\>
2、为我们的数据写入器添加一个类:<br\>

```cs
using System;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace Ribbon
{
    public class DataWriter
    {
        public static void WriteData()
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null)
                return;

            Worksheet ws = wb.Worksheets.Add(Type: XlSheetType.xlWorksheet);
            ws.Range["A1"].Value = "Date";
            ws.Range["B1"].Value = "Value";

            Range headerRow = ws.Range["A1", "B1"];
            headerRow.Font.Size = 12;
            headerRow.Font.Bold = true;

            // Generally it's faster to write an array to a range
            var values = new object[100, 2];
            var startDate = new DateTime(2007, 1, 1);
            var rand = new Random();
            for (int i = 0; i < 100; i++)
            {
                values[i, 0] = startDate.AddDays(i);
                values[i, 1] = rand.NextDouble();
            }

            ws.Range["A2"].Resize[100, 2].Value = values;
            ws.Columns["A:A"].EntireColumn.AutoFit();

            // Add a chart
            Range dataRange= ws.Range["A1:B101"];
            dataRange.Select();
            ws.Shapes.AddChart(XlChartType.xlLineMarkers).Select();
            xlApp.ActiveChart.SetSourceData(Source: dataRange);
        }
    }
}
```

3. Update the ribbon handler to call our data writer:<br\>
3。更新ribbon处理程序来调用我们的数据写入器:<br\>

```cs
        public void OnButtonPressed(IRibbonControl control)
        {
            MessageBox.Show("Hello from control " + control.Id);
            DataWriter.WriteData();
        }
```

4. Press F5 to load and press the ribbon button to run the `WriteData` code.<br\>
4。按F5加载并按ribbon按钮运行WriteData代码。<br\>

#### Getting the root `Application` object
#### 获取根应用程序对象

A key call in the above code is to retrieve the root `Application` object that matches the Excel instance that is hosting the add-in. We call `ExcelDnaUtil.Application`, which returns an object that is always the correct `Application` COM object. Code that attempts to get the `Application` object in other ways, e.g. by calling `new Application()` might work in some cases, but there is a danger that the Excel instance returned is not that instance hosting the add-in.  <br\>
上述代码中的一个关键调用是检索与托管加载项的Excel实例相匹配的根应用程序对象。我们称之为“ExcelDnaUtil。应用程序，它返回一个始终是正确的应用程序COM对象的对象。试图以其他方式获取应用程序对象的代码，例如调用新应用程序()在某些情况下可能会工作，但是有一个危险是，Excel实例返回的不是托管插件的实例。<br\>

Once the root `Application` object is retrieved, the object model is accessed normally as it would be from VBA.<br\>
一旦检索了根应用程序对象，就会像从VBA中那样访问对象模型。<br\>

* Don't confuse the types `Microsoft.Office.Interop.Excel.Application` that we use here with the WinForms type `System.Windows.Forms.Application`. You might use a namespace alias to distinguish these in your code.<br\>
* 不要混淆“Microsoft.Office.Interop.Excel类型。与WinForms应用程序”,这里我们使用类型的System.Windows.Forms.Application”。您可以使用名称空间别名来区分这些代码。<br\>

#### Interop assembly versions
#### 互操作组装版本

Each version of Excel adds some extensions to the object model (and rarely, but sometimes removes some parts). The changes might be entire classes and interfaces, methods on an interface or parameters or a method. Most add-ins are expected to run on different Excel versions, so some care is needed to make sure only object models features available on all the target versions are used.<br\>
每个版本的Excel都为对象模型添加了一些扩展(很少，但有时会删除某些部分)。更改可能是整个类和接口、接口、参数或方法的方法。大多数插件都将运行在不同的Excel版本上，因此需要一些注意来确保所有目标版本都可用的对象模型特性可用。<br\>

The simplest approach is to pick a minimum Excel version to support, and use the COM object model definitions (PIA asemblies) from that version. Such code will work against the chosen version, and any any other version (newer or older) that implements the same parts of the object model. Since most Excel versions only add to the object model, this means that add-in will work correctly with newer versions too. This is similar to developing a VBA extension on Excel 2010, which might then fail on older versions if the VBA code uses methods not available on the running version.<br\>
最简单的方法是选择一个最小的Excel版本来支持，并使用该版本的COM对象模型定义(PIA asemexcel)。这样的代码将针对所选的版本，以及实现对象模型相同部分的任何其他版本(更新的或更旧的)。由于大多数Excel版本只添加到对象模型中，这意味着插件也可以正确地使用新版本。这类似于在Excel 2010中开发VBA扩展，如果VBA代码使用了在运行版本中不可用的方法，那么在旧版本上可能会失败。<br\>

In this example we've installed the 'ExcelDna.Interop' NuGet package, which includes the interop assemblies for Excel 2010. This means features added in Excel 2013 and later will not be shown in the object model IntelliSense, ensuring that the add-in only uses features available on the minimum version. <br\>
在本例中，我们已经安装了“ExcelDna”。Interop的NuGet包，它包含了Excel 2010的互操作程序集。这意味着在Excel 2013中添加的特性将不会在对象模型智能感知中显示，确保插件只使用在最小版本中可用的特性。<br\>

#### Correct COM / .NET interop usage

There is a lot of misinformation on the web about using the COM object model from .NET.<br\>
网络上有很多关于使用COM对象模型的错误信息。<br\>

* To ensure that the Excel process always correctly exits, Excel add-ins should *only call the Excel COM object model from the main Excel thread, in a macro of callback context*. Never attempt to access the COM object model from multiple threads - since the Excel COM object model is single-threaded (technically a Single-Threaded Apartment) there can be no performance benefit in trying to access Excel from multiple threads.

* An Excel add-in should never call `Marshal.ReleaseComObject(...)` or `Marshal.FinalReleaseComObject(...)` when doing Excel interop. It is a confusing anti-pattern, but any information about this, including from Microsoft, that indicates one should manually release COM references from .NET is incorrect. The .NET runtime and garbage collector correctly keep track of and clean up COM references.

* Any guidance that mentions 'double-dots' is misleading. Sometimes this indicates that expressions calling into the object model should not chain object model access, i.e. to avoid code like `myWorkbook.Sheets[1].Range["A1"].Value = "abc". Such code is fine - just ignore any 'two dots' guidance.

* I've posted some more details on these issues (in the context of automating Excel from another application) in a [Stack Overflow answer](http://stackoverflow.com/a/38111294/44264).

## Further ribbon topics

These are some more aspects of the ribbon extensions and COM object model, not yet dealt with:

* Updating the ribbon, e.g. to trigger a `getEnabled` callback - the `onLoad` callback must be implemented to capture the ribbon interface during loading.

* Adding images to the ribbon. Note the the `IPictureDisp` interface does not have to be used - any `Bitmap` type can be returned from the `getImage` callbacks. Excel-DNA has some support for packing image files into the .xll.

* Using COM object model events.

* Transitioning to the main thread (or a macro context) from another thread. Excel-DNA has a helper method `ExcelAsyncUtil.QueueAsMacro` that can be called from another thread or a timer event, to transition to a context where the object model can be reliably used.
