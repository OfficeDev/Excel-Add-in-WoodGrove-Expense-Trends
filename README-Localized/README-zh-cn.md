# <a name="excel-add-in-js-woodgrove-expense-trends"></a>Excel-Add-in-JS-WoodGrove-Expense-Trends

WoodGrove Bank 支出趋势外接程序演示如何使用适用于 Microsoft Excel 2016 的新 JavaScript API 创建引人注目的 Excel 外接程序。通过支出趋势，你可以将支出交易导入工作簿、创建仪表板和跟踪器、查看并分析趋势，以及跟踪特殊交易（如慈善捐款）和跟进项目。该示例提供了两种体验：使用任务窗格的体验和使用外接程序命令的体验。下图显示了该外接程序的主屏幕。

![WoodGrove Bank 支出趋势加载项 - 功能区](../images/woodgrove_taskpane_ribbon.PNG)

![WoodGrove Bank 支出趋势加载项 - 初始任务窗格](../images/woodgrove_taskpane_import.PNG)

![WoodGrove Bank 支出趋势加载项 - 交易工作表](../images/woodgrove_taskpane_data.PNG)

![WoodGrove Bank 支出趋势加载项 - 仪表板](../images/woodgrove_taskpane_dashboard.PNG)

![WoodGrove Bank 支出趋势加载项 - 捐款跟踪器](../images/woodgrove_taskpane_donations.PNG)

## <a name="table-of-contents"></a>目录

* [先决条件](#prerequisites)
* [运行项目](#run-the-project)
* [其他资源](#additional-resources)

## <a name="prerequisites"></a>先决条件

需要具备以下条件：

* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Visual Studio 的 Office 开发人员工具](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
* Excel 2016，版本 6769.2011 或更高版本

## <a name="run-the-project"></a>运行项目

1. 将项目复制到本地文件夹。确保文件路径并未太长，否则在尝试为项目安装必要的 NuGet 程序包时可能在 Visual Studio 中出现错误。 
2. 然后打开 Visual Studio 中的 `WoodGrove Expense Trends.sln`。 
3. 按 F5 生成并部署示例外接程序。Excel 将启动，根据你拥有的 Excel 2016 版本，外接程序在功能区中加载名为 WoodGrove 的自定义选项卡或在工作表右侧的任务窗格中打开，如下图中所示。

![WoodGrove Bank 支出趋势外接程序 - 初始任务窗格] (../images/woodgrove_taskpane_ribbon.PNG)

![WoodGrove Bank 支出趋势外接程序 - 初始任务窗格] (../images/woodgrove_taskpane_import.PNG)

## <a name="additional-resources"></a>其他资源

* [Office 开发人员中心](http://dev.office.com/)

## <a name="copyright"></a>版权
版权所有 (c) 2016 Microsoft。保留所有权利。



此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
