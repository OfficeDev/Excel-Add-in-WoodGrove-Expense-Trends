# <a name="excel-add-in-js-woodgrove-expense-trends"></a>Excel-Add-in-JS-WoodGrove-Expense-Trends

WoodGrove Bank Expense Trends アドインには、Microsoft Excel 2016 用の新しい JavaScript API を使用して、魅力的な Excel アドインを作成する方法が示されます。Expense Trends では、ブックへの経費トランザクションのインポート、ダッシュボードや追跡ツールの作成、傾向の表示と分析、慈善寄付やフォロー アップ項目などの特殊なトランザクションの追跡を実行できます。サンプルには、2 つのエクスペリエンスが用意されています。1 つは作業ウィンドウに関するもので、もう 1 つはアドイン コマンドに関するものです。次の図は、このアドインのメイン画面を示しています。

![WoodGrove Bank Expense Trends アドイン - リボン] (../images/woodgrove_taskpane_ribbon.PNG)

![WoodGrove Bank Expense Trends アドイン - 初期作業ウィンドウ] (../images/woodgrove_taskpane_import.PNG)

![WoodGrove Bank Expense Trends アドイン - トランザクション シート] (../images/woodgrove_taskpane_data.PNG)

![WoodGrove Bank Expense Trendsアドイン - ダッシュボード] (../images/woodgrove_taskpane_dashboard.PNG)

![WoodGrove Bank Expense Trends アドイン - 寄付追跡ツール] (../images/woodgrove_taskpane_donations.PNG)

## <a name="table-of-contents"></a>目次

* [前提条件](#prerequisites)
* [プロジェクトを実行する](#run-the-project)
* [その他のリソース](#additional-resources)

## <a name="prerequisites"></a>前提条件

以下のものが必要です。

* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
* [Office Developer Tools for Visual Studio](https://www.visualstudio.com/ja/vs/office-tools/)
* Excel 2016 バージョン 6769.2011 以降

## <a name="run-the-project"></a>プロジェクトを実行する

1. プロジェクトをローカル フォルダーにコピーします。ファイル パスが長すぎないか確認します。ファイル パスが長すぎる場合、プロジェクトに必要な NuGet パッケージをインストールしようとすると、Visual Studio でエラーが発生します。 
2. Visual Studio で `WoodGrove Expense Trends.sln` を開きます。 
3. F5 キーを押して、サンプル アドインをビルドおよび展開します。Excel が起動し、Excel 2016 のバージョンに応じて、以下の 2 つの図に示すように、アドインはリボンに WoodGrove というカスタム タブをロードするか、アドインがワークシートの右側に作業ウィンドウ内に開きます。

![WoodGrove Bank Expense Trends アドイン - 初期作業ウィンドウ] (../images/woodgrove_taskpane_ribbon.PNG)

![WoodGrove Bank Expense Trends アドイン - 初期作業ウィンドウ] (../images/woodgrove_taskpane_import.PNG)

## <a name="additional-resources"></a>その他のリソース

* [Office デベロッパー センター](http://dev.office.com/)

## <a name="copyright"></a>著作権
Copyright (c) 2016 Microsoft. All rights reserved.

