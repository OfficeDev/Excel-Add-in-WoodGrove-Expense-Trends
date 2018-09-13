# <a name="excel-add-in-js-woodgrove-expense-trends"></a>Excel-Add-in-JS-WoodGrove-Expense-Trends

Этот пример надстройки, позволяющей анализировать направления расходных операций через WoodGrove Bank, показывает создание привлекательной надстройки Excel с помощью API JavaScript для Microsoft Excel 2016. С помощью этой надстройки вы можете импортировать данные о расходных операциях в книгу, создать панель мониторинга и средства отслеживания, просматривать и анализировать направления, а также отслеживать специальные операции, например благотворительные пожертвования и последующие действия. В этом примере представлено два варианта. Один предназначен для области задач, а другой предполагает использование команд надстроек. На приведенных ниже рисунках показаны основные экраны этой надстройки.

![Надстройка WoodGrove Bank Expense Trends — лента](../images/woodgrove_taskpane_ribbon.PNG)

![Надстройка WoodGrove Bank Expense Trends — первоначальная область задач](../images/woodgrove_taskpane_import.PNG)

![Надстройка WoodGrove Bank Expense Trends — ведомость учета проводок](../images/woodgrove_taskpane_data.PNG)

![Надстройка WoodGrove Bank Expense Trends — панель мониторинга](../images/woodgrove_taskpane_dashboard.PNG)

![Надстройка WoodGrove Bank Expense Trends — трекер пожертвований](../images/woodgrove_taskpane_donations.PNG)

## <a name="table-of-contents"></a>Содержание

* [Необходимые компоненты](#prerequisites)
* [Запуск проекта](#run-the-project)
* [Дополнительные ресурсы](#additional-resources)

## <a name="prerequisites"></a>Необходимые компоненты

Вам понадобится следующее:

* [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx).
* [Инструменты разработчика Office для Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx);
* Excel 2016 версии не ниже 6769.2011.

## <a name="run-the-project"></a>Запуск проекта

1. Скопируйте проект в локальную папку. Убедитесь, что путь к файлу не слишком длинный. В противном случае в Visual Studio может возникнуть ошибка при попытке установить пакеты NuGet, необходимые для проекта. 
2. Затем откройте файл `WoodGrove Expense Trends.sln` в Visual Studio. 
3. Нажмите клавишу F5, чтобы собрать и развернуть пример надстройки. Запустится приложение Excel 2016. В зависимости от его версии далее надстройка загрузит пользовательскую вкладку WoodGrove на ленте или откроется сама в области задач справа от листа, как показано на приведенных ниже рисунках.

![Надстройка, позволяющая анализировать направления расходных операций через WoodGrove Bank (исходная область задач)] (../images/woodgrove_taskpane_ribbon.PNG)

![Надстройка, позволяющая анализировать направления расходных операций через WoodGrove Bank (исходная область задач)] (../images/woodgrove_taskpane_import.PNG)

## <a name="additional-resources"></a>Дополнительные ресурсы

* [Центр разработки для Office](http://dev.office.com/)

## <a name="copyright"></a>Авторское право
(c) Корпорация Майкрософт (Microsoft Corporation), 2016. Все права защищены.



Этот проект соответствует [правилам поведения Майкрософт, касающимся обращения с открытым кодом](https://opensource.microsoft.com/codeofconduct/). Дополнительную информацию см. в разделе [часто задаваемых вопросов по правилам поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).
