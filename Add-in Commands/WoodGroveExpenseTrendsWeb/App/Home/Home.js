/// <reference path="../App.js" />

(function () {
    "use strict";

    // Declare global variables
    var myBankTransactions;
    var spinnerComponent;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {

            app.initialize();

            // If not using Excel 2016, return
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.2')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
                return;
            }

            $('.ms-CommandBar').CommandBar();

            // Initialize spinner
            var element = document.querySelector('.ms-Spinner');
            spinnerComponent = new fabric['Spinner'](element);

            // Hide the spinner initially
            $(".ms-Spinner").hide();

            // Hide the analyze section until transactions has been imported
            $("#analyze-commandbar").hide();
            $("#create-dashboards-section").hide();
            $("#analyze-transactions-section").hide();

            // Attach button click event handlers
            $('#import-transactions').click(importTransactions);
            $('#show-ytd-transactions').click(showYTDTransactions);
            $('#show-last-year-transactions').click(showLastYearTransactions);
            $('#show-all-transactions').click(showAllTransactions);
            $('#show-selected-categories').click(showSelectedCategories);
            $('#show-all-categories').click(showAllCategories);
            $('#create-dashboards').click(createDashboardAndTrackers);
            $('#view-transactions').click(viewTransactionsSheet);
            $('#view-dashboard').click(viewDashboard);
            $('#add-donation').click(addToTaxDeductibleItems);
            $('#add-followup').click(addToFollowUpItems);

            // Load the sample transactions data from the included JSON file
            $.getJSON('SampleBankData.json', function (data) {
                myBankTransactions = data.Transactions;
            });

            //Finally, create the Welcome sheet
            createWelcomeSheet();
        });
    };

    // Create the Welcome sheet 
    function createWelcomeSheet() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the worksheet collection of existing sheets
            var worksheetsBefore = ctx.workbook.worksheets;

            // Queue a command to load the name property of each worksheet in the collection
            // We will use this later to hide all of the existing sheets from view
            worksheetsBefore.load("name");

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync()
				.then(function () {

				    // Queue a command to add a new Welcome sheet to the workbook
				    var welcomeSheet = ctx.workbook.worksheets.add("Welcome");

				    // Create strings to store all the static content to display in the Welcome sheet
				    var sheetTitle = "WoodGrove Bank";
				    var sheetHeading1 = "With Expense Trends, you can...";
				    var sheetDesc1 = "1.  Import expense transactions into the workbook using the WoodGrove Trends task pane.";
				    var sheetDesc2 = "2.  Create dashboards and trackers.";
				    var sheetDesc3 = "3.  View and analyze trends.";
				    var sheetDesc4 = "4.  Select a transaction in the Transactions sheet and add it as a charitable donation or a follow up item.";

				    //Queue a command to fill white color in the sheet to remove gridlines
				    welcomeSheet.getRange().format.fill.color = "white";

				    // Add all the intro content to the Welcome sheet and format the text
				    addContentToWorksheet(welcomeSheet, "B1:K1", sheetTitle, "SheetTitle");
				    addContentToWorksheet(welcomeSheet, "B5:K5", sheetHeading1, "SheetHeading");
				    addContentToWorksheet(welcomeSheet, "C6:K6", sheetDesc1, "SheetHeadingDesc");
				    addContentToWorksheet(welcomeSheet, "C7:K7", sheetDesc2, "SheetHeadingDesc");
				    addContentToWorksheet(welcomeSheet, "C8:K8", sheetDesc3, "SheetHeadingDesc");
				    addContentToWorksheet(welcomeSheet, "C9:K9", sheetDesc4, "SheetHeadingDesc");

				    //Queue commands to autofit rows and columns in the sheet
				    welcomeSheet.getUsedRange().getEntireColumn().format.autofitColumns();
				    welcomeSheet.getUsedRange().getEntireRow().format.autofitRows();

				    // Queue a command to protect the sheet from user interaction
				    welcomeSheet.protection.protect();

				    // Queue a command to activate Welcome sheet
				    welcomeSheet.activate();

				    // Now queue commands to rename and hide all the sheets that were previously in the workbook. 
				    // We are not overwriting or deleting any existing sheets
				    for (var i = 0; i < worksheetsBefore.items.length; i++) {
				        worksheetsBefore.items[i].name += "Before";
				        worksheetsBefore.items[i].visibility = "hidden";
				    }

				    //Run the queued-up commands, and return a promise to indicate task completion
				    return ctx.sync();
				});
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Import sample transactions into the workbook
    function importTransactions() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Unhide and start the spinner
            $(".ms-Spinner").show();
            spinnerComponent.start();

            // Queue a command to add a new worksheet to store the transactions
            var dataSheet = ctx.workbook.worksheets.add("Transactions");

            // Create strings to store all the static content to display in the Transactions sheet
            var sheetTitle = "WoodGrove Bank";
            var sheetHeading1 = "Expense Transactions - Master List";
            var sheetDesc1 = "This is the master list of your spending activity.";
            var sheetDesc2 = "Filter transactions using the task pane to get insights.";
            var sheetDesc3 = "Track donations and flag items that need follow up.";
            var tableHeading = "Transactions";

            //Queue a command to fill white color in the sheet to remove gridlines from view
            dataSheet.getRange().format.fill.color = "white";

            // Add all the static content to the Transactions sheet and format the text
            addContentToWorksheet(dataSheet, "B1:E1", sheetTitle, "SheetTitle");
            addContentToWorksheet(dataSheet, "B3:E3", sheetHeading1, "SheetHeading");
            addContentToWorksheet(dataSheet, "C4:E4", sheetDesc1, "SheetHeadingDesc");
            addContentToWorksheet(dataSheet, "C5:E5", sheetDesc2, "SheetHeadingDesc");
            addContentToWorksheet(dataSheet, "C6:E6", sheetDesc3, "SheetHeadingDesc");
            addContentToWorksheet(dataSheet, "B19:B19", tableHeading, "TableHeading");

            // Queue a command to add a new table
            var startRowNumber = 20;
            var masterTableAddress = 'Transactions!B' + startRowNumber + ':G' + (startRowNumber + myBankTransactions.length);
            var masterTable = ctx.workbook.tables.add(masterTableAddress, true);
            masterTable.name = "TransactionsTable";

            // Queue a command to set the header row
            masterTable.getHeaderRowRange().values = [["DATE", "AMOUNT", "MERCHANT", "CATEGORY", "TYPEOFDAY", "MONTH"]];

            // Create an array of table data row values
            var tableBodyData = [];
            for (var i in myBankTransactions) {
                // Store all the data in an array
                // Also add calculated columns, typeOfDay and Month
                tableBodyData.push([
					myBankTransactions[i].DATE,
					myBankTransactions[i].AMOUNT,
					myBankTransactions[i].MERCHANT,
					myBankTransactions[i].CATEGORY,
					'=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")',
					'=TEXT([DATE], "mmm - yyyy")'
                ]);
            }

            // Queue a command to set the data body range of the table with the sample data
            masterTable.getDataBodyRange().formulas = tableBodyData;

            // Format the table header and data rows
            addContentToWorksheet(dataSheet, "B" + startRowNumber + ":G" + startRowNumber, "", "TableHeaderRow");
            addContentToWorksheet(dataSheet, "B" + (startRowNumber + 1) + ":G" + (startRowNumber + myBankTransactions.length), "", "TableDataRows");

            // Queue a command to set the number format of the Amount column
            masterTable.columns.getItem("AMOUNT").numberFormat = "$#";

            // Queue a command to sort by most recent transactions at the top (Date, descending order)
            var sortRange = masterTable.getDataBodyRange();
            sortRange.sort.apply([
				{
				    key: 0,
				    ascending: false,
				},
            ]);

            //Queue a command to add the new chart
            var chartDataRangeColumn1 = masterTable.columns.getItemAt(0).getDataBodyRange();
            var chartDataRangeColumn2 = masterTable.columns.getItemAt(1).getDataBodyRange();
            var chartDataRange = chartDataRangeColumn1.getBoundingRect(chartDataRangeColumn2);
            var chart = dataSheet.charts.add("Line", chartDataRange, Excel.ChartSeriesBy.auto);
            chart.setPosition("B8", "G17");
            chart.title.text = "Expense Trends";
            chart.title.format.font.color = "#41AEBD";
            chart.series.getItemAt(0).format.line.color = "#2E81AD";

            // Queue commands to auto-fit columns and rows
            dataSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            dataSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Queue a command to set this sheet as active
            dataSheet.activate();

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .then(function () {
            // Show the Dashboards section and hide the rest
            $("#create-dashboards-section").show();
            $("#import-transactions-section").hide();
            $("#analyze-commandbar").hide();
            $("#analyze-transactions-section").hide();

            // Stop and hide the spinner
            spinnerComponent.stop();
            $(".ms-Spinner").hide();
        })
        .catch(function (error) {
            handleError(error);
        });

    }

    // Filter the transactions table to show YTD transactions
    function showYTDTransactions() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions sheet
            var transactionsSheet = ctx.workbook.worksheets.getItem("Transactions");

            // Queue a command to activate the transactions sheet
            transactionsSheet.activate();

            // Queue a command to get the transactions table
            var table = ctx.workbook.tables.getItem("TransactionsTable");

            // Queue a command to filter the data for the chosen time period
            var filter = table.columns.getItem("DATE").filter;
            filter.apply({
                filterOn: Excel.FilterOn.dynamic,
                dynamicCriteria: Excel.DynamicFilterCriteria.yearToDate
            });

            // Queue a command to force a refresh of the chart
            table.getDataBodyRange().getLastRow().getCell(0, 1).values = 40;

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });

    }

    // Filter the transactions table to show last year's transactions
    function showLastYearTransactions() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions sheet
            var transactionsSheet = ctx.workbook.worksheets.getItem("Transactions");

            // Queue a command to activate the transactions sheet
            transactionsSheet.activate();

            // Queue a command to get the transactions table
            var table = ctx.workbook.tables.getItem("TransactionsTable");

            // Queue a command to filter the data for the chosen time period
            var filter = table.columns.getItem("DATE").filter;
            filter.apply({
                filterOn: Excel.FilterOn.dynamic,
                dynamicCriteria: Excel.DynamicFilterCriteria.lastYear
            });

            // Queue a command to force the refresh of the chart
            table.getDataBodyRange().getLastRow().getCell(0, 1).values = 40;

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Clear the filters to show all transactions
    function showAllTransactions() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions sheet
            var transactionsSheet = ctx.workbook.worksheets.getItem("Transactions");

            // Queue a command to activate the transactions sheet
            transactionsSheet.activate();

            // Queue a command to get the transactions table
            var table = ctx.workbook.tables.getItem("TransactionsTable");

            // Queue a command to clear all filters to show all transactions
            table.clearFilters();

            // Call a function to show selected categories
            showSelectedCategories();

            // Queue a command to force the refresh of the chart
            table.getDataBodyRange().getLastRow().getCell(0, 1).values = 40;

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Filter the transactions table to show selected categories
    function showSelectedCategories() {

        // First get the selected categories from the task pane
        var selectedCategoriesArray = $('input:checkbox:checked.category').map(function () {
            return this.value;
        }).get();

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions sheet
            var transactionsSheet = ctx.workbook.worksheets.getItem("Transactions");

            // Queue a command to activate the transactions sheet
            transactionsSheet.activate();

            // Queue a command to get the transactions table
            var table = ctx.workbook.tables.getItem("TransactionsTable");

            // Queue a command to filter the data for the chosen categories
            var filter = table.columns.getItem("CATEGORY").filter;
            filter.apply({
                filterOn: Excel.FilterOn.values,
                values: selectedCategoriesArray
            });

            // Queue a command to force the refresh of the chart
            table.getDataBodyRange().getLastRow().getCell(0, 1).values = 40;

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Filter the transactions table to show selected categories
    function showAllCategories() {
        // Select all the checkboxes in the UI
        $('.category').prop('checked', true);

        // Call this function to filter the data
        showSelectedCategories();
    }

    // Create the dashboard/trackers
    function createDashboardAndTrackers() {
        // Call the three functions
        createDashboard();
        createDonationsTracker();
        createFollowupItemsTracker();

        // Hide and show the necessary UI elements in the task pane
        $(".jumbotron").hide();
        $("#import-transactions-section").hide();
        $("#create-dashboards-section").hide();
        $("#analyze-commandbar").show();
        $("#analyze-transactions-section").show();
    }

    // Create the dashboard  with summary charts and tables
    function createDashboard() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to add a new worksheet to show the dashboard
            var dashboardSheet = ctx.workbook.worksheets.add("Dashboard");

            // Create strings to store all the static content to display in the Dashboard sheet
            var sheetTitle = "WoodGrove Bank";
            var sheetHeading1 = "Expense Trends Dashboard";
            var sheetDesc1 = "See a summary of your spending activity by category and by month.";
            var sheetDesc2 = "Analyze weekend vs weekday spending.";
            var sheetHeading2 = "Summary";
            var summaryDataHeader1 = "Total Spent";
            var summaryDataHeader2 = "Total Transactions";
            var summaryDataHeader3 = "Average Spent per Transaction";
            var tableHeading1 = "Expenses By Category";
            var tableHeading2 = "Expenses By Month";

            //Queue a command to fill white color in the sheet to remove gridlines
            dashboardSheet.getRange().format.fill.color = "white";

            // Add all the intro content to the Dashboard sheet and format the text
            addContentToWorksheet(dashboardSheet, "B1:E1", sheetTitle, "SheetTitle");
            addContentToWorksheet(dashboardSheet, "B3:E3", sheetHeading1, "SheetHeading");
            addContentToWorksheet(dashboardSheet, "C4:G4", sheetDesc1, "SheetHeadingDesc");
            addContentToWorksheet(dashboardSheet, "C5:E5", sheetDesc2, "SheetHeadingDesc");
            addContentToWorksheet(dashboardSheet, "B7:B7", sheetHeading2, "SheetHeading");
            addContentToWorksheet(dashboardSheet, "B9:B9", summaryDataHeader1, "SummaryDataHeader");
            addContentToWorksheet(dashboardSheet, "B10:B10", summaryDataHeader2, "SummaryDataHeader");
            addContentToWorksheet(dashboardSheet, "B11:B11", summaryDataHeader3, "SummaryDataHeader");
            addContentToWorksheet(dashboardSheet, "B28:D28", tableHeading1, "TableHeading");
            addContentToWorksheet(dashboardSheet, "F28:F28", tableHeading2, "TableHeading");

            // Queue a command to add the Expenses by Category table
            var expensesByCategoryTable = ctx.workbook.tables.add('Dashboard!B29:C29', true);
            expensesByCategoryTable.name = "ExpensesByCategoryTable";

            // Queue a command to set the header row
            expensesByCategoryTable.getHeaderRowRange().values = [["EXPENSE CATEGORY", "TOTAL SPENT"]];

            // Queue a command to add the Expenses by Month table
            var expensesByMonthTable = ctx.workbook.tables.add('Dashboard!F29:I29', true);
            expensesByMonthTable.name = "ExpensesByMonthTable";

            // Queue a command to set the header row
            expensesByMonthTable.getHeaderRowRange().values = [["MONTH", "WEEKEND SPEND", "WEEKDAY SPEND", "TOTAL SPENT"]];

            // Format the table header and data rows
            addContentToWorksheet(dashboardSheet, "B29:C29", "", "TableHeaderRow");

            addContentToWorksheet(dashboardSheet, "B30:C350", "", "TableDataRows");

            addContentToWorksheet(dashboardSheet, "F29:I29", "", "TableHeaderRow");

            addContentToWorksheet(dashboardSheet, "F30:I350", "", "TableDataRows");

            // Queue commands to set the number format of the Amount column in both the summary tables
            expensesByCategoryTable.columns.getItemAt(1).numberFormat = "$#";
            expensesByMonthTable.columns.getItemAt(1).numberFormat = "$#";

            // Next we need to queue commands to add rows to the summary tables

            // First, queue a command to get the transactions table in the Data sheet
            var transactionsTable = ctx.workbook.tables.getItem('TransactionsTable', true);

            // Queue commands to extract the Category and Month column values from the Transactions table and load the column values
            var categoryColumn = transactionsTable.columns.getItem("CATEGORY").getDataBodyRange().load("values");
            var monthColumn = transactionsTable.columns.getItem("MONTH").getDataBodyRange().load("values");

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {

                    // Store the Category column values from the transactions table
                    var categoryColumnValues = categoryColumn.values;

                    // Using a helper function, get the unique categories in the columns
                    var uniqueCategories = getUnique(categoryColumnValues);

                    // For each unique category, queue a command to add a row to the Expenses By Category table with the total amount spent for that category
                    for (var i in uniqueCategories) {
                        expensesByCategoryTable.rows.add(null, [[uniqueCategories[i], 5]]).getRange().getLastColumn().formulas =
							'=SUMIF(transactionsTable[CATEGORY], [@[Expense Category]],transactionsTable[AMOUNT])';
                    }

                    // Store the Month column values from the transactions table
                    var monthColumnValues = monthColumn.values;

                    // Using a helper function, get the unique months in the columns
                    var uniqueMonths = getUnique(monthColumnValues);

                    // for each unique category, queue a command to add a row to the Expenses By Month table with the total amount spent in that month
                    for (var i in uniqueMonths) {
                        expensesByMonthTable.rows.add(null, [[
							uniqueMonths[i],
							'=SUMIFS(TransactionsTable[AMOUNT], TransactionsTable[MONTH],[@MONTH],TransactionsTable[TYPEOFDAY],"Weekend")',
							'=SUMIFS(TransactionsTable[AMOUNT], TransactionsTable[MONTH],[@MONTH],TransactionsTable[TYPEOFDAY],"Weekday")',
							'=SUMIF(transactionsTable[MONTH], [@[MONTH]],transactionsTable[AMOUNT])'
                        ]]);
                    }

                    // Queue commands to show Totals row and set the value
                    expensesByCategoryTable.showTotals = true;
                    expensesByCategoryTable.getTotalRowRange().getLastCell().values = [["=SUM([TOTAL SPENT]"]];
                    expensesByCategoryTable.getTotalRowRange().getLastCell().numberFormat = "$#";

                    // Queue commands to show Totals row and set the value
                    expensesByMonthTable.showTotals = true;
                    expensesByMonthTable.getTotalRowRange().getLastCell().values = [["=SUM([TOTAL SPENT]"]];
                    expensesByMonthTable.getTotalRowRange().getLastCell().numberFormat = "$#";

                    // Queue commands to set the number format for Date and Currency columns
                    dashboardSheet.getRange("C30:C200").numberFormat = "$#";
                    dashboardSheet.getRange("G30:G200").numberFormat = "$#";
                    dashboardSheet.getRange("H30:H200").numberFormat = "$#";
                    dashboardSheet.getRange("I30:I200").numberFormat = "$#";

                    dashboardSheet.getRange("C9").numberFormat = "$#.##";
                    dashboardSheet.getRange("C11").numberFormat = "$#.##";

                    // Queue commands to set the summary data values
                    var rangeTotalSpent = dashboardSheet.getRange("C9");
                    rangeTotalSpent.numberFormat = "$#";
                    rangeTotalSpent.formulas = [["=SUM(ExpensesByCategoryTable[TOTAL SPENT])"]];
                    rangeTotalSpent.format.font.name = "Corbel";
                    rangeTotalSpent.format.font.size = 12;

                    var rangeCountTrans = dashboardSheet.getRange("C10");
                    rangeCountTrans.formulas = [["=COUNT(TransactionsTable[DATE])"]];
                    rangeCountTrans.format.font.name = "Corbel";
                    rangeCountTrans.format.font.size = 12;

                    var rangeAvgSpentPerTrans = dashboardSheet.getRange("C11");
                    rangeAvgSpentPerTrans.numberFormat = "$#.##";
                    rangeAvgSpentPerTrans.formulas = [["=C9/C10"]];
                    rangeAvgSpentPerTrans.format.font.name = "Corbel";
                    rangeAvgSpentPerTrans.format.font.size = 12;

                    // Queue commands to create a doughnut chart for showing % spent on expenses by category
                    var categoryChartDataRange = expensesByCategoryTable.getDataBodyRange();
                    var categoryChart = dashboardSheet.charts.add("3dpie", categoryChartDataRange, Excel.ChartSeriesBy.auto);
                    categoryChart.setPosition("B13", "D25");
                    categoryChart.title.text = "Expenses By Category";
                    categoryChart.title.format.font.size = 10;
                    categoryChart.title.format.font.name = "Corbel";
                    categoryChart.title.format.font.color = "#41AEBD";
                    categoryChart.legend.format.font.name = "Corbel";
                    categoryChart.legend.format.font.size = 8;
                    categoryChart.legend.position = "right";
                    categoryChart.dataLabels.showPercentage = true;
                    categoryChart.dataLabels.format.font.size = 8;
                    categoryChart.dataLabels.format.font.color = "white";
                    var points = categoryChart.series.getItemAt(0).points;
                    points.getItemAt(0).format.fill.setSolidColor("#0C8DB9");
                    points.getItemAt(1).format.fill.setSolidColor("#B1D9F7");
                    points.getItemAt(2).format.fill.setSolidColor("#4C66C5");
                    points.getItemAt(3).format.fill.setSolidColor("#5CC9EF");
                    points.getItemAt(4).format.fill.setSolidColor("#5CCBAD");
                    points.getItemAt(5).format.fill.setSolidColor("#A5E750");
                    points.getItemAt(6).format.fill.setSolidColor("#2E81AD");
                    points.getItemAt(7).format.fill.setSolidColor("#128FEB");
                    points.getItemAt(8).format.fill.setSolidColor("#26387E");

                    // Queue commands to create a column chart and format it
                    var monthChartDataRange = expensesByMonthTable.getRange();
                    var monthChart = dashboardSheet.charts.add("ColumnClustered", monthChartDataRange, Excel.ChartSeriesBy.auto);
                    monthChart.setPosition("F13", "J25");
                    monthChart.title.text = "Expenses By Month";
                    monthChart.title.format.font.size = 10;
                    monthChart.title.format.font.name = "Corbel";;
                    categoryChart.title.format.font.color = "#41AEBD";
                    monthChart.dataLabels.showValue = true;
                    monthChart.legend.position = "right";
                    monthChart.dataLabels.format.font.size = 8;
                    monthChart.dataLabels.format.font.color = "white";
                    monthChart.series.getItemAt(0).format.fill.setSolidColor("#A5E750");
                    monthChart.series.getItemAt(1).format.fill.setSolidColor("#4C66C5");
                    monthChart.series.getItemAt(2).format.fill.setSolidColor("#5CC9EF");

                    // Queue commands to auto-fit columns and rows
                    dashboardSheet.getUsedRange().getEntireColumn().format.autofitColumns();
                    dashboardSheet.getUsedRange().getEntireRow().format.autofitRows();

                    // Queue a command to protect the report
                    dashboardSheet.protection.protect();

                    // Queue a command to set the sheet as active
                    dashboardSheet.activate();

                    //Run the queued-up commands, and return a promise to indicate task completion
                    return ctx.sync();
                })
        })
            .catch(function (error) {
                handleError(error);
            });
    }

    // Create the charitable donations tracker
    function createDonationsTracker() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to add a new worksheet to track the donations
            var donationsSheet = ctx.workbook.worksheets.add("Donations");

            // Create strings to store all the static content to display in the Donations sheet
            var sheetTitle = "WoodGrove Bank";
            var sheetHeading1 = "Donations Tracker";
            var sheetDesc1 = "Track your charitable contributions throughout the year.";
            var sheetDesc2 = "Use this data at the end of the year to report your tax deductions.";
            var sheetHeading2 = "Summary";
            var summaryDataHeader1 = "Total Donations";
            var tableHeading1 = "Donations By Organization";
            var tableHeading2 = "Donations By Month";
            var tableHeading3 = "Transaction Details";

            //Queue a commad to fill white color in the sheet to remove gridlines
            donationsSheet.getRange().format.fill.color = "white";

            // Add all the intro content to the Donations sheet and format the text
            addContentToWorksheet(donationsSheet, "B1:G1", sheetTitle, "SheetTitle");
            addContentToWorksheet(donationsSheet, "B3:C3", sheetHeading1, "SheetHeading");
            addContentToWorksheet(donationsSheet, "C4:G4", sheetDesc1, "SheetHeadingDesc");
            addContentToWorksheet(donationsSheet, "C5:J5", sheetDesc2, "SheetHeadingDesc");
            addContentToWorksheet(donationsSheet, "B7:B7", sheetHeading2, "SheetHeading");
            addContentToWorksheet(donationsSheet, "B9:C9", summaryDataHeader1, "SummaryDataHeader");
            addContentToWorksheet(donationsSheet, "B11:D11", tableHeading1, "TableHeading");
            addContentToWorksheet(donationsSheet, "E11:F11", tableHeading2, "TableHeading");
            addContentToWorksheet(donationsSheet, "H11:K11", tableHeading3, "TableHeading");

            // Queue a command to add the Transaction Details table
            var donationsTable = ctx.workbook.tables.add('Donations!H12:K12', true);
            donationsTable.name = "DonationsTable";

            // Queue a command to set the header row
            donationsTable.getHeaderRowRange().values = [["DATE", "AMOUNT", "ORGANIZATION", "MONTH"]];

            // Queue a command to add the Donations by Organization table
            var donationsByOrgTable = ctx.workbook.tables.add('Donations!B12:C12', true);
            donationsByOrgTable.name = "DonationsByOrgTable";

            // Queue commands to set the header, show Totals row
            donationsByOrgTable.getHeaderRowRange().values = [["ORGANIZATION", "AMOUNT"]];
            donationsByOrgTable.showTotals = true;
            donationsByOrgTable.getTotalRowRange().getLastCell().values = [["=SUM([AMOUNT]"]];

            // Queue a command to add the Summary Donations by Month table
            var donationsByMonthTable = ctx.workbook.tables.add('Donations!E12:F12', true);
            donationsByMonthTable.name = "DonationsByMonthTable";

            // Queue commands to set the header, show Totals row
            donationsByMonthTable.getHeaderRowRange().values = [["MONTH", "AMOUNT"]];
            donationsByMonthTable.showTotals = true;
            donationsByMonthTable.getTotalRowRange().getLastCell().values = [["=SUM([AMOUNT]"]];

            // Format the header and data rows of both the summary tables
            addContentToWorksheet(donationsSheet, "B12:C12", "", "TableHeaderRow");
            addContentToWorksheet(donationsSheet, "B13:C250", "", "TableDataRows");
            addContentToWorksheet(donationsSheet, "E12:F12", "", "TableHeaderRow");
            addContentToWorksheet(donationsSheet, "E13:F250", "", "TableDataRows");
            addContentToWorksheet(donationsSheet, "H12:K12", "", "TableHeaderRow");
            addContentToWorksheet(donationsSheet, "G13:J250", "", "TableDataRows");

            // Queue commands to set the number format for Date and Currency columns
            donationsSheet.getRange("C13:C200").numberFormat = "$#";
            donationsSheet.getRange("F13:F200").numberFormat = "$#";
            donationsSheet.getRange("I13:I200").numberFormat = "$#";
            donationsSheet.getRange("E13:E200").numberFormat = "@";
            donationsSheet.getRange("K13:K200").numberFormat = "mmm-yyyy";
            donationsSheet.getRange("H13:H200").numberFormat = "mm/dd/yyyy";
            donationsByOrgTable.getTotalRowRange().getLastCell().numberFormat = "$#";
            donationsByMonthTable.getTotalRowRange().getLastCell().numberFormat = "$#";

            // Queue commands to set the value of Total Donations at the top
            var rangetotalDonated = donationsSheet.getRange("D9");
            rangetotalDonated.formulas = [["=SUM(DonationsTable[AMOUNT])"]];
            rangetotalDonated.format.font.name = "Corbel";
            rangetotalDonated.format.font.size = 18;
            rangetotalDonated.numberFormat = "$#";
            rangetotalDonated.merge();

            // Queue commands to auto-fit columns and rows
            donationsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            donationsSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
		.catch(function (error) {
		    handleError(error);
		});
    }

    // Create the follow up items tracker
    function createFollowupItemsTracker() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to add a new worksheet to store the follow up items
            var followupSheet = ctx.workbook.worksheets.add("FollowUp");

            // Create strings to store all the static content to display in the Follow Up sheet
            var sheetTitle = "WoodGrove Bank";
            var sheetHeading1 = "Follow Up Items Tracker";
            var sheetDesc1 = "Track the transactions that need follow up in this tracker.";
            var sheetDesc2 = "After follow up, mark it as Complete so it is filtered out of the view.";
            var sheetHeading2 = "Summary";
            var summaryDataHeader1 = "Items Pending";
            var tableHeading1 = "Follow Up Items List";

            //Queue a commad to fill white color in the sheet to remove gridlines
            followupSheet.getRange().format.fill.color = "white";

            // Add all the static content to the Follow Up sheet and format the text
            addContentToWorksheet(followupSheet, "B1:E1", sheetTitle, "SheetTitle");
            addContentToWorksheet(followupSheet, "B3:C3", sheetHeading1, "SheetHeading");
            addContentToWorksheet(followupSheet, "C4:G4", sheetDesc1, "SheetHeadingDesc");
            addContentToWorksheet(followupSheet, "C5:G5", sheetDesc2, "SheetHeadingDesc");
            addContentToWorksheet(followupSheet, "B7:B7", sheetHeading2, "SheetHeading");
            addContentToWorksheet(followupSheet, "B9:C9", summaryDataHeader1, "SummaryDataHeader");
            addContentToWorksheet(followupSheet, "B11:D11", tableHeading1, "TableHeading");

            // Queue a command to add a new table
            var followupTable = ctx.workbook.tables.add('FollowUp!B12:E12', true);
            followupTable.name = "FollowUpTable";

            // Queue a command to set the header row
            followupTable.getHeaderRowRange().values = [["TRANSACTION DATE", "ORGANIZATION", "AMOUNT", "STATUS"]];

            // Format the table header and data rows
            addContentToWorksheet(followupSheet, "B12:E12", "", "TableHeaderRow");
            addContentToWorksheet(followupSheet, "B13:E250", "", "TableDataRows");

            // Queue commands to set the number format
            followupSheet.getRange("B12:B200").numberFormat = "mm/dd/yyyy";
            followupSheet.getRange("D12:D200").numberFormat = "$#";

            // Queue commands to set the value of Items Pending at the top
            var rangetotalItems = followupSheet.getRange("D9");
            rangetotalItems.formulas = [['=COUNTIF(FollowUpTable[STATUS], "pending")']];
            rangetotalItems.format.font.name = "Corbel";
            rangetotalItems.format.font.size = 18;

            // Queue commands to auto-fit columns and rows
            followupSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            followupSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();

        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Switch to the Transactions sheet
    function viewTransactionsSheet() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the transactions sheet
            var transactionsSheet = ctx.workbook.worksheets.getItem("Transactions");

            // Queue a command to activate the transactions sheet
            transactionsSheet.activate();

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Switch to the Transactions sheet
    function viewDashboard() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Queue a command to get the Dashboard sheet
            var dashboardSheet = ctx.workbook.worksheets.getItem("Dashboard");

            // Queue a command to activate the Dashboard sheet
            dashboardSheet.activate();

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Add the selected transaction to donations tracker
    function addToTaxDeductibleItems() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // First switch to the Transactions sheet in case it's not already active
            viewTransactionsSheet();

            // The new row to be added
            var rowToAdd;

            // Create a proxy object for the selected range and load its address and values properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, address");

            // Queue a command to get the Transactions table
            var transactionsTable = ctx.workbook.tables.getItem("TransactionsTable");

            // Maintain a variable to ensure that a valid table cell is selected
            var keepGoing = true;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    rowToAdd = sourceRange.getEntireRow().getIntersection(transactionsTable.getRange());
                    rowToAdd.load("values");
                    rowToAdd.format.fill.color = "#92E4F0";
                })
                // Then run the queued-up commands, and return a promise to indicate task completion
                .then(ctx.sync)
                .catch(function () {
                    keepGoing = false;
                })
                .then(function () {
                    if (!keepGoing) {
                        return;
                    }

                    // Get the donations sheet
                    var donationsSheet = ctx.workbook.worksheets.getItem("Donations");

                    // Get the donations table
                    var donationsTable = ctx.workbook.tables.getItem("DonationsTable");

                    // Create a proxy object for the table rows
                    var tableRows = donationsTable.rows;

                    // Queue commands to add some sample rows to the donations table
                    tableRows.add(null, [[rowToAdd.values[0][0],
                                            rowToAdd.values[0][1],
                                            rowToAdd.values[0][2],
                                            '=(TEXT([@DATE], "mmm - yyyy"))']]);

                    // Auto-fit columns and rows
                    donationsSheet.getUsedRange().getEntireColumn().format.autofitColumns();
                    donationsSheet.getUsedRange().getEntireRow().format.autofitRows();

                    // Set the sheet as active
                    donationsSheet.activate();

                    //Run the queued-up commands, and return a promise to indicate task completion
                    return ctx.sync();
                })
                .then(function () {
                    if (keepGoing) {
                        updateDonationTrackerSummaryTables();
                    }
                });
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Update the summary tables in the donations tracker
    function updateDonationTrackerSummaryTables() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Get the tables
            var donationsTable = ctx.workbook.tables.getItem("DonationsTable");
            var donationsByOrgTable = ctx.workbook.tables.getItem("DonationsByOrgTable");
            var donationsByMonthTable = ctx.workbook.tables.getItem("DonationsByMonthTable");

            // Get the columns and the last row added and load their values
            var orgColumn = donationsByOrgTable.columns.getItemAt(0).load("values");
            var monthColumn = donationsByMonthTable.columns.getItemAt(0).load("values");
            var lastRowAdded = donationsTable.getDataBodyRange().getLastRow().load("values");

            //Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {

                    // Convert the organization column values from the 2d array
                    var orgColumnValueArrays = orgColumn.values;
                    var orgColumnValueArray = orgColumnValueArrays.map(function (item) { return item[0] });

                    // Check if you've already donated to this organization. 
                    // If so, get the index of the item in the summary table so we can update that row.
                    // change to currentorg
                    var indexOfOrg = orgColumnValueArray.indexOf(lastRowAdded.values[0][2]);

                    // If the organization is new, add a new row to the summary table
                    if (indexOfOrg === -1) {
                        donationsByOrgTable.rows.add(null, [[lastRowAdded.values[0][2], lastRowAdded.values[0][1]]]);
                        // return promise here as well
                    }
                        // If this organization has been donated to already, update the existing row
                    else {
                        var amountColumn = donationsByOrgTable.columns.getItemAt(1).load("values");
                        return ctx.sync()
							.then(function () {
							    donationsByOrgTable.getDataBodyRange().getCell(indexOfOrg - 1, 1).values =
									[[amountColumn.values[indexOfOrg][0] + lastRowAdded.values[0][1]]];
							});
                    }
                })
                // Do the same for the Donations By Month table
                .then(function () {
                    var monthColumnValueArrays = monthColumn.values;
                    var monthColumnValueArray = monthColumnValueArrays.map(function (item) { return item[0] });

                    var indexOfMonth = monthColumnValueArray.indexOf(lastRowAdded.values[0][3]);

                    if (indexOfMonth === -1) {
                        donationsByMonthTable.rows.add(null, [[lastRowAdded.values[0][3], lastRowAdded.values[0][1]]]);
                    }
                    else {
                        var amountColumn = donationsByMonthTable.columns.getItemAt(1).load("values");
                        return ctx.sync()
                            .then(function () {
                                donationsByMonthTable.getDataBodyRange().getCell(indexOfMonth - 1, 1).values =
									[[amountColumn.values[indexOfMonth][0] + lastRowAdded.values[0][1]]];
                            });
                    }
                })
            .then(ctx.sync)
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Add the selected transaction to Follow Up Items tracker
    function addToFollowUpItems() {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // First switch to the Transactions sheet in case it's not already active
            viewTransactionsSheet();

            var rowToAdd;

            // Create a proxy object for the selected range and load its address and values properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, address");

            // Get the table
            var transactionsTable = ctx.workbook.tables.getItem("TransactionsTable");

            // Maintain a variable to ensure that a valid table cell
            var keepGoing = true;

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    rowToAdd = sourceRange.getEntireRow().getIntersection(transactionsTable.getRange());
                    rowToAdd.load("values");
                    rowToAdd.format.fill.color = "#54ACB8";
                })
                .then(ctx.sync)
                .catch(function () {
                    keepGoing = false;
                })
                .then(function () {

                    if (!keepGoing) {
                        keepGoing = true;
                        return;
                    }

                    // Get the followup sheet
                    var followupSheet = ctx.workbook.worksheets.getItem("FollowUp");
                    var followupTable = ctx.workbook.tables.getItem("FollowUpTable");

                    // Create a proxy object for the table rows
                    var tableRows = followupTable.rows;

                    // Queue commands to add some sample rows to the donations table
                    tableRows.add(null, [[rowToAdd.values[0][0], rowToAdd.values[0][2], rowToAdd.values[0][1], 'Pending']]);

                    // Auto-fit columns and rows
                    followupSheet.getUsedRange().getEntireColumn().format.autofitColumns();
                    followupSheet.getUsedRange().getEntireRow().format.autofitRows();

                    // Set the sheet as active
                    followupSheet.activate();

                })

                // Run the queued-up commands
            .then(ctx.sync)
            .then(function () {
                var followupTable = ctx.workbook.tables.getItem("FollowUpTable");
                // After adding the row, apply the Pending filter
                var filter = followupTable.columns.getItemAt(3).filter;
                filter.apply({
                    filterOn: Excel.FilterOn.values,
                    values: ["Pending"]
                });
            })
        })
        .catch(function (error) {
            handleError(error);
        });
    }

    // Helper function that returns unique items in an array
    function getUnique(inputArray) {
        var outputArray = [];

        for (var i = 1; i < inputArray.length; i++) {
            if (($.inArray(inputArray[i][0], outputArray)) == -1) {
                outputArray.push(inputArray[i][0]);
            }
        }

        return outputArray;
    }

    // Handle errors
    function handleError(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        app.showNotification("Error: " + error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function to add and format content in the workbook
    function addContentToWorksheet(sheetObject, rangeAddress, displayText, typeOfText) {

        // Format differently by the type of content
        switch (typeOfText) {
            case "SheetTitle":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 30;
                range.format.font.color = "white";
                range.merge();
                //Fill color in the brand bar
                sheetObject.getRange("A1:M1").format.fill.color = "#41AEBD";
                break;
            case "SheetHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 18;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "SheetHeadingDesc":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.merge();
                break;
            case "SummaryDataHeader":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 13;
                range.merge();
                break;
            case "SummaryDataValue":
                var range = sheetObject.getRange(rangeAddress);
                range.numberFormat = numberFormat;
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 13;
                range.merge();
                break;
            case "TableHeading":
                var range = sheetObject.getRange(rangeAddress);
                range.values = displayText;
                range.format.font.name = "Corbel";
                range.format.font.size = 12;
                range.format.font.color = "#00b3b3";
                range.merge();
                break;
            case "TableHeaderRow":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.format.font.bold = true;
                range.format.font.color = "black";
                break;
            case "TableDataRows":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                sheetObject.getRange(rangeAddress).format.borders.getItem('EdgeBottom').style = 'Continuous';
                sheetObject.getRange(rangeAddress).format.borders.getItem('EdgeTop').style = 'Continuous';
                break;
            case "TableTotalsRow":
                var range = sheetObject.getRange(rangeAddress);
                range.format.font.name = "Corbel";
                range.format.font.size = 10;
                range.format.font.bold = true;
                break;
        }
    }
})();