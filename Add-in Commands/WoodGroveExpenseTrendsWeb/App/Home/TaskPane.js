/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // If not using Excel 2016, return
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this add-in only works with newer versions of Excel.");
                return;
            }

            // Invoke the NavBar jQuery plugin if it's available
            if ($.fn.CommandBar) {
                $('.ms-CommandBar').CommandBar();
            }

            // Attach button click event handlers
            $('#show-ytd-transactions').click(showYTDTransactions);
            $('#show-last-year-transactions').click(showLastYearTransactions);
            $('#show-all-transactions').click(showAllTransactions);
            $('#show-selected-categories').click(showSelectedCategories);
            $('#show-all-categories').click(showAllCategories);
            $('#view-transactions').click(viewTransactionsSheet);
            $('#view-dashboard').click(viewDashboard);
            $('#add-donation').click(addToTaxDeductibleItems);
            $('#add-followup').click(addToFollowUpItems);

        });
    }; 
})();

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

        // Queue a command to force a chart refresh
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

        // Queue a command to force a refresh of the chart
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

        // Queue a command to force a chart refresh
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

        // Queue a command to force a refresh of the chart
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

    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {

        // Select all the checkboxes
        $('.category').prop('checked', true);

        // Show selected categories
        showSelectedCategories();
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