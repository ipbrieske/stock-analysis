## Update Log

2/22/2022 update

Modified All Stocks Analysis to accept user input of ticker and year, verify input, then gather total daily volume and annual return for given inputs. 

2/23/22 Updates

Cleaned message boxes with new lines, fixed Or statement evaluating invalid years. Next step is to create dynamic arrays for tickers() and years() to scan new worksheets for data and populate arrays. 

Created new branch to hold tinkering with dynamic arrays within stockAnalysis(). Comments on Draft Pull Request.

Updated stockAnalysis() to output to new sheet Stock Analysis()

Completed Module 2.3 All Stocks Analysis added Module 2.4 Formatting

Added Module 2.5.2 with year verification recycled from stockAnalysis(). Added interactive buttons on multiple sheets.

Added Module 2.5.3 timer to allStocksAnalysis()

Got significant experience with Git, including creating new branches, merging branches, pulling, committing, pushing, and debugging. 
Questions for later: 
- How do I implement dynamic arrays in VBA? 
- What other GUIs are available on Excel
- How does Excel handle random number generation and time seeding? 
- Is recording questions like this in the readme appropriate?

2/24/22 Updates

Began Module 2 Challenge, renamed primary file to VBA_Challenge.xlsm. Using a for loop, we iterate through an index number to fill four arrays. These values are stored then all analyzed and printed together using another for loop. Periodic save. 

Completed Module 2 Challenge. Also added a table to store program runtimes, need to expand table to separate runtimes by year analyzed and by method of analysis. 

Finished Table recording all runtime tests, ready for analysis. 

2/26/22 Updates

Attempted to finalize table recording runtime test average times using multiple VBA methods, none were yielding success. I was attempting to take the average of the dynamic range of runtime tests, had issues selecting a dynamic range of cells for the .Average() function in VBA. Finally decided to just give the output cells a "=AVERAGE()" function, and this gives me the desired result. Also added screenshots of final run times and averages, as well as screenshots of the main loops driving the AllStocksAnalysis() and AllStocksAnalysisRefactored() subroutines. 

3/1/22
Completed readme and renamed files. Last steps are to insert images, final proofread, and submission. 
