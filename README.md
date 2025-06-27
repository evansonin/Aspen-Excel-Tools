# Aspen Excel Tools
Aspen Excel Tools is a Microsoft Excel extension for Aspen that, as of now, removes some steps in the accounting process.
It has three tools:
1. BOK bank tool
2. ANB bank tool
3. Divvy tool

From my understanding, since around 2016, Microsoft has been moving Office extensions from Visual Basic scripts that run natively in the Office application to integrated webpages rendered using Microsoft Edge WebView2 that interact with the application via the Office JavaScript API. The HTML/JS/CSS files for these webpages cannot be stored locally, and must be requested from a server via HTTPS, which is why it's currently stored here on [GitHub Pages](https://evansonin.github.io/Aspen-Excel-Tools/taskpane.html).

The source code is in the main branch, the actual files Excel loads are in the gh-pages branch.

This readme will explain roughly what each tool does and how it interacts with the network.

## BOK Bank Tool
The BOK bank tool is fairly straightforward. The user downloads a list of transactions as a CSV from BOK, and then imports it into an existing workbook following pre-defined, hard-coded formatting.

Since the banks don't have public APIs, the simplest way to implement this tool was just to open a dialog containing the bank's site where the user can log in and download the CSV export, and then select it as a local file on their PC. From there, it parses the CSV and pastes it into Excel. The only network requests made here are are standard web requests for the bank website through Microsoft Edge.
## ANB Bank Tool
The ANB tool is identical to the BOK tool, except for that it opens ANB's webpage, not BOK.
## Divvy Tool
The Divvy tool is where it gets slightly more interesting. It makes API calls to the Divvy (which I think recently renamed itself to Bill, which is why it's referred to as both) to retrieve transaction lists using [their API](https://developer.bill.com/docs/home).

Since, from what I could tell, it is not possible to directly make calls from the Excel add-in to Divvy, it first goes through a proxy server. This server isn't deployed anywhere, so this tool doesn't work (I don't have the correct API key anyway). A password is required to get a response from the proxy server.

I'll look into directly making API calls from the extension again, since that would certainly make things simpler. Although the proxy server wouldn't be accessible from outside of the network, I still would want it to use HTTPS (in addition to the password) so random people on Aspen's network can't receive lists of Aspen's credit card transactions (or intercept packets containing them).

## The bottom line
The most important thing right now is the hosting of the extension itself. It's currently done on GitHub pages, which isn't ideal, since anyone can access it. Hosting it locally at Aspen or a similar solution where HTTPS can be utilized but it cannot be accessed from the Internet would be ideal. I would also need to be able to publish my updated HTML/JS/CSS files to wherever it's stored, since this extension will probably continue to get updates while I'm here this summer.

Note: The source code is probably full of irrelevant comments, since a lot of this was written by LLMs and then tweaked by me, or vice versa.