# Excel VBA DRB Rounding
This project provides an alternative mechanism to round data contained within Excel to comply with the Disclosure Review Board (DRB) rounding rules for both unweighted counts (i.e., observations) and estimates (weighted or modeled estimates, regression coefficients, etc.). The base premise is that the user highlights the data of a single type, counts or estimates but not both, and then runs the appropriate macro for the data.

This project contains the Visual Basic for Applications (VBA) script to be used to define an Excel Add-in. The actual Add-In is also included.

For assistance in installing this Excel Add-in, please see the INSTALL.md file for complete step-by-step instructions. The remainder of this document will assume that you have installed the Add-in according to the INSTALL document. Furthermore, I will assume that the Ribbon tab and groups are named as in the INSTALL document. If you used another name, you will need to adjust acccordingly.

## Rounding Unweighted Observation Counts
To round observation counts, including counts of unweighted distributions from survey data, use the DRB_Round_Counts macro. You simply need to select the data in the worksheet, click on the "Round Counts" button in your Ribbon under the "DRB" tab. The counts will be rounded according to DRB requirements.

Note: If the original data is negative or a non-integer, the macro will not round these data since counts should be non-negative integers. It will also reject text.

## Rounding Estimates
To round estimates including weighted estimates, regression coefficients, etc., use the DRB_Round_Estimates macro. The procedure is similar to the one used to round observation counts. You would select just the estimates portion of you data and click on the "Round Estimates" button in your Ribbon. 

Note: Presently, the macro will only reject rounding cells that contain non-numeric data since estimates can be be a number of any form.

## Restoring Original Data
If you find that you rounded the data in error, for example you used the "Round Counts" macro instead of "Round Estimates" macro, then simply click on "Restore Original" and your data should be restored to its unrounded state. In certain error conditions, this may fail. For this reason, you should always work on a copy of your data to prevent the potential loss of information.


