# Automation of yield curve estimation with Nelson-Siegel model
This is a python code for the automation of the estimation of Yield Curve via Nelson-Siegel model. The code allows for creating a basic plug-in program to estimate the yield curve, calibrate the parameters of the model and save the results in excel spreadsheets.
These are the important points to take into account while running the code:
1) The code requires uploading an excel file, which contains two columns: 'Maturity' (in days) and 'Yield' (of the benchmark bonds). The sample excel spreadsheet is provided as an example.   
2) A basic plug-in program allows for running, calibrating the parameters of the model and visualizing the constructed curves (i.e., zero coupon yield curve, par curve and forward curve).
3) The final results are then saved in excel spreadsheets. Here the code permits to accumulate the results with their respective estimation date. 
