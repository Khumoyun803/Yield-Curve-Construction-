# Automation of yield curve estimation with Nelson-Siegel model
This is a python code for the automation of the estimation of Yield Curve via Nelson-Siegel model. The code allows for creating a basic plug-in program to estimate the yield curve, calibrate the parameters of the model and save the results in excel spreadsheets.
These are the important points to take into account while running the code:
1) The following packages are needed to be installed to run the code smoothly:
   ```
   pip install thinker, numpy, pandas, matplotlib
   ```
   An executable file can also be created with typing the following command in your terminal:
   First install 'pyinstaller' to create an 'exe' file from the python script:
   ```
   pip install pyinstaller
   ```
    Then type the following command to create an executable file:
   ```
   pyinstaller yc_prog.py --onefile --noconsole
   ```
3) The code requires uploading an excel file, which contains two columns: 'Maturity' (in days) and 'Yield' (of the benchmark bonds). The sample excel spreadsheet is provided as an example.   
4) A basic plug-in program allows for running, calibrating the parameters of the model and visualizing the constructed curves (i.e., zero coupon yield curve, par curve and forward curve).
5) The final results are then saved in excel spreadsheets. Here the code permits to accumulate the results with their respective estimation dates. 
