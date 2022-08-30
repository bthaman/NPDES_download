For one or more NPDES IDs, downloads an Excel file for each, reads the file, and writes results to an output
Excel file. The url is: https://echo.epa.gov/trends/loading-tool/get-data/monitoring-data-download

When the code was initially written in 2019, only Excel output was available. CSV output is now
available (8/2022) and looks like it would be easier to read (TODO is to read from CSV).

--  twdb_asr.py :           downloads Excel from above url. Run from _WWTP_download_tool.xlsm
--  npdes_processing.py     Reads individual Excel files from a hard-coded directory and
                            writes the data to an output Excel file. Run from the vs code terminal.
--  WRAP_read.py            reads a WRAP output file. Run from _WWTP_download_tool.xlsm
--  npdes_describe.py       generates statistics.

requirements:
    selenium==4.4.3
    xlwings==0.15.5
    pandas==1.1.4

    chromedriver.exe ---> updated ~2 months by Google. using version as of 8/30/22