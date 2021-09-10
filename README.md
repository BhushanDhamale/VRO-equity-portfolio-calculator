# VRO-stock-portfolio-calculator
This Python script will enable you to get a simple summary of your equity portfolio from the PPTX file downloaded from the Value Research Online (vro.in) portfolio management tool.



Prerequisites for running the script:
• Python 3
• Pandas
• Portfolio XLSX file

If you have your stock portfolio uploaded on the VRO portfolio management tool, you can download it in XLSX format as follows:
1. Go to vro.in and log into your account.
2. Navigate to 'My Investments' tab on the top.
3. Navigate to 'Overview' tab.
4. Click on the orange '⤓Excel' hyperlink in the 'Mutual Funds' or 'Stocks' header.
Your XLSX file should be downloaded now. 

You need the pandas library to run the script. You can install it by running the following command in your terminal:
pip install pandas
or
conda install pandas

Now, you can run the script by running the following command in your terminal:
python3 vro_equity_portfolio_calculator.py <xlsx_from_vro>

The script will automatically open the output file in LibreOffice Calc if you have it installed.
