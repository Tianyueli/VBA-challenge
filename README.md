# VBA-challenge
Mod 2 Challenge

* Context: This stock market dataset contains an extensive list of stocks and their corresponding  transaction pricing and volume changes on a daily basis.
* Objective: The primary goal of this analysis was to identify the stock worth investing in.
* Solution: I used the VBA scripting to analyze the stock market trend at individual stock level by tracking stocks' performance and price changes.

* Findings
    * 2018 summary
   ![image](https://github.com/Tianyueli/VBA-challenge/assets/42381263/e675f987-89d8-4a02-8f47-8703bcdd6d51)
    * 2019 summary
   ![image](https://github.com/Tianyueli/VBA-challenge/assets/42381263/e3111a44-6868-4805-94b3-391af1b28223)
    * 2020 summary
   ![image](https://github.com/Tianyueli/VBA-challenge/assets/42381263/99a81459-b2df-4a92-a1a8-636845a7e6de)

* Finding Explanation: 
    * This analysis pinpointed the stocks with the greatest increase and the greatest decrease within year 2018, 2019 and 2020. In addition, it highlighted the specific stocks generating the most transaction volume in each year.
       * For example, in year 2020, the stock with the highest growth was Ticker YDI. It increased 188.76% in value and indicates high profitability.
       * Ticker QKN reached the highest total transaction volume among all stocks in year 2020, which could indicate high liquidity and a high interest from buyers competing for this asset.
* Finding Limitations:
    * For the purpose of the assignment, findings are targeted to track the yearly open / close pricing change between the beginning of the year to the end of the year.
    * It could provide additional valuable insight if we also analyzed the annual pricing variance between the highest and lowest pricing points for each stock.
* Citation:
    * When troubleshooting to allow my code to run on all sheets successfully, our TA Randy helped point out I was missing a "ws.Activate" syntax within the For Each loop.
