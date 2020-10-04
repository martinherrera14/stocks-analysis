# Stocks-Analysis On Green Energy Company's
## **A Quick Overview**
 **Steve** ,*a recent college graduate recently came to us so we can help him and his parent's look into inevsting into* **"green energy company's"**. ***Steve's** parent's are decinding to invest a huge sum of money into "**DAQO**". **Steve** wants us to help them divided up that money to invest into other **"green energy company's"** similar to "**DAQO**". The purpose of this project is to create a **VBA Script** to be able to analyze **"green energy"** stock's for **Steve's** Parent's. However, also creating a **VBA Script** that runs more efficently that can be reused for a large sum of stock's.*
## **Result's** ##
Comparing "**DAQO**" return's from  **2017** and **2018**. "**DAQO**" isn't the right invesment **Steve's** parent's should invest in. Although in **2017** **DAQO** had a succesful year coming in at **("199.4%")** return . In **2018** **DAQO** had a signifciant loss of **("-62.2%")** in stock's.

![VBA_Challenge_2017_2018_Results](https://user-images.githubusercontent.com/71118429/95020568-c93efa80-0620-11eb-8a9d-cd46f4b6280e.png)


Based on my analysis and research . I would highly recommened **Steve's** parent's not to invest in **DAQO** . Although most of the stock's we analyzed including **DAQO** didn't thrive in **2018** in particular there were two stock's or **"green energy company's"** that profited throughout the year and prospered in return's. In **2018** **"ENPH"** return's were **("81.9%")** and **RUN** return's were **("84%")**. These two stock's are worth investing.

![VBA_Challenge_2017_INv](https://user-images.githubusercontent.com/71118429/95006798-05407400-05bd-11eb-8122-e53d2fa835b4.png)
##  **Refactored Execution Time's Compared To Original Execution Time's** ##

**Our New VBA Code**

    Worksheets("2018").Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    Dim startingPrice(12) As Single
    Dim endingPrice(12) As Single
    Dim totalVolume(12) As Long
    
    tickerIndex = 0

    Worksheets("2018").Activate

    'loop over all the rows
    For i = 2 To RowCount

        If Cells(i, 1).Value = tickers(tickerIndex) Then

            'increase totalVolume by the value in the current row
            totalVolume(tickerIndex) = totalVolume(tickerIndex) + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            startingPrice(tickerIndex) = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

            endingPrice(tickerIndex) = Cells(i, 6).Value
            
            Worksheets("All Stocks Analysis").Activate
            
            Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
            Cells(4 + tickerIndex, 2).Value = totalVolume(tickerIndex)
            Cells(4 + tickerIndex, 3).Value = (endingPrice(tickerIndex) / startingPrice(tickerIndex)) - 1
            
            tickerIndex = tickerIndex + 1
            
            Worksheets("2018").Activate
            
        End If

**Refactored Execution Time**

![VBA_Challenge_2017_Time_Elapsed](https://user-images.githubusercontent.com/71118429/95008561-d1bb1500-05cf-11eb-9861-b00f3883744e.png)
![VBA_Challenge_2018_Time_Elapsed](https://user-images.githubusercontent.com/71118429/95008563-d54e9c00-05cf-11eb-9b1a-d3b14f1c99d3.png)

 ***VS***
 
**Original Execution Time**

![original_Time_For2017](https://user-images.githubusercontent.com/71118429/95008490-32961d80-05cf-11eb-8f30-02284fbbbd49.png)
![original_Time_For2018](https://user-images.githubusercontent.com/71118429/95008491-375ad180-05cf-11eb-96bf-96b69fc07c4d.png)

 Overall **refactoring** our *VBA Script/Code* will compile the results more efficently and effectively . In simpler terms our refactored *VBA Code* execute's at a much rapid rate compared to the *Original Script* that took a bit longer to compile.



## **Summary**
*What are the advantages or disadvantages of refactoring code?*
### ***Advantages*** ###
- Refactoring the **code** help's us comprehend and understand the **code** better, But not only understand it we can find **bug's** within the **code** and fix them. 
- Also refactoring the **code** clean's up the overview look of our script and helps the **code** run more efficently .
### ***Disadvantages*** ###
- Refactoring the **code** could be too time consuming due to sum of the new **code's** creating new bug's.
- Adding on to my previous statement about "**Time Consuming**". What if your job give's you a deadline that you must complete with in hour's this could set you back hour's. Plus why ? should a **code** that is running perfectly fine be refactored. In my opinion you should only be refactoring a code if you have plenty of time on your hands and if your job require's you to .
### ***How do these pros and cons apply to refactoring the original VBA script?*** ###
In Conclusion, based on my expierence theres a bit of time difference between starting a script from scratch to refactoring it. Refactoring a script takes up more time due to formatting, Comment's , and creating better **if**-statement's to equation's . Refactoring the **code** may create new bugs where they weren't previously exsisting but refactoring the **code** those execute the script swiftly .
