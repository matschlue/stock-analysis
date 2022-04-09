# Stock market performance study
## An Analysis of the stock market during the years 2017 and 2018
For this examination stock market data was used to determine the performance of several companies. In order to analyze the market data, Excel macros were programmed in VBA (Visual Basic for Applications) to automate tasks.

## Results
### Stock market performance in 2017
![Alt text](/VBA_Challenge_2017_re.png "Image")

In 2017, almost all stocks achieved a positive return. The best performing stock was _"DQ"_ with a return of almost 200%. Only _"TERP"_ was not able to return value to their shareholders. 

### Stock market performance in 2018
![Alt text](/VBA_Challenge_2018_re.png "Image")

In 2018, stock market returns were significantly lower than in 2017. In fact, all despite two companies (_"ENPH"_ and _"RUN"_) show a negative return for their investors. Worst performers in 2018 are _"DQ"_, _"JKS"_ and _"SPWR"_.  

## Refactoring VBA code
Since there were many steps involved to create this code, an important part was to "refactor" the original code. This has several advantages:
- clean code is easier to improve and update 
- potential time and money saving aspects in future
- reduced complexity for better understanding

Nevertheless there are some drawbacks as well:
- timeline uncertainty, since refactoring can be time consuming
- higher personnel expenses due to time consumption for recoding
- possibility of making mistakes in the new code

### Refactoring Stock market performance macro

Just by refactoring the current stock performance marco, savings in the time to excecute the program were achieved. 

#### Original code execution time for 2017 and 2018
![Alt text](/VBA_Challenge_2017.png "Image")
![Alt text](/VBA_Challenge_2018.png "Image")

Before refactoring, the marco had an execution time of around 1.44 seconds for both years. After modifying the code, both years ran in abount 0.2 seconds.


