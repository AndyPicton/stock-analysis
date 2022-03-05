# stock-analysis
## Overview of Project
### Purpose
The purpose of this project was to create a more efficient way to analyze stock data between 2017 and 2018 in Excel using VBA.  

### Background
Steve created an excel file to help his parents invest in stocks in green energy. He needs help analyzing all the data because he is worried his parents are not diversifying their stocks enough. 

## Results
The results show that the 12 stocks I analyzed had a much better year in 2017 than 2018. All but one stock, TERP, had a positive return in 2017 and only two stocks, ENPH & RUN, had a positive return in 2018. In 2017, four stocks had returns over 100% and five over 20%, while 2018 had six over -20% returns. ENPH is the stock to invest in, its reutrn was 129.5% in 2017 and then 81.9% in 2018.

![image](https://user-images.githubusercontent.com/99369565/156901448-9eb6a25b-0fef-4ed0-8f92-9c4babcd91dc.png)

In the original script it took 1.11 seconds to run analysis for all stocks in 2017 and 2018.

![image](https://user-images.githubusercontent.com/99369565/156901619-0631ad8d-5ab5-4803-bb02-17fb3e99be46.png)

![image](https://user-images.githubusercontent.com/99369565/156901641-edca44f1-29a3-493a-b9d1-068936d57daa.png)

In the refactored code it took .125 seconds to run analysis for all stocks in 2017 and 2018.

![image](https://user-images.githubusercontent.com/99369565/156901493-42c54d96-28be-4285-9747-16671aca8097.png)

![image](https://user-images.githubusercontent.com/99369565/156901513-ee561c8a-853f-4370-b1cc-d79a5eead7d7.png)

I was able to make the code run faster by changing the nested for loops by adding four arrays. Tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices all allowed me to assign each stock to each of the arrays.

After initiailizing an array of all tickers this is what the starting code looked like:
![image](https://user-images.githubusercontent.com/99369565/156902111-beb44d32-15d1-4377-a787-ccec87b0c0cb.png)

After initiailizing an array of all tickers in the refactored code:
![image](https://user-images.githubusercontent.com/99369565/156902134-6b830175-b478-4910-981a-cdb565185a29.png)

## Summary
There seems to be more advantages than disadvantages in this project in refactoring the original VBA code. The first advantage I saw is that it made my code run faster but it also took more time to write the code. I am willing to invest more time in at first if I know the product will be better, if I scale the code up to look at more stocks the time to run the code will go up as well, being able to cut the run time by 1 whole second is huge.
