# Stock-Analysis
## Analyzing financial data using VBA

### Overview of Project: Explain the purpose of this analysis.

In this analysis we will help Steve with some advices for his parents, after doing a deep reasearch regarding the behaviour of some green energy stock by analizing their financial information through VBA, building some automated tasks; in order for them could see the outcomes in an easy way to read and then they would be able to decide on which of those stocks would be the best option to invest. 

### Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Therefore we first look into the outcomes on 2018 for DQ, which was the corporation that awake their interest to invest. As result, we could see that maybe they had to consider different options to analize before take their final decision. As you can see on the next image:


![DQ 2018](https://user-images.githubusercontent.com/90433064/135387746-2110c5dd-afbf-4b5f-a0ac-c8ec85556d45.png)


Conclusion: DQ traded the total volume of 107,873,900 shares in 2018, but not enough to mantein the return in a positive number, since on 2018 it end up on -63% of return.

How did we code the previous?


![Code DQ 1](https://user-images.githubusercontent.com/90433064/135396326-3c680b5d-2fef-4b69-bffc-506117df1e5a.png)
![Code DQ 2](https://user-images.githubusercontent.com/90433064/135396334-8078d5ab-6297-4e34-a35e-05a13c72e26c.png)

So then, we suggest to analize 12 possible options of stocks instead of 1 or 2; so Steve's parents could have enough information to know what to do next. Since tha main idea was make the things easy to handle for Steve's parents, so we though in building a way that they would only need to click a button and insert the year they want to visualize the results so they would have them. For that, we create a report for them to see the differences on the outcomes for these 12 stocks. As the next image shows:


![CLICK](https://user-images.githubusercontent.com/90433064/135391386-e9c78972-97e2-461c-9bdc-5fa5f2217ae3.png)

![CLICK 2](https://user-images.githubusercontent.com/90433064/135396772-194d5203-fdaa-4943-bc84-e28a8b780c9b.png)


 For that, we create a nested loop that run over the 12 tickers, give us the sum the total volume for each stock and the return for each as its correspond, with the following code:
 
![NSTD LOOP 1](https://user-images.githubusercontent.com/90433064/135397321-81751085-b206-4ef3-9025-6d4c5d9fb12d.png)
![NSTD LOOP 2](https://user-images.githubusercontent.com/90433064/135397337-81ef9e91-1540-4dcb-8b30-4b821e0a3f6f.png)
![NSTD LOOP 3](https://user-images.githubusercontent.com/90433064/135397344-e91dc90b-0e16-44a0-a850-3b7206d6c794.png)

We also added some formatting to our table to make it easier to read. And since Steve was worried about how fast the VBA would compile the results, to ensure being able in the future, to analyze a bigger dataset, so we started measuring them as well. And after that, we refactored the code to loop trough all the data one time, make it cleaner and actually get better timing in the compiletion of results.


![Refactored code 1](https://user-images.githubusercontent.com/90433064/135399597-f5ec3033-0eba-4610-ab42-cb35e546714c.png)
![Refactored code 2](https://user-images.githubusercontent.com/90433064/135399605-194a279e-f558-4d18-9b67-5029f2c6a47c.png)
![Refactored code 3](https://user-images.githubusercontent.com/90433064/135399627-f2d043a7-0c36-47ae-a8b8-eb03b5cc52fc.png)
![Refactored code 4](https://user-images.githubusercontent.com/90433064/135399631-a18e0e23-30d1-46e7-81de-cc9911a6ef87.png)


So then, at the end of the comparasion of the yearly results of the stocks under the analysis, as we can see on the following image; it would be a good opportunity to advice Steve's parents to take ENPH as the first option to invest, based on the results for both years, this stock behaved pretty well, having almost 130% of return on 2017 and even though on 2018 the % of return decrease to 81.9%, it's still one of the better outcomes of return that we can see on our table. Nevertheless, we could suggest RUN as the second option, having both years in positive numbers of returns, but taking into account that its biggest growth was during 2018, because it went from 5.5% to 84% of return.

On the contrary, the option that does not look good to invest in this moment would be TERP, even when the results for both years are not too far from the positive numbers of return, neither of both years it have reached a positive % of return, meaning that if Steve's parents invest in this stock, they would have good chances to lose money instead of making some profit. 


![Comparasion 17_18](https://user-images.githubusercontent.com/90433064/135402409-6624fbfc-1b30-43f3-a40d-ab784f28985a.png)


## What are the advantages or disadvantages of refactoring code?

As we previously explained, once the code has been refactored the code would loop trough all the data one time, so that means mainly that it would reduce the time for compiling results. And that would be a huge advantage when working with a bigger dataset. 

The only disadvantage that I can see, it would be that you can got lost from the original code and end up with a totally different result with a different purpose.  

## How do these pros and cons apply to refactoring the original VBA script?

  - Pros: I got this improvement on the timing on the compilation of results.

    - **Original VBA script:**
    
    ![Before_refactor 2017](https://user-images.githubusercontent.com/90433064/135405305-00125137-efc9-42f8-b879-6d234feb9d8f.png)
    ![Before_refactor 2018](https://user-images.githubusercontent.com/90433064/135405324-4f2071e3-7d9c-4a6b-b050-c9a9f03822c6.png)
    
    - **Refactoring the script**
    
    ![VBA_Challenge_2017](https://user-images.githubusercontent.com/90433064/135405397-e45884b2-2aae-4d38-a8c0-9aa9c6789c60.png)
    ![VBA_Challenge_2018](https://user-images.githubusercontent.com/90433064/135405425-4cf79acc-aedd-45e8-b391-49bc705fc0fd.png)

   
  - Cons: I could say that I found myself trying to refactor the code, and having as a result a slower code. Finally after a debuggin, I found the missing piece.
