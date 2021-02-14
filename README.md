# VBA of Wallstreet
---
## Overview of Project

### Purpose
Using a nested loop in VBA code, it was possible to tally and report investment returns for a handful of stocks over a two-year period for analysis for Steve’s parents. VBA code can complete this task more efficiently than trying to replicate the same results using functions in an Excel worksheet. However, as Steve would like to evaluate a larger number of stocks in the future, using nested loops on a bigger data set might take significantly more time to execute. Refactoring the original code by incorporating the use of arrays, it is possible to filter through more data in less time. 

## Results

### Stock Performance
The goal of this exercise was to determine which, if any, of the green-energy stocks provided a good investment. In 2017, this group of stocks did very well with an average return of 67.3%. Eleven of the 12 stocks reviewed had positive returns for the year including four that more than doubled in value. Contrast this to 2018 where the results were nearly the opposite - ten out of 12 had negative returns. 

Of the two stocks that had positive results both years, ENPH, was the more impressive with growth at 129 and 82% respectively. The other stock, RUN, while having had only a modest 5.5% return in 2017 also boasted a strong return of 84 percent in 2018. Both stocks appear to be solid picks as not only did they have positive returns both years but were also among the top-three traded stocks in the group for 2018.

[![2017-Query-Results.png](https://i.postimg.cc/d136LHwh/2017-Query-Results.png)](https://postimg.cc/S20CPGX4)

[![2018-Query-Results.png](https://i.postimg.cc/4yh11YKZ/2018-Query-Results.png)](https://postimg.cc/PvTZXrzF)

### Script Refactoring
Refactoring involves using already-written script as a template to build a similar but different and/or more efficient query. In this case, refactoring the original code was necessary to be able to quickly analyze a larger portfolio of stocks. The original code used a query that cycled through two separate loops and while the results were available in relatively quick fashion, the data set only included twelve stocks. If the goal is to evaluate hundreds or even thousands of stocks, the code needed to be upgraded.

This was accomplished by removing one of the loops and replacing it with a variable that when combined with use of arrays, meant that the query only had to loop through the data one time (as opposed to 12 times with the un-refactored code). Ultimately, the updated coding allowed the query to run approximately six times faster:

[![2017-Run-Time.png](https://i.postimg.cc/65zQ0tZ4/2017-Run-Time.png)](https://postimg.cc/vcgM8R3G)

[![2018-Run-Time.png](https://i.postimg.cc/jS5S3m6g/2018-Run-Time.png)](https://postimg.cc/SJB4JZvC)

### Summary
As was the case in this exercise, refactoring code can result in better and more efficient queries. It can also be a time-saver as code can be ‘borrowed” and not have to be re-written from scratch. However, refactoring may not always produce the desired short-cut. 

In this exercise, we had two major advantages: 1) a small and relatively simply original query and 2) familiarity with the original query. As such, refactoring the data only included erasing a loop, adding a variable, adding three arrays that tracked with the new variable and one additional if statement. We also had the benefit of notes that more than adequately described the original code. In many cases, with longer, more complicated queries, especially if we are copying the code from another source, it might require more time and work to update the copied code than to create our own. In the end, refactoring can be the right solution if it can be incorporated relatively quickly into our code. 
