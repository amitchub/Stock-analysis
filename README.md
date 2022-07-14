#  Kickstarting with Excel

## Overview of Project
A visualization and interpretation of Kickstarter campaign outcomes.

### Purpose
Using a curated Kickstarter campaign dataset that includes fields such as success/failure results, categories, subcategories, and pledged dollar amounts, this report will attempt to illustrate the relationship between a project's success or failure with respect to both its launch date and its funding goal.  The focus of the report will be specific to theatrical plays.

---

## Analysis and Challenges

### Analysis of Outcomes Based on Launch Date

Please refer to the following graph:

![Theater Outcomes vs Launch](/Theater_Outcomes_vs_Launch.png)

- The gross amount of launches generally ramp up from January through May, and then begin a decline from there.
- Failed launches remain relatively flat with respect to Successful launches.
- Two thirds of the data is culled from the United States alone for Theater Kickstarter campaigns.  The trendlines of just US data matches that of the whole dataset.  One theory as to why Successful launches peak in May is that many small dollar donors receive their IRS tax refund checks in this timeframe, and have more disposable income than at other times of the year.
- One other theory as to why all launches decline in Q4 is that small dollar donors are saving up to purchase gifts for the holidays.

 ### Analysis of Outcomes Based on Goals
 
 Please refer to the following graph:
 
 ![Outcomes vs Goals](/Outcomes_vs_Goals.png)
 
 - 50 percent of all Theater Play Kickstarter campaigns that ask for less than $15,000 as a goal are funded.
 - Somewhat surprisingly, Theater Play Kickstarter campaigns that ask for anywhere from $35,000 to $45,000 as a goal are also funded.
 - One should take into account, however, that approximately 90% of the total projects were in the 0-to-$14,999 range, whereas less than one percent of all Theater Play campaigns were asking for $35,000-to-$45,000.

### Challenges and Difficulties Encountered

- When creating the table that drives the Outcomes Based on Goals graph, a significant amount of manual labor was required to derive the values.  It was difficult to automate the process of collecting the data for the dataset.  This method is prone to a significant amount of errors, as other peers would attest to.

---

## Results

### Outcomes Based on Launch Date
1. Launching a campaign in the April-June timeframe is the most advantageous for garnering funding.
2. The smallest chance to have a failed campaign is in the November-December timeframe.

### Outcomes Based on Goals
1. Significant success can be found in setting a modest goal (less than $15,000) for funding.
2. Failure of funding rises precipitously as goal amount increases.

### Limitations of the dataset
- Clearly, not having more current data may skew the numbers.  Having just come out the other side of a global pandemic, spending habits for small dollar donors may have changed.
- Different genres of plays / musicals / etc. would be helpful, allowing us to drill down even further to reach conclusions about different types of live performances.

### Data or Tables that should be added
- One enhancement to the Outcomes Based on Goals should be a weight given to the amount of data in the lower ranges.  Over 90% of all goals were under $15,000, so making sweeping conclusions about more expensive or ambitious plays is difficult.
