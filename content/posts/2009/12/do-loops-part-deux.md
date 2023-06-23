+++
date = "2009-12-11T18:34:04+00:00"
title = "DO Loops: Part Deux (It's a pun. Get it?)"
draft = false
tags = ["SAS", "Tech", "Featured"]
+++

For Part II in this series of posts on DO Loops in SAS, I wanted to discuss an application slightly more complex than the previous example. Loops can not only be used to perform data correction (as shown in the previous post), but can also be very useful in creating new records where none exist. We’ll discuss using order data again as our example. Your client would like you to create a daily trend report showing **YTD Quantity Sold by Product by Date**. In our made-up table, let’s say for simplicity there are four fields: Order Number, Order Date, Product and Quantity. This fictitious company sells two products (Widgets and Super-Widgets) and follows a calendar fiscal year. The data might look something like this.

| OrderNumber | OrderDate | Product      | Quantity |
|-------------|-----------|--------------|----------|
| 2135468     | 1/2/09    | Widget       | 10       |
| 2135469     | 1/3/09    | Widget       | 17       |
| 2135470     | 1/4/09    | Super-Widget | 50       |

You probably noticed, in its current state, the table above cannot generate this report. We don’t have an appropriate date field and quantities aren’t cumulative. The good news is that, utilizing a loop, with a few lines of code we can create a table that will serve our purpose. In order to generate the requested report, we need to create a table that looks like this. _Notice that quantities are cumulative and we have a one record for each product every day._

| RunDate | Product      | Quantity |
|---------|--------------|----------|
| 1/1/09  | Widget       | 0        |
| 1/1/09  | Super-Widget | 0        |
| 1/2/09  | Widget       | 10       |
| 1/2/09  | Super-Widget | 0        |
| 1/3/09  | Widget       | 27       |
| 1/3/09  | Super-Widget | 0        |
| 1/4/09  | Widget       | 27       |
| 1/4/09  | Super-Widget | 50       |

So how do we get from Point A to Point B? Good question. There are two steps. Let’s dive in.

#### STEP #1: It's loop time. 
Using a combination of a DO UNTIL loop and the OUTPUT statement, we can generate a new order record for every date that each order existed within the fiscal year. "What?" You might say. Take the first order (2135468). That order was placed on 1/2/09, which means it should be included in the cumulative quantity from 1/2/09 – 12/31/09. Therefore, we need to create 363 new records for that order which contain the dates 1/3 – 12/31, so when we summarize that quantity is included in the total. This can be accomplished by using the following code. 
```
*Fiscal Year start date;
%LET BEGDTE = ‘01JAN2009’D;

*Date of current reporting;
%LET RPTDTE = ‘04JAN2009’D;

DATA TEST_2;
	SET TEST_1;
	RUNDATE = &RPTDTE.;
	DO WHILE (RUNDATE GE &BEGDTE.);
		IF ORDERDATE le RUNDATE;
		OUTPUT;
		RUNDATE = RUNDATE - 1;
	END;
RUN;
```

This will create a table that contains an individual record for every order on every date that order existed between the reporting date and the order date. If you have a large orders table this step could result in millions of records. Literally.

#### STEP #2: Summarize the data. 

After that we can use a simple PROC SUMMARY to summarize the data and we have successfully moved from Point A to Point B.
```
PROC SUMMARY DATA=TEST_2 NWAY MISSING;
    CLASS RUNDATE PRODUCT;
	VAR QUANTITY;
    OUTPUT OUT=TEST_3 (DROP=_TYPE_ _FREQ_) SUM=;
RUN;
```

We can now generate the requested trend report and the client is happy. What more could you ask.

