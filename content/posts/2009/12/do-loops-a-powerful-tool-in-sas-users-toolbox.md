+++
date = "2009-12-09T15:21:29+00:00"
title = "DO Loops: Useful Tool in SAS User's Toolbox"
draft = true
tags = ["SAS", "Tech", "Featured"]
+++

DO Loops in SAS are outside the scope of what most beginning users feel comfortable tackling. However, they are a powerful, time-saving tool in the report writer’s toolbox. For those unfamiliar with programming, loops are a set of instructions that the programmer tells the computer to repeat until predefined condition(s) are true. It’s a simple concept but powerful in execution. 

To illustrate, here is an example of where a loop might be used. Say, you have a table that contains order data. This particular company has decided that its Order Number is seven digits long and starts with “0000000”. The next order in sequence would be “0000001” and so forth. Well, the person who sent you the data decided it would be a good idea to put it into Microsoft Excel, which then reads the Order Number field as numeric and promptly removes all the leading zeros. Perfect. (This has never happened to me…can you tell?) 

However, with three lines of code a DO Loop can easily correct this issue. Here we go. In SAS, there are two types of loops, DO WHILE and DO UNTIL. The main difference is when the logical condition is evaluated. 

#### DO WHILE 
The condition is evaluated at the top of the loop before the statements within the DO loop are executed. If the expression is true, the DO loop iterates. If the expression is false the first time it is evaluated, the DO loop does not iterate. Not even once. 

#### DO UNTIL
The condition is evaluated at the bottom of the loop after the statements in the DO loop have been executed. If the expression is true, the DO loop does not iterate again. The DO loop always iterates at least once. So, using our example above we want to use a DO WHILE loop, which will evaluate the condition before execution and then run if the condition is met. Here is our code. The loop statement is italicized in the middle of the data step.

```
DATA TEST_2;
SET TEST_1;
	DO WHILE (LENGTH(COMPRESS(OrderNumber)) LT 7);
    	OrderNumber="0"||OrderNumber;
	END;
RUN;
```

This code is performing the following steps. 

1. Reads in the first record. 
2. DO Loop looks at the Order Number field. If the length of the number is less than seven digits, then a “0” is appended to the front. 
3. DO Loop checks the condition again. 

If the length is still less than seven digits, the loop iterates again. If not, the loop ends, the next record is read-in, and the process repeats itself. This example gives you an idea of how DO Loops might be used, but their application can go far beyond simple data correction. Until next time. Happy coding.