+++
date = "2011-02-11T17:12:37+00:00"
title = "Excel: VLOOKUP vs. MATCH / INDEX. No Contest."
draft = true
tags = ["Excel", "Tech"]
+++

Last week, I posted a [tweet](http://twitter.com/#!/jmonger/status/35012005412741120) declaring that VLOOKUP (the popular joining function in Excel) was for [newbs](http://www.urbandictionary.com/define.php?term=newb) and that cool kids used MATCH and INDEX. I got some interesting responses, so I thought I would take a moment to defend my position and maybe gain some converts. 

#### VLOOKUP in Action 
![](http://numbermonger.files.wordpress.com/2011/02/untitled.jpg)  
Essentially, both sets of functions are simple ways to join two tables that share a common field. Here's an example of VLOOKUP in action. As you can see, we are trying to join the customer name to a table that contains individual orders using the customer number. The function basically says, "Hey Excel, go to this array. Find the first row that has my value in the first column of that array. Return the value from "x" columns over in that same row." Pretty simple. 

#### The Problem 
What if the key column you are using is in the middle of the array and you need to pull a value to its left? Well...you've got a problem. VLOOKUP will only index from left to right. If you could simply enter negative numbers in the column index value there would be no issue...but you can't. 

#### The Solution: MATCH and INDEX
![](http://numbermonger.files.wordpress.com/2011/02/untitled1.jpg)  
MATCH and INDEX are two separate functions in Excel, but used in concert they can be very powerful. Here's how it works. INDEX requires two inputs: array and row number. (It can also accept a column number, but that's a post for another day.) In this example, the array represents the values you would like to populate in your new table. The MATCH function is inserted into the row number section of the INDEX function to supply the appropriate row. As you can see in the example, it doesn't matter if the array is left or right of the needed data, since all INDEX needs is a row number. This is a powerful advantage that, in my estimation, makes it the superior choice.
