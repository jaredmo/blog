+++
date = "2009-11-28T19:13:22+00:00"
title = "Actuarial Triangle: Powerful Tool for Accounting Estimates"
draft = false
tags = ["Financial Reporting", "SAS", "Tech", "Featured"]
+++

When individuals outside the profession think of accounting, they don't often consider its subjective nature. Most people assume there are bright line rules and it's the CPA's job to make sure the rules are adhered to. In some instances, there is exact guidance (which makes our lives easier), but in many others CPAs are required to make estimates. Estimates being a nice way of saying, "Mr. CPA, please attempt to predict the future." One of the most useful models I have found in this inexact world is called an actuarial triangle. Actuaries are those fun (usually insurance) people whose job it is to figure out how likely we are to get cancer, have a car accident, or become disabled. For those unfamiliar, a triangle model attempts to predict future activity based on prior history using two related variables. For example, if a CPA is trying to estimate a bad debt reserve, there would be two inputs: sales and write-offs. The data is collected and modeled into a 'triangle'. See example below.

|           | Sales     | 10_2007   | 11_2007   | 12_2007   |
|-----------|-----------|-----------|-----------|-----------|
| WOs Per 1 | 1,000,000 | 13,654    | 12,481    | 18,057    |
| WOs Per 2 | 2,000,000 | 13,954    | 17,417    |           |
| WOs Per 3 | 1,500,000 | 6,056     |           |           |

  
In a triangle, the periods are relative depending on the column being viewed. For example, 'Period 1' in column 11\_2007 represents November's write-offs for items sold in November, Period 2 would be December's write-offs for items sold in November, and the last cell is empty and will eventually contain January's write-offs for November's sales. Notice how the triangle is gradually filled in as time passes. 

Did I confuse you yet? If not, keep reading. It's the CPA's job to turn this triangle into a square by filling in the missing pieces. In the example, this could be accomplished by taking an average percentage of write-offs to sales for each period and then using that percentage to estimate the blank spaces. Once the blank spaces are filled in with estimates, the bad debt reserve would be equal to all of the previously empty cells. This is easier to see in with numbers than with words, so I have attached a [PDF example](/files/triangle_example.pdf) to illustrate. 

If you work with triangles regularly, as I did at one point in my career, then you know these calculations can get intense. Oftentimes it is essential to break estimates down by product, model, or, in the case of accounts receivable, by credit tiers. If you have access to SAS, [here is a Base SAS macro](/files/triangle_macro.doc) which I wrote a few years ago to automate the process. There may be some data manipulation/code tweaking necessary to make it fit your particular needs, but there you go. 

**Attachments:**   
[Triangle Example (PDF)](/files/triangle_example.pdf)   
[Base SAS Macro (Word)](/files/triangle_macro.doc)
