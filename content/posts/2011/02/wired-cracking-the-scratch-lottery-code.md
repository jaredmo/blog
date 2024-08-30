+++
date = "2011-02-07T14:30:43+00:00"
title = "Wired: Cracking the Scratch Lottery Code"
draft = true
tags = ["Fun", "News"]
+++

[This](http://www.wired.com/magazine/2011/01/ff_lottery/all/1) was an interesting story from Wired about a Canadian statistician who cracked the Ontario lottery scratch-off algorithm.

> The apparent randomness of the scratch ticket was just a facade, a mathematical lie. And this meant that the lottery system might actually be solvable, just like those mining samples. “At the time, I had no intention of cracking the tickets,” he says. He was just curious about the algorithm that produced the numbers. Walking back from the gas station with the chips and coffee he’d bought with his winnings, he turned the problem over in his mind. By the time he reached the office, he was confident that he knew how the software might work, how it could precisely control the number of winners while still appearing random. “It wasn’t that hard,” Srivastava says. “I do the same kind of math all day long.” ... 

>The trick itself is ridiculously simple. (Srivastava would later teach it to his 8-year-old daughter.) Each ticket contained eight tic-tac-toe boards, and each space on those boards—72 in all—contained an exposed number from 1 to 39. As a result, some of these numbers were repeated multiple times. Perhaps the number 17 was repeated three times, and the number 38 was repeated twice. And a few numbers appeared only once on the entire card. Srivastava’s startling insight was that he could separate the winning tickets from the losing tickets by looking at the number of times each of the digits occurred on the tic-tac-toe boards. In other words, he didn’t look at the ticket as a sequence of 72 random digits. Instead, he categorized each number according to its frequency, counting how many times a given number showed up on a given ticket. “The numbers themselves couldn’t have been more meaningless,” he says. “But whether or not they were repeated told me nearly everything I needed to know.” Srivastava was looking for singletons, numbers that appear only a single time on the visible tic-tac-toe boards. He realized that the singletons were almost always repeated under the latex coating. If three singletons appeared in a row on one of the eight boards, that ticket was probably a winner.

---
### Comments:

#### 
[Chris Hemedinger](http://blogs.sas.com/sasdummy "chris.hemedinger@sas.com") - <time datetime="2011-02-13 16:32:46">Feb 0, 2011</time>

Jared, if you haven't seen it, you might be interested in this series of posted by Rick Wicklin, SAS/IML expert and blogger: http://blogs.sas.com/iml/index.php?/archives/93-Scratch-Off-Lottery-Games-One-Way-to-Design-Them.html
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2011-02-13 17:27:12">Feb 0, 2011</time>

Very cool! No, I hadn't seen those posts yet. Just goes to show how complicated the problem is. This was my favorite quote from the article. “I remember telling myself that the Ontario Lottery is a multibillion-dollar-a- year business,” he says. “They must know what they’re doing, right?” Nope.
<hr />
