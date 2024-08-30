+++
date = "2010-12-07T17:29:11+00:00"
title = "Should Accountants Be Required to Code"
draft = true
tags = ["Profession", "SAS", "Featured"]
+++

Sometimes, I wonder where I fit on the career spectrum. I am trained as a CPA and 90% of my daily work involves accounting in some form. However, I am also a programmer who geeks out when I find a new way to automate processes or transform data into knowledge. My boss made a comment a few weeks ago in a meeting, “If you want to get Jared excited, ask him about two things: 1) His family or 2) SAS.” 

Sometimes, I feel like my two halves are exclusive. I don’t know how they fit together. But, then there are days when someone comes to me looking for information (which they have been told was impossible to get), and I can write some code to pull it together in an hour. Those are good days. Here’s the thing. I think programming makes me a better accountant. There is something about starting with a blank slate, then outlining, writing, and debugging code that forces you to think logically. It also helps me to work more efficiently. I had coffee with a former Deloitte colleague over the weekend. We reminisced about 80 hour workweeks and what we learned from those experiences. We concluded that if there is one thing our Big 4 experience taught us to be, it’s efficient. If you were not efficient, you drowned from the sheer volume of work. Efficiency meant I got to have dinner with my wife, instead of staring at workpapers. 

There are certain tasks that humans are good at (recognizing patterns), but there are others where computers are far superior. Combing through thousands of records and joining two tables by hand is something I could do. But, I will probably miss something, and it would take days (if not weeks). OR, I could write something that looks like this and let the computer do it in less than a minute with 100% accuracy. 

```
PROC SQL; 
CREATE TABLE &Table. AS 
SELECT A.\*, B.\* 
FROM &Table1. A 
INNER JOIN &Table2. B 
ON A.&Field1. = B.&Field1.; 
QUIT;
```

Each day, accountants are increasingly asked to take on responsibilities that used to be confined to IT (report writing, data analysis, etc.). However, most of my peers rely on Excel for all but the most complicated analysis. How much time and effort is wasted in performing tasks that a computer should really be doing? This week is Computer Science Education Week. While I don’t think all accountants need to learn C++, I think it’s time our colleges added basic courses in programming to accounting degrees. Like it or not, our profession is shaped by technology. Simple AIS and MIS courses just aren’t enough to prepare students for what is required of them in the workplace, and it will only get more complex from here.


---
### Comments:

#### 
[Chris Hemedinger](http://blogs.sas.com/sasdummy "chris.hemedinger@sas.com") - <time datetime="2010-12-07 15:05:27">Dec 7, 2010</time>

Hear, Hear! In fact, I think you can fill in the blank for lots of different professions: "Programming makes me a better \_\_\_\_\_." Thinking logically about a problem...it never hurts!
<hr />

#### 
[Diane]( "dmiller91@charter.net") - <time datetime="2011-11-22 16:50:13">Nov 22, 2011</time>

I'm an accountant realizing I need more of a programming background, but I don't know where to start! Is SQL for beginners?
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2011-11-22 19:03:57">Nov 22, 2011</time>

Thanks for the comment! SQL is a great place to start. We accountants spend most of our days querying and reconciling data, so this skill set is definitely useful.
<hr />

#### 
[Mel]( "vomel12012@gmail.com") - <time datetime="2012-09-19 22:02:02">Sep 19, 2012</time>

You'll probably do well in accounting. Like programming, the tedium of the work becomes kind of enjoyable.
<hr />

#### 
[Den]( "cheatx@gmail.com") - <time datetime="2012-09-19 20:20:55">Sep 19, 2012</time>

i'm a programmer and i would like to learn accountants stuff. :)
<hr />

#### 
[Chris Hooper](http://www.cirillohooper.com "chris@cirillohooper.com") - <time datetime="2013-03-17 02:03:17">Mar 17, 2013</time>

Really glad I found this post. All of our accountants have basic programming skills, but I'm really looking at ramping it up so we can start building our own proprietary software in house!
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2014-07-14 09:09:00">Jul 14, 2014</time>

Have you thought about going for an entry level accounting position and learning programming on the job? That route worked well for me.
<hr />

#### 
[Ralph]( "ralphhut@gmx.de") - <time datetime="2014-08-25 05:47:23">Aug 25, 2014</time>

I am 39 and started to learn SQL two years ago. The only source for learning was the local library and online resources such as http://www.w3schools.com/sql/ on weekends. This is the first language I have learned. My background is finance only (Master in Finance). After half a year I started creating a data warehouse for our company which holds now 16 entities with eight different currencies, their entire bookings, and a substantial amount of non-financial data. Initially it was only meant for me to learn and to give my superiors fast and more accurate answers. There was only one user to this database (me). Think of it as your own Excel file which you use only for yourself. Now, I do dashboards for the executive board with some basic BI analytics such as predictive analytics, selective reporting, etc. I guess I don't have to mention that my company liked that very much and increased my salary + position. I also like to thing that this helped my company save lots of money on software (+ consulting) to implement all of this. Especially, because it is now multi-user with all of our companies connecting to the database (+100 users). What I am trying to say is: give it a try and don't list to people who are trying to stifle your enthusiasm. Worst case scenario is that you learned something.
<hr />

#### 
[Ruth]( "r.ellenbryant@gmail.com") - <time datetime="2014-07-06 19:01:28">Jul 6, 2014</time>

I'm 33, have an MBA and going for an MS in accounting. My education is so worthless that I leave it off of a lot of job applications. (Because makes me overqualified, and the HR people tell me that to my face.) I tell people that I'm going to learn programming (SQL, Ruby) and they tell me not to bother since I'll be competing with people who have 15 years of experience as programmers. I hate everyone.
<hr />

#### 
[Mathew Heggem](http://www.suminnovation.com "mathew.heggem@suminnovation.com") - <time datetime="2015-01-24 11:34:54">Jan 24, 2015</time>

Hello there! I read this article and was very interested in where you are today with regards to your life as a coder/accountant. I am currently learning to code, HTML/CSS and Javascript, and plan to soon check out Ruby on Rails. I know these are web languages, but I think with the move towards QBO/Xero, these languages may be useful -- or at least interesting and relevant (as if they aren't already....ha!) to accountants. I wonder, though, to what degree this will be valuable. On one hand, it's smart to be able to speak to your customers -- and tech customers could be a great focus, especially with coupled by the fact that I'd be able to 'speak the clients's langugae'; but, how much do I actually need to know in order to 'do something' with it? I will continue to explore, but am curious what your thoughts are too -- as I assume much has happened between today and December 7, 2010.
<hr />

#### 
[Numbermonger]( "jared@monger.cc") - <time datetime="2015-01-26 13:46:09">Jan 26, 2015</time>

Thanks for the comment! Indeed, a lot has happened in four years. My role now is a more people focused and less technical, but I still find myself in situations where I need to bridge the gap between accounting and IT. It's really helpful to speak both languages. Accountants that can't do that are going to be left behind. I don't think you need to be an expert. That's what IT is for. But understanding the constraints and how development works saves time (and helps you know when they are sandbagging). Best of luck.
<hr />
