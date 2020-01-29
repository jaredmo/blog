+++
date = "2011-03-26T16:20:41+00:00"
title = "Mapping an MS Access Database to SAS Library"
draft = false
tags = ["SAS", "Tech"]
+++

This post is for people, like myself, who assume things are harder than they usually are. Worked on an analysis yesterday where I needed to map an Access database to a SAS library so I could use the tables in some joins. My first thought was [ODBC](http://en.wikipedia.org/wiki/Open_Database_Connectivity). I proceeded to setup the connection, only to find out my company didn't give everyone that ability. Bummer. 15 minutes wasted. 

Searched around and couldn't find any other ideas. Then I thought, I'm going to try to just reference the database directly in the LIBNAME statement and see what happens. It worked. So, this post is to save people who, like myself, usually try to hard way first, only to find their way back to simplicity. 

#### Example:
`LIBNAME TEST "C:\test\database.accdb";`

---
### Comments:

#### 
[Pramod](http://www.pramod-r.blogspot.com "getpramod.r@gmail.com") - <time datetime="2011-03-28 00:09:21">Mar 1, 2011</time>

Hey! This is an interesting way to connect to the Access database... All this while, I thought that we would need the Access database engine like the way we do for Oracle/DB2. I had a quick question here.. Would you know how would we connect to the Access/SQL server database which resides in a server through SAS on UNIX? Thanks in Advance!!
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2011-03-28 15:01:56">Mar 1, 2011</time>

Pramod, thanks for your question. I believe you still need the Access engine. I just didn't realize that my company had it, so I was trying to go through ODBC. I don't have much experience with UNIX, but I can recommend some places to ask. There is a Usenet group at http://groups.google.com/group/comp.soft-sys.sas, which has proven very useful when I have a technical question. The forums at http://support.sas.com/forums/index.jspa are also very helpful. Good luck!
<hr />

#### 
[Chris Hemedinger](http://blogs.sas.com/sasdummy "chris.hemedinger@sas.com") - <time datetime="2011-03-29 18:21:28">Mar 2, 2011</time>

Jared, great questions on your blog here! And thanks for referencing the support.sas.com/forums site. You DO need SAS/ACCESS to PC Files to use the LIBNAME assignment to an MS Access database. It also works for Excel files. Using SAS Enterprise Guide, you can import and export MS Access tables and Excel spreadsheets via point-and-click, and SAS/ACCESS is not required. (This method does not use/generate a SAS program -- and it does not support "update-in-place" for an existing table.)
<hr />

#### 
[William Krebs]( "wkrebs@pacbell.net") - <time datetime="2011-03-26 14:06:01">Mar 6, 2011</time>

For my information, do you license any SAS connectivity software, like SAS/ACCESS?
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2011-03-28 14:57:19">Mar 6, 2011</time>

Thanks for reading, William. I don't sell any services through this site, but I'm sure the good folks at SAS would love to talk to you.
<hr />


