+++
date = "2010-11-04T15:55:36+00:00"
title = "SAS EG: Filtering a Dataset via Prompt Manager Variable"
draft = true
tags = ["SAS", "Tech"]
+++

Nobody who reads this blog is likely a SAS programmer. However, on the off chance this would be useful to someone who finds it via Google, I wanted to post. Currently, I am working on a project in SAS Enterprise Guide where I needed to allow the user to provide variable value(s) via prompt manager and then filter the resulting dataset based on the user's input. After rubbing a few brain cells together, created a macro that does the trick. It works by embedding a loop inside of a WHERE statement in SQL. If the user selects one value the loop ends and returns that value into the WHERE statement. However, if multiple values are selected, the loop iterates until all values have been returned. There's probably a better way, but at least this works. Here's a quick outline of the variables. 

`PRODUCTCATEGORY:` This is the name of the prompt variable I setup. Replace it with your variable of choice. 

`IN:` Input dataset 

`OUT:` Output dataset **FIELD:** Variable in the dataset on which the filtering is performed.

```
%MACRO PROCESS(IN,OUT,FIELD); 
PROC SQL; 
CREATE TABLE &OUT. AS 
SELECT * 
FROM &IN. 
WHERE &FIELD. IN ( 
	%IF &PRODUCTCATEGORY_COUNT = 1 %THEN "&PRODUCTCATEGORY"; 
	%ELSE %DO I=1 %TO &PRODUCTCATEGORY_COUNT; 
	"&&PRODUCTCATEGORY&I" %END; 
	); 
QUIT; 
%MEND;
```

---
### Comments:

#### 
[Chris Hemedinger](http://blogs.sas.com/sasdummy "chris.hemedinger@sas.com") - <time datetime="2010-11-09 14:18:05">Nov 2, 2010</time>

Jared, You might be able use a built-in EG macro this with code. Suppose your prompt is named Country, and you've supplied a list of countries to include within the prompt selection. Instead of referencing &Country in your code, you would need to use: %\_eg\_WhereParam( COUNTRY, Country, IN, TYPE=N ) This macro function is supplied by EG 4.2 and should resolve to the proper construct for your program. Your example might turn into: PROC SQL; CREATE TABLE &OUT. AS SELECT \* FROM &IN. WHERE %\_eg\_WhereParam( FIELD, PRODUCTCATEGORY, IN, TYPE=C) ; QUIT; Chris
<hr />

#### 
[Jared]( "jared@monger.cc") - <time datetime="2010-11-09 15:00:25">Nov 2, 2010</time>

Thanks! I wasn't aware that built-in macro existed. Code is much tidier than my original. There's always a better way.
<hr />
