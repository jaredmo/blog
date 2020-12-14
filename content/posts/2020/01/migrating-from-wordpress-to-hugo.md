+++
date = "2020-01-24T10:28:28-06:00"
title = "Migrating from Wordpress to Hugo"
draft = true 
tags = ["Tech"]
+++

For around 6 years, I ran a [Wordpress](https://en.wikipedia.org/wiki/WordPress) site called _Numbermonger_ that focused on accounting and technology. The site averaged a few thousand visitors a month - mostly SAS programmers and other technology-minded accountants. Wordpress had a number of features that I enjoyed using, but I eventually came to the conclusion that it was overkill for my needs. Work and personal life also became busier, and my appetite to sit down and write dried up. In 2013, I took the site down, but saved an XML export of all the posts.

Fast forward to late 2019, I started to get interested in new [static site generators](https://en.wikipedia.org/wiki/Web_template_system#Static_site_generators) (SSGs). Rather than require a database and hosting infrastructure, SSGs are used to generate an entire site from content written in [Markdown](https://en.wikipedia.org/wiki/Markdown). Metadata is handled with [TOML](https://en.wikipedia.org/wiki/TOML)/[YAML](https://en.wikipedia.org/wiki/YAML) front matter. Projects such as [Jekyll](https://jekyllrb.com/) and [Hugo](https://gohugo.io/) seemed like a way to bootstrap a personal homepage very quickly, and I could host for free on [GitHub Pages](https://pages.github.com/). 

I was especially drawn to Hugo, since the only setup is downloading a binary and setting the PATH variable in Windows. I watched a few [YouTube videos](https://www.youtube.com/watch?v=qtIqKaDlqXo&list=PLLAZ4kZ9dFpOnyRlyS-liKL5ReHDcj4G3), and found a [project on GitHub](https://github.com/palaniraja/blog2md) which could convert my Wordpress XML export to markdown files.

Overall, the experience was very good. Hugo is fast, and fits my needs better than a full-blown Wordpress installation. I also like that my content is stored in human readable files. If you are thinking about experimenting with a static site generator, give it try. 
