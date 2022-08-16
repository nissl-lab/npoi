This repository is fork of [NPOI](https://github.com/nissl-lab/npoi). Syncing with upstream repository will be implemented like in the following [article](https://medium.com/sweetmeat/how-to-keep-a-downstream-git-repository-current-with-upstream-repository-changes-10b76fad6d97).

To work with the repository on local machine:

Run Git bash or Command Line. In it run commands:

`cd c:/your_folder_with/repos`

`git clone git@github.com:jake-codes-at-5-am/npoi.git # clones this fork to the location c:/your_folder_with/repos/npoi`

`cd npoi # to get into the folder with the repo`

`git remote add upstream git@github.com:nissl-lab/npoi.git # this will add an upstream remote to our repo`

`git pull upstream master # to pull all the changes from an upstream remote, branch master`



NPOI
===================
[![NuGet Badge](https://buildstats.info/nuget/NPOI)](https://www.nuget.org/packages/NPOI)
[![Ko-Fi](https://img.shields.io/static/v1?style=flat-square&message=Support%20the%20Project&color=success&style=plastic&logo=ko-fi&label=$$)](https://ko-fi.com/tonyqus)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg?style=flat-square&logo=Apache)](LICENSE)
[![traffic](https://api.segment.io/v1/pixel/track?data=ewogICJ3cml0ZUtleSI6ICJBV2NjaWd1UkhKODBuNkJ4WlI4cHRaRzBINzY0RmJObCIsCiAgInVzZXJJZCI6ICJ0b255cXVzIiwKICAiZXZlbnQiOiAiTlBPSSBIb21lcGFnZSIKfQ==
)](#)
<a href="https://github.com/nissl-lab/npoi/graphs/contributors">
    <img
      src="https://img.shields.io/github/contributors/nissl-lab/npoi?logo=github&label=contributors"
      alt="GitHub contributors"
    />
  </a>
<br />
This project is the .NET version of POI Java project. With NPOI, you can read/write Office 2003/2007 files very easily.<br />

[Who is using NPOI?](https://github.com/nissl-lab/npoi/issues/705)

Donation
============
ERC20: 0xD9CA5B6F3BcEa3f10b7C3B92f0EC783FbB47cBE1

![image](https://user-images.githubusercontent.com/772561/184463909-01562041-215a-4eb4-8806-d4128b0d3783.png)

Telegram User Group
================
Join us on telegram: https://t.me/npoidevs

NOTE: QQ or wechat is not recommended. 

Get Started with NPOI
============
[Getting Started with NPOI](https://github.com/nissl-lab/npoi/wiki/Getting-Started-with-NPOI)

[How to use NPOI on Linux](https://github.com/nissl-lab/npoi/wiki/How-to-use-NPOI-on-Linux)

[ORM on NPOI](https://github.com/nissl-lab/npoi/wiki/ORM-on-NPOI)

[NPOI Changelog](https://github.com/nissl-lab/npoi/wiki/Changelog)

Replace DotnetCore.NPOI with NPOI ASAP
===========
NPOI NEVER joins China NCC (.NET Core Community) group. They are cheating. The readme.md in Dotnetcore.npoi repo is full of lies. What they want to do is just to destory NPOI since they cannot make use of reputation of this component any more. That's why I'm always saying they are evil. The whole NCC group is evil. 

NPOI从未加入过中国NCC开源组织，他们在欺骗公众！Dotnetcore.npoi的readme.md完全是诽谤，一堆谎言。他们想做的无非就是想毁掉NPOI，因为他们不能再用NPOI来行骗了。这也是我为什么一直说他们很邪恶，整个NCC组织就是一个邪教。

[Stop using Dotnetcore/NPOI nuget package. It’s obsolete!](https://tonyqus.medium.com/stop-using-dotnetcore-npoi-nuget-package-its-too-obsolete-6d0aeedb3319)

[The real history of Dotnetcore.NPOI](https://tonyqus.medium.com/the-real-history-of-dotnetcore-npoi-999bb5e140c7) [中文版](https://zhuanlan.zhihu.com/p/506975972)

Advantage of NPOI
=================
a. It's totally free to use

b. Cover most features of Excel (cell style, data format, formula and so on)

c. Supported formats: xls, xlsx, docx.

d. Designed to be interface-oriented (take a look at NPOI.SS namespace)

e. Support not only export but also import

f. Real successful cases all over the world

g. [huge amount of basic examples](https://github.com/nissl-lab/npoi-examples)

h. Works on both Windows and Linux 


System Requirement
===================
.NET Standard 2.1 (.NET Core 3.x)

.NET Standard 2.0 (.NET Core 2.x)

.NET Framework 4.0 and above

No Loongson CPU (NPOI will NOT support any issues running on LoognArch architecture. This CPU is a shit)
