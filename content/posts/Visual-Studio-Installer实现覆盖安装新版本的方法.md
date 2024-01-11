+++
title = 'Visual Studio Installer实现覆盖安装新版本的方法'
date = 2023-12-28T15:13:34+08:00
draft = false
categories = ['部署']
+++

当使用Visual Studio Installer进行打包时，要实现安装时自动卸载旧版本然后安装新版本，需要同时设置以下各点：

1. Deployment Project Properties -> DetectNewerInstalledVersion -> True
2. Deployment Project Properties -> RemovePreviousVersions -> True
3. Deployment Project Properties -> Version

![deployment-project-properties.png](https://s2.loli.net/2023/12/28/5ToctnJNzqDbaks.png)

其中，Version被安装程序用来判断是否继续执行安装程序，所以Version值应大于上一个版本。

满足以上条件时，执行安装程序可以顺利执行，且控制面板中可以看到更新后的版本号。但是，安装程序仍然可能没有正确执行。这是因为项目的主输出并没有被正确拷贝。这往往是因为没有正确设置项目（程序项目，非部署项目）的AssemblyInfo。主输出中的Version将被用来比较是否需要拷贝主输出到安装目录，所以当上一版本的主输出版本为0.2.1.0时，即使已经设置部署项目的Version为0.2.2.0，由于此时主输出的版本仍然为0.2.1.0，安装过程中不会拷贝新生成的主输出到安装目录。

![assembly-info.png](https://s2.loli.net/2023/12/28/BipWZAytP5qGQrl.png)

主输出的Version设置位于程序项目的AssemblyInfo.cs文件中。只有当此文件中的AssemblyVersion或AssemblyFileVersion的值大于上一个版本的值时，才会覆盖原安装目录的dll。
