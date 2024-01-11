+++
title = '修复Visio容器调整尺寸引发的意外的文件尾错误：正确的容器定义方法'
date = 2023-06-01T10:12:51+08:00
draft = false
categories = ['Visio']
+++

## 引言

容器是 Visio 中一类可以帮助实现结构化组织和管理其它图形元素的特殊对象。通过容器，您可以将相关的子图形放置在一个框架内，以便更清晰地展示信息和关系。在管理员定义方面，与组合相比，容器可以实现更丰富的预定义内容：形状结构、外观样式、连接点定义等，以满足不同公司对图表样式的要求。在用户操作方面，用户可以通过将子图形拖拽在容器内或容器外轻松实现将子图形加入或移出容器，通过拖拽容器可以实现容器内子图形的一起移动。因此，容器尤其适用于图表、流程图、组织结构图的设计。

但是，容器作为 Visio2010 的新增特性，Visio 并没有提供关于自定义容器的相关教程及建议，这也使得当前 Visio 容器的定义存在极高的自由度。通常，若将一个对象定义为容器，只需要在该对象 Spreadsheet 中的 User Section 内增加`msvStructureType`属性，并设置其值为`="Container"`。

![create-a-container.png](https://s2.loli.net/2023/06/01/F7hI19NnCyGziPt.png)
但是，这种简单的定义可能会在特定情况下引发`意外的文件尾`错误。

![unexpected-end-of-file.png](https://s2.loli.net/2023/06/01/TYaG18djwNZW2sL.png)

## 问题描述

星期一的时候，我的同事向我演示了他是如何引起这个错误的。首先，他将一条表示管路的线段放在了表示功能单元的容器内。由于这条管路表示外部进来的管路，所以他希望这条线段一端恰好贴在容器边框，另一端落在容器内。随后，他又删除了这条线段。在此之后，当他想要调整容器的尺寸时，出现了这个错误。
![a-line-in-the-container.png](https://s2.loli.net/2023/06/01/7PwEGvabXVqdio6.png)

![error-when-resize-container-after-remove-line.png.png](https://s2.loli.net/2023/06/01/bDJ3cExyn1Qehsk.png)

当我尝试复现这个错误时，我发现使用 Visio 预定义的容器并不会引发该现象。

## 原因分析

通过仔细对比我们自定义的容器和 Visio 预定义的容器，我们确定了以下两个前置条件：

1. 线段一端必须吸附在容器的几何形状上。当线段被吸附至容器的几何形状上时，点击容器，连接处可以看到如图所示的高亮的连接点。

![glue-line-to-container-geometry.png](https://s2.loli.net/2023/06/01/CDl6AM3PNg4HtLG.png)

在默认状态下，拖拽线段至容器的边界不会引发吸附，除非文档模板中预定义或用户手动开启了该功能。该功能位于 `View`-`Visual aids`-`Snap & Glue`-`General`-`Glue to`节，勾选 Shape geometery 可以启用该功能。

![snap-and-glue.png](https://s2.loli.net/2023/06/01/RVMuCWfdjAyXeiE.png)

2. 被吸附的容器的几何形状必须定义在容器本身，而不是容器组合的子对象。也就是说，表示几何形状的 Geometry Section 位于容器的 Spreadsheet 内。

![geometry-defined-in-container.png](https://s2.loli.net/2023/06/01/imehkFnI6H94NzU.png)

## 解决方法

要解决这个问题，我们需要将容器的几何形状定义在容器组合的子对象中。也就是说，在创建自定义容器时，不应在容器自身的 Spreadsheet 中定义 Geometry，而应将容器转换为组合，并将表示容器几何形状的对象作为容器组合的子对象。

以下给出了建议的自定义容器步骤：

1. 插入表示容器造型的形状对象；
2. 将这些形状对象使用 `Ctrl+G`组合成一个组合对象；
3. 打开组合对象的 Spreadsheet，并在 User Section 内增加`msvStructureType`属性，并设置其值为`="Container"`。