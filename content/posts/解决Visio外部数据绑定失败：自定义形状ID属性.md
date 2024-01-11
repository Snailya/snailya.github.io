+++
title = '解决Visio外部数据绑定失败：自定义形状ID属性'
date = 2023-06-25T10:52:20+08:00
draft = false
categories = ['Visio']
+++

## 引言

在 Visio 中，可以通过使用外部数据来创建和更新图标、图形和其它可视化元素。例如，可以通过外部数据绑定，建立形状与数据表之间的同步关系，实现当外部数据发生变化时，Visio 图形的自动更新。利用这个方法，既可以解决人工通过 Excel 批量修改形状数据的需求，又可以实现数据库自动更新数据的需求。

但是，目前针对外部数据功能的教程非常少，均是以组织架构图或时序图为例，以人工拖拽的方式实现数据链接到形状，鲜有如何自动将导入的数据连接到形状的教程。尽管 Microsoft 官方提供了[自动链接向导将导入的数据连接到形状](https://support.microsoft.com/zh-cn/office/%E8%87%AA%E5%8A%A8%E5%B0%86%E5%AF%BC%E5%85%A5%E7%9A%84%E6%95%B0%E6%8D%AE%E9%93%BE%E6%8E%A5%E5%88%B0%E5%BD%A2%E7%8A%B6-be56f5ff-9b13-4311-9a6c-b27dd243dbea)的支持文档，但是仅参照该文档并不能实现我们如下的需求：

1. 通过数据导出功能导出形状的数据至 Excel 表格
2. 通过数据导入及自动链接功能将修改后的 Excel 表格与 Visio 形状数据的自动绑定，实现形状数据的更新。

该需求应于我们的业务过程存在的批量修改某类设备型号、供应商的情况。面对大量的设备，如果通过 Visio 逐次去修改每个形状的数据，一是操作繁琐耗时，二是容易产生遗漏，以上两个问题均可以得到解决。

## 问题描述

借助数据库导出向导，我们可以将指定图层上的形状数据导出至数据库（或 Excel 文档），导出过程中将使用形状的 ID 作为数据表的 Key。因此，在将数据表导回 Visio 时，只需将 Key 值与形状的 ID 关联，理论上应该可以实现自动将数据链接至形状。

![key-field.png](https://s2.loli.net/2023/06/25/KcMqIOdNn2SVwLi.png)

![link-to-object-id.png](https://s2.loli.net/2023/06/25/FguYD1bfpV2OBlo.png)

但是，在实际操作中，外部数据窗口中并没有出现链接符号，即绑定失败了。

![failed-on-linking-shape-id.png](https://s2.loli.net/2023/06/25/lC1HozFsIpYK5rw.png)

此时，如果使用拖拽的方式将数据连接至形状，还会发现 Visio 并没有如期更新已有的属性，而是将表格中的属性作为新的属性写入形状。

![auto-create-properties.png](https://s2.loli.net/2023/06/25/gz3oXulY4riAytB.png)

也就是说，要实现我们的需求需要解决如下两个问题：

1. 自动链接数据至形状数据
2. 更新已有形状数据而不是新建数据

## 原因分析

针对问题一，使用录制宏功能，可以发现当设置形状的 ObjectID 与数据源中的 ShapeKey 匹配时，没有录制上任何有效代码，而其它选项均有代码被录制。因此推测 ID(Object Info)不是有效选项。

![vba-code.png](https://s2.loli.net/2023/06/25/ZrU6nLWj3AzpqKI.png)

针对问题二，由于在导出数据时，允许重新定义属性的名称，因此推测属性名称在更新数据时起关联作用，即表格中的列名称应与属性的 Name 或 Label 一致。经过多次尝试，应设置列名称为属性的 Label。

## 解决方法

因此，要在 Visio 中实现外部数据绑定与自动更新，需要完成以下操作:

1. 在 User Section 中添加新的 ShapeID 属性，并设置其公式为"`=GUARD(ID())`"。
   ![shape-id.png](https://s2.loli.net/2023/06/25/ytJuX4R3IxpNgrs.png)
2. 在数据导出窗口中，依次修改导出属性的 Field Name 为其 Label。
   ![define-field-name.png](https://s2.loli.net/2023/06/25/MRSHE8ay3e2jmoO.png)
3. 在自动链接窗口中，选择 Shape Field 的值为 User.ShapeID.
   ![link-to-shape-id.png](https://s2.loli.net/2023/06/25/RjM9OdT1m4nqphN.png)

此时，修改后的数据已与形状自动关联。

![linked-pump.png](https://s2.loli.net/2023/06/25/JCXdDK13pUy42xH.png)
