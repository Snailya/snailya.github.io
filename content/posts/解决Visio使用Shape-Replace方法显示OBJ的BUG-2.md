+++
title = '解决Visio使用Shape Replace方法显示OBJ的BUG 2'
date = 2023-10-11T14:37:16+08:00
draft = false
categories = ['Visio']
+++

很不幸，又遇到了同样的问题。但这次问题涉及到一个容器对象，该容器内的某个形状链接了容器的某个属性值。当使用Shape.Replace()方法进行更新时，该属性又被更新成“OBJ”。

因此，按照之前提到的方法，我们首先删除了我们自定义的容器的基本属性，例如SelectMode、DisplayMode、CalWH等。然后重新执行更新程序。此时，“OBJ”被正确的属性值所取代。

通过Visio应用程序的编辑模具功能，重新定义先前的基本属性，再次执行更新程序。不幸的是OBJ又出现了。

最终我们发现，当同时满足以下两个条件时，会产生上述的BUG：

1. IsTextEditTarget=False
2. GlueType=8

![image.png](https://s2.loli.net/2023/10/11/W17f94tyPncGdTC.png)
![image.png](https://s2.loli.net/2023/10/13/dq1KO5MIRJYrBDs.png)