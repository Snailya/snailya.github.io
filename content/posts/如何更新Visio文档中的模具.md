+++
title = '如何更新Visio文档中的模具'
date = 2023-12-25T11:33:58+08:00
draft = false
categories = ['Visio']
+++

在工程设计领域，流程图设计是许多企业不可或缺的一环。Visio，作为备受推崇的流程图设计工具，正在成为越来越多企业的首选，以期通过规范化的流程设计来提升整体工作效率。

尽管Visio提供了丰富的内置模具库，但在实际应用中，通用模具往往难以满足企业独特的业务需求。为了更好地适应企业的特殊流程和标准，许多企业纷纷转向定制模具库的方向，这与我们从外企的学习中所观察到的趋势是一致的。然而，由于缺乏指导性的方法，企业在推广和运用Visio时可能会面临一系列挑战。其中一个普遍存在的问题是在初版模具库建立后，如何有效地进行迭代更新。

本文将讨论Visio文档中模具更新的实现。首先介绍用户从模具库拖拽至绘图页时的背后过程，揭示为什么文档中的形状实例不会随模具库的更新而更新。随后介绍手动更新的方法。最后，提供比较两种自动化更新的方法及实现刚方法模具需满足的条件。

## 拖拽背后的故事

[当用户首次从模具库拖拽模具到绘图页（Page）时，Visio在后台完成了多个操作](https://techcommunity.microsoft.com/t5/microsoft-365-blog/drag-drop-done/ba-p/237244)。首先，Visio会在文档的文档模具（Document Stencil）中创建该模具的副本，然后再在绘图页上创建针对该模具形状的实例。由于文档模具默认是隐藏的，所以用户可能无法察觉到这一点。（要显示文档模具，首先需要在“选项”-“自定义功能区”中启用“开发者”选项卡。然后，在“开发者”选项卡-“显示/隐藏”分组中勾选“文档模具”。）
当用户再一次从模具库拖拽同一个模具时，Visio将检查文档模具中是否已存在该模具的副本。如果副本已经存在，Visio将直接创建实例。那么Visio是如何判断模具已存在的？在默认情况下Visio会比较模具的UniqueID属性。因此，即使两个模具具有相同的名称，Visio也可以通过UniqueID判断它们的对应关系。每当用户编辑并保存模具时，模具的UniqueID会发生变化。所以拖拽修改后的模具到绘图页时，可以观察到文档模具中出现了新的副本。这也就是为什么修改了模具库中的模具，绘图页中的实例没有被更新。

这显然与我们的期望不符。我们希望图纸中的实例永远与最新的模具一致。

## 手动更新

要将实例引用的模具修改为最新的模具，一种已知的方法是使用“主页”-“更改形状”对图纸中的实例进行更改。但是，对于已包含大量实例的文档，这个操作费时费力。尤其是当新版模具与旧模具的差异并不大时，用户很容易发生遗漏。

![change-shape.png](https://s2.loli.net/2023/12/26/YSPq96nDTp2vEib.png)

## 使用COM批量更新

借助COM组件，我们可以通过创建自动化程序的方式批量选择某一模具的实例，然后调用Shape.ReplaceShape()方法，实现批量更改这些实例的形状。

要获取文档模具中模具在文档中的所有实例，我们可以调用遍历文档中的所有形状，并筛选出Shape.Master等于文档模具中的Master情况。关键代码如下：

``` C#
public IEnumerable<IVShape> GetInstances(IVMaster master) {
    var instances = document.Pages.OfType<IVPage>()
        .SelectMany(x => x.Shapes.OfType<IVShape>()).Where(x => x.Master == master).ToList();
    return instances;
}
```

要找出模具库中对应的新模具，需要使用到模具的另一个ID属性————BaseID。模具的BaseID是在模具被创建的时候生成的，随后不会发生改变。因此，可以通过BaseID找到模具库中的同源模具。但是，使用这种方式时，要求模具库中的BaseID具有唯一性。一种常见的错误是管理员在创建模具时，采用的不是首先在绘图页绘制模具形状再拖拽至模具库，而是直接将模具库中的模具复制成了新的模具并编辑该模具。此时，模具库中的代表不通类型的模具具有相同的BaseID。

关键代码：

``` C#
public IVMaster GetLatestMaster(IVDocument document, string baseID) {
    var latestMaster = document.Masters.OfType<IVMaster>()
        .SingleOrDefault(x => x.BaseID == baseID);
    return latestMaster;
}
```

然后，遍历这些事例，并将形状替换为新版本的模具。

``` C#
public void Replace(IEnumerable<IVShape> instances, IVMaster latestMaster) {
    foreach (var instance in instances) {
        instance.ReplaceShape(latestMaster);
    }
}
```

使用COM方式更新的好处是可以直接在原文件中进行修改。然而，由于UI的频繁更新可能导致方法执行时间较长，特别是在复杂的涂装车间原理图中，可能需要数分钟。因此，为了提升用户的使用体验，开发者可能会考虑加入进度条，以直观地显示更新进度。但是Visio使用STA模型且UI更新过程会向主线程封送消息，如果使用WPF组件，UI线程会发生阻塞，因此应使用WinForm。

另一个不利因素是，更新可能会引发连接线（Connector）的几何属性重新计算。也就是说，直角型连接线的折点位置可能会发生改变。

## 使用OpenXML批量更新

当不要求在原文件中完成更新时，可以考虑直接修改[OpenXML](https://learn.microsoft.com/en-us/office/client-developer/visio/introduction-to-the-visio-file-formatvsdx)文件。OpenXML是一种基于XML的文件格式，在2013年被引入Visio。OpenXML格式的Visio文档后缀为".vsdx"。要查看OpenXML格式的详细内容，可以修改文档后缀为".zip"后打开压缩包。

![open-vsdx-as-zip.png](https://s2.loli.net/2023/12/26/1dKCDJwfYpV6B79.png)

.\visio\masters目录下存储了文档模具的相关内容。masters.xml文件中列出了文档模具中的模具的部分属性。

![master.png](https://s2.loli.net/2023/12/26/BLtjEZDhK4a3crA.png)

![master-content.png](https://s2.loli.net/2023/12/26/38Wx795LYotHfdk.png)

其中，对我们有用的是Master节点的BaseID属性和Rel子节点的r:id属性。前者的作用已在前文中提及。r:id属性可以通过查看_rels文件夹下的masters.xml.rels确定与此Master关联的MasterContents文件。MasterContents文件定义了模具的形状。关键代码（XmlHelper部分的代码参考[以编程方式处理Visio文件格式](https://learn.microsoft.com/en-us/office/client-developer/visio/how-to-manipulate-the-visio-file-format-programmatically)):

``` C#
public IEnumerable<XlElement> GetMasterElements(Package package) {
    // mastersPart指masters.xml
    var mastersPart = package.GetPart(XmlHelper.MastersPartUri);

    // 筛选出BaseID属性为baseID的Master节点
    var masterElements = XmlHelper.GetXElementsByName(mastersPart, "Master").

    return masterElements;
}

public PackagePart GetMasterContentsFile(Package package, string baseID){
    var masterElement = GetMasterElement(package).SingleOrDefault(x=>x.Attribute("BaseID")!.Value == baseID);

    // 通过子节点Rel获取r:id
    var relElement = masterElement.Descendants(XmlHelper.MainNs + "Rel").First();
    var relId = relElement.Attribute(XmlHelper.RelNs + "id")!.Value;
    var rel = mastersPart.GetRelationship(relId);

    // masterPart指master{i}.xml
    var masterPart = package.GetPart(PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
    return masterPart;
}

```

当我们更新文档中的模具时，实际上只需要将Master节点和MasterContents节点替换为修改后的模具库中的对应内容。关键代码：

``` C#
public void Replace(Package drawingDoc, Package stencilDoc, string baseID){
    var mastersPartDrawing = sourcePackage.GetPart(XmlHelper.MastersPartUri);

    foreach (var masterEleDrawing in GetMasterElements(drawingDoc)) {
        // 查看模具库中是否存在对应的Master
        var masterEleStencil = GetMasterElements(stencilDoc).SingleOrDefault(x=>x.Attribute("BaseID")!.Value == baseID);
        if (masterEleStencil == null) continue;

        // 使用模具库中的Master节点替换文档中的Master节点。但是由于模具库中的Rel关系和可能与文档中的不一致，所以为了不去修改masters.xml.rel文件，仍使用原文档中的Rel节点
        var relEleDrawing = masterEleDrawing.Descendants(XmlHelper.MainNs + "Rel").First();
        masterEleStencil.Descendants(XmlHelper.MainNs + "Rel").First().ReplaceWith(relEleDrawing);
        masterEleDrawing.ReplaceWith(masterEleStencil);

        // 替换MasterContents
        var contentsPartDrawing = GetMasterContentsFile(drawingDoc, baseID);
        var contentsPartStencil = GetMasterContentsFile(stencilDoc, baseID);
        XmlHelper.SaveXDocumentToPart(contentsPartDrawing, XmlHelper.GetXmlFromPart(contentsPartStencil));
    }

    XmlHelper.RecalculateDocument(drawingDoc);
    XmlHelper.SaveXDocumentToPart(mastersPartDrawing, XmlHelper.GetXmlFromPart(mastersPartDrawing));
}
```

[查看完整代码](https://github.com/Snailya/AE.PID/blob/main/PID.VisioAddIn/Controllers/Services/DocumentUpdater.cs)
