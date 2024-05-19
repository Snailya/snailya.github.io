+++
title = 'AE PID快速入门'
date = 2024-05-19T08:30:21+08:00
draft = false
categories = ['Visio']
+++

本文以中涂洁净间循环风机为例介绍AE PID绘制设备原理图的基本流程。

![sample.png](https://s2.loli.net/2024/05/19/u3pJOTt2UEXwFvj.png)

## 详细步骤

1. 在Visio中使用空白模板创建新的绘图，单位选择“公制单位”；

    ![create-document.png](https://s2.loli.net/2024/05/19/JpaVdy54qF7wQOs.png)

2. 在功能区处切换至AE PID选项卡；

    ![ribbon.png](https://s2.loli.net/2024/05/19/hK6ZgQisEmPIHlX.png)

3. 点击“编辑”-“初始化”按钮对当前绘图进行初始化。初始化操作将在当前页面中插入一个A0图框，并设置文档页面的网格宽度为2.5 mm。

    ![frame-and-grid.png.png](https://s2.loli.net/2024/05/19/RLeACZV3Yu7btdq.png)

    此外，为文档增加2个AE指定样式。其中，“AE Normal”样式是设备单元的默认样式，“AE Pipeline”是管线的默认样式。

   > AE样式将使用“思源黑体”作为指定字体，在AE PID插件的安装过程中会自动为系统安装思源黑体。若未能正确显示思源黑体，请尝试手动安装。

    ![styles.png](https://s2.loli.net/2024/05/19/UG3uiSCWmwhMy92.png)

4. 若要修改图框大小，用鼠标右键点击图框，并在子类中选择合适的尺寸;

    ![context-menu-frame.png](https://s2.loli.net/2024/05/19/fVHZdDqYFS2xyhr.png)

5. 在选项卡中点击“编辑”-“库”加载模具库。

    ![libraries.png](https://s2.loli.net/2024/05/19/IHUv9JCYzA1bQ2h.png)

6. 从“AE逻辑”库中拖拽“功能单元”至绘图区；

    ![drage-functional-group.png](https://s2.loli.net/2024/05/19/QNKqU3kwch1ldtV.png)

   > 当您不确定所需的设备对象属于哪一个类型时，可以借助Visio模具的搜索栏进行搜索。这将帮助您快速找到所需的设备对象，无需事先知道其具体类型。
   >
   > 如果您正在使用Windows11系统，搜索功能可能未能如期使用，请按照[修复了Visio桌面应用中的形状Windows 11](https://support.microsoft.com/zh-cn/office/%E4%BD%BF%E7%94%A8-%E5%BD%A2%E7%8A%B6-%E7%AA%97%E5%8F%A3%E7%BB%84%E7%BB%87%E5%92%8C%E6%9F%A5%E6%89%BE%E5%BD%A2%E7%8A%B6-2e468457-1059-49d3-8955-32b2527cce98)的方法修补。

    用鼠标右键点击该“功能单元”，打开“形状数据”面板;

    ![open-shape-data.png](https://s2.loli.net/2024/05/19/keFbHPUrGioymDt.png)

    在“形状数据”面板中输入“功能组”值：“GF612”，“功能组名称”值：“中途洁净间循环风机”；

    ![funtioanl-group-shape-data.png](https://s2.loli.net/2024/05/19/TXdHVCcjabxorDN.png)

    此时，功能单元显示如下：

    ![functioanl-group.png](https://s2.loli.net/2024/05/19/8TZ4tODMoQUHfB2.png)

7. 使用相同的方法从“AE基础”库中拖拽“鼓风机”至功能组内，此时可以看到功能单元边框被高亮为绿色，表示该风机已被加入功能单元。

    ![fan-in-the-functional-group.png](https://s2.loli.net/2024/05/19/muyJnQ5FlzhdtBW.png)

    保持风机的选中状态，在“形状数据”面板中键入“功能元件”值：11，按下回车后该值将显示为“GQ11”；

    ![fan-shape-data.png](https://s2.loli.net/2024/05/19/yVXTCDESfLc4tja.png)

    拖动GQ11上的黄色控制点，可以移动功能元件标签的位置；

    ![control-of-functioanl-element.png](https://s2.loli.net/2024/05/19/hKoL1NycvArb9tF.png)

8. 为了表示鼓风机配备的电机，从“AE逻辑”库中拖拽“代理功能元件”至功能组内；

    ![functional-element.png](https://s2.loli.net/2024/05/19/fgt8zQnVpBiKs12.png)

    选中该代理功能元件，将黄色控制点拖拽至鼓风机上，使代理功能元件与鼓风机相关联；

    ![assing-functional-element-to-fun.png](https://s2.loli.net/2024/05/19/5Emv1w7ChGMRWxs.png)

    被关联后，代理功能元件的形状数据处可以看见被关联设备的位号；

    ![parent-designation.png](https://s2.loli.net/2024/05/19/Y74CWkbzVtxwidJ.png)

    当被关联设备发生移动时，关联元件会跟随移动；

    ![move-along-with-parent.png](https://s2.loli.net/2024/05/19/R1yqsAmtx7joGYI.png)

    保持代理功能元件的选中状态，补充“元件位号”值：“MA01”，描述：“电机”，并将代理功能元件拖拽至鼓风机附近。

    ![move-functional-element-near-fun.png](https://s2.loli.net/2024/05/19/QcezgOjJqMtTdZx.png)

9. 继续从“AE基础”库中拖拽“仪表”至功能单元中，并在右键菜单中选择“子类”-“本地面板监视仪表”；

    ![selection-of-subclass-for-instrument.png](https://s2.loli.net/2024/05/19/dfZirLgUcHtnNYT.png)

    补充形状数据并将仪表移动至合适的位置；

    ![move-and-supplyment.png](https://s2.loli.net/2024/05/19/I2BPZobUzh1Sxen.png)

    将仪表正中的控制点拖拽至鼓风机上；

    ![move-line.png](https://s2.loli.net/2024/05/19/DcxZOjh5HkY1VPz.png)

10. 使用同样的方法绘制下方的“中控监视操作仪表”；

    ![another-instrument.png](https://s2.loli.net/2024/05/19/oghrQZ63BX8abPx.png)

11. 从“AE管线”库中拖拽“管路”至绘图区，并在“子类”中选择“排出空气”使其显示为柠黄色；

    ![pipe.png](https://s2.loli.net/2024/05/19/vR5ClPgoiUcsn1V.png)

    将管路一段连在鼓风机的连接点上，另一端连接至其他对象，并在右键菜单中点击“改变箭头方向”；

    ![change-arrow-direction.png](https://s2.loli.net/2024/05/19/Vy83WblwJ2LKzDp.png)

    多次点击直到箭头朝向如图所示；

    ![arrow-direction.png](https://s2.loli.net/2024/05/19/Sqj7GgtNnkOCJXZ.png)

12. 使用相同的方法完成另一根管路的绘制；

    ![other-pipe.png](https://s2.loli.net/2024/05/19/h24bCodJ3pYZcLB.png)

13. 在选项卡中点击“导出”-“BOM”查看当前图纸的BOM结构；

    ![bom-structure.png](https://s2.loli.net/2024/05/19/GRrxfnHbIw1yz9K.png)

14. 在选项卡中点击“编辑”-“图例”将在图签上方生成图例；

    ![legend-button.png](https://s2.loli.net/2024/05/19/uemMgcUQsxJPWL7.png)

    图例仅显示图纸中出现的设备。

    ![legend.png](https://s2.loli.net/2024/05/19/IDmfyi8QaurEZkU.png)
