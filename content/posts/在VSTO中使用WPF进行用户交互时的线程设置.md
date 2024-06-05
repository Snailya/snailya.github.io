+++
title = '在VSTO中使用WPF进行用户交互时的线程设置'
date = 2024-06-05T09:48:52+08:00
draft = false
categories = ['Visio']
+++

许多.NET桌面开发者习惯使用WPF开发用户界面。然而，在VSTO应用程序中使用WPF时，与传统桌面开发相比，必须正确设置线程以防止UI阻塞。

在传统桌面开发中，UI线程即是主线程。因此，我们通常在主线程中实例化Window并调用Show()方法来显示Window。但在VSTO应用程序中，主线程是VSTO_Main线程, Office外接程序将在VSTO_Main中通过COM组件访问Office中的数据以及执行操作。如果在该线程中实例化Window，那么所有用户与Window的交互及程序对COM组件的调用都将共享此线程。试想一下，如果我们打算在Window中展示Office中已有的数据，并在加载数据过程中显示等待动画，此时Window和数据读取程序将轮流使用VSTO_Main线程。当读取数据时，等待动画必须等待数据读取完成才能刷新界面，所以我们会看到等待动画卡顿。

为了解决这个问题，开发者可能会尝试将所有对COM组件的调用放在TaskPool中。然而，由于Office应用程序的单线程单元（STA）特性，在后台线程中调用COM组件会显著影响性能。例如，在一个程序中，需要读取Visio页面中的约300个形状并显示在UI中。如果在VSTO_Main线程中执行读取操作，约需10秒钟；而在后台线程中执行则需要约2分钟。这是因为后台线程在访问COM组件时，需要进行对象封送处理，极大降低了效率。

因此，最佳的解决方案是为WPF页面创建独立的线程。这可以通过以下步骤实现：

1. 创建新线程：在新线程中实例化WPF窗口。
2. 设置线程属性：确保新线程为STA线程，并启动消息循环。
3. 管理线程间通信：通过Dispatcher将COM相关操作放在VSTO_Main中进行。
这样可以避免在主线程中进行繁重的UI操作，同时保持对COM组件的高效访问。

``` CSharp
public partial class ThisAddIn
{
    public static Dispatcher? VSTODispatcher { get; private set; }
    public static Dispatcher? UIDispatcher { get; private set; }

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        // 记录下VSTO_Main线程的Dispatcher，以方便将相COM操作放在VSTO_Main中执行
        VSTODispatcher = Dispatcher.CurrentDispatcher;

        // declare a UI thread to display wpf window
        var uiThread = new Thread(()=>{
                // 记录下UI线程的Dispatcher，以方便在VSTO_Main线程中更新UI
                UIDispatcher = Dispatcher.CurrentDispatcher;

                var window = new MainWindow();
                window.Show();

                // 开启消息循环
                System.Windows.Threading.Dispatcher.Run();
        }) { Name = "UI Thread" };

        // 设置新线程为单线程套间
        uiThread.SetApartmentState(ApartmentState.STA);
        // 启动线程
        uiThread.Start();
    }
    
    // other codes
}

```

现在，我们可以按照传统桌面程序开发的方式，将复杂且耗时的任务放在TaskPool或ThreadPool中执行。当需要访问COM数据或调用COM接口时，通过`VSTODispatcher.Invoke()`和`InvokeAsnyc()`方法将部分代码切换到VSTO_Main线程中执行。

``` CSharp
public async Task SomeTask(){
    // 一些简单的耗时操作但是不使用到COM
    ...

    // 耗时操作需要从COM返回数据
    var shapes = await VSTODispatcher.InvokeAsync(()=>{
        return Globals.ThisAddIn.Application.ActivePage.Shapes.OfType<Shape>().Select(x=> new MyShape(x)).ToList();
    })

    UIDispatcher.Invoke(()=>{
        // 更新UI
        ...
    })
}
```

这样做的好处在于可以充分利用多线程的优势，将繁重的计算和数据处理任务放在后台线程中进行，减少主线程的负载，提高应用程序的响应速度。同时，通过VSTODispatcher.Invoke()方法，可以确保COM数据访问和接口调用在主线程中进行，避免因线程切换而导致的性能问题。
