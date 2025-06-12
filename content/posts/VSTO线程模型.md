+++
title = 'VSTO线程模型'
date = 2024-06-05T09:48:52+08:00
draft = false
categories = ['Visio']
+++

在VSTO开发领域，每个开发者都曾经历过这样的挫败：精心设计的插件界面在数据加载时陷入卡顿，复杂的计算过程使整个Office程序失去响应。这不是代码质量问题，而是VSTO特有的线程模型带来的挑战。本文将围绕如何利用独立线程和异步调度解决该问题进行剖析。

## VSTO线程模型的本质矛盾

Office的COM对象是基于单线程单元（STA, Single Threaded Apartment）模型的。这个模型的好处是保障了线程安全，但也造成了性能瓶颈。具体表现：

* UI线程和Office线程绑在一起，做耗时数据读取或计算时，界面会卡死甚至完全无响应。
* 想用多线程访问COM对象？ 不行，需要走代理（Proxy）封送（marshalling）调用，开销大且复杂。

这种设计就像双刃剑，让我们陷入两难：一边要保证线程安全，一边又想流畅不卡顿。

## 突破瓶颈的思路：线程分离 + 异步计算

要打破这个瓶颈，必须实现UI线程与Office线程的分离，并确保所有COM对象的访问都发生在Office线程上。同时，为避免业务逻辑阻塞UI或Office线程，计算应异步在线程池（TaskPool）中进行。

这带来关键问题：如何在不同线程间高效且安全地传递数据？

* 如何在后台任务中发起Office数据读取，并调度到VSTO主线程执行？
* 数据处理完成后，如何将结果发回UI线程，满足UI线程禁止跨线程更新界面的要求？
 
那具体怎么调度线程？如何跨线程安全调用Office对象并更新UI？这里，SynchronizationContext和调度器（Scheduler）派上用场了。

下面用时序图简单描述整体流程：

``` mermaid
sequenceDiagram
    UI线程->>+线程池: 发起计算请求
    线程池->>+VSTO主线程: 通过Office的SyncContext或Scheduler获取数据
    VSTO主线程-->>-线程池: 返回封送后的数据
    线程池->>线程池: 执行计算
    线程池->>+UI线程: 通过UI的SyncContext或Dispatcher更新界面
```

## SynchronizationContext 的选型与实践

### 为 VSTO 主线程设置上下文

SynchronizationContext是一种线程抽象机制，提供“将任务调度到指定线程”的能力。VSTO项目启动时默认无上下文，无法通过SynchronizationContext.Current直接获取VSTO主线程上下文。

一种解决方案是自定义上下文，但更简便的做法是利用Office主线程已有的Win32消息泵，使用WindowsFormsSynchronizationContext。尽管名字带“Windows Forms”，它不仅限于WinForms应用，只要线程是STA且运行Win32消息循环，它都能正常工作。

若需更细粒度的任务优先级控制，可考虑DispatcherSynchronizationContext，但它实现复杂且需手动启动消息泵。因职责已隔离，通常用WindowsFormsSynchronizationContext即可。

示例：

``` CSharp
public partial class ThisAddIn
{
    private static SynchronizationContext _officeSyncContext;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        // 为 VSTO 主线程设置上下文，供后续调度使用
        _officeSyncContext = new WindowsFormsSynchronizationContext();
        SynchronizationContext.SetSynchronizationContext(_officeSyncContext);

        // 其他初始化代码
    }
}
```

### 在自建 UI 线程中设置上下文

WPF和WinForms要求UI线程必须是STA，然而默认线程池线程为MTA，不适合做UI线程。为避免阻塞Office线程（VSTO_Main，唯一STA线程），只能新建一个专用STA线程运行UI。

Avalonia允许非STA线程，但考虑到VSTO依赖Windows平台和消息泵机制，建议UI线程依然使用STA模式。

WPF/WinForms依赖Win32消息泵，同样可用WindowsFormsSynchronizationContext。

示例：

``` CSharp
public partial class ThisAddIn
{
    private static SynchronizationContext _uiSyncContext;

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        // 其它代码

        // 创建专用 UI 线程
        _uiThread = new Thread(() =>
        {
            _uiSyncContext = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(_uiSyncContext);
        })
        {
            Name = "UI Thread",
            IsBackground = true
        };
        _uiThread.SetApartmentState(ApartmentState.STA);
        _uiThread.Start();
    }
}
```

利用SynchronizationContext.Post可在线程池、VSTO主线程与UI线程间安全传递数据。

实际上，对于UI线程，我们更常用Dispatcher而非直接使用SynchronizationContext。

## Reactive Extensions (Rx) 的调度器集成

Rx (Reactive Extensions) 的核心是调度异步事件和操作。它并不直接使用 SynchronizationContext，而是通过 IScheduler 来管理任务调度。

幸运的是，Rx 为我们提供了 SynchronizationContextScheduler 类，可以轻松地将同步上下文封装成调度器：

``` CSharp
// 假设你已经有了某个线程的 SynchronizationContext，如 Office 线程或 UI 线程
var scheduler = new SynchronizationContextScheduler(SynchronizationContext.Current);
```

例如，在你的 VSTO 项目中：

``` CSharp
// 为 COM 线程创建调度器
SchedulerManager.COM = new SynchronizationContextScheduler(_officeSyncContext);

// 为 UI 线程创建调度器
SchedulerManager.UI = new SynchronizationContextScheduler(_uiSyncContext);
```

这样，你就可以把这个 scheduler 传给 Rx 操作符，确保相关操作在正确的线程上执行。

使用时，在Rx操作链中：

``` CSharp
observable
    .ObserveOn(SchedulerManager.UI)  // UI线程更新界面
    .Subscribe(...);

observable
    .ObserveOn(SchedulerManager.COM) // COM线程安全调用Office对象
    .Subscribe(...);
```

对于Avalonia + ReactiveUI项目，Avalonia会自动初始化RxApp.MainThreadScheduler为UI线程调度器，无需额外设置。
