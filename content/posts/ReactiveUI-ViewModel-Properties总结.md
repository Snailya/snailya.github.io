+++
title = 'ReactiveUI ViewModel Properties总结'
date = 2024-01-11T14:36:16+08:00
draft = false
categories = ['ReactiveUI']
+++

## 属性类型

在ReactiveUI中，ViewModel中的属性可以根据其用途划分为三种情况：读写属性（Read-Write Properties）、只读属性（Read-Only Properties）和输出属性（Output Properties）。

读写属性 (Read-Write Properties):

读写属性是可以被服务修改，也可以被用户在View中修改的属性。这类属性是我们通常比较熟悉的普通属性。

例如，用户的姓名可能是被服务加载的，而在加载之后又被用户修改。

只读属性 (Read-Only Properties):

只读属性是在构造函数中被初始化且在之后不再变化的属性。

例如， 用户的ID在一般情况下时不允许变化的。

输出属性 (Output Properties):

输出属性是ReactiveUI中新提出的概念，初次接触ReactiveUI时，可能会将输出属性与只读属性混为一谈。尽管输出属性对于用户而言是只读的，但是对于属性本身是可变的。这类属性通常由Observable变化而成，表示属性值可能随时间变化。
例如，用户负债率随用户的总资产和总负债变化，但负债率属性本身是不允许被用户修改的。

在ReactiveUI中，使用这三种属性类型可以更清晰地表示属性的特性和用途。读写属性用于需要双向绑定的数据，只读属性用于一次性初始化后不再改变的数据，而输出属性用于表示可能随时间变化的数据流。这种划分有助于更好地理解和管理ViewModel中的属性。

## 属性的声明与绑定方法

在明确了属性类型的基础上，ViewModel中所有`非集合类型`的属性都可以按照下面固定的方式进行声明与绑定。

读写属性的声明需要调用ReactiveObject的RaiseAndSetIfChanged方法，该方法实现了INotifydPropertyChanged。

``` C#
private string name;
public string Name 
{
    get => name;
    set => this.RaiseAndSetIfChanged(ref name, value);
}

```

读写属性的绑定使用TView的Bind方法进行双向绑定。

``` C#
this.WhenActivated(disposable =>
{
    this.Bind(ViewModel, vm => vm.Name, v => v.NameTextBox.Text)
        .DisposeWith(disposable);
});
```

只读属性使用一般的声明方式即可。

``` C#
public int Id {get;}
```

绑定时，使用单向绑定。

``` C#
this.WhenActivated(disposable =>
{
    this.OneWayBind(ViewModel, vm => vm.Id, v => v.IdTextBlock.Text)
        .DisposeWith(disposable);
});
```

输出属性也是使用固定的方式进行声明，但是在初始化时需要注意。

``` C#
// 声明
private readonly ObservableAsPropertyHelper<double> _debtAssetRatio;
public string DebtAssetRatio => _debtAssetRatio.Value;

// WhenAnyValue产生一个Observable，当Debt或Asset变化时，会发出新的Debt/Asset的值。这里需要注意Debt和Asset必须也实现了INotifyPropertyChanged，否则无法观察到它们的变化。
// ToProperty将Observable转变为ObservableAsPropertyHelper。
UserAccountViewModel(){
    this.WhenAnyValue(x => x.Debt, x => x.Asset, (debt, asset) => debt/asset)
        .ToProperty(this, x => x.DebtAssetRatio, out _debtAssetRatio);
}
```

由于输出属性对用户而言也是只读的，所以使用单向绑定。

``` C#
this.WhenActivated(disposable =>
{
    this.OneWayBind(ViewModel, vm => vm.DebtAssetRatio, v => v.DebtAssetRatioTextBlock.Text)
        .DisposeWith(disposable);
});
```

## 集合

在使用集合类型的数据，最简单的情况是View中使用的是不可变的数据集合，例如显示Blog中已归档的文章列表，显示用户银行账户的历史交易信息等。这种情况下由于数据只在构造函数中被初始化，所以可以声明为任意集合类型，例如`IEnmerable<T>`、 `IList<T>`、`ObservableCollection<T>`。再在View中使用OneWayBind绑定。

``` C#
// ViewModel
public IEnumerable<Article> Articles {get;}

// View
this.WhenActivated(disposable =>
{
    this.OneWayBind(ViewModel, vm => vm.Articles, v => v.DataGrid.ItemsSource)
        .DisposeWith(disposable);
});

```

对于可变数据集合，则需要将集合声明为`ObservableCollection<T>`，然后在View中使用OneWayBind绑定。要注意的是`T`也需要实现`INotifydPropertyChanged`，否则T属性的变化无法触发View更新。

到目前为止这和我们在以往使用WPF进行数据绑定时没有任何差异，那为什么还要将集合单拎出来特别说明。这是因为在ReactiveUI中推荐使用Dynamic Data来操作集合。ReactiveUI的文档中写的比较抽象，这里我将举例说明为什么使用Dynamic Data可以帮助我们更好的维护代码。

在这个例子中，我将使用到两个DataGrid，其中SourceDataGrid中显示带有复选框的用户数据，SourceDataGrid中被选中的数据将被显示在OutputDataGrid中，也就是说OutputDataGrid是SourceDataGrid的一个视图。View的代码如下：

``` C#

// in MainWindow.xaml.xaml
<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition />
        <ColumnDefinition />
    </Grid.ColumnDefinitions>
    <Grid Margin="0 0 4 0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>
        <TextBlock Text="Original DataGrid with Read-Write Items" />
        <DataGrid Grid.Row="1" x:Name="SourceDataGrid" AutoGenerateColumns="False">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Binding="{Binding IsSelected, UpdateSourceTrigger=PropertyChanged}" />
                <DataGridTextColumn Binding="{Binding Id}" />
                <DataGridTextColumn Binding="{Binding Name}" />
                <DataGridTextColumn Binding="{Binding Age}" />
            </DataGrid.Columns>
        </DataGrid>
    </Grid>
    <Grid Grid.Column="1" Margin="4 0 0 0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>
        <TextBlock Text="DataGrid Shows Output Properties that Filters IsSelected" />
        <DataGrid Grid.Row="1" x:Name="OutputDataGrid" AutoGenerateColumns="False">
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Id}" />
                <DataGridTextColumn Binding="{Binding Name}" />
                <DataGridTextColumn Binding="{Binding Age}" />
            </DataGrid.Columns>
        </DataGrid>
    </Grid>
</Grid>

//in MainWindow.xaml.cs
this.WhenActivated(disposable =>
{
    this.OneWayBind(ViewModel, 
        vm => vm.ReadWriteItems, 
        v => v.SourceDataGrid.ItemsSource)
        .DisposeWith(disposable);
    
    this.OneWayBind(ViewModel, 
            vm => vm.OutputItems, 
            v => v.OutputDataGrid.ItemsSource)
        .DisposeWith(disposable);
});

```

在不使用Dynamic Data时，要使`OutputItems`在`ReadWriteItems`发生变化时同步刷新，可以选在在`ReadWriteItems`属性变化时，调用`ReactiveObject`的`RaisePropertyChanged`方法。但是，因为`ObservableCollection`只提供了`CollectionChanged`事件，这个事件只有当集合增删时会被触发，所以必须遍历`ReadWriteItems`的对象并在`Item`对象的`PropertyChanged`事件上增加处理方法。

```C#

public ObservableCollection<ItemViewModel> ReadWriteItems { get; }
public IEnumerable<ItemViewModel> OutputItems => ReadWriteItems.Where(x => x.IsSelected);

public MainWindowViewModel{
    ReadWriteItems = new ObservableCollection<ItemViewModel>(ItemViewModel.GetItems());
        foreach (var item in ReadWriteItems)
            item.PropertyChanged += (sender, args) => { this.RaisePropertyChanged(nameof(OutputItems)); };
}

```

这种做法将业务逻辑分散在两个地方，表示筛选的部分放在了属性的声明中，而表示刷新的部分放在了构造函数中。当代码体量增加时，分散的代码会增加运维的成本。而使用Dynamic Data，可以将上述分散的代码合并在构造函数中。

```C#
private readonly ReadOnlyObservableCollection<ItemViewModel> _outputItems;

public ObservableCollection<ItemViewModel> ReadWriteItems { get; }
public ReadOnlyObservableCollection<ItemViewModel> OutputItems => _outputItems;

public MainWindowViewModel{
    ReadWriteItems = new ObservableCollection<ItemViewModel>(ItemViewModel.GetItems());
    var disposable =  ReadWriteItems.ToObservableChangeSet()
        .AutoRefresh()
        // .AutoRefreshOnObservable(x=>x.WhenPropertyChanged(i=>i.IsSelected))
        .Filter(x=>x.IsSelected)
        .Bind(out _outputItems)
        .Subscribe();
}

```

在这段代码中，我们首先将`ReadWriteItems`转化为`IObservable<IChangeSet<ItemViewModel>>`，然后使用`AutoRefresh`将`ReadWriteItems`的变化向下传递，最终绑定到`_outputItems`中。在过程中我们调用`Filter`操作符对数据进行筛选，选出被选中的对象，还可以在后面串联更多的操作符对数据进行变换。需要注意结尾的`Subscribe`调用是不可缺少的，因为只有在被订阅后，上述操作才会生效。

同样的原理，有时我们需要根据集合中数据的状态来切换按钮的状态，例如只有当全被选中时，才允许按钮被点击。

``` C#
private readonly ObservableAsPropertyHelper<bool> _isEnabled;

public ObservableCollection<ItemViewModel> ReadWriteItems { get; }
public bool IsEnabled => _isEnabled.Value;

public MainWindowViewModel{
    ReadWriteItems = new ObservableCollection<ItemViewModel>(ItemViewModel.GetItems());
    _isEnabled =  ReadWriteItems.ToObservableChangeSet()
        .AutoRefresh()
        .ToCollection()
        .Select(x => x.All(i => i.IsSelected))
        .ToProperty(this, x => x.IsEnabled);
}
```
