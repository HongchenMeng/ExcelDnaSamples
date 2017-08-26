AsyncAwaitMacro sample 异步等待宏观样本
---

Shows how the ExcelAsyncUtil.QueueAsMacro mechanism can be used to implement async macros with the C# async/await mechanism.<br>
显示了ExcelAsyncUtil.QueueAsMacro机制可以使用C#异步/等候机制来实现异步宏。

### Warning 警告

This sample does not represent 'best practice'. It just explores how the C# async/await features interact with the Excel-DNA async mechanism, and with the Excel hosting environment.<br>
这个示例不代表“最佳实践”。它只是探讨了casync/坐等特性与Excel-dna异步机制以及Excel托管环境的交互方式。<br>

Trying to run async macros as in this example will interfere with an interactive user busy with Excel: <br>
 their undo stack and copy selection will get cleared at unexpected times, and any other add-ins or macros being run might be
 interleaved with the async code.<br>
 在本例中尝试运行异步宏将会干扰与Excel交互的交互式用户:<br>
它们的撤销栈和复制选择将在不预期的时间被清除，而运行的任何其他插件或宏都可能是
与异步代码交叉。
 
#### `ExcelAsyncTask`

The sample includes a helper class called `ExcelAsyncTask` with a single `Run` method.
In turn, this starts the `Task` with a TaskScheduler which just enqueues `Task`s to run in the macro context, 
ensuring that the async/await continuations are again scheduled in a macro context.
这个示例包括一个名为ExcelAsyncTask的助手类，它带有一个单独的Run方法。<br>
反过来，这将启动任务调度程序，该任务调度程序只在宏上下文中运行任务s，
确保在宏上下文中再次调度异步/等待延续。<br>

```c#

public static void MacroToRunSlowWork()
{
    // Starts running SlowWork in a context where async/await will return to the macro context on the main Excel thread.
    ExcelTaskAsync.Run(SlowWork);
}

static async Task SlowWork()
{
    // All the code here, before and after the awaits, will run on the main thread in a macro context
    // (where C API calls and the COM object model is safe to access).
    await SomeWorkAsync();
    Application.Range["A1"].Value = "abc;
    await OtherWorkAsync();
    Application.Range["A2"].Value = "xyz;
}

```


#### `ExcelSynchronizationContext`

The first implementation I attempted was run the async/await code in a context where `SynchronizationContext.Current` was set to an `ExcelSynchronizationContext`.<br>
我尝试的第一个实现是在同步上下文环境中运行异步/等待代码。当前设置为一个“ExcelSynchronizationContext”。<br>

There is a problem I don't yet understand when trying to use a `SynchronizationContext` in Excel-DNA.
Somehow the `SynchronizationContext.Current` is cleared in the SyncWindow or macro running process.
I have found references to `WindowsFormsSynchronizationContext.AutoInstall` causing trouble, but could not see how that applies in our case.
当我试图在excel dna中使用同步上下文时，我还不明白这个问题。<br>
“SynchronizationContext。当前在SyncWindow或宏运行过程中被清除。<br>
我发现WindowsFormsSynchronizationContext引用”。自动安装引起了麻烦，但在我们的例子中却看不到这一点。<br>

It could be that the unmanaged -> managed transition interferes with the thread-based context that stores the SynchronizationContext.Current.
As an alternative, we use the TaskScheduler-based approach.<br>
可能是过渡干扰管理的非托管- >存储SynchronizationContext.Current线程上下文。
作为另一种选择，我们使用基于任务调度的方法。

