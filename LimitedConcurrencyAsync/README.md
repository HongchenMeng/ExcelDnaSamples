#LimitedConcurrencyAsync sample 并发异步有限样本

[Related to this discussion: https://groups.google.com/d/topic/exceldna/tCbtb2zmQrs/discussion]<br>
[与此相关的讨论: https://groups.google.com/d/topic/exceldna/tCbtb2zmQrs/discussion]<br>

This sample shows how the async function support can be customised by using the .NET 4 Task-based functionality.<br>
In particular, we create a limited concurrency scheduler, that will restrict the number of async threads that are used to run Tasks.<br>
本示例显示如何异步功能的支持，可以通过使用.NET 4基于任务的功能。<br>
特别是，我们创建了一个有限的并发调度，这将限制异步线程来运行任务数。<br>

Some hard-coded paths are set in the project properties in the Debug tab - the exact Excel version and command line arguments. These must be fixed before running the sample. One way is to run "Uninstall-Package Excel-DNA" and then "Install-Package Excel-DNA" in the NuGet Package Manager Console.<br>
某些硬编码路径设置在“调试”选项卡中的项目属性中，即精确的Excel版本和命令行参数。这些必须运行在固定的样品。一个方法是运行“卸载软件Excel DNA”和“安装包Excel DNA“NuGet包管理器控制台。<br>

When running, there should be two new functions in Excel - "Sleep" and "SleepPerCaller", taking the number of seconds to sleep.<br>
在跑步时，在Excel中应该有两种新功能——“睡眠”和“睡眠者”，它们会花几秒钟的时间来睡觉。<br>

Some details on the code:<br>
关于代码的一些细节:<br>

## AsyncFunctions.cs

This is the user code part of the sample. A custom TaskScheduler and related TaskFactory is initialized, and some async Excel-DNA functions defined that will create Tasks using that TaskFactory.<br>
这是示例的用户代码部分。初始化了一个自定义任务调度器和相关的任务工厂，并定义了一些async excel-dna函数，它将使用这个TaskFactory创建任务。<br>

There are two versions of the Sleep function:<br>
  * Sleep - different calls to Sleep with the same timeout parameter will be combined and run as the same Task.<br>
  * SleepPerCaller - calls from different cells will create separate Task, making the concurrency behaviour easier to see.<br>
睡眠功能有两个版本:<br>
睡眠——使用相同的超时参数的不同调用将被组合起来并作为相同的任务运行。<br>
睡眠调用者-来自不同单元的调用将创建单独的任务，使并发行为更容易看到。<br>

## AsyncTaskUtil.cs

This file contains some helpers to integrate the Task-based API with Excel-DNA async support. The main helper function is AsyncTaskUtil, which takes the async call identifiers (the callerFunctionName and callerParameters) as well as an Action<Task> that will create the async Task on the first call. Internally, an ExcelTaskObservable is created, which converts the Task completion result into the appropriate IObservable interface to register with Excel-DNA.<br>
该文件包含一些帮助将基于任务的API与excel-dna异步支持集成在一起的助手。主要的助手函数是AsyncTaskUtil，它使用异步调用标识符(callerFunctionName和callerParameters)，以及在第一个调用中创建async任务的操作任务。在内部，可以创建一个exceltask观测，它将任务完成结果转换为适当的i可视接口，以注册为excel-dna。<br>

There is also an overload that supports cancellation.<br>
还有一个过载支持取消。
<br>
## LimitedConcurrencyLevelTaskScheduler.cs

This file is taken from the "Samples for Parallel Programming" on MSDN (https://code.msdn.microsoft.com/Samples-for-Parallel-b4b76364/sourcecode?fileId=44488&pathId=2044791305). There are a number of custom TaskScheduler samples, including a very flexible QueuedTaskScheduler. The TaskScheduler samples are discussed in detail by Stephen Taub here: <br>http://blogs.msdn.com/b/pfxteam/archive/2010/04/09/9990424.aspx .<br>
这个文件是从“MSDN上的并行编程”(https://code.msdn.microsoft.com/Samples-for-Parallel-b4b76364/sourcecode?fileId=44488&pathId=2044791305。有一些自定义TaskScheduler样品，包括一个非常灵活的queuedtaskscheduler。任务调度器的样品进行详细的讨论：<br>
Stephen Taub在这里http://blogs.msdn.com/b/pfxteam/archive/2010/04/09/9990424.aspx .<br>

The LimitedConcurrencyLevelTaskScheduler uses the .NET ThreadPool threads to run the Tasks, but limits the number of concurrent Tasks that can be running.<br>
LimitedConcurrencyLevelTaskScheduler使用。NET ThreadPool线程来运行这些任务，但是限制了可以运行的并发任务的数量。



