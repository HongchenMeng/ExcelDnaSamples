# Simple RTD server with array function wrapper

This sample was created to provide a simple RTD server to check the behaviour of RTD functions when called from array formulas.

The wrapper function has signature
   ```cs
	public static object RtdArrayTest(string prefix)
   ```
and returns a 2x1 array.

The function can be called from Excel as an array formula (with Ctrl+Shift+Enter) using another cell as the "prefix":
   ```
    {=RtdArrayTest(D1)}
   ```

The implementation of the RTD server is based on the Excel-DNA base class `ExcelRtdServer`, and just uses a Timer to update the topics.

# RTD服务器阵列功能简单的包装
创建这个示例是为了提供一个简单的RTD服务器来检查当从数组公式调用时RTD函数的行为。
包装器函数签名
   ```cs
	public static object RtdArrayTest(string prefix)
   ```
并返回一个2X1阵列。
该函数可以从Excel调用为数组公式（用Ctrl + Shift + Enter），使用另一个单元格作为前缀：
   ```
    {=RtdArrayTest(D1)}
   ```

RTD的服务器的实现是基于Excel-DNA 的基类` excelrtdserver `，只使用一个定时器来更新内容。