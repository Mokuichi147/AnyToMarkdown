namespace MarkdownConverter.Tests;

public class UnitTest1
{
    [Fact]
    public void Test1()
    {
        string hello = MarkdownConverter.Class1.HelloWorld();
        Assert.Equal("Hello World", hello);
    }
}
