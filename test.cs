public class Foo {

    private Handler _h;

    public Foo(Handler h)
    {
        _h = h;
    }

    public void Bar(int i)
    {
        _h.AsyncHandle(CalcOn(i));
    }

    private SomeResponse CalcOn(int i)
    {
        ...;
    }
}
==================
var mock = new Mock<Handler>();
SomeResponse result = null;
mock.Setup(h => h.AsyncHandle(It.IsAny<SomeResponse>()))
    .Callback<SomeResponse>(r => result = r);

// do your test
new Foo(mock.Object).Bar(22);
Assert.NotNull(result);

==================
mock.Setup(h => h.AnsyncHandle(It.Is<SomeResponse>(response => response != null)));
