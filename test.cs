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
==================
delegate void DoSomethingCallback(ref string value);

var newValue = "This is the new referenced value";
mock.Setup(p => p.DoSomething(It.Ref<string>.IsAny))
    .Callback(new DoSomethingCallback((ref string value) => value = newValue));

==================
coreMethodsMock
  .Setup(m => m.PushPageModel<It.IsSubtype<SecondPageModel>>(
    It.Is<bool>(b => b)))
  .Returns(Task.CompletedTask)
  .Verifiable();
