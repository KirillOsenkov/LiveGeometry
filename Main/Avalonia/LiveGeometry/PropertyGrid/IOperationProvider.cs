namespace LiveGeometry;

public interface IOperationProvider
{
    IOperationDescription ProvideOperation(object instance);
}
