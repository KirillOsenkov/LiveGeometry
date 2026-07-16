namespace DynamicGeometry
{
    public interface IOperationProvider
    {
        IOperationDescription ProvideOperation(object instance);
    }
}
