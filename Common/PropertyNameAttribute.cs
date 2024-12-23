namespace Common
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public class PropertyNameAttribute : Attribute
    {
        public string Name { get; }

        public PropertyNameAttribute(string name)
        {
            Name = name;
        }
    }
}
