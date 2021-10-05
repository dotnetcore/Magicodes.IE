namespace OfficeOpenXml.LoadFunctions.Params
{
    /// <summary>
    /// Declares how headers should be parsed before they are added to the worksheet
    /// </summary>
    public enum HeaderParsingTypes
    {
        /// <summary>
        /// Leaves the header as it is
        /// </summary>
        Preserve,
        /// <summary>
        /// Replaces any underscore characters with a space
        /// </summary>
        UnderscoreToSpace,
        /// <summary>
        /// Adds a space between camel cased words ('MyProp' => 'My Prop')
        /// </summary>
        CamelCaseToSpace,
        /// <summary>
        /// Replaces any underscore characters with a space and adds a space between camel cased words ('MyProp' => 'My Prop')
        /// </summary>
        UnderscoreAndCamelCaseToSpace
    }
}
