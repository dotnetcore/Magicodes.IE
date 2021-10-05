namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public interface ITokenIndexProvider
    {
        int Index { get; }

        void MoveIndexPointerForward();
    }
}
